/**
 * Tiny safe-by-construction Markdown renderer.
 *
 * Scope: covers the subset of GitHub-Flavored Markdown the LLM
 * actually emits in a chat reply. Not aiming for full GFM —
 * we deliberately stay smaller so we can hand-roll the parser
 * in a few hundred lines and never pull in a 200 KB markdown
 * pipeline. The features that matter here:
 *
 *   - Paragraphs (double newlines)
 *   - Headings (ATX-style: #, ##, …)
 *   - Unordered (-, *) and ordered (1., 2.) lists, including
 *     arbitrarily-nested children
 *   - Fenced code blocks with optional language hint (rendered
 *     as a small label above the code)
 *   - Block quotes (> …) including multi-line
 *   - Horizontal rule (---, ***, ___) on its own line
 *   - GFM tables (| col | col | + |---| separator |)
 *   - Inline `code`
 *   - **bold**, *italic*, ~~strikethrough~~
 *   - [text](url) links
 *   - Hard line breaks (two trailing spaces)
 *   - The literal sequences `&`, `<`, `>` escaped to entities
 *
 * Deliberately NOT supported (and silently rendered as plain
 * text if the LLM emits them):
 *
 *   - HTML inside markdown (rehype-style). A user message can
 *     not inject a `<script>` or `<iframe>` even if they
 *     write the markup; we never call `dangerouslySetInnerHTML`.
 *   - Reference-style links
 *   - Images (![alt](src)) — not useful in a task pane reply
 *   - Task-list checkboxes (- [ ] task) — easy to add later
 *
 * Why a custom renderer at all: react-markdown ships with a
 * full unified/remark/rehype pipeline and dozens of transitive
 * dependencies. We're already over the 500 KB bundle size
 * warning, and we don't need image processing or footnotes.
 * This module adds ~8 KB and zero new dependencies.
 */
import { Fragment, type ReactNode } from 'react'

// ---------- Block parser ----------

type ListItem = {
  text: string
  children: ListItem[]
}

type Block =
  | { kind: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; text: string }
  | { kind: 'paragraph'; text: string }
  | { kind: 'code'; lang: string; text: string }
  | { kind: 'list'; ordered: boolean; items: ListItem[] }
  | { kind: 'blockquote'; text: string }
  | { kind: 'hr' }
  | {
      kind: 'table'
      header: string[]
      rows: string[][]
      align: ('left' | 'center' | 'right')[]
    }

function parseBlocks(input: string): Block[] {
  const blocks: Block[] = []
  // Normalise Windows line endings so the same parser works on
  // every platform the add-in runs on.
  const src = input.replace(/\r\n?/g, '\n')
  const lines = src.split('\n')
  let i = 0
  while (i < lines.length) {
    const line = lines[i]

    // Skip blank lines between blocks.
    if (line.trim() === '') {
      i++
      continue
    }

    // Fenced code block — ```lang\n…\n```
    const fence = /^```([^\s`]*)?\s*$/.exec(line)
    if (fence) {
      const lang = fence[1] ?? ''
      const buf: string[] = []
      i++
      while (i < lines.length && !/^```\s*$/.test(lines[i])) {
        buf.push(lines[i])
        i++
      }
      // Eat the closing fence if present (be lenient if missing).
      if (i < lines.length) i++
      blocks.push({ kind: 'code', lang, text: buf.join('\n') })
      continue
    }

    // ATX heading — #, ##, … up to ######
    const heading = /^(#{1,6})\s+(.*)$/.exec(line)
    if (heading) {
      const level = heading[1].length as 1 | 2 | 3 | 4 | 5 | 6
      blocks.push({ kind: 'heading', level, text: heading[2] })
      i++
      continue
    }

    // Horizontal rule — ---, ***, or ___ on its own line,
    // surrounded only by whitespace. We require 3+ chars
    // and the same character to avoid eating Markdown setext
    // heading underlines (=== for h1, --- for h2), which the
    // heading branch above has already consumed.
    if (/^\s*([-*_])\s*\1\s*\1[\s\S]*$/.test(line)) {
      blocks.push({ kind: 'hr' })
      i++
      continue
    }

    // Unordered list — `- …` or `* …`
    // Ordered list — `1. …` / `1) …`
    // Lists are recursive: an indented child line (2+ spaces,
    // 4 spaces, or a tab) belongs to the most recent list
    // item. This is GFM behaviour and matches what the LLM
    // emits when laying out a multi-step procedure.
    if (/^([-*]|\d+[.)])\s+/.test(line)) {
      const ordered = /^\d+[.)]\s+/.test(line)
      const { items, end } = parseListBlock(lines, i, ordered)
      blocks.push({ kind: 'list', ordered, items })
      i = end
      continue
    }

    // Block quote — `> …` per line, multi-line, may contain
    // blank lines. We strip the leading `> ` and join the
    // remaining text with newlines so the inline parser
    // gets a clean string.
    if (/^>\s?/.test(line)) {
      const buf: string[] = []
      while (i < lines.length && /^>\s?/.test(lines[i])) {
        buf.push(lines[i].replace(/^>\s?/, ''))
        i++
      }
      blocks.push({ kind: 'blockquote', text: buf.join('\n') })
      continue
    }

    // GFM table — header row, separator row, then body rows.
    if (
      line.includes('|') &&
      i + 1 < lines.length &&
      /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/.test(lines[i + 1])
    ) {
      const header = splitTableRow(line)
      const sep = lines[i + 1]
      const align = parseTableAlign(sep)
      i += 2
      const rows: string[][] = []
      while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
        rows.push(splitTableRow(lines[i]))
        i++
      }
      blocks.push({ kind: 'table', header, rows, align })
      continue
    }

    // Paragraph — collect contiguous non-blank, non-special lines.
    const para: string[] = [line]
    i++
    while (
      i < lines.length &&
      lines[i].trim() !== '' &&
      !/^```/.test(lines[i]) &&
      !/^#{1,6}\s+/.test(lines[i]) &&
      !/^([-*]|\d+[.)])\s+/.test(lines[i]) &&
      !/^>\s?/.test(lines[i]) &&
      !/^\s*([-*_])\s*\1\s*\1/.test(lines[i])
    ) {
      para.push(lines[i])
      i++
    }
    blocks.push({ kind: 'paragraph', text: para.join('\n') })
  }
  return blocks
}

// ---------- List helpers ----------

/**
 * Walk lines starting at `start`, consuming a list block
 * (and any indented child lists inside it). Returns the flat
 * item array plus the line index where the list ended so the
 * outer loop can continue.
 *
 * The recursion is what gives us nested lists. The "indent
 * belongs to the previous item" rule is the same one GFM
 * uses: a line indented at least 2 spaces relative to the
 * parent's marker (`- ` or `1. `) is a child of that item.
 * Children may be either the same kind (more `-` items
 * underneath a `-` item) or the opposite kind (`1.` underneath
 * a `-`); we accept both because the LLM is inconsistent
 * about which it picks.
 */
function parseListBlock(
  lines: string[],
  start: number,
  ordered: boolean,
): { items: ListItem[]; end: number } {
  // Indent unit: 2 spaces. Anything >= 2 spaces is a child of
  // the previous item at that level. Tabs count as 2 spaces,
  // matching what most editors emit when the LLM pads with
  // \t.
  const INDENT = 2

  // We need the indent depth of the *first* item to anchor
  // the rest of the block. Capture it, then walk the same
  // depth plus deeper.
  const baseIndent = measureIndent(lines[start])

  const items: ListItem[] = []
  let i = start
  let current: ListItem | null = null
  let currentItemIndent = 0

  while (i < lines.length) {
    const line = lines[i]
    if (line.trim() === '') {
      // Blank line inside a list: peek to see if the next
      // non-blank line is still a list item at the same level
      // (lazy continuation). If not, the list ends here.
      let j = i + 1
      while (j < lines.length && lines[j].trim() === '') j++
      const next = j < lines.length ? lines[j] : ''
      if (next === '' || measureIndent(next) < baseIndent) {
        break
      }
      i = j
      continue
    }

    const indent = measureIndent(line)
    if (indent < baseIndent) break
    const marker = isListMarker(line)
    if (!marker) {
      // Could be a continuation of the current item's text
      // (e.g. a wrapped sentence under `- foo`). The
      // strict-GFM rule is to keep eating until indent drops
      // or a new marker appears. We just push the rest onto
      // the current item's text.
      if (current && indent >= currentItemIndent) {
        current.text =
          current.text.length > 0
            ? `${current.text}\n${line.trim()}`
            : line.trim()
        i++
        continue
      }
      break
    }

    const text = line.slice(indent + marker.length).trim()
    if (indent === baseIndent) {
      current = { text, children: [] }
      items.push(current)
      currentItemIndent = indent + INDENT
      i++
      continue
    }

    // Deeper indent — child of `current`. Recurse.
    if (current) {
      const child = parseListBlock(lines, i, ordered)
      // Adopt the first top-level child(ren) as current's
      // children, then drop the (now-empty) wrapper list.
      for (const c of child.items) current.children.push(c)
      i = child.end
      continue
    }

    // Deeper indent but no parent to attach to — bail.
    break
  }
  return { items, end: i }
}

function measureIndent(line: string): number {
  // Count leading spaces; tabs are worth 2 (matches the
  // GFM rule of thumb for what most editors produce when
  // you press Tab while writing a list).
  let n = 0
  for (const ch of line) {
    if (ch === ' ') n++
    else if (ch === '\t') n += 2
    else break
  }
  return n
}

function isListMarker(
  line: string,
): { length: number; ordered: boolean } | null {
  // We look anywhere along the leading-whitespace run, not
  // just at position 0 — indented children of a parent list
  // item look like "  - mid" and the parser still needs to
  // recognise them as a marker (with the leading indent
  // counted separately by the caller).
  const ul = /^[-*]\s+/.exec(line.trimStart())
  if (ul) return { length: ul[0].length, ordered: false }
  const ol = /^\d+[.)]\s+/.exec(line.trimStart())
  if (ol) return { length: ol[0].length, ordered: true }
  return null
}

// ---------- Table helpers ----------

/**
 * Split a single `| a | b | c |` row into the cell strings.
 */
function splitTableRow(line: string): string[] {
  let s = line.trim()
  if (s.startsWith('|')) s = s.slice(1)
  if (s.endsWith('|')) s = s.slice(0, -1)
  return s.split('|').map((c) => c.trim())
}

function parseTableAlign(
  sep: string,
): ('left' | 'center' | 'right')[] {
  return splitTableRow(sep).map((cell) => {
    const t = cell.trim()
    const left = t.startsWith(':')
    const right = t.endsWith(':')
    if (left && right) return 'center'
    if (right) return 'right'
    return 'left'
  })
}

// ---------- Inline parser ----------
//
// Runs on a single line (or single paragraph blob). We apply
// the replacements in a fixed sequence on the already-`&<>`-
// escaped string so the earlier passes can't accidentally eat
// tokens later passes would have matched.

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function parseInline(input: string): ReactNode {
  // First, pull out fenced inline `code` spans so their content
  // is left untouched by the bold/italic/link/strike passes
  // below.
  const parts: Array<{ kind: 'text' | 'code'; value: string }> = []
  let last = 0
  const codeRe = /`([^`\n]+)`/g
  let m: RegExpExecArray | null
  while ((m = codeRe.exec(input)) !== null) {
    if (m.index > last) {
      parts.push({ kind: 'text', value: input.slice(last, m.index) })
    }
    parts.push({ kind: 'code', value: m[1] })
    last = m.index + m[0].length
  }
  if (last < input.length) {
    parts.push({ kind: 'text', value: input.slice(last) })
  }

  return parts.map((p, i) => {
    if (p.kind === 'code') {
      return (
        <code
          key={i}
          className="rounded bg-neutral-100 px-1 py-0.5 font-mono text-[0.85em] text-neutral-800"
        >
          {p.value}
        </code>
      )
    }
    return <Fragment key={i}>{renderText(p.value)}</Fragment>
  })
}

function renderText(s: string): ReactNode {
  // Match order matters: `**` before `*` (so bold grabs its
  // delimiters first), `~~` is single-character-delimited so
  // it can sit anywhere.
  const nodes: ReactNode[] = []
  const re =
    /(\*\*([^*\n]+)\*\*|\*([^*\n]+)\*|~~([^~\n]+)~~|\[([^\]\n]+)\]\((https?:\/\/[^\s)]+)\))/g
  let last = 0
  let m: RegExpExecArray | null
  let key = 0
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) {
      nodes.push(<Fragment key={key++}>{s.slice(last, m.index)}</Fragment>)
    }
    if (m[2] != null) {
      nodes.push(<strong key={key++}>{m[2]}</strong>)
    } else if (m[3] != null) {
      nodes.push(<em key={key++}>{m[3]}</em>)
    } else if (m[4] != null) {
      nodes.push(<del key={key++}>{m[4]}</del>)
    } else if (m[5] != null && m[6] != null) {
      nodes.push(
        <a
          key={key++}
          href={m[6]}
          target="_blank"
          rel="noreferrer noopener"
          className="text-excel-green underline underline-offset-2 hover:text-excel-green-light"
        >
          {m[5]}
        </a>,
      )
    }
    last = m.index + m[0].length
  }
  if (last < s.length) {
    nodes.push(<Fragment key={key++}>{s.slice(last)}</Fragment>)
  }
  // Hard line breaks (two trailing spaces) get rendered as <br/>.
  const result: ReactNode[] = []
  nodes.forEach((node, i) => {
    if (typeof node === 'string' || typeof node === 'number') {
      const text = String(node)
      const segments = text.split(/ {2}\n|\n/)
      segments.forEach((seg, j) => {
        result.push(
          <Fragment key={`${i}-${j}`}>
            {seg}
            {j < segments.length - 1 ? <br /> : null}
          </Fragment>,
        )
      })
    } else {
      result.push(node)
    }
  })
  return result
}

// ---------- Thinking extraction ----------

const THINKING_RE = /<thinking>([\s\S]*?)<\/thinking>/i

/**
 * Pull a `<thinking>…</thinking>` block out of a chat message
 * and return the block + the remaining text.
 */
export function splitThinking(
  raw: string,
): { thinking: string | null; text: string | null } {
  const m = THINKING_RE.exec(raw)
  if (!m) return { thinking: null, text: raw.length > 0 ? raw : null }
  const thinking = m[1].trim()
  // Collapse the gap the <thinking> tag leaves behind — raw
  // text often reads "hello <thinking>...</thinking> world",
  // and we don't want a double-space to leak into the
  // rendered message.
  const rest = raw.replace(m[0], ' ').replace(/\s+/g, ' ').trim()
  return {
    thinking: thinking.length > 0 ? thinking : null,
    text: rest.length > 0 ? rest : null,
  }
}

// ---------- Renderer ----------

export function Markdown({ source }: { source: string }) {
  const blocks = parseBlocks(source)
  return (
    <>
      {blocks.map((b, i) => {
        switch (b.kind) {
          case 'heading': {
            const cls =
              b.level === 1
                ? 'mt-3 mb-1 text-base font-semibold text-neutral-900'
                : b.level === 2
                  ? 'mt-3 mb-1 text-sm font-semibold text-neutral-900'
                  : 'mt-2 mb-1 text-sm font-medium text-neutral-800'
            const inner = parseInline(escapeHtml(b.text))
            switch (b.level) {
              case 1:
                return <h1 key={i} className={cls}>{inner}</h1>
              case 2:
                return <h2 key={i} className={cls}>{inner}</h2>
              case 3:
                return <h3 key={i} className={cls}>{inner}</h3>
              case 4:
                return <h4 key={i} className={cls}>{inner}</h4>
              case 5:
                return <h5 key={i} className={cls}>{inner}</h5>
              case 6:
                return <h6 key={i} className={cls}>{inner}</h6>
            }
            return null
          }
          case 'paragraph': {
            return (
              <p
                key={i}
                className="text-sm leading-relaxed text-neutral-800 [&:not(:last-child)]:mb-3"
              >
                {parseInline(escapeHtml(b.text))}
              </p>
            )
          }
          case 'code': {
            return (
              <div
                key={i}
                className="my-3 overflow-hidden rounded-lg border border-neutral-800/40 bg-neutral-900"
              >
                {b.lang ? (
                  <div className="border-b border-neutral-700 bg-neutral-800 px-3 py-1 text-[10px] font-medium uppercase tracking-wide text-neutral-400">
                    {b.lang}
                  </div>
                ) : null}
                <pre className="overflow-x-auto px-3 py-2 text-xs leading-relaxed text-neutral-100">
                  <code className="font-mono">{b.text}</code>
                </pre>
              </div>
            )
          }
          case 'list': {
            return (
              <ListView
                key={i}
                items={b.items}
                ordered={b.ordered}
                depth={0}
              />
            )
          }
          case 'blockquote': {
            // Multi-line block quote: split on newlines and
            // re-parse each line so inline emphasis and code
            // spans still apply inside the quote.
            const inner = b.text
              .split('\n')
              .map((line, j) => (
                <Fragment key={j}>
                  {j > 0 ? <br /> : null}
                  {parseInline(escapeHtml(line))}
                </Fragment>
              ))
            return (
              <blockquote
                key={i}
                className="my-3 border-l-4 border-excel-green-tint-2 bg-excel-green-tint/40 px-3 py-1.5 text-sm leading-relaxed text-neutral-700"
              >
                {inner}
              </blockquote>
            )
          }
          case 'hr': {
            return <hr key={i} className="my-3 border-neutral-200" />
          }
          case 'table': {
            const alignCls = (a: 'left' | 'center' | 'right'): string =>
              a === 'center'
                ? 'text-center'
                : a === 'right'
                  ? 'text-right'
                  : 'text-left'
            return (
              <div
                key={i}
                className="my-3 overflow-x-auto rounded-lg border border-neutral-200"
              >
                <table className="w-full text-sm text-neutral-800">
                  <thead className="bg-neutral-50">
                    <tr>
                      {b.header.map((cell, j) => (
                        <th
                          key={j}
                          className={`border-b border-neutral-200 px-2.5 py-1.5 font-medium ${alignCls(b.align[j] ?? 'left')}`}
                        >
                          {parseInline(escapeHtml(cell))}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {b.rows.map((row, j) => (
                      <tr key={j} className="even:bg-neutral-50/50">
                        {row.map((cell, k) => (
                          <td
                            key={k}
                            className={`border-t border-neutral-100 px-2.5 py-1.5 align-top ${alignCls(b.align[k] ?? 'left')}`}
                          >
                            {parseInline(escapeHtml(cell))}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          }
        }
      })}
    </>
  )
}

// ---------- Recursive list renderer ----------

function ListView({
  items,
  ordered,
  depth,
}: {
  items: ListItem[]
  ordered: boolean
  depth: number
}): ReactNode {
  const Tag = ordered ? 'ol' : 'ul'
  // Indent nested lists with extra left padding; Tailwind's
  // arbitrary-value syntax lets us tie the indent to the
  // recursion depth.
  const padCls = depth === 0 ? 'pl-5' : 'pl-5'
  return (
    <Tag
      className={`my-2 space-y-1 ${padCls} text-sm leading-relaxed text-neutral-800 ${
        ordered ? 'list-decimal' : 'list-disc'
      }`}
    >
      {items.map((it, j) => (
        <li key={j}>
          {parseInline(escapeHtml(it.text))}
          {it.children.length > 0 ? (
            <ListView
              items={it.children}
              ordered={ordered}
              depth={depth + 1}
            />
          ) : null}
        </li>
      ))}
    </Tag>
  )
}
