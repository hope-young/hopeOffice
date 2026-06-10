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
 *   - Unordered (-, *) and ordered (1., 2.) lists
 *   - Fenced code blocks with optional language hint
 *   - Inline `code`
 *   - **bold** and *italic*
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
 *   - Tables (the LLM rarely tables in chat; if it does, the
 *     pipe-and-dash syntax passes through as plain text)
 *   - Images (![alt](src)) — not useful in a task pane reply
 *
 * Why a custom renderer at all: react-markdown ships with a
 * full unified/remark/rehype pipeline and dozens of transitive
 * dependencies. We're already over the 500 KB bundle size
 * warning, and we don't need image processing or footnotes.
 * This module adds ~6 KB and zero new dependencies.
 */
import { Fragment, type ReactNode } from 'react'

// ---------- Block parser ----------

type Block =
  | { kind: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; text: string }
  | { kind: 'paragraph'; text: string }
  | { kind: 'code'; lang: string; text: string }
  | { kind: 'list'; ordered: boolean; items: string[] }
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

    // Unordered list — `- …` or `* …`
    const ul = /^[-*]\s+(.*)$/.test(line)
    // Ordered list — `1. …` / `1) …`
    const ol = /^\d+[.)]\s+(.*)$/.test(line)
    if (ul || ol) {
      const items: string[] = []
      const ordered = ol
      const re = ordered
        ? /^\d+[.)]\s+(.*)$/
        : /^[-*]\s+(.*)$/
      while (i < lines.length && re.test(lines[i])) {
        items.push(re.exec(lines[i])![1])
        i++
      }
      blocks.push({ kind: 'list', ordered, items })
      continue
    }

    // GFM table — header row, separator row, then body rows.
    // The separator row is `| --- | :---: | ---: |` and is the
    // anchor we sniff to tell a real table apart from a paragraph
    // that happens to contain a pipe (rare but possible).
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
      !/^[-*]\s+/.test(lines[i]) &&
      !/^\d+[.)]\s+/.test(lines[i])
    ) {
      para.push(lines[i])
      i++
    }
    blocks.push({ kind: 'paragraph', text: para.join('\n') })
  }
  return blocks
}

// ---------- Table helpers ----------

/**
 * Split a single `| a | b | c |` row into the cell strings.
 * Leading and trailing pipes are tolerated but ignored, and
 * cells with surrounding whitespace get trimmed. Empty
 * cells (`||`) are preserved as `''` so the column count
 * stays consistent with the header.
 */
function splitTableRow(line: string): string[] {
  let s = line.trim()
  if (s.startsWith('|')) s = s.slice(1)
  if (s.endsWith('|')) s = s.slice(0, -1)
  return s.split('|').map((c) => c.trim())
}

/**
 * Parse the alignment markers from the separator row. Each
 * column gets `left` (`:---`), `center` (`:---:`), `right`
 * (`---:`), or `left` (plain `---`). Anything we don't
 * understand falls back to left, which matches GFM.
 */
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
// Runs on a single line (or single paragraph blob). Tokenises
// out of order, deliberately — we apply the replacements in a
// fixed sequence on the already-`&<>`-escaped string so the
// earlier passes can't accidentally eat tokens later passes
// would have matched.

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function parseInline(input: string): ReactNode {
  // First, pull out fenced inline `code` spans so their content
  // is left untouched by the bold/italic/link passes below.
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
  // We render bold then italic then links. `parseInline` is
  // already HTML-escaped, so each match is a literal pattern
  // against the escaped text. The risk of recursive replacement
  // is small because none of the markers (**, *, [, ]) get
  // escaped by `escapeHtml`.
  const nodes: ReactNode[] = []
  // Bold: **text** (greedy but not across newlines)
  // Italic: *text* (single asterisk; double-star must come first)
  // Link:  [text](url)
  const re = /(\*\*([^*\n]+)\*\*|\*([^*\n]+)\*|\[([^\]\n]+)\]\((https?:\/\/[^\s)]+)\))/g
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
    } else if (m[4] != null && m[5] != null) {
      nodes.push(
        <a
          key={key++}
          href={m[5]}
          target="_blank"
          rel="noreferrer noopener"
          className="text-excel-green underline underline-offset-2 hover:text-excel-green-light"
        >
          {m[4]}
        </a>,
      )
    }
    last = m.index + m[0].length
  }
  if (last < s.length) {
    nodes.push(<Fragment key={key++}>{s.slice(last)}</Fragment>)
  }
  // Hard line breaks (two trailing spaces) get rendered as <br/>.
  // We split on that pattern after the bold/italic pass to keep
  // the regexes above single-line.
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
 * and return the block + the remaining text. The LLM is
 * allowed to write its private reasoning inline (as opposed
 * to using the structured `reasoning` field) and the task
 * pane shouldn't render the literal tags.
 *
 * Returns null for `text` when the message was nothing but a
 * thinking block (caller can then choose to hide the message
 * body entirely).
 */
export function splitThinking(
  raw: string,
): { thinking: string | null; text: string | null } {
  const m = THINKING_RE.exec(raw)
  if (!m) return { thinking: null, text: raw.length > 0 ? raw : null }
  const thinking = m[1].trim()
  const rest = raw.replace(m[0], '').trim()
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
                ? 'text-base font-semibold text-neutral-900'
                : b.level === 2
                  ? 'text-sm font-semibold text-neutral-900'
                  : 'text-sm font-medium text-neutral-800'
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
            // Convert two-or-more trailing spaces + newline into
            // hard breaks. parseInline already handles single
            // newlines as <br>; here we only need to split out
            // the actual paragraphs from the joined block.
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
              <pre
                key={i}
                className="my-3 overflow-x-auto rounded-lg bg-neutral-900 px-3 py-2 text-xs leading-relaxed text-neutral-100"
              >
                <code className="font-mono">{b.text}</code>
              </pre>
            )
          }
          case 'list': {
            const Tag = b.ordered ? 'ol' : 'ul'
            return (
              <Tag
                key={i}
                className={`my-2 space-y-1 pl-5 text-sm leading-relaxed text-neutral-800 ${
                  b.ordered ? 'list-decimal' : 'list-disc'
                }`}
              >
                {b.items.map((it, j) => (
                  <li key={j}>{parseInline(escapeHtml(it))}</li>
                ))}
              </Tag>
            )
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
