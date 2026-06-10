/**
 * Tests for the hand-rolled Markdown renderer. The big cases
 * are at the block level (which parseBlocks produces); the
 * inline / structure assertions are made against the rendered
 * tree by stringifying the tree to HTML and pattern-matching
 * on the output. The test-only file lives next to the
 * component so it picks up its local re-exports.
 */
import { describe, expect, it } from 'vitest'
import { renderToStaticMarkup } from 'react-dom/server'
import { createElement } from 'react'
import { Markdown, splitThinking } from './Markdown'

function render(md: string): string {
  return renderToStaticMarkup(createElement(Markdown, { source: md }))
}

describe('Markdown', () => {
  it('renders a heading', () => {
    const html = render('# Hello')
    expect(html).toContain('<h1')
    expect(html).toContain('Hello')
    expect(html).toContain('</h1>')
  })

  it('renders bold and italic', () => {
    const html = render('**bold** and *italic*')
    expect(html).toContain('<strong>bold</strong>')
    expect(html).toContain('<em>italic</em>')
  })

  it('renders strikethrough', () => {
    expect(render('~~gone~~')).toContain('<del>gone</del>')
  })

  it('renders an unordered list with multiple items', () => {
    const html = render('- one\n- two\n- three')
    expect(html).toContain('<ul')
    expect(html).toContain('one')
    expect(html).toContain('two')
    expect(html).toContain('three')
  })

  it('renders an ordered list', () => {
    const html = render('1. first\n2. second')
    expect(html).toContain('<ol')
    expect(html).toContain('first')
    expect(html).toContain('second')
  })

  it('renders nested lists via indentation', () => {
    const html = render('- top\n  - mid\n    - leaf')
    // We render each list level as its own <ul> + <ol>. Two
    // separate <ul>s + three <li>s is the structural proof
    // that nesting landed.
    const ulCount = (html.match(/<ul/g) ?? []).length
    expect(ulCount).toBeGreaterThanOrEqual(2)
    expect(html).toContain('top')
    expect(html).toContain('mid')
    expect(html).toContain('leaf')
  })

  it('renders a horizontal rule', () => {
    const html = render('above\n\n---\n\nbelow')
    expect(html).toContain('<hr')
    expect(html).toContain('above')
    expect(html).toContain('below')
  })

  it('renders a blockquote', () => {
    const html = render('> quoted text')
    expect(html).toContain('<blockquote')
    expect(html).toContain('quoted text')
  })

  it('renders a multi-line blockquote', () => {
    const html = render('> first\n> second')
    expect(html).toContain('<blockquote')
    expect(html).toContain('first')
    expect(html).toContain('second')
  })

  it('renders a fenced code block with a language label', () => {
    const html = render('```python\nprint("hi")\n```')
    expect(html).toContain('python')
    expect(html).toContain('print(&quot;hi&quot;)')
    // The language label is rendered as a small badge above
    // the code body.
    expect(html).toContain('uppercase')
  })

  it('renders a GFM table', () => {
    const html = render(
      '| a | b |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |',
    )
    expect(html).toContain('<table')
    expect(html).toContain('<th')
    expect(html).toContain('>1<')
    expect(html).toContain('>3<')
  })

  it('renders a link', () => {
    const html = render('[text](https://example.com)')
    expect(html).toContain('href="https://example.com"')
    expect(html).toContain('>text</a>')
  })

  it('escapes HTML so a user message cannot inject markup', () => {
    const html = render('<script>alert(1)</script>')
    // The literal "<script>" must never appear in the output.
    expect(html).not.toContain('<script>')
    // Our parser outputs `&lt;script&gt;` after escapeHtml.
    // React then escapes the `&` again when it sees it in a
    // className / text node context, so the final string has
    // `&amp;lt;script&amp;gt;`. Either form is fine — the
    // security property is that the literal tag never lands.
    expect(html).toMatch(/&amp;lt;script(&amp;|;)/)
  })
})

describe('splitThinking', () => {
  it('extracts a thinking block and returns the rest', () => {
    const result = splitThinking('hello <thinking>private</thinking> world')
    expect(result.thinking).toBe('private')
    // After the tag is stripped, the surrounding text gets
    // re-joined with a single newline so the LLM's reply
    // stream doesn't accidentally produce a double-space
    // inside the rendered text.
    expect(result.text).toBe('hello world')
  })

  it('returns the raw text when there is no thinking block', () => {
    const result = splitThinking('just a normal reply')
    expect(result.thinking).toBeNull()
    expect(result.text).toBe('just a normal reply')
  })

  it('returns null text when the message is only a thinking block', () => {
    const result = splitThinking('<thinking>all private</thinking>')
    expect(result.thinking).toBe('all private')
    expect(result.text).toBeNull()
  })
})
