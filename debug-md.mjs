import { parseBlocks, parseInline, splitThinking } from './src/taskpane/components/Markdown.tsx'

const md = '- top\n  - mid\n    - leaf'
console.log(JSON.stringify(parseBlocks(md), null, 2))
