// One-off probe — hit each module and report status.
const paths = [
  '/src/taskpane/index.tsx',
  '/src/taskpane/App.tsx',
  '/src/taskpane/store/chat.ts',
  '/src/taskpane/store/settings.ts',
  '/src/taskpane/components/ChatPanel.tsx',
  '/src/taskpane/components/SettingsPanel.tsx',
  '/src/core/providers/registry.ts',
  '/src/core/providers/anthropic.ts',
  '/src/core/agent/orchestrator.ts',
  '/src/core/agent/reducer.ts',
  '/node_modules/.vite/deps/ai.js',
  '/node_modules/.vite/deps/@ai-sdk_anthropic.js',
]
const BASE = 'https://localhost:3721'
const out = {}
for (const p of paths) {
  try {
    const r = await fetch(BASE + p)
    const t = await r.text()
    out[p] = { status: r.status, len: t.length, head: t.slice(0, 80) }
  } catch (e) {
    out[p] = 'ERR: ' + e.message
  }
}
console.log(JSON.stringify(out, null, 2))
