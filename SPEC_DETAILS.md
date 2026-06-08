# hope-Office SPEC — Implementation Details

> SPEC.md 给的是架构骨架（做什么、为什么、怎么分模块）。
> 本附录给的是**实现细节**（具体类型、协议、算法、示例），足够另一个 AI 接手按图施工。
> SPEC.md 里没展开的 6 块全在这里。

---

## 1. 基础类型（Message / StreamChunk / Cost / ToolCall）

`src/core/types.ts`

```typescript
// ---------- Conversation ----------

type Role = 'user' | 'assistant' | 'tool'

type Message =
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string; toolCalls?: ToolCall[] }
  | { role: 'tool'; toolCallId: string; content: string }

// 一次 LLM 流式响应里的一个 chunk
type StreamChunk =
  | { type: 'text-delta';     delta: string }
  | { type: 'reasoning-delta'; delta: string }                       // thinking 模型
  | { type: 'tool-call-start'; toolCall: ToolCall }
  | { type: 'tool-call-args';  toolCallId: string; delta: unknown }   // JSON 增量
  | { type: 'finish';          reason: 'stop' | 'tool-calls' | 'length' | 'error'; error?: string }
  | { type: 'usage';           inputTokens: number; outputTokens: number }

// ---------- Tool call（LLM 决定调什么工具） ----------

type ToolCall = {
  id: string                                                       // 一次对话内唯一
  name: string                                                     // 'execute_skill' | 'execute_code' | '<mcp-tool-name>'
  args: unknown                                                    // Zod 校验后用
  result?: unknown
  status: 'pending' | 'running' | 'success' | 'error'
  error?: { message: string; stack?: string }
  startedAt?: number
  completedAt?: number
}

// ---------- Cost（每次 LLM 调用的成本） ----------

type CostSource =
  | 'gateway-exact'         // Vercel AI Gateway 准确
  | 'openrouter-exact'      // OpenRouter 准确
  | 'estimated'             // 用本地价目表估算
  | 'tokens-only'           // 没有价目表
  | 'local-free'            // Ollama / LM Studio

type Cost = {
  input: number
  output: number
  total: number
  source: CostSource
  pricingVersion?: string   // 价目表版本（"estimated" 时必填）
}
```

**为什么这样切**：`Message` 只描述已收齐的对话；`StreamChunk` 描述正在收的增量；`ToolCall` 是 LLM 决策的"工单"。三者**不可混用**——常见 bug 就是把 `Message` 当 `StreamChunk` 累加。

---

## 2. Sandbox iframe 的 postMessage 协议

`src/taskpane/executor/sandbox.ts`（parent 端）

```typescript
type ParentToIframe =
  | { type: 'execute'; id: string; code: string; host: 'word' | 'excel' | 'powerpoint' }
  | { type: 'ping';    id: string }

type IframeToParent =
  | { type: 'ready' }
  | { type: 'pong';    id: string }
  | { type: 'host';    host: 'word' | 'excel' | 'powerpoint' | 'unsupported' }
  | { type: 'result';  id: string; ok: true;  value: unknown }
  | { type: 'result';  id: string; ok: false; error: { message: string; stack?: string } }

const PENDING_TIMEOUT_MS = 30_000
```

`src/taskpane/executor/iframe.html`（iframe 端，简化骨架）

```html
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
<script>
  const pending = new Map()  // id -> { resolve, reject, timer }

  window.addEventListener('message', async (e) => {
    if (e.source !== window.parent) return
    const msg = e.data
    if (!msg) return

    if (msg.type === 'ping') {
      window.parent.postMessage({ type: 'pong', id: msg.id }, '*')
      return
    }

    if (msg.type === 'execute') {
      const { id, code, host } = msg
      // 真实 host 在 Office 加载后才能知道，告诉 parent
      const realHost = Office.context.host === Office.HostType.Word    ? 'word'
                     : Office.context.host === Office.HostType.Excel   ? 'excel'
                     : Office.context.host === Office.HostType.PowerPoint ? 'powerpoint'
                     : 'unsupported'
      window.parent.postMessage({ type: 'host', host: realHost }, '*')

      if (realHost !== host) {
        window.parent.postMessage({
          type: 'result', id, ok: false,
          error: { message: `Host mismatch: manifest says ${host}, Office says ${realHost}` }
        }, '*')
        return
      }

      try {
        const runner = (realHost === 'word' ? Word.run
                      : realHost === 'excel' ? Excel.run
                      : PowerPoint.run).bind(null, async (ctx) => {
          // 动态执行用户代码
          // eslint-disable-next-line no-new-func
          const userFn = new Function('context', `return (async () => { ${code} })()`)
          const value = await userFn(ctx)
          await ctx.sync()
          return value
        })
        const value = await runner()
        window.parent.postMessage({ type: 'result', id, ok: true, value }, '*')
      } catch (err) {
        window.parent.postMessage({
          type: 'result', id, ok: false,
          error: { message: String(err?.message ?? err), stack: err?.stack }
        }, '*')
      }
    }
  })

  // 告诉 parent iframe 已就绪
  window.parent.postMessage({ type: 'ready' }, '*')
</script>
```

**Parent 端的 `runInSandbox` 收尾**：

```typescript
async function runInSandbox(code: string, host: HostKind, signal: AbortSignal): Promise<unknown> {
  const id = crypto.randomUUID()
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      window.removeEventListener('message', listener)
      reject(new Error('Sandbox timeout after 30s'))
    }, PENDING_TIMEOUT_MS)

    const listener = (e: MessageEvent) => {
      const msg = e.data as IframeToParent
      if (msg?.type !== 'result' || msg.id !== id) return
      clearTimeout(timer)
      window.removeEventListener('message', listener)
      if (msg.ok) resolve(msg.value)
      else reject(new Error(msg.error.message))
    }
    window.addEventListener('message', listener)

    signal.addEventListener('abort', () => {
      clearTimeout(timer)
      window.removeEventListener('message', listener)
      reject(new Error('Aborted'))
    })

    iframeRef.contentWindow!.postMessage({ type: 'execute', id, code, host }, '*')
  })
}
```

**为什么 host 校验在 iframe 里**：Office.js 实际 host 在 parent 里看不到（parent 只看到 task pane URL），只有 iframe 自己能拿到 `Office.context.host`。所以 `host` mismatch 检测放在 iframe。

---

## 3. Host detection

`src/taskpane/lib/host.ts`

```typescript
import type { HostKind } from '../../core/types'

const HOST_MAP: Record<number, HostKind> = {
  [Office.HostType.Word]:       'word',
  [Office.HostType.Excel]:      'excel',
  [Office.HostType.PowerPoint]: 'powerpoint',
}

export function detectHost(): HostKind {
  const numeric = Office.context.host
  const kind = HOST_MAP[numeric]
  if (!kind) throw new Error(`hope-Office does not support this host (Office.host=${numeric})`)
  return kind
}

// 单例：app 启动时调用一次，存到 React context
export const HostContext = createContext<HostKind>(/* never */ (undefined as any))
```

**用 host 的地方**（3 个）：
1. `core/skills/index.ts` 注册时按 host 过滤（Word skill 不给 Excel 用）
2. `core/agent/system-prompt.ts` 选 host-specific 的 system prompt
3. `manifest.xml` 的 `Hosts` 字段（决定 add-in 在哪些 host 注册）

---

## 4. Self-healing retry 算法

`src/core/agent/orchestrator.ts`

```typescript
const MAX_RETRIES = 3
const RETRY_DELAYS_MS = [0, 1_000, 3_000]   // 第 1 次立刻试，失败后等 1s，再失败等 3s

type RetryOutcome =
  | { kind: 'success'; result: unknown }
  | { kind: 'exhausted'; error: Error }
  | { kind: 'unretriable'; error: Error }   // 错误类型不可重试

async function executeWithRetry(
  toolCall: ToolCall,
  run: () => Promise<unknown>,
  signal: AbortSignal
): Promise<RetryOutcome> {
  const err = classifyError
  let lastError: Error | undefined

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (signal.aborted) return { kind: 'unretriable', error: new Error('Aborted') }

    if (RETRY_DELAYS_MS[attempt] > 0) {
      await sleep(RETRY_DELAYS_MS[attempt], signal)
    }

    try {
      const result = await run()
      return { kind: 'success', result }
    } catch (e) {
      lastError = e as Error
      const classified = err(e)
      if (!classified.retriable) return { kind: 'unretriable', error: classified }
    }
  }
  return { kind: 'exhausted', error: lastError ?? new Error('Unknown') }
}
```

**retry 完之后怎么把错误喂回 LLM**：

```typescript
// orchestrator 里
const outcome = await executeWithRetry(toolCall, () => runInSandbox(...), signal)
if (outcome.kind !== 'success') {
  // 把失败作为 tool result 推回对话，让 LLM 决定下一步
  dispatch({
    type: 'tool-call-result',
    toolCallId: toolCall.id,
    result: { error: outcome.error.message, stack: outcome.error.stack },
  })
  // 同时也 dispatch 一个 'user' 消息吗？不要——LLM 已经看到 tool result，
  // 它会自然产生下一步的 'tool-calls' 或 'finish'。
  // 如果 3 次都失败且 LLM 也没好主意，user 会看到 'agent 给最终错误'。
}
```

**为什么 3 次**：经验值。再多就让用户等了。

---

## 5. Provider 错误处理

`src/core/providers/errors.ts`

```typescript
type ProviderErrorKind =
  | 'network'           // fetch failed, DNS, CORS
  | 'rate-limit'        // 429
  | 'auth'              // 401 / 403
  | 'context-overflow'  // 413, "context_length_exceeded"
  | 'model-not-found'   // 404
  | 'server'            // 5xx
  | 'unknown'

type ProviderError = {
  kind: ProviderErrorKind
  message: string
  retriable: boolean
  retryAfterMs?: number    // 429 时从 Retry-After header 读
  raw?: unknown            // 原始错误对象，给 UI / log 用
}

function classifyProviderError(err: unknown): ProviderError {
  // 各 SDK 抛错格式不一样，但都有 status code 或 message 关键字可识别
  const anyErr = err as any
  const status: number | undefined = anyErr?.status ?? anyErr?.statusCode
  const msg: string = anyErr?.message ?? String(err)

  // 1. 网络层
  if (anyErr instanceof TypeError && /fetch|network/i.test(msg)) {
    return { kind: 'network', message: msg, retriable: true }
  }

  // 2. HTTP 状态码
  if (typeof status === 'number') {
    if (status === 401 || status === 403) {
      return { kind: 'auth', message: msg, retriable: false, raw: err }
    }
    if (status === 429) {
      const ra = anyErr?.headers?.['retry-after'] ?? anyErr?.response?.headers?.get?.('retry-after')
      const retryAfterMs = ra ? Number(ra) * 1000 : undefined
      return { kind: 'rate-limit', message: msg, retriable: true, retryAfterMs, raw: err }
    }
    if (status === 413 || /context.?length|too.?long/i.test(msg)) {
      return { kind: 'context-overflow', message: msg, retriable: false, raw: err }
    }
    if (status === 404 || /model.?not.?found|unknown.?model/i.test(msg)) {
      return { kind: 'model-not-found', message: msg, retriable: false, raw: err }
    }
    if (status >= 500) {
      return { kind: 'server', message: msg, retriable: true, raw: err }
    }
  }

  // 3. Provider 特定错误字面量（Anthropic / OpenAI 等 SDK 各有自己的）
  if (/api[_-]?key|authentication/i.test(msg)) {
    return { kind: 'auth', message: msg, retriable: false, raw: err }
  }

  return { kind: 'unknown', message: msg, retriable: false, raw: err }
}
```

**用户 UI 上**：

| 错误类型 | 显示 |
|---|---|
| `auth` | "API key 无效，请在设置里检查" |
| `rate-limit` | "请求太快（X 秒后重试）" |
| `context-overflow` | "对话太长，请 /clear 或开新对话" |
| `network` | "网络问题，正在重试…"（自动 retry） |
| `model-not-found` | "模型 {model} 不存在，请检查设置" |
| `server` | "Provider 服务异常，正在重试…"（自动 retry） |
| `unknown` | 原始 message |

`auth` / `model-not-found` / `context-overflow` **不重试**（重试也是同样的错），直接 dispatch error event 落到 chat 里。

---

## 6. Skill 示例：Word `add-table`

`src/core/skills/word/add-table.ts`

```typescript
import { z } from 'zod'
import type { Skill } from '../types'

const args = z.object({
  rows: z.number().int().min(1).max(50),
  columns: z.number().int().min(1).max(20),
  data: z.array(z.array(z.string())).optional(),     // 可选的预填内容
  hasHeader: z.boolean().default(true),
  position: z.enum(['start', 'end', 'cursor']).default('end'),
})

type AddTableArgs = z.infer<typeof args>
type AddTableResult = { tableId: string; rows: number; columns: number }

export const addTable: Skill<AddTableArgs, AddTableResult> = {
  name: 'add-table',
  description:
    'Insert a table into the current Word document. Optionally pre-fill cells. ' +
    'Returns the table id so follow-up skills can reference it.',
  host: ['word'],
  args,

  async execute(input, ctx) {
    // ctx.executor 是个抽象层：开发时直跑，生产时 postMessage 到 sandbox iframe
    return await ctx.executor.runOnHost('word', async (Word) => {
      return await Word.run(async (context) => {
        const target = (() => {
          if (input.position === 'start') return context.document.body
          if (input.position === 'end')   return context.document.body
          return context.document.getSelection()  // 'cursor'
        })()

        const ref = (input.position === 'start') ? 'Start' : 'End'
        const table = target.insertTable(input.rows, input.columns, ref, input.data ?? [])
        table.styleBuiltIn = Word.Style.wellLightShadingAccent1

        if (input.hasHeader && input.data === undefined) {
          // 标记第一行为表头（Word 表格对象）
          table.rows.getFirst().font.bold = true
        }

        await context.sync()
        return { tableId: String(table.id), rows: input.rows, columns: input.columns }
      })
    })
  },
}
```

**注册到 skill registry**（`src/core/skills/index.ts`）：

```typescript
import { addTable } from './word/add-table'

export const SKILL_REGISTRY = {
  'add-table': addTable,
  // 后面加：'insert-image', 'apply-style', 'create-list', ...
} as const

export type SkillName = keyof typeof SKILL_REGISTRY
```

**LLM 怎么看到这些 skill**（`src/core/agent/system-prompt.ts`）：

```typescript
export function buildSystemPrompt(host: HostKind): string {
  const skills = Object.values(SKILL_REGISTRY)
    .filter(s => s.host.includes(host))
    .map(s => `- **${s.name}**: ${s.description}\n  args: ${s.args.description}`)

  return `You are hope-Office, an AI assistant for ${host} that controls the document by calling skills.

## Available skills
${skills.join('\n')}

## Calling convention
Use the \`execute_skill\` tool with the skill name and a JSON object of args.
For tasks not covered by a skill, use \`execute_code\` to run arbitrary office.js code.`
}
```

**为什么 args 是 Zod**：
- LLM 给的 args 跑前要 Zod parse，错了直接告诉 LLM 哪里错（"row 超过 50"），LLM 自己改
- 沙箱里执行的是已经校验过的数据，不会因为 LLM 瞎写崩

---

## 7. 落地 checklist（按这个顺序写，AI 不会乱）

```
Phase 1 — 骨架
  □ manifest.xml (dev) with Custom Tab
  □ vite.config.ts with HTTPS
  □ tsconfig.json strict
  □ package.json deps per SPEC §3
  □ src/taskpane/index.tsx (Office.onReady → React)
  □ src/taskpane/App.tsx (空壳)
  □ 能 npm run sideload 进 Word，看到空 task pane + Custom Tab

Phase 2 — 类型
  □ src/core/types.ts (per §1)
  □ src/core/providers/interface.ts (per SPEC §6)
  □ src/core/skills/types.ts

Phase 3 — 沙箱
  □ src/taskpane/executor/iframe.html (per §2)
  □ src/taskpane/executor/sandbox.ts (per §2)
  □ 写一个 'echo' 测试 skill，验证 postMessage 跑通

Phase 4 — 一条 LLM
  □ src/core/providers/anthropic.ts (用 @ai-sdk/anthropic 薄包)
  □ src/core/providers/registry.ts
  □ src/core/agent/reducer.ts (per SPEC §5)
  □ src/core/agent/orchestrator.ts (调度 reducer + provider + sandbox)
  □ 跑通：发 "你好" → 流式显示回复

Phase 5 — Skill
  □ 写 add-table (per §6)
  □ 注册到 SKILL_REGISTRY
  □ LLM 调通：发 "在文档末尾插入 3x4 表格" → 看到表格出现

Phase 6 — 三个 host
  □ src/core/skills/word/   (10+ skills)
  □ src/core/skills/excel/  (10+ skills)
  □ src/core/skills/powerpoint/  (10+ skills)
  □ system-prompt 按 host 切换

Phase 7 — 错误处理 + retry
  □ src/core/providers/errors.ts (per §5)
  □ executeWithRetry (per §4)

Phase 8 — Settings + History
  □ taskpane/store/settings.ts (Zustand + 持久化)
  □ taskpane/store/history.ts
  □ taskpane/components/SettingsPanel.tsx
  □ taskpane/components/HistoryPanel.tsx

Phase 9 — i18n
  □ taskpane/i18n/registry.ts
  □ taskpane/i18n/context.tsx
  □ taskpane/i18n/locales/{zh-CN,en}.json
  □ 所有 UI 字符串走 t() 函数

Phase 10 — MCP
  □ src/core/mcp/client.ts (用 @ai-sdk/mcp)
  □ Settings 面板加 MCP server 配置

Phase 11 — Ribbon 完整化
  □ manifest.xml 加 ExecuteFunction 按钮
  □ src/taskpane/ribbon/commands.ts (5 个 command handlers)

Phase 12 — Polish
  □ Export markdown
  □ Slash commands
  □ 性能优化（虚拟滚动、token 预算警告）
```
