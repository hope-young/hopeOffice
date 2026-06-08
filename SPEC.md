# hope-Office SPEC

> 从零重写的 Office (Word/Excel/PPT) AI 助手 add-in。本文档是 2026-06-08 起开发的源真相。

## 1. 目标

本地优先、BYO API key 的 Office AI 助手。用户聊天描述想做什么，agent 写 office.js 代码并 sandbox 执行。比 Copilot 自由，比传统 wrapper 覆盖面广。

## 2. 非目标

- 不做 M365 账号 / SSO 集成
- 不打包成 Windows 安装器（sideload manifest 即可）
- 不接 stdio MCP（只 HTTP/SSE）
- 不做跨应用上下文（单 add-in，三 host 各跑各的）

## 3. 技术栈

| 维度 | 选 | 版本 |
|---|---|---|
| 语言 | TypeScript（strict）| 5.x |
| UI | React | 19 |
| 构建 | Vite + HTTPS dev cert | 8.x |
| 样式 | Tailwind + 自写 primitives（暂不引 shadcn，避免重）| 4.x |
| 状态 | Zustand | 5.x |
| LLM | Vercel AI SDK | 6.x |
| Provider adapters | `@ai-sdk/anthropic` `@ai-sdk/openai` `@ai-sdk/openai-compatible` | 各自最新版 |
| MCP | `@ai-sdk/mcp` | 1.x |
| 校验 | Zod | 4.x |
| Office | `@types/office-js` + `@microsoft/office-js` CDN | 1.x |
| 测试 | Vitest + jsdom | 3.x |

## 4. 目录结构

```
src/
├── taskpane/                React 应用，DOM-only
│   ├── index.tsx            入口，Office.onReady → ReactDOM
│   ├── App.tsx
│   ├── components/          ChatPanel, HistoryPanel, SettingsPanel, CodeBlock, MessageBubble, ToolActivity
│   ├── executor/            sandbox.ts (postMessage 桥) + iframe.html
│   ├── ribbon/              commands.ts (从 ribbon 触发的动作 dispatch)
│   ├── i18n/                registry / context / locales/{en,zh,he}.json
│   └── store/               chat.ts settings.ts history.ts  (Zustand)
├── core/                    纯逻辑，无 React 无 Office
│   ├── agent/
│   │   ├── reducer.ts       纯函数 reducer (核心循环)
│   │   ├── orchestrator.ts  side-effect 编排（订阅 reducer，驱动 SDK 流）
│   │   └── types.ts
│   ├── providers/
│   │   ├── interface.ts     ChatProvider 统一接口
│   │   ├── anthropic.ts openai.ts openai-compatible.ts
│   │   └── registry.ts
│   ├── skills/
│   │   ├── types.ts         Skill<TArgs,TResult> 接口
│   │   ├── word/  excel/  powerpoint/   各自 TS 函数
│   │   └── index.ts
│   ├── mcp/client.ts
│   └── schema/              Zod
└── ...

manifest.xml                 dev (localhost:3721)
manifest.production.xml      prod (hope-young.github.io/hope-Office/)
```

**关键分割**：`taskpane/` = 浏览器世界，`core/` = 纯逻辑（可被 reducer、ribbon commands、test 同时调用）。`core/` 不 import `@fluentui/*` 或任何 DOM-only API。

## 5. Agent 循环（核心 reducer 风格）

跟旧参考**不同**的关键点——拒绝大 orchestrator.ts。

```typescript
// core/agent/types.ts
type AgentState = {
  messages: Message[]            // 已收齐的对话
  draft: string                  // 当前流式增量
  toolCalls: ToolCall[]
  status: 'idle' | 'streaming' | 'awaiting-approval' | 'executing' | 'error'
  cost: Cost                     // 本轮累计
  approval?: { code: string; resolve: (ok: boolean) => void }
  error?: Error
}

type AgentEvent =
  | { type: 'user-send'; text: string }
  | { type: 'stream-token'; token: string }
  | { type: 'stream-end' }
  | { type: 'tool-call-start'; tool: string; args: unknown }
  | { type: 'tool-call-result'; result: unknown }
  | { type: 'approval-requested'; code: string }
  | { type: 'approval-decided'; approved: boolean }
  | { type: 'error'; error: Error }
  | { type: 'reset' }

function reduce(state: AgentState, event: AgentEvent): AgentState
```

**reducer 是纯函数**。副作用（调 SDK、postMessage 到 sandbox、dispatch 下一个 event）由 `orchestrator.ts` 编排。reducer 单测极其便宜，state transitions 完全可追溯。

## 6. Provider 抽象

```typescript
// core/providers/interface.ts
interface ChatProvider {
  id: string                   // 'anthropic' | 'openai' | 'openai-compatible'
  name: string
  streamChat(opts: {
    messages: Message[]
    tools: ToolDef[]
    signal: AbortSignal
  }): AsyncIterable<StreamChunk>

  listModels(opts: { apiKey: string; baseUrl?: string }): Promise<ModelInfo[]>
  estimateCost(usage: TokenUsage, model: string): Cost
}
```

每个 provider 一个薄 adapter 文件，把 SDK 的 stream 转成统一 `StreamChunk`。UI 只看 `ChatProvider` 接口，不认 SDK 名字。

## 7. 技能系统

```typescript
// core/skills/types.ts
interface Skill<TArgs, TResult> {
  name: string                          // 'add-table'
  description: string                   // 给 LLM 看的
  host: HostKind[]                      // ['word'] 或 ['word','excel']
  args: ZodSchema<TArgs>
  execute(args: TArgs, ctx: Ctx): Promise<TResult>
}
```

**不是 markdown 文档 + 字符串检索**，是**类型化 TS 函数**。LLM 通过 `execute_skill(name, args)` 调用，参数经过 Zod 校验。Word / Excel / PowerPoint 各有 `skills/<host>/<name>.ts` 文件，函数体直接 import `@microsoft/office-js`。

MCP server 工具由 orchestrator 启动时注册，行为跟 Skill 一样对外。

## 8. Custom Tab Ribbon

manifest 顶层用 `<ExtensionPoint xsi:type="CustomTab">`，独立 `hopeOfficeTab`，3 个 group：

```xml
<CustomTab id="hopeOfficeTab">
  <Group id="grpMain">
    <Label resid="grpMainLabel" />
    <Control xsi:type="Button" id="openTaskPane">
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>hopeOfficeTaskPane</TaskpaneId>
        <SourceLocation resid="taskpaneUrl" />
      </Action>
    </Control>
    <Control xsi:type="Button" id="newChat"><Action xsi:type="ExecuteFunction"><FunctionName>newChat</FunctionName></Action></Control>
  </Group>
  <Group id="grpHistory">
    <Label resid="grpHistoryLabel" />
    <Control xsi:type="Button" id="openHistory" />
    <Control xsi:type="Button" id="exportCurrent" />
  </Group>
  <Group id="grpSettings">
    <Label resid="grpSettingsLabel" />
    <Control xsi:type="Button" id="openSettings" />
    <Control xsi:type="Button" id="manageMcp" />
  </Control>
</CustomTab>
```

`ExecuteFunction` 触发的 handler 写在 `src/taskpane/ribbon/commands.ts`，通过 `Office.actions.associate` 绑定。`ShowTaskpane` 之外的所有按钮都用 `ExecuteFunction`（不是 `ShowTaskpane`），因为它们要在 task pane 已开的情况下做事。

## 9. 状态（Zustand）

3 个 store，分别管一件事：

```typescript
// taskpane/store/chat.ts
type ChatStore = {
  state: AgentState
  dispatch: (e: AgentEvent) => void
  // 内部: subscribe to orchestrator, drive reducer
}

// taskpane/store/settings.ts
type SettingsStore = {
  providerId: string
  apiKey: string
  model: string
  baseUrl: string
  autoApprove: boolean
  mcpServers: McpServer[]
  // 持久化到 Office.context.roamingSettings
}

// taskpane/store/history.ts
type HistoryStore = {
  index: ConversationSummary[]
  activeId: string | null
  // localStorage blob
}
```

## 10. i18n

自写最小化。2 locale：`zh-CN`（首选）`en`。
- `registry.ts` 注册 locale + fallback
- `context.tsx` React context 暴露 `t(key, vars)`
- `locales/<lang>.json` 扁平 key（`chat.welcomeTitle`）
- 持久化键 `hope-office_language`

`gen:i18n` 脚本从 locales 生成 `keys.generated.ts` 类型。**只用 key，不引 i18next、react-intl 之类的库**。

## 11. 持久化

| 数据 | 位置 |
|---|---|
| Settings | `Office.context.roamingSettings`（生产）/`localStorage`（dev） |
| Chat history index | `localStorage` 键 `hope-office_history_index` |
| Chat history blob | `localStorage` 键 `hope-office_history_conv_<id>` |
| Language | `localStorage` 键 `hope-office_language` |

## 12. 开发

```
npm run dev          vite + HTTPS on :3721
npm run sideload     npm run dev + office-addin-debugging start
npm run check:i18n   校验 3 个 locale key shape 一致
npm test             vitest run
npm run build        tsc + vite build → dist/
```

`.github/workflows/deploy.yml` push master 自动 build + deploy 到 GitHub Pages。

## 13. Roadmap

- [ ] **W24** SPEC + 骨架（manifest + vite + ts + 空 App + 能 sideload）
- [ ] **W24** Provider 配置 + 一条 LLM 打通
- [ ] **W25** Chat 流（user send → stream → display） + Sandbox iframe
- [ ] **W25-W26** Word 技能（先接 Word，验证整条链）
- [ ] **W26** Excel + PowerPoint 技能
- [ ] **W27** MCP 接入
- [ ] **W27** i18n (zh-CN/en) + 完整设置面板
- [ ] **W28+** 多会话、历史面板、Slash 命令、导出 md

## 14. 风险

- Custom Tab 在不同 Office 版本渲染差异（实测后再修）
- sandbox iframe 跨域限制：必须用 `postMessage` 桥，不能直接调 office.js
- reducer 风格简洁但异步编排（orchestrator）需要写得更小心，state 不会"自动更新"
- 自写 primitives 比 Fluent UI 慢，但更小、更"不像是 fork"

## 15. 不归档

- `docs/superpowers/` —— 上游内部设计文档，不属于本项目
- `installer/` —— Inno Setup 脚本，本项目不打安装器
- Co-Authored-By trailer —— commit 只署 `hope_young`
