/**
 * ChatPanel — main task pane surface.
 *
 * Layout follows the M365 Copilot "AI Assistant" pattern: messages
 * are split into AI (left, with avatar + name + small action row
 * below) and user (right, single bubble), input is a single tall
 * rounded box with the send button tucked into the bottom-right
 * corner, a small footer disclaimer anchors the bottom.
 */
import { useState } from 'react'
import type { Message, ToolCall } from '@core/types'
import { clearChatHistory, useChatStore } from '../store/chat'
import { useSettingsStore } from '../store/settings'

export function ChatPanel() {
  const state = useChatStore((s) => s.state)
  const send = useChatStore((s) => s.send)
  const abort = useChatStore((s) => s.abort)

  const apiKey = useSettingsStore((s) => s.apiKey)

  const [input, setInput] = useState('')

  const isStreaming = state.status === 'streaming' || state.status === 'executing'

  const canSend =
    !isStreaming && input.trim().length > 0 && apiKey.length > 0

  const onSend = (): void => {
    const text = input.trim()
    if (!text || isStreaming) return
    setInput('')
    void send(text)
  }

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      onSend()
    }
  }

  return (
    <div className="flex h-full flex-col">
      {/* Message list */}
      <div className="flex-1 overflow-y-auto px-4 py-4 space-y-5">
        {state.messages.length === 0 && !state.draft && (
          <EmptyState hasApiKey={apiKey.length > 0} />
        )}

        {state.messages.map((m, i) => (
          <MessageBubble key={i} message={m} />
        ))}

        {state.draft && (
          <MessageBubble
            message={{
              role: 'assistant',
              content: state.draft,
              ...(state.draftReasoning
                ? { reasoning: state.draftReasoning }
                : {}),
            }}
            streaming
          />
        )}

        {state.status === 'error' && state.error && (
          <div className="rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
            {state.error.message}
          </div>
        )}

        {state.toolCalls.length > 0 && (
          <div className="space-y-1.5">
            {state.toolCalls.map((tc) => (
              <ToolCallCard key={tc.id} tc={tc} />
            ))}
          </div>
        )}

        {state.messages.length > 0 && !isStreaming && (
          <div className="pt-2 text-center">
            <button
              type="button"
              onClick={() => {
                if (window.confirm('Clear the entire conversation history?')) {
                  clearChatHistory()
                }
              }}
              className="text-[11px] text-neutral-400 underline-offset-2 hover:text-neutral-600 hover:underline"
            >
              Clear history
            </button>
          </div>
        )}
      </div>

      {/* Input + footer */}
      <div className="border-t border-neutral-100 bg-white px-3 pb-2 pt-3">
        <form
          onSubmit={(e) => {
            e.preventDefault()
            onSend()
          }}
          className="relative"
        >
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={onKeyDown}
            placeholder={
              !apiKey
                ? 'Add an API key in Settings first…'
                : isStreaming
                  ? 'Streaming…'
                  : 'Ask anything about your document…'
            }
            disabled={isStreaming || !apiKey}
            rows={3}
            className="w-full resize-none rounded-2xl border border-neutral-300 bg-white px-3.5 py-2.5 pb-10 pr-12 text-sm leading-relaxed text-neutral-800 shadow-sm transition-colors placeholder:text-neutral-400 hover:border-neutral-400 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:bg-neutral-50 disabled:text-neutral-400"
          />
          {/* Floating action buttons in the bottom-right of the textarea */}
          <div className="pointer-events-none absolute bottom-2 right-2 flex items-center gap-1">
            {isStreaming ? (
              <button
                type="button"
                onClick={abort}
                className="pointer-events-auto rounded-full bg-neutral-700 px-3 py-1 text-xs font-medium text-white shadow hover:bg-neutral-800"
              >
                Stop
              </button>
            ) : (
              <button
                type="submit"
                disabled={!canSend}
                aria-label="Send"
                className="pointer-events-auto flex h-7 w-7 items-center justify-center rounded-full bg-blue-500 text-white shadow-sm transition-opacity hover:bg-blue-600 disabled:cursor-not-allowed disabled:bg-neutral-300"
              >
                <IconSend />
              </button>
            )}
          </div>
        </form>
        <p className="mt-2 text-center text-[10px] text-neutral-400">
          AI can make mistakes. Verify important information.
        </p>
      </div>
    </div>
  )
}

// ---------- Message bubble ----------

function MessageBubble({
  message,
  streaming,
}: {
  message: Message
  streaming?: boolean
}) {
  const isUser = message.role === 'user'
  // `reasoning` only lives on assistant messages. Narrow the union
  // before reading it so TS doesn't complain about the tool variant.
  const reasoning =
    message.role === 'assistant' ? message.reasoning : undefined

  if (isUser) {
    return (
      <div className="flex justify-end">
        <div className="max-w-[85%] rounded-2xl rounded-br-md bg-blue-500 px-3.5 py-2 text-sm leading-relaxed text-white">
          <p className="whitespace-pre-wrap break-words">
            {message.content}
            {streaming ? (
              <span className="ml-0.5 animate-pulse">▍</span>
            ) : null}
          </p>
        </div>
      </div>
    )
  }

  return (
    <div className="flex items-start gap-2.5">
      <AssistantAvatar />
      <div className="min-w-0 flex-1">
        <div className="text-xs font-medium text-neutral-700">
          AI Assistant
        </div>
        {reasoning ? (
          <details className="mb-2 mt-1 rounded-lg border border-neutral-200 bg-neutral-50 px-2.5 py-1.5 text-xs text-neutral-600">
            <summary className="cursor-pointer select-none font-medium">
              thinking
            </summary>
            <pre className="mt-1.5 whitespace-pre-wrap break-words font-sans">
              {reasoning}
            </pre>
          </details>
        ) : null}
        <div className="text-sm leading-relaxed text-neutral-800">
          <p className="whitespace-pre-wrap break-words">
            {message.content}
            {streaming ? (
              <span className="ml-0.5 animate-pulse">▍</span>
            ) : null}
          </p>
        </div>
        {!streaming && message.content ? (
          <div className="mt-1.5 flex items-center gap-0.5 text-neutral-400">
            <IconButton label="Copy">
              <IconCopy />
            </IconButton>
            <IconButton label="Good response">
              <IconThumbUp />
            </IconButton>
            <IconButton label="Bad response">
              <IconThumbDown />
            </IconButton>
            <IconButton label="Regenerate">
              <IconRefresh />
            </IconButton>
          </div>
        ) : null}
      </div>
    </div>
  )
}

function AssistantAvatar() {
  return (
    <div className="mt-0.5 flex h-7 w-7 flex-shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-blue-500 to-blue-600 text-white">
      <svg
        width="14"
        height="14"
        viewBox="0 0 16 16"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        aria-hidden="true"
      >
        <path
          d="M3 11.5C3 11.5 4.5 12 6 12C7.5 12 8 11 8 11C8 11 8.5 9 10 8.5C11.5 8 13 8 13 8M5 5.5C5 4.67 5.67 4 6.5 4C7.33 4 8 4.67 8 5.5M3 8.5L5 6.5L7 8.5"
          stroke="currentColor"
          strokeWidth="1.6"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    </div>
  )
}

function IconButton({
  label,
  children,
}: {
  label: string
  children: React.ReactNode
}) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      className="rounded p-1 transition-colors hover:bg-neutral-100 hover:text-neutral-700"
    >
      {children}
    </button>
  )
}

// ---------- Tool call card ----------

function ToolCallCard({ tc }: { tc: ToolCall }) {
  return (
    <div className="ml-9 flex items-center gap-2 rounded-lg border border-neutral-200 bg-white px-2.5 py-1.5 text-xs">
      <span className="font-mono text-neutral-700">{tc.name}</span>
      <ToolStatusPill status={tc.status} />
      {tc.status === 'error' && tc.error ? (
        <span className="truncate text-red-600">{tc.error.message}</span>
      ) : tc.status === 'success' ? (
        <span className="truncate text-neutral-500">
          {summarizeResult(tc.result)}
        </span>
      ) : null}
    </div>
  )
}

function ToolStatusPill({ status }: { status: ToolCall['status'] }) {
  const label =
    status === 'pending'
      ? 'pending'
      : status === 'running'
        ? 'running…'
        : status === 'success'
          ? 'ok'
          : 'error'
  const colour =
    status === 'pending'
      ? 'bg-neutral-200 text-neutral-600'
      : status === 'running'
        ? 'bg-blue-100 text-blue-700'
        : status === 'success'
          ? 'bg-green-100 text-green-700'
          : 'bg-red-100 text-red-700'
  return (
    <span
      className={`rounded-full px-2 py-0.5 text-[10px] font-medium ${colour}`}
    >
      {label}
    </span>
  )
}

function summarizeResult(result: unknown): string {
  if (result == null) return '—'
  if (typeof result === 'string') return result
  if (typeof result !== 'object') return String(result)
  try {
    const json = JSON.stringify(result)
    if (json.length > 80) return json.slice(0, 77) + '…'
    return json
  } catch {
    return '[unserialisable]'
  }
}

// ---------- Empty state ----------

function EmptyState({ hasApiKey }: { hasApiKey: boolean }) {
  return (
    <div className="mt-6 px-2 text-center">
      <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-blue-50 text-blue-500">
        <svg
          width="24"
          height="24"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.6"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          <path d="M12 2L9 9l-7 1 5 5-1.5 7L12 18l6.5 4L17 15l5-5-7-1z" />
        </svg>
      </div>
      <h2 className="text-sm font-semibold text-neutral-800">
        {hasApiKey ? 'How can I help?' : 'Set up your API key'}
      </h2>
      <p className="mt-1 text-xs text-neutral-500">
        {hasApiKey
          ? 'Ask me to write formulas, analyze data, create charts, or explain cell ranges.'
          : 'Open Settings (gear icon, top right) to add an API key, then come back.'}
      </p>
    </div>
  )
}

// ---------- Icons ----------

function IconSend() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M22 2L11 13" />
      <path d="M22 2L15 22L11 13L2 9L22 2Z" />
    </svg>
  )
}

function IconCopy() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <rect x="9" y="9" width="13" height="13" rx="2" />
      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
    </svg>
  )
}

function IconThumbUp() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M7 10v12" />
      <path d="M15 5.88L14 10h5.83a2 2 0 0 1 1.92 2.56l-2.33 8A2 2 0 0 1 17.5 22H7V10l4.34-9.66a1 1 0 0 1 1.85.4L15 5.88Z" />
    </svg>
  )
}

function IconThumbDown() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M17 14V2" />
      <path d="M9 18.12L10 14H4.17a2 2 0 0 1-1.92-2.56l2.33-8A2 2 0 0 1 6.5 2H17v12l-4.34 9.66a1 1 0 0 1-1.85-.4L9 18.12Z" />
    </svg>
  )
}

function IconRefresh() {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M21 12a9 9 0 1 1-3-6.7L21 8" />
      <path d="M21 3v5h-5" />
    </svg>
  )
}
