/**
 * ChatPanel — the main task pane surface. Renders the message list
 * (committed + streaming draft), a status hint, an error banner,
 * and the input form.
 *
 * Phase 4: text-only chat. Tool calls, slash commands, history,
 * approval UI etc. land in later phases per SPEC §13.
 */
import { useState } from 'react'
import type { Message, ToolCall } from '@core/types'
import { useChatStore } from '../store/chat'
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
      <div className="flex-1 overflow-y-auto px-3 py-2 space-y-3">
        {state.messages.length === 0 && !state.draft && (
          <p className="text-sm text-neutral-400 text-center mt-8">
            {apiKey
              ? 'Send a message to start.'
              : 'Open Settings to add an API key, then come back.'}
          </p>
        )}

        {state.messages.map((m, i) => (
          <MessageBubble key={i} message={m} />
        ))}

        {state.draft && (
          <MessageBubble
            message={{ role: 'assistant', content: state.draft }}
            streaming
          />
        )}

        {state.status === 'error' && state.error && (
          <div className="rounded border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-700">
            {state.error.message}
          </div>
        )}

        {isStreaming && (
          <p className="text-xs text-neutral-400 text-center">
            {state.status === 'executing' ? 'Executing tool…' : 'Streaming…'}
          </p>
        )}

        {state.toolCalls.length > 0 && (
          <div className="space-y-1">
            {state.toolCalls.map((tc) => (
              <div
                key={tc.id}
                className="flex items-center gap-2 rounded border border-neutral-200 bg-white px-2 py-1 text-xs"
              >
                <span className="font-mono text-neutral-700">{tc.name}</span>
                <ToolStatusPill status={tc.status} />
                {tc.status === 'error' && tc.error ? (
                  <span className="text-red-600 truncate">{tc.error.message}</span>
                ) : tc.status === 'success' ? (
                  <span className="text-neutral-500 truncate">
                    {summarizeResult(tc.result)}
                  </span>
                ) : null}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Input */}
      <form
        onSubmit={(e) => {
          e.preventDefault()
          onSend()
        }}
        className="border-t border-neutral-200 p-2 flex gap-2"
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
                : 'Send a message (Enter to send, Shift+Enter for newline)'
          }
          disabled={isStreaming || !apiKey}
          rows={2}
          className="flex-1 resize-none rounded border border-neutral-300 px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-400 disabled:bg-neutral-50 disabled:text-neutral-400"
        />
        {isStreaming ? (
          <button
            type="button"
            onClick={abort}
            className="rounded bg-neutral-200 px-3 py-1 text-sm hover:bg-neutral-300"
          >
            Stop
          </button>
        ) : (
          <button
            type="submit"
            disabled={!canSend}
            className="rounded bg-blue-500 px-3 py-1 text-sm text-white hover:bg-blue-600 disabled:bg-neutral-300 disabled:cursor-not-allowed"
          >
            Send
          </button>
        )}
      </form>
    </div>
  )
}

function MessageBubble({
  message,
  streaming,
}: {
  message: Message
  streaming?: boolean
}) {
  const isUser = message.role === 'user'
  return (
    <div className={isUser ? 'flex justify-end' : 'flex justify-start'}>
      <div
        className={
          isUser
            ? 'max-w-[80%] rounded-2xl bg-blue-500 px-3 py-2 text-sm text-white'
            : 'max-w-[80%] rounded-2xl bg-neutral-100 px-3 py-2 text-sm text-neutral-800'
        }
      >
        <p className="whitespace-pre-wrap break-words">
          {message.content}
          {streaming ? <span className="ml-0.5 animate-pulse">▍</span> : null}
        </p>
      </div>
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
