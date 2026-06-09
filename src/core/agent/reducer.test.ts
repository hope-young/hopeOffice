import { describe, expect, it } from 'vitest'
import { INITIAL_AGENT_STATE, reduce } from './reducer'
import type { Message, ToolCall } from '../types'

function tc(overrides: Partial<ToolCall> = {}): ToolCall {
  return {
    id: 'tc-1',
    name: 'execute_skill',
    args: { skill: 'add-table' },
    status: 'pending',
    ...overrides,
  }
}

describe('reduce — user-send', () => {
  it('appends a user message and switches to streaming', () => {
    const next = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'hello' })
    expect(next.messages).toEqual([{ role: 'user', content: 'hello' }])
    expect(next.status).toBe('streaming')
    expect(next.draft).toBe('')
    expect(next.error).toBeUndefined()
  })

  it('clears any prior toolCalls and error', () => {
    const dirty = { ...INITIAL_AGENT_STATE, toolCalls: [tc()], error: new Error('x') }
    const next = reduce(dirty, { type: 'user-send', text: 'again' })
    expect(next.toolCalls).toEqual([])
    expect(next.error).toBeUndefined()
  })
})

describe('reduce — streaming', () => {
  it('appends tokens to the draft while streaming', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'hi' })
    s = reduce(s, { type: 'stream-token', token: 'Hel' })
    s = reduce(s, { type: 'stream-token', token: 'lo' })
    expect(s.draft).toBe('Hello')
  })

  it('ignores stream-token when not streaming', () => {
    const s = reduce(INITIAL_AGENT_STATE, { type: 'stream-token', token: 'x' })
    expect(s).toBe(INITIAL_AGENT_STATE)
  })
})

describe('reduce — stream-end', () => {
  it('commits the draft as an assistant message', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'hi' })
    s = reduce(s, { type: 'stream-token', token: 'Hello' })
    s = reduce(s, { type: 'stream-end' })
    expect(s.messages).toEqual([
      { role: 'user', content: 'hi' },
      { role: 'assistant', content: 'Hello' },
    ])
    expect(s.draft).toBe('')
    expect(s.status).toBe('idle')
  })

  it('includes tool calls on the assistant message when present', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'go' })
    s = reduce(s, {
      type: 'tool-call-start',
      toolCall: tc({ id: 'tc-1', name: 'add-table' }),
    })
    s = reduce(s, { type: 'stream-token', token: 'doing it' })
    s = reduce(s, { type: 'stream-end' })
    const assistant = s.messages.at(-1)
    expect(assistant?.role).toBe('assistant')
    if (assistant?.role === 'assistant') {
      expect(assistant.toolCalls).toHaveLength(1)
      expect(assistant.toolCalls?.[0]?.id).toBe('tc-1')
    }
  })
})

describe('reduce — tool calls', () => {
  it('tool-call-start appends and sets status executing', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'go' })
    s = reduce(s, { type: 'tool-call-start', toolCall: tc({ id: 'tc-1' }) })
    expect(s.toolCalls).toHaveLength(1)
    expect(s.toolCalls[0]?.status).toBe('running')
    expect(s.status).toBe('executing')
  })

  it('tool-call-result marks success and returns to streaming', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'go' })
    s = reduce(s, { type: 'tool-call-start', toolCall: tc({ id: 'tc-1' }) })
    s = reduce(s, { type: 'tool-call-result', toolCallId: 'tc-1', result: { ok: true } })
    expect(s.toolCalls[0]?.status).toBe('success')
    expect(s.status).toBe('streaming')
  })

  it('tool-call-error marks the tool as errored', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'go' })
    s = reduce(s, { type: 'tool-call-start', toolCall: tc({ id: 'tc-1' }) })
    s = reduce(s, { type: 'tool-call-error', toolCallId: 'tc-1', error: 'boom' })
    expect(s.toolCalls[0]?.status).toBe('error')
    expect(s.toolCalls[0]?.error?.message).toBe('boom')
  })
})

describe('reduce — approval', () => {
  it('approval-requested sets awaiting-approval', () => {
    const s = reduce(INITIAL_AGENT_STATE, { type: 'approval-requested', code: 'xyz' })
    expect(s.status).toBe('awaiting-approval')
  })

  it('approval-decided moves awaiting → executing', () => {
    const pending = { ...INITIAL_AGENT_STATE, status: 'awaiting-approval' as const }
    const next = reduce(pending, { type: 'approval-decided', approved: true })
    expect(next.status).toBe('executing')
  })

  it('approval-decided is a no-op when not awaiting', () => {
    const next = reduce(INITIAL_AGENT_STATE, { type: 'approval-decided', approved: true })
    expect(next).toBe(INITIAL_AGENT_STATE)
  })
})

describe('reduce — reasoning', () => {
  it('reasoning-delta accumulates into draftReasoning while streaming', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'hi' })
    s = reduce(s, { type: 'reasoning-delta', delta: 'Plan: ' })
    s = reduce(s, { type: 'reasoning-delta', delta: 'compute x.' })
    expect(s.draftReasoning).toBe('Plan: compute x.')
    expect(s.draft).toBe('')
  })

  it('reasoning-delta is a no-op when not streaming', () => {
    const s = reduce(INITIAL_AGENT_STATE, {
      type: 'reasoning-delta',
      delta: 'x',
    })
    expect(s).toBe(INITIAL_AGENT_STATE)
  })

  it('stream-end folds draftReasoning into the assistant message', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'q' })
    s = reduce(s, { type: 'reasoning-delta', delta: 'thinking…' })
    s = reduce(s, { type: 'stream-token', token: 'answer' })
    s = reduce(s, { type: 'stream-end' })
    const assistant = s.messages.at(-1)
    expect(assistant?.role).toBe('assistant')
    if (assistant?.role === 'assistant') {
      expect(assistant.content).toBe('answer')
      expect(assistant.reasoning).toBe('thinking…')
    }
  })

  it('stream-end omits reasoning when draftReasoning is empty', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'q' })
    s = reduce(s, { type: 'stream-token', token: 'plain' })
    s = reduce(s, { type: 'stream-end' })
    const assistant = s.messages.at(-1)
    if (assistant?.role === 'assistant') {
      expect(assistant.reasoning).toBeUndefined()
    } else {
      throw new Error('expected assistant message')
    }
  })

  it('user-send clears the prior draftReasoning', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'a' })
    s = reduce(s, { type: 'reasoning-delta', delta: 'x' })
    s = reduce(s, { type: 'stream-end' })
    s = reduce(s, { type: 'user-send', text: 'b' })
    expect(s.draftReasoning).toBe('')
  })
})

describe('reduce — restore-messages', () => {
  it('wholesale replaces messages', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'a' })
    s = reduce(s, { type: 'stream-token', token: 'r' })
    s = reduce(s, { type: 'stream-end' })
    const replaced: Message[] = [
      { role: 'user', content: 'history' },
      { role: 'assistant', content: 'preserved' },
    ]
    const next = reduce(s, { type: 'restore-messages', messages: replaced })
    expect(next.messages).toEqual(replaced)
    expect(next.draft).toBe('')
    expect(next.draftReasoning).toBe('')
  })
})

describe('reduce — error / reset', () => {
  it('error sets status and keeps messages', () => {
    let s = reduce(INITIAL_AGENT_STATE, { type: 'user-send', text: 'hi' })
    s = reduce(s, { type: 'error', error: new Error('llm down') })
    expect(s.status).toBe('error')
    expect(s.error?.message).toBe('llm down')
    expect(s.messages).toHaveLength(1)
  })

  it('reset clears draft + toolCalls + error but keeps messages', () => {
    const dirty = {
      ...INITIAL_AGENT_STATE,
      messages: [{ role: 'user' as const, content: 'a' }],
      draft: 'leftover',
      toolCalls: [tc()],
      error: new Error('x'),
    }
    const next = reduce(dirty, { type: 'reset' })
    expect(next.messages).toEqual(dirty.messages)
    expect(next.draft).toBe('')
    expect(next.toolCalls).toEqual([])
    expect(next.error).toBeUndefined()
  })
})
