/**
 * System prompt builder. Source of truth: SPEC_DETAILS §6.
 *
 * Pulls the current host's skills from the registry and lists
 * them so the LLM knows what's available. Each tool's JSON
 * Schema (derived from the Zod args via `z.toJSONSchema`) is
 * inlined so the LLM can also see the exact arg shape — the
 * tool definitions the AI SDK sends are authoritative, but
 * surfacing the same info in the system prompt makes the model
 * more reliable in our testing.
 */
import { z } from 'zod'
import type { HostKind } from '../types'
import { listSkillsForHost } from '../skills'

export function buildSystemPrompt(host: HostKind): string {
  const skills = listSkillsForHost(host)
  const skillList =
    skills.length === 0
      ? '_No skills are available for this host yet._'
      : skills
          .map((s) => {
            const shape = z.toJSONSchema(s.args, { target: 'draft-07' })
            return [
              `- **${s.name}** — ${s.description}`,
              `  args: \`${JSON.stringify(shape)}\``,
            ].join('\n')
          })
          .join('\n\n')

  return `You are hope-Office, an AI assistant for ${host} that controls the document by calling skills.

## Available skills

${skillList}

## Calling convention

Call the skill tools directly. Each tool's args are validated by Zod; a malformed call returns a clear error and you can retry with corrected args. After a tool runs you'll get a \`tool-result\` event — incorporate the result (e.g. an inserted object's id) into the next assistant turn.

When the user asks for something the available skills can do, prefer calling the skill over describing how the user could do it manually. When a task needs multiple skills, chain them: the first skill's result feeds into the next.
`
}
