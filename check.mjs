import { SKILL_REGISTRY, listSkillsForHost } from './src/core/skills/index.ts'
const names = listSkillsForHost('excel').map(s => s.name).sort()
console.log('count:', names.length)
console.log('names:', names)
