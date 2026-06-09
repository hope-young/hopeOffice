import { execSync } from 'node:child_process'
const out = execSync('cd E:/OH-workspace/hopeOffice && npx vitest run 2>&1', { encoding: 'utf8' })
const lines = out.split('\n')
let printing = false
for (const line of lines) {
  if (line.includes('FAIL')) printing = true
  if (printing) console.log(line)
  if (printing && line.trim() === '') printing = false
}
