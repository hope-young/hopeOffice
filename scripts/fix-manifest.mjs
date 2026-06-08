// One-off fix: change ExtensionPoint xsi:type from "CustomTab" to "PrimaryCommandSurface".
// Per Microsoft docs, the parent ExtensionPoint type is PrimaryCommandSurface;
// CustomTab is the child element name, not an xsi:type value.
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const path = resolve(__dirname, '..', 'manifest.xml')

let s = readFileSync(path, 'utf8')
const before = (s.match(/xsi:type="CustomTab"/g) || []).length
s = s.replace(/<ExtensionPoint xsi:type="CustomTab">/g, '<ExtensionPoint xsi:type="PrimaryCommandSurface">')
const after = (s.match(/xsi:type="CustomTab"/g) || []).length
writeFileSync(path, s)
console.log(`replaced ${before - after} occurrence(s) of ExtensionPoint xsi:type="CustomTab" → "PrimaryCommandSurface"`)
