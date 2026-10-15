// Guards the two loader contracts the desktop app enforces at load time and the
// theming rule from DESIGN.md:
//   1. desktop/plugin.js imports only @hermes/plugin-sdk, react, react/jsx-runtime;
//   2. no hard-coded colours (hex, rgb(), hsl()) in desktop/src or desktop/plugin.js.
// desktop/dev/ is a dev-only harness and is deliberately not scanned.
import { readdir, readFile } from 'node:fs/promises'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const ALLOWED = new Set(['@hermes/plugin-sdk', 'react', 'react/jsx-runtime'])

const SPECIFIER_PATTERNS = [
  /\bimport\s+(?:[\w$*{}\s,]+\s+from\s+)?["']([^"']+)["']/g,
  /\bexport\s+(?:\*|{[^}]*})\s*(?:as\s+[\w$]+\s+)?from\s+["']([^"']+)["']/g,
  /\bimport\s*\(\s*["'`]([^"'`]+)["'`]\s*\)/g
]

const COLOUR_PATTERNS = [
  { name: 'hex colour', re: /(?<![\w&$])#(?:[0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{3,4})(?![\w-])/g },
  { name: 'rgb() literal', re: /\brgba?\(/g },
  { name: 'hsl() literal', re: /\bhsla?\(/g }
]

async function sourceFiles(dir) {
  const out = []

  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name)

    if (entry.isDirectory()) {
      out.push(...(await sourceFiles(full)))
    } else if (/\.(jsx?|mjs|css)$/.test(entry.name)) {
      out.push(full)
    }
  }

  return out
}

function lineOf(text, index) {
  return text.slice(0, index).split('\n').length
}

const problems = []
const pluginPath = join(here, 'plugin.js')
const plugin = await readFile(pluginPath, 'utf8')

for (const pattern of SPECIFIER_PATTERNS) {
  for (const match of plugin.matchAll(pattern)) {
    if (!ALLOWED.has(match[1])) {
      problems.push(`plugin.js:${lineOf(plugin, match.index)} imports "${match[1]}" (allowed: ${[...ALLOWED].join(', ')})`)
    }
  }
}

if (!/\bexport\s*{[^}]*\bas\s+default\b|\bexport\s+default\b/.test(plugin)) {
  problems.push('plugin.js has no default export (the HermesPlugin)')
}

for (const file of [pluginPath, ...(await sourceFiles(join(here, 'src')))]) {
  const text = await readFile(file, 'utf8')

  for (const { name, re } of COLOUR_PATTERNS) {
    for (const match of text.matchAll(re)) {
      problems.push(`${relative(here, file)}:${lineOf(text, match.index)} ${name} "${match[0]}"; use a var(--ui-*) token`)
    }
  }
}

if (problems.length) {
  console.error(`check-imports: ${problems.length} problem(s)\n  ${problems.join('\n  ')}`)
  process.exit(1)
}

console.log('check-imports: ok (imports limited to the SDK and React; no hard-coded colours)')
