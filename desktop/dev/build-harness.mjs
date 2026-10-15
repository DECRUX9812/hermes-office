// Bundles the REAL built desktop/plugin.js plus the SDK shim and the real
// React/ReactDOM (read-only from the Hermes checkout's node_modules) into
// desktop/dev/dist/harness.js. The generated entry calls the plugin's
// register() with a minimal PluginContext and mounts the /desk page.
import { mkdir, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const hermesModules = '/home/user/Code/hermes-next/node_modules'
const esbuild = await import(
  pathToFileURL(process.env.HERMES_ESBUILD ?? join(hermesModules, 'esbuild/lib/main.js')).href
)

const dist = join(here, 'dist')
await mkdir(dist, { recursive: true })

const entry = join(dist, 'entry.jsx')
await writeFile(
  entry,
  `import { createElement, StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import * as sdk from '../sdk-shim'
import plugin from '../../plugin'

const disposers = []
const ctx = {
  source: 'plugin:hermes-office',
  register: contribution => {
    if (contribution.area === sdk.ROUTES_AREA) {
      ctx.page = contribution
    }
    return () => undefined
  },
  registerMany(contributions) {
    const disposersFor = contributions.map(c => ctx.register(c))
    return () => disposersFor.forEach(dispose => dispose())
  },
  rest: (path, opts = {}) =>
    fetch('/api/plugins/hermes-office' + (path.startsWith('/') ? path : '/' + path), {
      method: opts.method ?? 'GET',
      headers: opts.body === undefined ? undefined : { 'content-type': 'application/json' },
      body: opts.body === undefined ? undefined : JSON.stringify(opts.body)
    }).then(async res => {
      if (!res.ok) {
        throw new Error(res.status + ': ' + (await res.text()))
      }
      return res.json()
    }),
  socket: () => () => undefined,
  onEvent: () => () => undefined,
  onDispose: fn => disposers.push(fn),
  setTimeout: (fn, ms) => { const id = setTimeout(fn, ms); return () => clearTimeout(id) },
  setInterval: (fn, ms) => { const id = setInterval(fn, ms); return () => clearInterval(id) },
  addEventListener: (target, type, listener, options) => {
    target.addEventListener(type, listener, options)
    return () => target.removeEventListener(type, listener, options)
  },
  os: {
    openExternal: () => Promise.resolve(false),
    revealPath: () => Promise.resolve(false),
    writeClipboard: text => navigator.clipboard.writeText(text).then(() => true, () => false)
  },
  storage: {
    get: (key, fallback) => {
      try {
        const raw = localStorage.getItem('hermes.plugin.hermes-office.' + key)
        return raw === null ? fallback : JSON.parse(raw)
      } catch {
        return fallback
      }
    },
    set: (key, value) => localStorage.setItem('hermes.plugin.hermes-office.' + key, JSON.stringify(value)),
    remove: key => localStorage.removeItem('hermes.plugin.hermes-office.' + key)
  }
}

plugin.register(ctx)

if (!ctx.page) {
  throw new Error('the plugin registered no ROUTES_AREA contribution')
}

createRoot(document.getElementById('root')).render(createElement(StrictMode, null, ctx.page.render()))
`
)

await esbuild.build({
  entryPoints: [entry],
  outfile: join(dist, 'harness.js'),
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  jsx: 'automatic',
  nodePaths: [hermesModules],
  plugins: [
    {
      name: 'desk-harness-sdk-shim',
      setup(build) {
        build.onResolve({ filter: /^@hermes\/plugin-sdk$/ }, () => ({ path: join(here, 'sdk-shim.js') }))
      }
    }
  ],
  logLevel: 'info'
})

console.log('dev harness → desktop/dev/dist/harness.js')
