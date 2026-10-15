// Builds desktop/src/plugin.jsx into the committed desktop/plugin.js (plain ESM).
// There is no local node_modules: esbuild is borrowed read-only from a Hermes
// Desktop checkout. Set HERMES_NODE_MODULES=/path/to/hermes-desktop/node_modules,
// or HERMES_ESBUILD=/path/to/esbuild/lib/main.js.
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const modules = process.env.HERMES_NODE_MODULES
const esbuildPath = process.env.HERMES_ESBUILD ?? (modules ? join(modules, 'esbuild/lib/main.js') : null)
if (!esbuildPath) {
  console.error('Set HERMES_NODE_MODULES (a Hermes Desktop node_modules dir) or HERMES_ESBUILD.')
  process.exit(2)
}
const esbuild = await import(pathToFileURL(esbuildPath).href)

const ALLOWED = ['@hermes/plugin-sdk', 'react', 'react/jsx-runtime']

await esbuild.build({
  entryPoints: [join(here, 'src', 'plugin.jsx')],
  outfile: join(here, 'plugin.js'),
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  jsx: 'automatic',
  external: ALLOWED,
  legalComments: 'none',
  charset: 'utf8',
  banner: {
    js: '// hermes-office Desk desktop plugin. Built from desktop/src by desktop/build.mjs; do not edit.'
  },
  logLevel: 'info'
})
