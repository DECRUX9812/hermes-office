// Builds desktop/src/plugin.jsx into the committed desktop/plugin.js (plain ESM).
// There is no local node_modules: esbuild is borrowed read-only from a Hermes
// checkout. Override with HERMES_ESBUILD=/path/to/esbuild/lib/main.js.
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const esbuildPath = process.env.HERMES_ESBUILD ?? '/home/user/Code/hermes-next/node_modules/esbuild/lib/main.js'
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
