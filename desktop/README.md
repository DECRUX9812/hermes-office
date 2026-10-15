# Desk — the hermes-office desktop page

`desktop/plugin.js` is the desktop half of the hermes-office Hermes plugin: a
`/desk` page (plus a sidebar nav row) in the native Hermes Desktop app. It is
plain ESM, imports only `@hermes/plugin-sdk`, `react` and `react/jsx-runtime`,
and talks to the plugin's backend (`dashboard/plugin_api.py`, mounted at
`/api/plugins/hermes-office`) through `ctx.rest`.

## Layout

- `src/` — the page source (`plugin.jsx` is the entry; JSX, automatic runtime).
- `build.mjs` — bundles `src/` into the committed `plugin.js`.
- `check-imports.mjs` — fails if `plugin.js` imports anything but the three
  allowed specifiers, or if `src/`/`plugin.js` contain a hard-coded colour
  (hex, `rgb()`, `hsl()`); the page styles itself with the app's `--ui-*`
  theme variables only.
- `plugin.js` — the built artifact, committed. The desktop app loads exactly
  this file.
- `dev/` — a browser harness that renders the real `plugin.js` against the
  real backend (screenshots / manual QA). Never shipped.

## Build and check

There is no local `node_modules`; esbuild is borrowed read-only from a Hermes
checkout (`HERMES_ESBUILD` overrides the path):

```bash
flock /tmp/hermes-build.lock node desktop/build.mjs
node desktop/check-imports.mjs
```

## Install

Copy (or symlink) the built file into the app's disk-plugin folder; the folder
name must equal the plugin id:

```bash
mkdir -p "$HERMES_HOME/desktop-plugins/hermes-office"
cp desktop/plugin.js "$HERMES_HOME/desktop-plugins/hermes-office/plugin.js"
```

The app hot-reloads on save (⌘K → “Reload desktop plugins” if it does not).
When hermes-office is installed as a unified package
(`$HERMES_HOME/plugins/hermes-office/`), the app copies `desktop/plugin.js`
into `desktop-plugins/` itself.

The page degrades honestly when the Python backend is not in
`plugins.enabled`: `ctx.rest` calls fail and the page shows the error inline.

## Dev harness (against the real backend)

```bash
flock /tmp/hermes-build.lock node desktop/dev/build-harness.mjs   # → dev/dist/harness.js (gitignored)
export HERMES_HOME=$(mktemp -d)                                   # never ~/.hermes
~/Code/ship-test/office-venv/.venv/bin/python desktop/dev/serve.py --repo <path to plugin repo> --port 9341
```

Then open `http://127.0.0.1:9341/` (`?dark` for dark mode). `serve.py` mounts
`<repo>/dashboard/plugin_api.py` at `/api/plugins/hermes-office` and serves
`dev/` statically. The harness shim (`dev/sdk-shim.js`) implements only the SDK
exports the plugin imports; `host.request` rejects with "no Hermes gateway in
the dev harness", so "New request" creates the workstream and honestly reports
that the chat session could not be started — everything else is live.
