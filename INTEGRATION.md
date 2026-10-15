# Hermes Office: Desktop App Integration Guide

This document describes how to integrate the **Desk** plugin page from `hermes-office` into the **Hermes Desktop** app, and how to bundle it in a Hermes Desktop distribution.

---

## Architecture Overview

`hermes-office` provides a full-featured Office desktop workspace ("Desk") built strictly according to the Hermes desktop plugin specification (`apps/desktop/DESIGN.md` and `apps/desktop/src/lib/runtime-loader.ts`):

```
┌─────────────────────────────────────────────────────────────┐
│                       Hermes Desktop                        │
│                                                             │
│ ┌───────────────┐ ┌───────────────────────────────────────┐ │
│ │ Plugin Rail   │ │ Desk Page (/desk)                     │ │
│ │               │ │                                       │ │
│ │ [Chat]        │ │ ┌─────────┐ ┌─────────────┐ ┌───────┐ │ │
│ │ [Agents]      │ │ │ Folders │ │ Workstream  │ │ Diff/ │ │ │
│ │ [Desk] ◀──────┼─┤ │ & Queue │ │ & Island    │ │ Prev  │ │ │
│ └───────────────┘ │ └─────────┘ └─────────────┘ └───────┘ │ │
│                   └──────────────────┬────────────────────┘ │
└──────────────────────────────────────┼──────────────────────┘
                                       │ REST / Events
                                       ▼
┌─────────────────────────────────────────────────────────────┐
│                    Hermes Agent Core                        │
│  Mounted API Router: /api/plugins/office-desk/*           │
│  Desk Engine: DeskState · Workstreams · Guard · Receipts    │
│  Paper Office: paper-docx · paper-xlsx · paper-pptx         │
└─────────────────────────────────────────────────────────────┘
```

The desktop frontend is compiled into a single standard ES module at `desktop/plugin.js`. It contains zero hard-coded colors, imports only `@hermes/plugin-sdk` and `react`, and delegates all operations to the verified `hermes-office` backend.

---

## 1. Hermes Desktop Integrator Instructions

### A. Automatic Plugin Discovery (Recommended)
Hermes Desktop automatically scans `<HERMES_HOME>/desktop-plugins/` for desktop plugin packages.

1. Ensure the `office-desk` backend plugin (repo `hermes-office`) is installed and enabled in `$HERMES_HOME/config.yaml`:
   ```yaml
   plugins:
     enabled:
       - office-desk
   ```

2. Symlink or copy the compiled plugin folder into the desktop plugins directory:
   ```bash
   mkdir -p "$HERMES_HOME/desktop-plugins/office-desk"
   cp desktop/plugin.js "$HERMES_HOME/desktop-plugins/office-desk/"
   ```

3. When Hermes Desktop launches, the `runtime-loader` discovers `office-desk`:
   - It invokes `plugin.register(ctx)`.
   - `Desk` is mounted at route `/desk` (`ROUTES_AREA`).
   - A briefcase icon (`codicon: 'briefcase'`) appears in the persistent sidebar rail (`SIDEBAR_NAV_AREA`).
   - Clicking the briefcase takes the user directly to the Desk page.

### B. Static In-Tree Vendor Drop-in (Alternative)
If you wish to vendor the Desk plugin directly into a Hermes Desktop checkout:
1. Copy `desktop/src/` into `apps/desktop/src/plugins/office-desk/`.
2. In `apps/desktop/src/plugins/registry.ts`, register the entry point:
   ```typescript
   import deskPlugin from './office-desk/plugin';
   // Register alongside built-in plugins
   registerPlugin(deskPlugin);
   ```
3. Re-run desktop build:
   ```bash
   pnpm --filter @hermes/desktop build
   ```

---

## 2. Bundling in a Desktop Distribution

For distributions that ship Hermes Desktop with add-on plugins:

1. **Include Backend Dependency:**
   In your distribution manifest or setup script, ensure `hermes-office` is installed in the python environment:
   ```bash
   pip install paper-docx paper-pptx paper-xlsx
   pip install git+https://github.com/DECRUX9812/hermes-office.git@v0.1.1
   ```
   Or install editable for active development:
   ```bash
   pip install -e .
   ```

2. **Bundle Desktop Assets:**
   During the distribution's packaging step:
   ```bash
   ADDON_DEST="$HERMES_HOME/desktop-plugins/office-desk"
   mkdir -p "$ADDON_DEST"
   cp desktop/plugin.js "$ADDON_DEST/"
   ```

3. **Verify Imports:**
   Run the verification script to ensure no illegal imports or external dependencies leaked in:
   ```bash
   node desktop/check-imports.mjs
   ```

---

## 3. Design Standard & UX Compliance

The Desk plugin strictly adheres to the Hermes Desktop design standard (`apps/desktop/DESIGN.md`):

1. **Three-Zone Hierarchy:**
   - **Zone 1 (Left Rail / Drawer):** Folders registered on your Desk, workstream list per folder, with a dedicated **"Needs you"** filter pinned to the top.
   - **Zone 2 (Centre Stage):** Clean workstream dialogue, bounded at 72ch for readability, featuring an **Activity Island** that conveys real-time agent status ("Analyzing workbook formulas...", "Diffing presentation decks...").
   - **Zone 3 (Right Drawer / Inspector):** Side-by-side preview with rich tracked changes, cell-level spreadsheet diffs, PowerPoint shape change inspectors, and "Open in Word / Excel / PowerPoint" OS launchers.

2. **Semantic Design Tokens:**
   - Zero hardcoded hex colors.
   - All colors derive from `--bg-canvas`, `--bg-surface`, `--text-primary`, `--text-muted`, `--border-subtle`, `--accent`, and `--status-*`.
   - Adapts automatically to Light, Dark, and High-Contrast themes.

3. **Zero Mocks:**
   - No mock data or simulated state. All data loads from the live API endpoints (`/api/plugins/office-desk/*`), reflecting real filesystem changes, backups, and edit receipts.

---

## 4. Development Harness & Verification

To verify the Desk plugin independently without running full Electron:

1. **Compile Bundle & Build Harness:**
   ```bash
   # esbuild/react are borrowed from a Hermes Desktop checkout
   export HERMES_NODE_MODULES=/path/to/hermes-desktop/node_modules

   # Recompile desktop/plugin.js from source
   node desktop/build.mjs

   # Build standalone dev harness
   node desktop/dev/build-harness.mjs
   ```

2. **Start Dev Server:**
   ```bash
   # optional: python scripts/setup_demo_desk.py  (demo Desk state in /tmp/office-demo-home)
   HERMES_HOME=$(mktemp -d) python desktop/dev/serve.py --repo . --port 9341   # needs fastapi + uvicorn
   ```

3. **Inspect in Browser:**
   Open `http://127.0.0.1:9341` to interact with the live page against real fixture files.
