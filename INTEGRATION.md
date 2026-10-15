# Hermes Office: Desktop App Integration Guide

This document describes how to integrate the **Desk** plugin page from `hermes-office` into the **Hermes Desktop** app (`hermes-next`), and how to package it for the `hermes-desktop-plus` add-on bundle.

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
│  Mounted API Router: /api/plugins/hermes-office/*           │
│  Desk Engine: DeskState · Workstreams · Guard · Receipts    │
│  Paper Office: paper-docx · paper-xlsx · paper-pptx         │
└─────────────────────────────────────────────────────────────┘
```

The desktop frontend is compiled into a single standard ES module at `desktop/plugin.js`. It contains zero hard-coded colors, imports only `@hermes/plugin-sdk` and `react`, and delegates all operations to the verified `hermes-office` backend.

---

## 1. hermes-next Integrator Instructions

### A. Automatic Plugin Discovery (Recommended)
Hermes Desktop automatically scans `<HERMES_HOME>/desktop-plugins/` for desktop plugin packages.

1. Ensure `hermes-office` backend plugin is installed and enabled in `$HERMES_HOME/config.yaml`:
   ```yaml
   plugins:
     enabled:
       - hermes-office
   ```

2. Symlink or copy the compiled plugin folder into the desktop plugins directory:
   ```bash
   mkdir -p "$HERMES_HOME/desktop-plugins/hermes-office"
   cp ~/Code/hermes-office/desktop/plugin.js "$HERMES_HOME/desktop-plugins/hermes-office/"
   cp ~/Code/hermes-office/desktop/.hermes-package.json "$HERMES_HOME/desktop-plugins/hermes-office/" 2>/dev/null || true
   ```

3. When Hermes Desktop launches, the `runtime-loader` discovers `hermes-office`:
   - It invokes `plugin.register(ctx)`.
   - `Desk` is mounted at route `/desk` (`ROUTES_AREA`).
   - A briefcase icon (`codicon: 'briefcase'`) appears in the persistent sidebar rail (`SIDEBAR_NAV_AREA`).
   - Clicking the briefcase takes the user directly to the Desk page.

### B. Static In-Tree Vendor Drop-in (Alternative)
If you wish to vendor the Desk plugin directly into the `hermes-next` repository:
1. Copy `desktop/src/` into `apps/desktop/src/plugins/hermes-office/`.
2. In `apps/desktop/src/plugins/registry.ts`, register the entry point:
   ```typescript
   import deskPlugin from './hermes-office/plugin';
   // Register alongside built-in plugins
   registerPlugin(deskPlugin);
   ```
3. Re-run desktop build:
   ```bash
   pnpm --filter @hermes/desktop build
   ```

---

## 2. hermes-desktop-plus Add-On Integration

For distributions bundling `hermes-desktop-plus`:

1. **Include Backend Dependency:**
   In your distribution manifest or setup script, ensure `hermes-office` is installed in the python environment:
   ```bash
   pip install ~/Code/hermes-office
   ```
   Or install editable for active development:
   ```bash
   pip install -e ~/Code/hermes-office
   ```

2. **Bundle Desktop Assets:**
   During the packaging step of `hermes-desktop-plus`:
   ```bash
   ADDON_DEST="$HERMES_HOME/desktop-plugins/hermes-office"
   mkdir -p "$ADDON_DEST"
   cp ~/Code/hermes-office/desktop/plugin.js "$ADDON_DEST/"
   ```

3. **Verify Imports:**
   Run the verification script to ensure no illegal imports or external dependencies leaked in:
   ```bash
   node ~/Code/hermes-office/desktop/check-imports.mjs
   ```

---

## 3. Design Standard & UX Compliance

The Desk plugin strictly adheres to `~/Code/hermes-next/DESIGN.md`:

1. **Three-Zone Hierarchy:**
   - **Zone 1 (Left Rail / Drawer):** Folders registered on your Desk, workstream list per folder, with a dedicated **"Needs you"** filter pinned to the top.
   - **Zone 2 (Centre Stage):** Clean workstream dialogue, bounded at 72ch for readability, featuring an **Activity Island** that conveys real-time agent status ("Analyzing workbook formulas...", "Diffing presentation decks...").
   - **Zone 3 (Right Drawer / Inspector):** Side-by-side preview with rich tracked changes, cell-level spreadsheet diffs, PowerPoint shape change inspectors, and "Open in Word / Excel / PowerPoint" OS launchers.

2. **Semantic Design Tokens:**
   - Zero hardcoded hex colors.
   - All colors derive from `--bg-canvas`, `--bg-surface`, `--text-primary`, `--text-muted`, `--border-subtle`, `--accent`, and `--status-*`.
   - Adapts automatically to Light, Dark, and High-Contrast themes.

3. **Zero Mocks:**
   - No mock data or simulated state. All data loads from the live API endpoints (`/api/plugins/hermes-office/*`), reflecting real filesystem changes, backups, and edit receipts.

---

## 4. Development Harness & Verification

To verify the Desk plugin independently without running full Electron:

1. **Compile Bundle & Build Harness:**
   ```bash
   # Recompile desktop/plugin.js from source
   node desktop/build.mjs

   # Build standalone dev harness
   node desktop/dev/build-harness.mjs
   ```

2. **Start Dev Server:**
   ```bash
   HERMES_HOME=/tmp/test-hermes-home ~/Code/ship-test/office-venv/.venv/bin/python \
       desktop/dev/serve.py --repo ~/Code/hermes-office --port 9341
   ```

3. **Inspect in Browser:**
   Open `http://127.0.0.1:9341` to interact with the live page against real fixture files.
