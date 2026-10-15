import { host, queryClient, ROUTES_AREA, SIDEBAR_NAV_AREA, useValue } from '@hermes/plugin-sdk'

import { bindContext, KEY } from './api'
import { DeskHome } from './home'
import { $selectedWorkstream } from './state'
import { injectStyle } from './styles'
import { WorkstreamView } from './workstream'

// broadcast_plugin_event("hermes-office", "workstream.updated", ...) arrives
// on the gateway stream under this wire name.
const WORKSTREAM_UPDATED = 'plugin.hermes-office.workstream.updated'

function DeskPage() {
  const selected = useValue($selectedWorkstream)

  return selected === null ? <DeskHome /> : <WorkstreamView id={selected} key={selected} />
}

export default {
  id: 'hermes-office',
  name: 'Desk',
  description: 'Office work on your folders: Word, Excel and PowerPoint edits with tracked changes, previews and receipts.',
  register(ctx) {
    ctx.onDispose(bindContext(ctx))
    ctx.onDispose(injectStyle())
    ctx.onDispose(() => $selectedWorkstream.set(null))
    ctx.onDispose(host.onEvent(WORKSTREAM_UPDATED, () => void queryClient.invalidateQueries({ queryKey: [KEY] })))

    ctx.registerMany([
      {
        id: 'page',
        area: ROUTES_AREA,
        data: { path: '/desk' },
        render: () => <DeskPage />
      },
      {
        id: 'nav',
        area: SIDEBAR_NAV_AREA,
        data: { path: '/desk', label: 'Desk', codicon: 'briefcase' }
      }
    ])
  }
}
