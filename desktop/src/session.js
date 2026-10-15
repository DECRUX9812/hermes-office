import { host } from '@hermes/plugin-sdk'

import { api, errorText } from './api'

export function requestPrompt(folder, workstreamId, request) {
  return `${request}\n\nWork only inside ${folder.path}. Use the office tools (toolset "office"). Workstream id: ${workstreamId}.`
}

// Records the workstream first so the request is never lost, then starts a real
// Hermes session for it. Any gateway failure leaves the workstream `queued` and
// is reported to the caller; it is never marked running without a live turn.
export async function startWorkstream(folder, request) {
  const workstream = await api.createWorkstream({ folder_id: folder.id, title: request.slice(0, 60), request })

  if (typeof host.request !== 'function') {
    return { workstream, started: false, error: 'This Hermes build has no gateway request door (host.request).' }
  }

  let created

  try {
    created = await host.request('session.create', { cwd: folder.path, cwd_explicit: true })
  } catch (error) {
    return { workstream, started: false, error: `Could not start a Hermes session: ${errorText(error)}` }
  }

  const storedId = created?.stored_session_id || created?.session_id

  try {
    await host.request('prompt.submit', {
      session_id: created.session_id,
      text: requestPrompt(folder, workstream.id, request)
    })
  } catch (error) {
    if (storedId) {
      await api.patchWorkstream(workstream.id, { session_id: storedId }).catch(() => undefined)
    }

    return { workstream, started: false, error: `The session started but the request was not sent: ${errorText(error)}` }
  }

  try {
    const updated = await api.patchWorkstream(workstream.id, { session_id: storedId, status: 'running' })

    return { workstream: updated, started: true }
  } catch (error) {
    return { workstream, started: true, error: `Started, but the Desk could not record the session: ${errorText(error)}` }
  }
}
