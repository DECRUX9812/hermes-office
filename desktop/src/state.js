import { atom } from '@hermes/plugin-sdk'

// One route (`/desk`); which workstream is open lives here, not in the URL.
export const $selectedWorkstream = atom(null)
