import { createRoot } from 'react-dom/client'
import { StrictMode } from 'react'
import { App } from './App'
import { setOrchestratorHost } from './store/chat'
import type { HostKind } from '../core/types'
import './styles.css'

// Office.js globals — the runtime can land in three different
// places depending on how the task pane is hosted. We probe all
// three before giving up, in order:

//   1. `Office` is a global on the task pane's own window
//      (desktop Office injects it that way).
//   2. `window.parent.Office` is where Excel online parks it —
//      the task pane runs in a sandboxed iframe, but the host
//      application keeps the Office.js globals on its parent
//      frame.
//   3. `window.top.Office` is a further-up fallback for
//      edge-case hosts.
type OfficeLike = {
  onReady: (
    cb: (info: { host: Office.HostType | string; platform?: unknown }) => void,
  ) => unknown
  ribbon?: {
    requestUpdate?: (
      payload: { tabs: Array<{ id: string; visible: boolean }> },
      cb: (
        result:
          | { status: 'succeeded' }
          | { status: 'failed'; error: { message: string } },
      ) => void,
    ) => void
  }
}

function findOffice(): OfficeLike | undefined {
  if (typeof Office !== 'undefined') {
    return Office as unknown as OfficeLike
  }
  if (typeof window === 'undefined') return undefined
  const w = window as unknown as {
    parent?: { Office?: OfficeLike }
    top?: { Office?: OfficeLike }
  }
  return w.parent?.Office ?? w.top?.Office
}

const officeApi = findOffice()

// Custom Tabs are hidden by default in sideloaded add-ins. The
// manifest declares our tab as 'hopeOfficeTab'; we explicitly flip
// it on once the task pane is mounted. (See SPEC §11 — this is
// the *only* part of ribbon registration Phase 1 needs.)
function requestRibbon(office: OfficeLike | undefined): void {
  if (!office?.ribbon?.requestUpdate) return
  // Cast: the TS overload ships without the callback variant
  // that the JS API actually supports. The callback only logs
  // failures; not having it doesn't change behavior.
  ;(
    office.ribbon.requestUpdate as (
      tabs: { tabs: Array<{ id: string; visible: boolean }> },
      cb: (result: { status: 'succeeded' } | { status: 'failed'; error: { message: string } }) => void,
    ) => void
  )(
    { tabs: [{ id: 'hopeOfficeTab', visible: true }] },
    (result) => {
      if (result.status === 'failed') {
        console.warn(
          '[hope-office] Office.ribbon.requestUpdate failed:',
          result.error.message,
        )
      }
    },
  )
}

const container = document.getElementById('root')
if (!container) throw new Error('#root element not found in taskpane index.html')

const root = createRoot(container)

if (officeApi) {
  officeApi.onReady((info) => {
    const host = mapHostType(info.host)
    setOrchestratorHost(host)

    root.render(
      <StrictMode>
        <App host={host} />
      </StrictMode>,
    )

    requestRibbon(officeApi)
  })
} else {
  // Standalone preview outside Office (e.g. plain browser for
  // design work). Office isn't defined on this window, on
  // window.parent, or on window.top — the App component handles
  // the unknown-host case.
  setOrchestratorHost('unsupported')
  root.render(
    <StrictMode>
      <App host={null} />
    </StrictMode>,
  )
}

// @types/office-js declares HostType as a *numeric* enum
// (Word = 0, Excel = 1, PowerPoint = 2, ...) but Excel online
// and PowerPoint online return the *string* form ('Word' /
// 'Excel' / 'PowerPoint'). Compare against the lower-cased name
// so both representations resolve to the same HostKind.
function mapHostType(h: Office.HostType | string | null | undefined): HostKind {
  if (h == null) return 'unsupported'
  const name = String(h).toLowerCase()
  if (name.includes('word')) return 'word'
  if (name.includes('excel')) return 'excel'
  if (name.includes('power')) return 'powerpoint'
  return 'unsupported'
}
