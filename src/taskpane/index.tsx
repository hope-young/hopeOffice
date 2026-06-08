import { createRoot } from 'react-dom/client'
import { StrictMode } from 'react'
import { App } from './App'
import './styles.css'

// Office.js is loaded by the manifest's <Runtime>; we wait for onReady before
// doing anything that touches Office.context. Until then, the React tree can
// still render the placeholder so the user sees something on first paint.

const container = document.getElementById('root')
if (!container) throw new Error('#root element not found in taskpane index.html')

const root = createRoot(container)

if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    root.render(
      <StrictMode>
        <App host={info.host} />
      </StrictMode>,
    )

    // Custom Tabs are hidden by default in sideloaded add-ins. The manifest
    // declares our tab as 'hopeOfficeTab'; we explicitly flip it on once the
    // task pane is mounted. (See SPEC §11 — this is the *only* part of
    // ribbon registration Phase 1 needs.)
    if (Office.ribbon) {
      // Cast: the TS overload ships without the callback variant that the
      // JS API actually supports. The callback only logs failures; not
      // having it doesn't change behavior.
      ;(Office.ribbon.requestUpdate as (
        tabs: { tabs: Array<{ id: string; visible: boolean }> },
        cb: (result: { status: 'succeeded' } | { status: 'failed'; error: { message: string } }) => void,
      ) => void)(
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
  })
} else {
  // Standalone preview outside Office (e.g. plain browser for design work).
  // Office isn't defined in this case — the App component handles unknown host.
  root.render(
    <StrictMode>
      <App host={null} />
    </StrictMode>,
  )
}
