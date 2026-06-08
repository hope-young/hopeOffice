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
