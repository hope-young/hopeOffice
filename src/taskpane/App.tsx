/**
 * App — task pane root.
 *
 * Layout follows the M365 Copilot "AI Assistant" pattern (header bar
 * with brand + host chip, single main surface, no extra nav rail).
 * Settings is reached via the gear icon in the header rather than a
 * top nav row, so the chat gets the full vertical space.
 */
import { useState } from 'react'
import { ChatPanel } from './components/ChatPanel'
import { SettingsPanel } from './components/SettingsPanel'

type AppProps = {
  host: Office.HostType | null
}

type View = 'chat' | 'settings'

export function App({ host }: AppProps) {
  const [view, setView] = useState<View>('chat')
  const hostLabel = hostToLabel(host)

  return (
    <div className="flex h-full flex-col bg-white">
      <header className="flex items-center justify-between border-b border-neutral-200 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <BrandMark />
          <div className="flex flex-col leading-tight">
            <span className="text-sm font-semibold text-neutral-900">
              AI Assistant
            </span>
            <span className="text-[10px] text-neutral-500">
              {hostLabel}
            </span>
          </div>
        </div>
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setView(view === 'settings' ? 'chat' : 'settings')}
            aria-label={view === 'settings' ? 'Back to chat' : 'Open settings'}
            title={view === 'settings' ? 'Back to chat' : 'Settings'}
            className={`rounded-md p-1.5 transition-colors ${
              view === 'settings'
                ? 'bg-blue-50 text-blue-600'
                : 'text-neutral-500 hover:bg-neutral-100 hover:text-neutral-700'
            }`}
          >
            {view === 'settings' ? <IconBack /> : <IconSettings />}
          </button>
        </div>
      </header>
      <main className="flex-1 overflow-hidden">
        {view === 'chat' ? <ChatPanel /> : <SettingsPanel />}
      </main>
    </div>
  )
}

// ---------- Brand ----------

function BrandMark() {
  // Squared mark, 24x24, blue-500. Mirrors the icon we ship in
  // public/assets so the header looks right at any size.
  return (
    <div className="flex h-7 w-7 items-center justify-center rounded-md bg-gradient-to-br from-blue-500 to-blue-600 text-white">
      <svg
        width="16"
        height="16"
        viewBox="0 0 16 16"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        aria-hidden="true"
      >
        <path
          d="M3 11.5C3 11.5 4.5 12 6 12C7.5 12 8 11 8 11C8 11 8.5 9 10 8.5C11.5 8 13 8 13 8M5 5.5C5 4.67 5.67 4 6.5 4C7.33 4 8 4.67 8 5.5M3 8.5L5 6.5L7 8.5"
          stroke="currentColor"
          strokeWidth="1.5"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    </div>
  )
}

// ---------- Icons (inline SVG, no icon font) ----------

function IconSettings() {
  return (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
    </svg>
  )
}

function IconBack() {
  return (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M19 12H5" />
      <path d="M12 19l-7-7 7-7" />
    </svg>
  )
}

// ---------- Host label ----------

function hostToLabel(host: Office.HostType | null): string {
  if (host === null) return 'browser (no Office)'
  switch (host) {
    case Office.HostType.Word:
      return 'Word'
    case Office.HostType.Excel:
      return 'Excel'
    case Office.HostType.PowerPoint:
      return 'PowerPoint'
    default:
      return `host(${host})`
  }
}
