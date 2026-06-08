/**
 * App — task pane root.
 *
 * Phase 4 replaces the Phase-1 placeholder with two real surfaces
 * (Chat / Settings) toggled by an in-app nav. The Ribbon Custom Tab
 * buttons that also flip this are Phase 11; for now the nav is
 * self-contained so the W24 acceptance can verify chat end-to-end.
 */
import { useState } from 'react'
import { ChatPanel } from './components/ChatPanel'
import { SettingsPanel } from './components/SettingsPanel'

type AppProps = {
  /**
   * The Office host, or `null` when running outside Office
   * (plain browser preview).
   */
  host: Office.HostType | null
}

type View = 'chat' | 'settings'

export function App({ host }: AppProps) {
  const [view, setView] = useState<View>('chat')
  const hostLabel = hostToLabel(host)

  return (
    <div className="flex h-full flex-col">
      <header className="flex items-center justify-between border-b border-neutral-200 px-3 py-2">
        <div>
          <h1 className="text-base font-semibold">hope-Office</h1>
          <p className="text-xs text-neutral-500">Running in {hostLabel}</p>
        </div>
        <nav className="flex gap-1">
          <NavButton active={view === 'chat'} onClick={() => setView('chat')}>
            Chat
          </NavButton>
          <NavButton
            active={view === 'settings'}
            onClick={() => setView('settings')}
          >
            Settings
          </NavButton>
        </nav>
      </header>
      <main className="flex-1 overflow-hidden">{view === 'chat' ? <ChatPanel /> : <SettingsPanel />}</main>
    </div>
  )
}

function NavButton({
  active,
  onClick,
  children,
}: {
  active: boolean
  onClick: () => void
  children: React.ReactNode
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={
        active
          ? 'rounded bg-blue-500 px-2 py-1 text-xs text-white'
          : 'rounded px-2 py-1 text-xs hover:bg-neutral-100'
      }
    >
      {children}
    </button>
  )
}

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
