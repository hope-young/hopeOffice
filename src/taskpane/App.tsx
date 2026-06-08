type AppProps = {
  /**
   * The Office host, or `null` when running outside Office
   * (plain browser preview).
   */
  host: Office.HostType | null
}

/**
 * Phase 1 skeleton. Wires Office.onReady to React and renders a placeholder.
 * Real chat / settings / history panels arrive in later phases per SPEC §13.
 */
export function App({ host }: AppProps) {
  const hostLabel = hostToLabel(host)

  return (
    <div className="h-full flex flex-col items-center justify-center p-6 gap-2 text-center">
      <h1 className="text-2xl font-semibold">hope-Office</h1>
      <p className="text-sm text-neutral-500">
        Running in <span className="font-mono">{hostLabel}</span>
      </p>
      <p className="text-xs text-neutral-400 mt-4">
        Skeleton · Phase 1 of the SPEC
      </p>
    </div>
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
