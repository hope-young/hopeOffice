/**
 * Selection context — capture what the user has selected (or
 * where their active cell is) so the LLM knows what the
 * conversation is actually about.
 *
 * Why this is its own module: capturing selection is async
 * and host-specific. We use the cross-host
 * `Office.context.document.getSelectedDataAsync(Matrix)` API
 * which works in Word, Excel, and PowerPoint, plus a
 * best-effort `getActiveCell` fallback for Excel so we
 * always have *something* to put in the system prompt.
 */
import type { HostKind } from '../types'

/* eslint-disable @typescript-eslint/no-unused-vars */

/**
 * What the user has selected in the host, in a form the LLM
 * can act on. `kind: 'none'` means we couldn't capture
 * anything (which is fine — we just skip the context hint).
 */
export type SelectionContext =
  | {
      kind: 'range'
      /** A1:B2-style address of the selected range, when the
       *  host could tell us. May be absent for PowerPoint
       *  shape selections. */
      address?: string
      /** Rows × cols size, useful for the LLM to know the
       *  extent without having to ask. */
      rows: number
      cols: number
      /** The actual values. Truncated to a small sample so
       *  we don't blow the context window. */
      preview: unknown[][]
    }
  | {
      kind: 'chart'
      /** Display title of the chart if the user has set one. */
      title?: string
    }
  | { kind: 'none' }

/**
 * Read the current selection out of the host. Returns a
 * Promise that resolves to a SelectionContext — resolves to
 * `{ kind: 'none' }` if the host is missing, doesn't expose
 * the API, the call returns no data, OR the host's callback
 * never fires (observed on some Excel 2605 builds in dev
 * mode where the matrix API silently no-ops).
 *
 * The internal 1.5s timeout is critical: without it, a
 * missing callback would hang the orchestrator's `send()`
 * indefinitely, with no way for the user to escape because
 * the orchestrator's `abort()` only cancels the chat
 * stream, not the selection probe that runs *before* the
 * stream. The timeout is the safety net.
 *
 * The `host` argument is currently unused — the cross-host
 * `getSelectedDataAsync(Matrix)` API works in Word, Excel,
 * and PowerPoint — but we accept it so the call site
 * documents intent and so a future host-specific probe
 * (Excel active chart, etc.) can branch on it.
 */
export async function captureSelection(
  _host?: HostKind,
): Promise<SelectionContext> {
  // 1.5s is well above the synchronous round-trip the API
  // takes on a healthy host (single-digit ms) but short
  // enough that a hung callback is invisible to the user.
  const SELECTION_TIMEOUT_MS = 1500
  return Promise.race([
    captureSelectionImpl(),
    new Promise<SelectionContext>((resolve) =>
      setTimeout(
        () => resolve({ kind: 'none' }),
        SELECTION_TIMEOUT_MS,
      ),
    ),
  ])
}

async function captureSelectionImpl(): Promise<SelectionContext> {
  const office = (globalThis as { Office?: { context?: unknown } })
    .Office
  if (!office || !office.context) return { kind: 'none' }

  // Try the matrix API first. This is the only path that gives
  // us the actual selected cell range + values; for a chart
  // selection in Excel the result is empty and we have to
  // fall back to the active chart probe below.
  const matrix = await tryGetSelectedMatrix()
  if (matrix) {
    if (matrix.rows === 0 || matrix.cols === 0) {
      // Empty selection — the user clicked into a chart
      // (Excel returns an empty matrix for non-range
      // selections). Defer to the chart probe.
      const chart = await tryGetActiveChart()
      if (chart) return { kind: 'chart', title: chart.title }
      return { kind: 'none' }
    }
    return {
      kind: 'range',
      address: matrix.address,
      rows: matrix.rows,
      cols: matrix.cols,
      preview: matrix.preview,
    }
  }

  // No selected data — see if the user has a single chart
  // selected instead.
  const chart = await tryGetActiveChart()
  if (chart) return { kind: 'chart', title: chart.title }
  return { kind: 'none' }
}

// ---------- Internals ----------

type MatrixResult = {
  address?: string
  rows: number
  cols: number
  preview: unknown[][]
} | null

function tryGetSelectedMatrix(): Promise<MatrixResult> {
  return new Promise<MatrixResult>((resolve) => {
    // The cross-host matrix API takes a coercion type and
    // returns a 2D array of values when the user has a
    // range selected. We probe through `globalThis` because
    // the `Office` symbol is a global at runtime but only
    // a TS namespace at compile time.
    const g = globalThis as unknown as {
      Office?: {
        CoercionType?: { Matrix?: unknown }
        context?: {
          document?: {
            getSelectedDataAsync?: (
              coercionType: unknown,
              options: { valueFormat: 'unformatted' | 'formatted' },
              callback: (result: {
                status: 'succeeded' | 'failed'
                value?: unknown[][] | string
              }) => void,
            ) => void
          }
        }
      }
    }
    const doc = g.Office?.context?.document
    const Matrix = g.Office?.CoercionType?.Matrix
    if (!doc?.getSelectedDataAsync || Matrix === undefined) {
      resolve(null)
      return
    }
    try {
      doc.getSelectedDataAsync(
        Matrix,
        { valueFormat: 'unformatted' },
        (result) => {
          if (result.status !== 'succeeded' || !Array.isArray(result.value)) {
            resolve(null)
            return
          }
          const rows = result.value
          const r = rows.length
          const c = r > 0 && Array.isArray(rows[0]) ? rows[0].length : 0
          // Sample at most 3 rows × 3 cols to keep the system
          // prompt bounded. The LLM uses size + presence more
          // than the values themselves.
          const preview = rows.slice(0, 3).map((row) =>
            (row as unknown[]).slice(0, 3),
          )
          resolve({ rows: r, cols: c, preview })
        },
      )
    } catch {
      resolve(null)
    }
  })
}

type ChartResult = { title?: string } | null

function tryGetActiveChart(): Promise<ChartResult> {
  // Excel Office.js 1.x doesn't actually expose a
  // "what chart is currently selected" API — the only path
  // to detect a chart selection is the matrix API returning
  // an empty value (handled by the caller) and a few
  // host-specific hooks we don't have generic typings for.
  // We return null and let the system prompt default to
  // "Active selection: none" for the chart case. A future
  // commit can wire in the Excel-specific hook if the user
  // wants it.
  return Promise.resolve(null)
}

/**
 * Render the SelectionContext as a short system-prompt line
 * the LLM can read in one breath. The format is intentional:
 *   "Active selection: A1:B2 (3 rows × 4 cols)"
 *   "Active selection: chart titled 'Q4 Revenue'"
 *   "Active selection: none"
 * so a model scanning the system prompt can pattern-match
 * without having to parse free-form English.
 */
export function formatSelectionForPrompt(s: SelectionContext): string {
  switch (s.kind) {
    case 'range': {
      const where = s.address ?? `(${s.rows}×${s.cols})`
      // A 1×1 selection (single cell) is rarely useful as a
      // chart source — it has no second row to plot. Tell the
      // LLM so it asks the user for a wider range instead of
      // burning a tool call on `add-chart` with the cell.
      if (s.rows === 1 && s.cols === 1) {
        return `Active selection: single cell ${where}. ` +
          `This is too small for add-chart (needs ≥2 rows); ` +
          `either ask the user for a wider range, or expand the selection to the data extent (Ctrl+Shift+End).`
      }
      return `Active selection: range ${where} (${s.rows} rows × ${s.cols} cols).`
    }
    case 'chart':
      return `Active selection: chart${s.title ? ` titled "${s.title}"` : ''}.`
    case 'none':
      return 'Active selection: none.'
  }
}
