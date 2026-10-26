const GOOGLE_FILE_ID = /^[a-zA-Z0-9_-]{20,}$/

export function splitInput(value: string): string[] {
  return value
    .split(/[\n,;，；\t ]+/)
    .map((item) => item.trim())
    .filter(Boolean)
}

export function parseGoogleFileIds(value: string, deduplicate = true): string[] {
  const ids = splitInput(value)
    .map((item) => {
      if (GOOGLE_FILE_ID.test(item)) return item

      try {
        const url = new URL(item)
        const pathMatch = url.pathname.match(
          /\/(?:spreadsheets\/d|document\/d|presentation\/d|file\/d|drive\/folders)\/([a-zA-Z0-9_-]+)/,
        )
        const queryId = url.searchParams.get('id')
        return pathMatch?.[1] || (queryId && GOOGLE_FILE_ID.test(queryId) ? queryId : '')
      } catch {
        const looseMatch = item.match(
          /\/(?:spreadsheets\/d|document\/d|presentation\/d|file\/d|drive\/folders)\/([a-zA-Z0-9_-]+)/,
        )
        return looseMatch?.[1] || ''
      }
    })
    .filter((item): item is string => Boolean(item))

  return deduplicate ? [...new Set(ids)] : ids
}

export function parseEmails(value: string, deduplicate = true): string[] {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
  const emails = splitInput(value).map((email) => email.toLowerCase()).filter((email) => emailPattern.test(email))
  return deduplicate ? [...new Set(emails)] : emails
}

export function countInvalidEmails(value: string): number {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
  return splitInput(value).filter((email) => !emailPattern.test(email)).length
}

export function escapeCsvCell(value: unknown): string {
  let cell = String(value ?? '')

  // Spreadsheet programs may evaluate imported cells as formulas. Logs contain
  // user-controlled file names and API error text, so neutralize formula sigils.
  if (/^[\t\r]/.test(cell) || /^\s*[=+\-@]/.test(cell)) cell = `'${cell}`

  return `"${cell.replaceAll('"', '""')}"`
}

export async function mapWithConcurrency<T, R>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const output = new Array<R>(items.length)
  let cursor = 0

  async function run() {
    while (cursor < items.length) {
      const index = cursor++
      output[index] = await worker(items[index], index)
    }
  }

  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, run))
  return output
}
