import type {
  Content,
  FileExtension,
  Sheet,
  Spreadsheet,
  SpreadsheetContent,
  SpreadsheetOptions,
  SpreadsheetType,
  Workbook,
} from './types'
import { Buffer } from 'node:buffer'
import { readFile, writeFile } from 'node:fs/promises'
import { deflateRawSync } from 'node:zlib'

export const spreadsheet: Spreadsheet = Object.assign(
  (data: Workbook) => ({
    csv: () => spreadsheet.generateCSV(data),
    excel: () => spreadsheet.generateExcel(data),
    store: async (path: string) => {
      const extension = path.slice(path.lastIndexOf('.')) as FileExtension
      const type = extension === '.csv' ? 'csv' : 'excel'
      const content = spreadsheet.generate(data, { type })
      await spreadsheet.store({ content, type }, path)
    },
    generateCSV: () => spreadsheet.generateCSV(data),
    generateExcel: () => spreadsheet.generateExcel(data),
  }),
  {
    generate: (data: Workbook, options: SpreadsheetOptions = { type: 'csv' }): string | Uint8Array => {
      const generators: Record<SpreadsheetType, (content: Workbook) => string | Uint8Array | SpreadsheetWrapper> = {
        csv: spreadsheet.generateCSV,
        excel: spreadsheet.generateExcel,
      }

      const generator = generators[options.type || 'csv']

      if (!generator) {
        throw new Error(`Unsupported spreadsheet type: ${options.type}`)
      }

      const result = generator(data)
      if (result instanceof SpreadsheetWrapper) {
        return result.getContent()
      }

      return result
    },

    create: (data: Workbook, options: SpreadsheetOptions = { type: 'csv' }): SpreadsheetContent => ({
      content: spreadsheet.generate(data, options),
      type: options.type || 'csv',
    }),

    generateCSV: (content: Workbook): SpreadsheetWrapper => {
      const csvContent = generateCSVContent(content)
      return new SpreadsheetWrapper(csvContent, 'csv')
    },

    generateExcel: (content: Workbook): SpreadsheetWrapper => {
      const excelContent = generateExcelContent(content)
      return new SpreadsheetWrapper(excelContent, 'excel')
    },

    store: async ({ content }: SpreadsheetContent, path: string): Promise<void> => {
      try {
        await writeFile(path, content)
      }
      catch (error) {
        throw new Error(`Failed to store spreadsheet: ${(error as Error).message}`)
      }
    },

    download: ({ content, type }: SpreadsheetContent, filename: string): Response => {
      const mimeType = type === 'csv' ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      const blob = new Blob([content], { type: mimeType })

      return new Response(blob, {
        headers: {
          'Content-Type': mimeType,
          'Content-Disposition': `attachment; filename="${filename}"`,
        },
      })
    },
  },
)

export class SpreadsheetWrapper {
  constructor(
    private content: string | Uint8Array,
    private type: SpreadsheetType,
  ) {}

  getContent(): string | Uint8Array {
    return this.content
  }

  download(filename: string): Response {
    return spreadsheet.download({ content: this.content, type: this.type }, filename)
  }

  store(path: string): Promise<void> {
    return spreadsheet.store({ content: this.content, type: this.type }, path)
  }
}

export function createSpreadsheet(data: Workbook, options: SpreadsheetOptions = { type: 'csv' }): SpreadsheetWrapper {
  const content = spreadsheet.generate(data, options)

  return new SpreadsheetWrapper(content, options.type || 'csv')
}

/**
 * One cell, quoted only where a reader would otherwise misparse it.
 *
 * A leading `=`, `+`, `-` or `@` is prefixed with a tab first. Those are the
 * four characters a spreadsheet treats as the start of a formula, and a CSV is
 * very often an export of data somebody else supplied: a display name of
 * `=HYPERLINK("http://…"&A1)` sits inertly in a database and runs the moment
 * the file is opened in Excel. The tab makes the cell text without changing
 * what it says, which is why it is preferred to stripping the character.
 *
 * The PHP sibling does exactly the same thing, and the two have to agree byte
 * for byte: they write the same product's exports, and a customer moving
 * between them should not get a different file.
 */
// eslint-disable-next-line pickier/no-unused-vars
function csvCell(cell: string | number): string {
  let text = String(cell ?? '')

  if (text !== '' && '=+-@'.includes(text[0]!))
    text = `\t${text}`

  // A lone carriage return breaks a row for the same reason a newline does,
  // and Excel writes them, so it is quoted too. The tab above is quoted for a
  // different reason: unquoted, a reader splitting on tabs would take it as a
  // column boundary.
  if (text.includes(',') || text.includes('"') || text.includes('\n') || text.includes('\r') || text.includes('\t'))
    return `"${text.replace(/"/g, '""')}"`

  return text
}

function csvRows(content: Content): string {
  return [content.headings, ...content.data]
    .map(row => row.map(csvCell).join(','))
    .join('\n')
}

/**
 * A workbook as CSV.
 *
 * CSV has no word for a tab, so several sheets are written one after another
 * with the sheet's name on a row of its own and a blank line between them. No
 * reader will split that back into tabs, and it is not meant to: it is the
 * arrangement somebody scrolling the file can follow, where three heading rows
 * silently concatenated is not.
 */
// eslint-disable-next-line pickier/no-unused-vars
export function generateCSVContent(content: Workbook): string {
  if (!Array.isArray(content))
    return csvRows(content)

  return content
    .map(sheet => `${csvCell(sheet.name)}\n${csvRows(sheet)}`)
    .join('\n\n')
}

/**
 * The sheets of a workbook, however it was described.
 *
 * A bare `Content` is one unnamed sheet, which is what every caller wrote
 * before workbooks had tabs, and it keeps the name Excel defaults to.
 */
function sheetsOf(content: Workbook): Sheet[] {
  if (Array.isArray(content))
    return content.length > 0 ? content : [{ name: 'Sheet1', headings: [], data: [] }]

  return [{ name: 'Sheet1', headings: content.headings, data: content.data }]
}

/**
 * A sheet name Excel will actually open.
 *
 * The rules are Excel's, not ours: at most 31 characters, none of `: \ / ? * [
 * ]`, not blank, and unique in the workbook. A name that breaks one is
 * corrected here rather than rejected, because a report called "Revenue / cost"
 * is a reasonable thing to call a report and losing the whole export over the
 * slash is not a reasonable response to it.
 */
function safeSheetName(name: string, taken: Set<string>): string {
  let cleaned = String(name ?? '').replace(/[:\\/?*[\]]/g, ' ').trim().slice(0, 31)

  if (!cleaned)
    cleaned = 'Sheet'

  let candidate = cleaned
  let suffix = 2

  // Excel compares sheet names case-insensitively, so the uniqueness check has
  // to as well or the workbook opens with a repair prompt.
  while (taken.has(candidate.toLowerCase())) {
    const room = 31 - String(suffix).length - 1
    candidate = `${cleaned.slice(0, room)} ${suffix}`
    suffix++
  }

  taken.add(candidate.toLowerCase())

  return candidate
}

/** `A`, `Z`, `AA`, `AB`... Past column 26 the old single-letter form collided. */
function columnRef(index: number): string {
  let ref = ''
  let n = index

  do {
    ref = String.fromCharCode(65 + (n % 26)) + ref
    n = Math.floor(n / 26) - 1
  } while (n >= 0)

  return ref
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // XML 1.0 has no way to represent these at all, and a single stray control
    // character makes the whole workbook unreadable rather than one cell.
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
}

/**
 * One worksheet part.
 *
 * Strings are written as `inlineStr`, not as a bare `<v>`. An untyped `<v>`
 * means "number" to Excel, so every text cell was either read as a shared
 * string index into a table that does not exist, or shown as a number: the old
 * writer emitted `<c r="A1"><v>Block</v></c>`, and the heading row came back as
 * `#VALUE!` or as whatever integer that index happened to hit.
 */
function worksheetXml(sheet: Sheet): string {
  const rows = [sheet.headings, ...sheet.data]

  const body = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((cell, cellIndex) => {
          const ref = `${columnRef(cellIndex)}${rowIndex + 1}`

          if (typeof cell === 'number' && Number.isFinite(cell))
            return `<c r="${ref}"><v>${cell}</v></c>`

          return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(cell ?? ''))}</t></is></c>`
        })
        .join('')

      return `<row r="${rowIndex + 1}">${cells}</row>`
    })
    .join('')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
    + `<sheetData>${body}</sheetData>`
    + `</worksheet>`
}

/** The CRC-32 every ZIP reader checks each entry against. */
const CRC_TABLE = (() => {
  const table = new Uint32Array(256)

  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++)
      c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1
    table[i] = c >>> 0
  }

  return table
})()

function crc32(bytes: Uint8Array): number {
  let crc = 0xFFFFFFFF

  for (let i = 0; i < bytes.length; i++)
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xFF]! ^ (crc >>> 8)

  return (crc ^ 0xFFFFFFFF) >>> 0
}

interface ZipEntry {
  name: string
  content: Uint8Array
}

/**
 * A ZIP archive, written properly.
 *
 * An xlsx *is* a zip, so this is the part that decides whether the file opens
 * at all, and the previous version did not produce one: it compressed with
 * `gzipSync` (which is deflate wrapped in a gzip header and trailer, not the
 * raw deflate stream a zip entry holds), wrote the compression method into the
 * wrong header offset so it read as "stored", left every CRC at zero, and
 * ended the archive with twenty-two zero bytes where the end-of-central-
 * directory record goes. `unzip -t` answered "cannot find zipfile directory",
 * and so did Excel, Numbers and LibreOffice. Every xlsx this library has ever
 * produced was unopenable.
 *
 * Deflate is used where it pays and the entry is stored otherwise, since a
 * "compressed" form larger than the original is just a slower way to be bigger.
 */
function zip(entries: ZipEntry[]): Uint8Array {
  const encoder = new TextEncoder()
  const locals: Uint8Array[] = []
  const centrals: Uint8Array[] = []

  let offset = 0

  for (const entry of entries) {
    const name = encoder.encode(entry.name)
    const crc = crc32(entry.content)
    const deflated = new Uint8Array(deflateRawSync(entry.content))
    const stored = deflated.length >= entry.content.length
    const payload = stored ? entry.content : deflated
    const method = stored ? 0 : 8

    const local = new Uint8Array(30 + name.length + payload.length)
    const localView = new DataView(local.buffer)

    localView.setUint32(0, 0x04034B50, true) // local file header signature
    localView.setUint16(4, 20, true) // version needed
    localView.setUint16(6, 0, true) // flags
    localView.setUint16(8, method, true)
    localView.setUint16(10, 0, true) // modification time
    localView.setUint16(12, 0x21, true) // modification date, 1 Jan 1980
    localView.setUint32(14, crc, true)
    localView.setUint32(18, payload.length, true)
    localView.setUint32(22, entry.content.length, true)
    localView.setUint16(26, name.length, true)
    localView.setUint16(28, 0, true) // extra field length
    local.set(name, 30)
    local.set(payload, 30 + name.length)

    const central = new Uint8Array(46 + name.length)
    const centralView = new DataView(central.buffer)

    centralView.setUint32(0, 0x02014B50, true) // central directory signature
    centralView.setUint16(4, 20, true) // version made by
    centralView.setUint16(6, 20, true) // version needed
    centralView.setUint16(8, 0, true) // flags
    centralView.setUint16(10, method, true)
    centralView.setUint16(12, 0, true) // modification time
    centralView.setUint16(14, 0x21, true) // modification date
    centralView.setUint32(16, crc, true)
    centralView.setUint32(20, payload.length, true)
    centralView.setUint32(24, entry.content.length, true)
    centralView.setUint16(28, name.length, true)
    centralView.setUint16(30, 0, true) // extra field length
    centralView.setUint16(32, 0, true) // comment length
    centralView.setUint16(34, 0, true) // disk number
    centralView.setUint16(36, 0, true) // internal attributes
    centralView.setUint32(38, 0, true) // external attributes
    centralView.setUint32(42, offset, true) // offset of the local header
    central.set(name, 46)

    locals.push(local)
    centrals.push(central)
    offset += local.length
  }

  const centralSize = centrals.reduce((sum, part) => sum + part.length, 0)
  const end = new Uint8Array(22)
  const endView = new DataView(end.buffer)

  endView.setUint32(0, 0x06054B50, true) // end of central directory signature
  endView.setUint16(4, 0, true) // this disk
  endView.setUint16(6, 0, true) // disk with the central directory
  endView.setUint16(8, centrals.length, true)
  endView.setUint16(10, centrals.length, true)
  endView.setUint32(12, centralSize, true)
  endView.setUint32(16, offset, true) // where the central directory starts
  endView.setUint16(20, 0, true) // comment length

  const total = offset + centralSize + end.length
  const archive = new Uint8Array(total)

  let cursor = 0
  for (const part of [...locals, ...centrals, end]) {
    archive.set(part, cursor)
    cursor += part.length
  }

  return archive
}

/**
 * A workbook, as the bytes of an `.xlsx` file.
 *
 * Accepts one sheet's `Content` or an array of named `Sheet`s. The parts are
 * the minimum a reader needs: the content types, the package relationships,
 * the workbook and its relationships, and one worksheet each.
 */
export function generateExcelContent(content: Workbook): Uint8Array {
  const taken = new Set<string>()
  const sheets = sheetsOf(content).map(sheet => ({
    ...sheet,
    name: safeSheetName(sheet.name, taken),
  }))

  const encoder = new TextEncoder()
  const part = (name: string, xml: string): ZipEntry => ({ name, content: encoder.encode(xml) })

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
    + sheets
      .map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join('')
    + `</Types>`

  const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`
    + `</Relationships>`

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<sheets>`
    + sheets
      .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
      .join('')
    + `</sheets>`
    + `</workbook>`

  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + sheets
      .map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`)
      .join('')
    + `</Relationships>`

  return zip([
    part('[Content_Types].xml', contentTypes),
    part('_rels/.rels', packageRels),
    part('xl/workbook.xml', workbook),
    part('xl/_rels/workbook.xml.rels', workbookRels),
    ...sheets.map((sheet, index) => part(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet))),
  ])
}

export async function csvToContent(csvPath: string): Promise<Content> {
  const csvContent = await readFile(csvPath, 'utf-8')

  // Split into lines and parse CSV
  const lines = csvContent.split('\n').map(line =>
    line.split(',').map((cell) => {
      const trimmed = cell.trim()
      // Remove quotes if present
      if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
        // Handle escaped quotes
        return trimmed.slice(1, -1).replace(/""/g, '"')
      }
      return trimmed
    }),
  )

  // First line is headers
  const [headings, ...data] = lines

  // Convert numeric strings to numbers
  const typedData = data.map(row =>
    row.map((cell) => {
      const num = Number(cell)
      return !Number.isNaN(num) && cell.trim() !== '' ? num : cell
    }),
  )

  return {
    headings,
    data: typedData,
  }
}

export * from './types'
