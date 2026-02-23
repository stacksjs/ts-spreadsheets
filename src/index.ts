import type {
  Content,
  FileExtension,
  Spreadsheet,
  SpreadsheetContent,
  SpreadsheetOptions,
  SpreadsheetType,
} from './types'
import { Buffer } from 'node:buffer'
import { readFile, writeFile } from 'node:fs/promises'
import { gzipSync } from 'node:zlib'

export const spreadsheet: Spreadsheet = Object.assign(
  (data: Content) => ({
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
    generate: (data: Content, options: SpreadsheetOptions = { type: 'csv' }): string | Uint8Array => {
      const generators: Record<SpreadsheetType, (content: Content) => string | Uint8Array | SpreadsheetWrapper> = {
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

    create: (data: Content, options: SpreadsheetOptions = { type: 'csv' }): SpreadsheetContent => ({
      content: spreadsheet.generate(data, options),
      type: options.type || 'csv',
    }),

    generateCSV: (content: Content): SpreadsheetWrapper => {
      const csvContent = generateCSVContent(content)
      return new SpreadsheetWrapper(csvContent, 'csv')
    },

    generateExcel: (content: Content): SpreadsheetWrapper => {
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

export function createSpreadsheet(data: Content, options: SpreadsheetOptions = { type: 'csv' }): SpreadsheetWrapper {
  const content = spreadsheet.generate(data, options)

  return new SpreadsheetWrapper(content, options.type || 'csv')
}

// eslint-disable-next-line pickier/no-unused-vars
export function generateCSVContent(content: Content): string {
  const rows = [content.headings, ...content.data]

  return rows
    .map(row =>
      row
        // eslint-disable-next-line pickier/no-unused-vars
        .map((cell) => {
          const cellString = String(cell)
          if (cellString.includes(',') || cellString.includes('"') || cellString.includes('\n')) {
            return `"${cellString.replace(/"/g, '""')}"`
          }
          return cellString
        })
        .join(','),
    )
    .join('\n')
}

export function generateExcelContent(content: Content): Uint8Array {
  const workbook = new Uint8Array(
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
      <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
  </workbook>`),
  )

  const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
      ${[content.headings, ...content.data]
        .map(
          (row, rowIndex) => `
      <row r="${rowIndex + 1}">
        ${row
          .map(
            (cell, cellIndex) => `
        <c r="${String.fromCharCode(65 + cellIndex)}${rowIndex + 1}" ${typeof cell === 'number' ? 't="n"' : ''}>
          <v>${typeof cell === 'string' ? cell.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') : cell}</v>
        </c>`,
          )
          .join('')}
      </row>`,
        )
        .join('')}
    </sheetData>
  </worksheet>`

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  </Types>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  </Relationships>`

  const files: Array<{ name: string, content: Uint8Array }> = [
    { name: '[Content_Types].xml', content: new Uint8Array(Buffer.from(contentTypes)) },
    { name: '_rels/.rels', content: new Uint8Array(Buffer.from(rels)) },
    { name: 'xl/workbook.xml', content: workbook },
    { name: 'xl/worksheets/sheet1.xml', content: new Uint8Array(Buffer.from(worksheet)) },
  ]

  const zipData = files.map((file) => {
    const compressedContent = new Uint8Array(gzipSync(file.content))
    const header = new Uint8Array(30 + file.name.length)
    const headerView = new DataView(header.buffer)

    headerView.setUint32(0, 0x04034B50, true) // 'PK\x03\x04'
    headerView.setUint32(4, 0x0008, true)
    headerView.setUint32(18, compressedContent.length, true)
    headerView.setUint32(22, file.content.length, true)
    headerView.setUint16(26, file.name.length, true)

    const encoder = new TextEncoder()
    header.set(encoder.encode(file.name), 30)

    return { header, compressedContent }
  })

  const centralDirectory = files.map((file, index) => {
    const header = new Uint8Array(46 + file.name.length)
    const headerView = new DataView(header.buffer)

    headerView.setUint32(0, 0x02014B50, true) // 'PK\x01\x02'
    headerView.setUint16(4, 0x0014, true)
    headerView.setUint16(6, 0x0008, true)
    headerView.setUint32(8, 0x0008, true)
    headerView.setUint32(20, zipData[index].compressedContent.length - 30 - file.name.length, true)
    headerView.setUint32(
      42,
      zipData.slice(0, index).reduce((acc, curr) => acc + curr.header.length + curr.compressedContent.length, 0),
      true,
    )
    headerView.setUint16(28, file.name.length, true)

    const encoder = new TextEncoder()
    header.set(encoder.encode(file.name), 46)

    return header
  })

  const endOfCentralDirectory = new Uint8Array(22)

  const totalSize
    = zipData.reduce((acc, { header, compressedContent }) => acc + header.length + compressedContent.length, 0)
      + centralDirectory.reduce((acc, header) => acc + header.length, 0)
      + endOfCentralDirectory.length

  // Create a single Uint8Array with the total size
  const result = new Uint8Array(totalSize)

  // Copy data into the result array
  let offset = 0
  for (const { header, compressedContent } of zipData) {
    result.set(header, offset)
    offset += header.length
    result.set(compressedContent, offset)
    offset += compressedContent.length
  }
  for (const header of centralDirectory) {
    result.set(header, offset)
    offset += header.length
  }
  result.set(endOfCentralDirectory, offset)

  return result
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
