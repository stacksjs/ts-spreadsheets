import type { Content } from '../src/types'
import { afterEach, beforeEach, describe, expect, it } from 'bun:test'
import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync, rmdirSync, unlinkSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import { createSpreadsheet, spreadsheet } from '../src/index'

function runCLI(command: string): string {
  try {
    // Use proper path handling
    const cliPath = resolve(__dirname, '../bin/cli.ts')

    // Don't split the command by spaces, pass it as a single argument
    const result = spawnSync('bun', [cliPath, ...command.match(/(?:[^\s"]|"[^"]*")+/g)?.map(arg =>
      arg.startsWith('"') && arg.endsWith('"') ? arg.slice(1, -1) : arg,
    ) || []], {
      encoding: 'utf8',
      env: {
        ...process.env,
        NODE_ENV: 'test',
        NO_COLOR: '1',
      },
      cwd: process.cwd(), // Ensure we're in the right directory
    })

    // Combine stdout and stderr
    const output = [
      result.stdout?.trim(),
      result.stderr?.trim(),
    ].filter(Boolean).join('\n')

    if (result.error) {
      console.error('CLI execution error:', result.error)
    }

    return output || ''
  }
  catch (error) {
    console.error('Unexpected error in runCLI:', error)
    return error instanceof Error ? error.message : String(error)
  }
}

// Helper function to wait for file to exist
async function waitForFile(filepath: string, timeoutMs = 2000): Promise<boolean> {
  const startTime = Date.now()
  while (Date.now() - startTime < timeoutMs) {
    if (existsSync(filepath))
      return true
    await new Promise(resolve => setTimeout(resolve, 100))
  }
  return false
}

describe('ts-spreadsheets', () => {
  let testData: Content

  beforeEach(() => {
    testData = {
      headings: ['Name', 'Age', 'City'],
      data: [
        ['John Doe', 30, 'New York'],
        ['Jane Smith', 25, 'London'],
        ['Bob Johnson', 35, 'Paris'],
      ],
    }
  })

  afterEach(() => {
    // Clean up any files created during tests
    const filesToDelete = ['output.csv', 'output.xlsx']
    filesToDelete.forEach((file) => {
      if (existsSync(file))
        unlinkSync(file)
    })
  })

  describe('Content Creation', () => {
    it('should create valid Content object', () => {
      expect(testData.headings.length).toBe(3)
      expect(testData.data.length).toBe(3)
      expect(testData.data[0].length).toBe(3)
    })

    it('should handle empty data', () => {
      const emptyData: Content = { headings: [], data: [] }
      expect(() => createSpreadsheet(emptyData)).not.toThrow()
    })

    it('should handle single row data', () => {
      const singleRowData: Content = {
        headings: ['Test'],
        data: [['Value']],
      }
      expect(() => createSpreadsheet(singleRowData)).not.toThrow()
    })
  })

  describe('CSV Generation', () => {
    it('should generate valid CSV content', () => {
      const csvContent = spreadsheet(testData).csv().getContent() as string
      const lines = csvContent.split('\n')
      expect(lines[0]).toBe('Name,Age,City')
      expect(lines[1]).toBe('John Doe,30,New York')
    })

    it('should handle special characters in CSV', () => {
      const specialData: Content = {
        headings: ['Name', 'Description'],
        data: [['John, Doe', 'Likes "quotes"']],
      }
      const csvContent = spreadsheet(specialData).csv().getContent() as string
      expect(csvContent).toBe('Name,Description\n"John, Doe","Likes ""quotes"""')
    })

    it('should correctly store numbers in CSV', () => {
      const numericData: Content = {
        headings: ['Name', 'Age', 'Score'],
        data: [
          ['Alice', 28, 95.5],
          ['Bob', 32, 88],
          ['Charlie', 45, 72.75],
        ],
      }
      const csvContent = spreadsheet(numericData).csv().getContent() as string
      const lines = csvContent.split('\n')
      expect(lines[0]).toBe('Name,Age,Score')
      expect(lines[1]).toBe('Alice,28,95.5')
      expect(lines[2]).toBe('Bob,32,88')
      expect(lines[3]).toBe('Charlie,45,72.75')
    })

    it('should handle empty cells', () => {
      const dataWithEmpty: Content = {
        headings: ['Col1', 'Col2'],
        data: [['', 'value'], ['value', '']],
      }
      const csvContent = spreadsheet(dataWithEmpty).csv().getContent() as string
      expect(csvContent).toBe('Col1,Col2\n,value\nvalue,')
    })
  })

  describe('Excel Generation', () => {
    it('should generate valid Excel content', () => {
      const excelContent = spreadsheet(testData).excel().getContent() as Uint8Array
      expect(excelContent).toBeInstanceOf(Uint8Array)
      expect(excelContent.length).toBeGreaterThan(0)
    })

    it('should generate larger Excel files correctly', () => {
      const largeData: Content = {
        headings: ['ID', 'Value'],
        data: Array.from({ length: 1000 }, (_, i) => [i, `Value ${i}`]),
      }
      const excelContent = spreadsheet(largeData).excel().getContent() as Uint8Array
      expect(excelContent).toBeInstanceOf(Uint8Array)
      expect(excelContent.length).toBeGreaterThan(0)
    })
  })

  describe('File Storage', () => {
    it('should store CSV file', async () => {
      await spreadsheet(testData).store('output.csv')
      expect(existsSync('output.csv')).toBe(true)
    })

    it('should store Excel file', async () => {
      await spreadsheet(testData).store('output.xlsx')
      expect(existsSync('output.xlsx')).toBe(true)
    })

    it('should handle file overwrite', async () => {
      await spreadsheet(testData).store('output.csv')
      await spreadsheet(testData).store('output.csv')
      expect(existsSync('output.csv')).toBe(true)
    })
  })

  describe('Download Response', () => {
    it('should create valid download response for CSV', () => {
      const response = spreadsheet(testData).csv().download('test.csv')
      expect(response).toBeInstanceOf(Response)
      expect(response.headers.get('Content-Type')).toBe('text/csv')
      expect(response.headers.get('Content-Disposition')).toBe('attachment; filename="test.csv"')
    })

    it('should create valid download response for Excel', () => {
      const response = spreadsheet(testData).excel().download('test.xlsx')
      expect(response).toBeInstanceOf(Response)
      expect(response.headers.get('Content-Type')).toBe(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      )
      expect(response.headers.get('Content-Disposition')).toBe('attachment; filename="test.xlsx"')
    })
  })

  describe('Method Chaining', () => {
    it('should support method chaining', async () => {
      await spreadsheet(testData).csv().store('output.csv')
      expect(existsSync('output.csv')).toBe(true)
    })

    it('should support multiple operations', async () => {
      const result = spreadsheet(testData)
      await result.csv().store('output.csv')
      await result.excel().store('output.xlsx')
      expect(existsSync('output.csv')).toBe(true)
      expect(existsSync('output.xlsx')).toBe(true)
    })
  })

  describe('Error Handling', () => {
    it('should throw error for unsupported spreadsheet type', () => {
      // @ts-expect-error: Testing invalid type
      expect(() => createSpreadsheet(testData, { type: 'pdf' })).toThrow()
    })

    it('should handle invalid file paths', async () => {
      await expect(
        spreadsheet(testData).store('/invalid/path/file.csv'),
      ).rejects.toThrow()
    })
  })

  describe('CLI Integration', () => {
    const testDir = resolve(process.cwd(), '.test-output')
    let testData: Content
    let testFilePath: string

    beforeEach(() => {
      // Ensure test directory exists
      mkdirSync(testDir, { recursive: true })

      testData = {
        headings: ['Name', 'Age', 'City'],
        data: [
          ['John Doe', 30, 'New York'],
          ['Jane Smith', 25, 'London'],
          ['Bob Johnson', 35, 'Paris'],
        ],
      }
      testFilePath = resolve(testDir, 'test-input.json')
      writeFileSync(testFilePath, JSON.stringify(testData, null, 2))
    })

    afterEach(() => {
      if (existsSync(testDir)) {
        readdirSync(testDir).forEach((file) => {
          unlinkSync(resolve(testDir, file))
        })
        rmdirSync(testDir)
      }
    })

    describe('create command', () => {
      it('should create CSV file from JSON input', async () => {
        const outputPath = resolve(testDir, 'output.csv')
        const output = runCLI(`create "${testFilePath}" -o "${outputPath}"`)

        await waitForFile(outputPath)
        expect(existsSync(outputPath)).toBe(true)
        expect(output).toContain('Spreadsheet saved')
      })

      it('should create Excel file from JSON input', async () => {
        const outputPath = resolve(testDir, 'output.xlsx')
        const output = runCLI(`create "${testFilePath}" --type excel -o "${outputPath}"`)

        await waitForFile(outputPath)
        expect(existsSync(outputPath)).toBe(true)
        expect(output).toContain('Spreadsheet saved')
      })

      it('should output to stdout when no output file specified', () => {
        const output = runCLI(`create "${testFilePath}"`)
        expect(output).toContain('Name,Age,City')
      })

      it('should handle invalid JSON input', () => {
        const invalidPath = resolve(testDir, 'invalid.json')
        writeFileSync(invalidPath, '{ invalid json }')
        const output = runCLI(`create "${invalidPath}"`)
        expect(output).toContain('Failed to create spreadsheet')
      })
    })

    describe('convert command', () => {
      it('should convert CSV to Excel', async () => {
        const csvPath = resolve(testDir, 'test-output.csv')
        const xlsxPath = resolve(testDir, 'test-output.xlsx')

        // First create CSV
        runCLI(`create "${testFilePath}" -o "${csvPath}"`)
        await waitForFile(csvPath)

        // Then convert
        const output = runCLI(`convert "${csvPath}" "${xlsxPath}"`)
        await waitForFile(xlsxPath)

        expect(existsSync(xlsxPath)).toBe(true)
        expect(output).toContain('Converted')
      })

      it('should warn when input and output formats are the same', async () => {
        const csvPath = resolve(testDir, 'test-output.csv')
        const samePath = resolve(testDir, 'same-output.csv')

        runCLI(`create "${testFilePath}" -o "${csvPath}"`)
        await waitForFile(csvPath)

        const output = runCLI(`convert "${csvPath}" "${samePath}"`)
        expect(output).toContain('same')
      })

      it('should handle invalid input file', () => {
        const nonexistentPath = resolve(testDir, 'nonexistent.csv')
        const outputPath = resolve(testDir, 'output.xlsx')
        const output = runCLI(`convert "${nonexistentPath}" "${outputPath}"`)
        expect(output).toContain('Failed to convert spreadsheet')
      })
    })

    describe('validate command', () => {
      it('should validate correct JSON format', () => {
        const output = runCLI(`validate "${testFilePath}"`)
        expect(output).toContain('Input JSON is valid')
      })

      it('should catch missing headings', () => {
        const invalidPath = resolve(testDir, 'invalid.json')
        const invalidData = { data: [['test']] }
        writeFileSync(invalidPath, JSON.stringify(invalidData))
        const output = runCLI(`validate "${invalidPath}"`)
        expect(output).toContain('Missing or invalid headings array')
      })

      it('should catch invalid data types', () => {
        const invalidPath = resolve(testDir, 'invalid.json')
        const invalidData = {
          headings: ['Test'],
          data: [[{ invalid: 'object' }]],
        }
        writeFileSync(invalidPath, JSON.stringify(invalidData))
        const output = runCLI(`validate "${invalidPath}"`)
        expect(output).toContain('Data must be an array of arrays')
      })

      it('should catch non-string headings', () => {
        const invalidPath = resolve(testDir, 'invalid.json')
        const invalidData = {
          headings: [42],
          data: [['test']],
        }
        writeFileSync(invalidPath, JSON.stringify(invalidData))
        const output = runCLI(`validate "${invalidPath}"`)
        expect(output).toContain('Headings must be strings')
      })
    })
  })
})

/**
 * Reading the workbook back.
 *
 * The excel tests above assert that some bytes came out and that there are more
 * than zero of them, which is what let a completely unopenable file pass for
 * years: the writer produced a gzip stream labelled as stored, no CRCs and no
 * end-of-central-directory record, so `unzip -t` said "cannot find zipfile
 * directory" and so did Excel. Nothing here had ever opened one.
 *
 * These do. The archive is parsed with the same rules a reader applies, which
 * is the only assertion that can tell a workbook from 1500 plausible bytes.
 */
describe('the xlsx is a workbook a reader can open', () => {
  /** The same CRC-32 a zip reader checks each entry against. */
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
      crc = CRC_TABLE[(crc ^ bytes[i]!) & 0xFF]! ^ (crc >>> 8)
    return (crc ^ 0xFFFFFFFF) >>> 0
  }

  /** The entries of a zip, by name, verifying each CRC as it goes. */
  function unzipEntries(archive: Uint8Array): Map<string, Uint8Array> {
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)

    // Find the end-of-central-directory record, which is what a reader looks
    // for first and what the old writer never wrote.
    let end = -1
    for (let i = archive.length - 22; i >= 0; i--) {
      if (view.getUint32(i, true) === 0x06054B50) {
        end = i
        break
      }
    }

    expect(end).toBeGreaterThanOrEqual(0)

    const count = view.getUint16(end + 10, true)
    let cursor = view.getUint32(end + 16, true)
    const entries = new Map<string, Uint8Array>()

    for (let i = 0; i < count; i++) {
      expect(view.getUint32(cursor, true)).toBe(0x02014B50)

      const method = view.getUint16(cursor + 10, true)
      const crc = view.getUint32(cursor + 16, true)
      const compressedSize = view.getUint32(cursor + 20, true)
      const nameLength = view.getUint16(cursor + 28, true)
      const extraLength = view.getUint16(cursor + 30, true)
      const commentLength = view.getUint16(cursor + 32, true)
      const localOffset = view.getUint32(cursor + 42, true)
      const name = new TextDecoder().decode(archive.slice(cursor + 46, cursor + 46 + nameLength))

      expect(view.getUint32(localOffset, true)).toBe(0x04034B50)

      const localNameLength = view.getUint16(localOffset + 26, true)
      const localExtraLength = view.getUint16(localOffset + 28, true)
      const dataStart = localOffset + 30 + localNameLength + localExtraLength
      const payload = archive.slice(dataStart, dataStart + compressedSize)

      const content = method === 0
        ? payload
        : new Uint8Array(inflateRawSync(payload))

      // The CRC is what proves the bytes survived, and it was zero throughout
      // the old archives.
      expect(crc32(content)).toBe(crc)

      entries.set(name, content)
      cursor += 46 + nameLength + extraLength + commentLength
    }

    return entries
  }

  /** Every row of a worksheet, as the values a reader would show. */
  function rowsOf(sheetXml: Uint8Array): (string | number)[][] {
    const text = new TextDecoder().decode(sheetXml)
    const rows: (string | number)[][] = []

    for (const rowMatch of text.matchAll(/<row[^>]*>(.*?)<\/row>/g)) {
      const cells: (string | number)[] = []

      for (const cellMatch of rowMatch[1]!.matchAll(/<c[^>]*?(t="inlineStr")?>(.*?)<\/c>/g)) {
        const body = cellMatch[2]!

        if (cellMatch[1]) {
          const inline = body.match(/<t[^>]*>(.*?)<\/t>/)
          cells.push((inline?.[1] ?? '')
            .replace(/&quot;/g, '"')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&'))
        }
        else {
          cells.push(Number(body.match(/<v>(.*?)<\/v>/)?.[1] ?? 0))
        }
      }

      rows.push(cells)
    }

    return rows
  }

  const sample: Content = {
    headings: ['Block', 'Point', 'Series', 'Value'],
    data: [
      ['Revenue per day', '2026-08-01T00:00:00Z', 'total', 4250],
      ['Orders', 'a, comma & an ampersand', 'say "hi"', 12],
    ],
  }

  it('produces an archive whose entries all pass their CRC', () => {
    const entries = unzipEntries(spreadsheet(sample).excel().getContent() as Uint8Array)

    expect([...entries.keys()]).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
    ])
  })

  it('keeps text as text and numbers as numbers', () => {
    const entries = unzipEntries(spreadsheet(sample).excel().getContent() as Uint8Array)

    // The heading row used to come back as numbers, because a string written
    // into a bare <v> is a shared-string index to Excel and there is no shared
    // string table here.
    expect(rowsOf(entries.get('xl/worksheets/sheet1.xml')!)).toEqual([
      ['Block', 'Point', 'Series', 'Value'],
      ['Revenue per day', '2026-08-01T00:00:00Z', 'total', 4250],
      ['Orders', 'a, comma & an ampersand', 'say "hi"', 12],
    ])
  })

  it('writes one tab per sheet, named and made safe', () => {
    const archive = spreadsheet([
      { name: 'Revenue per day', headings: ['Point', 'Value'], data: [['2026-08-01', 4250]] },
      { name: 'Orders by plan/status', headings: ['Series', 'Value'], data: [['processing', 120]] },
      { name: 'Orders by plan/status', headings: ['Series', 'Value'], data: [['cancelled', 8]] },
    ]).excel().getContent() as Uint8Array

    const entries = unzipEntries(archive)
    const workbook = new TextDecoder().decode(entries.get('xl/workbook.xml')!)
    const names = [...workbook.matchAll(/<sheet name="([^"]*)"/g)].map(match => match[1])

    // The slash is illegal in a sheet name and a repeat opens the workbook with
    // a repair prompt, so both are corrected rather than passed through.
    expect(names).toEqual(['Revenue per day', 'Orders by plan status', 'Orders by plan status 2'])
    expect(entries.has('xl/worksheets/sheet3.xml')).toBe(true)
    expect(rowsOf(entries.get('xl/worksheets/sheet2.xml')!)).toEqual([['Series', 'Value'], ['processing', 120]])
  })

  it('names columns past Z properly', () => {
    const wide: Content = {
      headings: Array.from({ length: 28 }, (_, index) => `H${index}`),
      data: [Array.from({ length: 28 }, (_, index) => index)],
    }

    const entries = unzipEntries(spreadsheet(wide).excel().getContent() as Uint8Array)
    const sheet = new TextDecoder().decode(entries.get('xl/worksheets/sheet1.xml')!)

    // `String.fromCharCode(65 + 26)` is '[', so column 27 used to collide with
    // the one before it and the row silently lost cells.
    expect(sheet).toContain('r="AA1"')
    expect(sheet).toContain('r="AB1"')
    expect(rowsOf(entries.get('xl/worksheets/sheet1.xml')!)[1]).toHaveLength(28)
  })

  it('writes several sheets to CSV one after another', () => {
    const csv = spreadsheet([
      { name: 'First', headings: ['A'], data: [['one']] },
      { name: 'Second', headings: ['B'], data: [['two']] },
    ]).csv().getContent() as string

    expect(csv).toBe('First\nA\none\n\nSecond\nB\ntwo')
  })
})

describe('a formula in a cell cannot run when the file is opened', () => {
  it('defuses the four characters a spreadsheet treats as a formula', () => {
    const csv = spreadsheet({
      headings: ['Name'],
      data: [['=1+1'], ['+1'], ['-1'], ['@SUM(A1)'], ['safe']],
    }).csv().getContent() as string

    // Tab-prefixed and therefore quoted, so the cell is text and says the same
    // thing. `safe` is untouched: this must not quote every cell in the file.
    expect(csv).toBe('Name\n"\t=1+1"\n"\t+1"\n"\t-1"\n"\t@SUM(A1)"\nsafe')
  })

  it('agrees with the PHP sibling byte for byte', () => {
    // stacksjs/php-spreadsheets prefixes then quotes in exactly this order.
    // The two write the same product's exports and a customer moving between
    // them should not get a different file.
    const csv = spreadsheet({ headings: ['A'], data: [['=HYPERLINK("http://x")']] })
      .csv().getContent() as string

    expect(csv).toBe('A\n"\t=HYPERLINK(""http://x"")"')
  })
})
