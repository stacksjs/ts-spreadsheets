import type { SpreadsheetWrapper } from './'

/**
 * Content is the data structure that represents the spreadsheet content.
 *
 * @example
 * const content: Content = {
 *   headings: ['Name', 'Age', 'City'],
 *   data: [
 *     ['John Doe', 30, 'New York'],
 *     ['Jane Smith', 25, 'London'],
 *     ['Bob Johnson', 35, 'Paris']
 *   ]
 * }
 */
export interface Content {
  headings: string[]
  data: (string | number)[][]
}

/**
 * One named tab in a workbook.
 *
 * A spreadsheet with several tabs is the ordinary shape of a report, and a
 * `Content` can only ever describe one. CSV has no concept of a tab, so an
 * export to CSV writes the sheets one after another with the name on its own
 * row between them, which is the closest honest reading of a workbook in a
 * format that has no word for it.
 *
 * Excel's own rules on the name are applied when the workbook is written: 31
 * characters, none of `: \ / ? * [ ]`, and unique within the workbook. A name
 * that breaks one is corrected rather than rejected, because losing an entire
 * export over a slash in a report title helps nobody.
 */
export interface Sheet extends Content {
  name: string
}

/** Either one sheet's content, or several named ones. */
export type Workbook = Content | Sheet[]

export type SpreadsheetType = 'csv' | 'excel'

export interface SpreadsheetContent {
  content: string | Uint8Array
  type: SpreadsheetType
}

export type SpreadsheetOptions = Partial<{
  type: SpreadsheetType
}>

export type FileExtension = '.csv' | '.xlsx'

export interface Spreadsheet {
  (
    data: Workbook,
  ): {
    csv: () => SpreadsheetWrapper
    excel: () => SpreadsheetWrapper
    store: (path: string) => Promise<void>
    generateCSV: () => SpreadsheetWrapper
    generateExcel: () => SpreadsheetWrapper
  }
  create: (data: Workbook, options: SpreadsheetOptions) => SpreadsheetContent
  generate: (data: Workbook, options: SpreadsheetOptions) => string | Uint8Array
  generateCSV: (content: Workbook) => SpreadsheetWrapper
  generateExcel: (content: Workbook) => SpreadsheetWrapper
  store: (spreadsheet: SpreadsheetContent, path: string) => Promise<void>
  download: (spreadsheet: SpreadsheetContent, filename: string) => Response
}
