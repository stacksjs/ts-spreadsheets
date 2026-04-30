#!/usr/bin/env bun
import { readFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import process from 'node:process'
import { CLI as CAC } from '@stacksjs/clapp'
import { type Content, createSpreadsheet, csvToContent } from '../src/index'

// Use sync version for version to avoid race conditions
const version = process.env.NODE_ENV === 'test'
  ? '0.0.0'
  : JSON.parse(
    readFileSync(resolve(__dirname, '../package.json'), 'utf-8'),
  ).version

interface CLIOptions {
  type?: 'csv' | 'excel'
  output?: string
  pretty?: boolean
}

const cli = new CAC('spreadsheets')

function logMessage(message: string) {
  if (process.env.NODE_ENV === 'test') {
    process.stderr.write(`${message}\n`)
  }
  else {
    console.log(message)
  }
}

cli
  .command('create <input>', 'Create a spreadsheet from JSON input file')
  .option('-t, --type <type>', 'Output type (csv or excel)', { default: 'csv' })
  .option('-o, --output <path>', 'Output file path')
  .action(async (input: string, options: CLIOptions) => {
    try {
      const inputPath = resolve(process.cwd(), input)
      const content = JSON.parse(await readFile(inputPath, 'utf-8')) as Content

      const result = createSpreadsheet(content, { type: options.type })

      if (options.output) {
        await result.store(options.output)
        logMessage(`Spreadsheet saved to ${options.output}`)
      }
      else {
        const output = result.getContent()
        if (typeof output === 'string') {
          process.stdout.write(output)
        }
        else {
          process.stdout.write(output)
        }
      }
    }
    catch (error) {
      logMessage(`Failed to create spreadsheet: ${(error as Error).message}`)
      process.exit(1)
    }
  })

cli
  .command('convert <input> <output>', 'Convert between spreadsheet formats')
  .action(async (input: string, output: string) => {
    try {
      const inputExt = input.slice(input.lastIndexOf('.')) as '.csv' | '.xlsx'
      const outputExt = output.slice(output.lastIndexOf('.')) as '.csv' | '.xlsx'

      if (inputExt === outputExt) {
        logMessage('Input and output formats are the same')
        return
      }

      // Handle CSV input
      let content: Content
      if (inputExt === '.csv') {
        content = await csvToContent(input)
      }
      else {
        // Handle JSON input
        content = JSON.parse(await readFile(input, 'utf-8')) as Content
      }

      const outputType = outputExt === '.csv' ? 'csv' : 'excel'
      const result = createSpreadsheet(content, { type: outputType })
      await result.store(output)

      logMessage(`Converted ${input} to ${output}`)
    }
    catch (error) {
      logMessage(`Failed to convert spreadsheet: ${(error as Error).message}`)
      process.exit(1)
    }
  })

cli
  .command('validate <input>', 'Validate JSON input format')
  .action(async (input: string) => {
    try {
      const content = JSON.parse(await readFile(input, 'utf-8'))

      if (!content.headings || !Array.isArray(content.headings)) {
        throw new Error('Missing or invalid headings array')
      }
      if (!content.data || !Array.isArray(content.data)) {
        throw new Error('Missing or invalid data array')
      }

      const invalidHeadings = content.headings.some((h: unknown) => typeof h !== 'string')
      if (invalidHeadings) {
        throw new Error('Headings must be strings')
      }

      const invalidData = content.data.some((row: unknown[]) =>
        !Array.isArray(row)
        || row.some((cell: unknown) => typeof cell !== 'string' && typeof cell !== 'number'),
      )
      if (invalidData) {
        throw new Error('Data must be an array of arrays containing only strings and numbers')
      }

      logMessage('Input JSON is valid')
    }
    catch (error) {
      logMessage(`Invalid input: ${(error as Error).message}`)
      process.exit(1)
    }
  })

cli.version(version)
cli.help()
cli.parse()
