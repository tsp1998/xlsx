import * as exceljs from 'exceljs'
import { homedir } from 'os'
import { join } from 'path'

const homeDir = homedir()
// const filePath = join(homeDir, 'excel-files', 'users.xlsx')
const filePath = join(homeDir, 'excel-files', 'QualitiaTestData_New_5k.xlsx')
const filePath2 = join(homeDir, 'excel-files', 'users2.xlsx')
const filePath3 = join(homeDir, 'excel-files', 'QualitiaTestData_New_5k-new.xlsx')
const filePath4 = join(homeDir, 'excel-files', 'users.xlsx')
const filePath15 = join(homeDir, 'excel-files', '87337.xlsx')
const filePath152 = join(homeDir, 'excel-files', '87337-2.xlsx')
const filePath16 = join(homeDir, 'excel-files', 'BLSE_CreditUnderwriting Stage_Enhacements1.xlsx')
const filePath162 = join(homeDir, 'excel-files', 'BLSE_CreditUnderwriting Stage_Enhacements1-2.xlsx')
const filePath17 = join(homeDir, 'excel-files', 'US_77809_Risk Segmentation and TTD Monitoring-Credit Underwriting-PSBL.xlsx')
const filePath172 = join(homeDir, 'excel-files', 'US_77809_Risk Segmentation and TTD Monitoring-Credit Underwriting-PSBL-2.xlsx')

const streamXLSX = async (params: {
  filePath: string,
  chunkSize?: number
}) => {
  const { filePath, chunkSize = 100 } = params
  const workbookReader = new exceljs.stream.xlsx.WorkbookReader(filePath, {})
  const rows: unknown[][] = []

  for await (const worksheetReader of workbookReader) {
    let _rows: string[][] = []
    for await (const row of worksheetReader) {
      _rows.push(row.values as string[])
      if (_rows.length === 100) {
        rows.push(..._rows)
        // console.log(rows.length)
        _rows = []
      }
    }
    rows.push(..._rows)
  }

  console.log(`streaming completed`)
  return rows
}

const writeXLSXWithStream = async (params: {
  filePath: string,
  rows: unknown[][]
}) => {
  const workbookWriter = new exceljs.stream.xlsx.WorkbookWriter({ filename: params.filePath })
  const worksheet = workbookWriter.addWorksheet('sheet1')
  console.log('stream write started')

  for (let i = 0; i < params.rows.length; i++) {
    worksheet.addRow(params.rows[i])
  }
  // worksheet.autoFilter = {
  //   from: { row: 1, column: 1 },
  //   to: { row: 2, column: 2 }
  // }
  // worksheet.autoFilter = 'A1:C1';
  worksheet.commit()
  await workbookWriter.commit()
  console.log('written')
}

const rowsCount = 1500
const columnCount = 5000
const cellContentPrefix = '$NULL'

const main = async () => {
  // const rows = [...new Array(rowsCount)].map((_, i) => {
  //   return [...new Array(columnCount)].map((_, j) => `${cellContentPrefix}${i}${j}`)
  // })
  const rows = await streamXLSX({ filePath: filePath16 })
  console.log(`rows.length`, rows.length)
  const writeFilePath = join(homeDir, 'excel-files', `e${rowsCount}x${columnCount}.xlsx`)
  // writeXLSXWithStream({ rows, filePath: writeFilePath })
  writeXLSXWithStream({ rows, filePath: filePath162 })
}

main()