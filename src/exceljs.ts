import { Workbook, stream } from 'exceljs'
import { homedir } from 'os'
import { join } from 'path'

const homeDir = homedir()
const filePath = join(homeDir, 'excel-files', 'users.xlsx')

const readXLSX = async () => {
  // const wb = new Workbook()
  // await wb.xlsx.readFile(filePath)
  const wb = new stream.xlsx.WorkbookReader(filePath, { styles: 'cache' })
  // const iterator = wb[Symbol.asyncIterator]()
  // const worksheet = (await iterator.next()).value
  // console.log(`worksheet`, worksheet.iterator)
  for await (const worksheetReader of wb) {
    for await (const row of worksheetReader) {
      row.eachCell(cell => {
        cell.value
        cell.style
      })
      // console.log(row.values)
      // const cell = row.getCell(2)
      // console.log(cell.value)
      // console.log(cell.style)
    }
  }
}

const main = async () => {
  readXLSX()
}

main()