import { homedir } from 'os'
import { join } from 'path'
import * as xlsx from 'xlsx'
import { createReadStream, createWriteStream, ReadStream } from 'fs'

const homeDir = homedir()
// const filePath = join(homeDir, 'excel-files', 'users.xlsx')
const filePath = join(homeDir, 'excel-files', 'QualitiaTestData_New_5k.xlsx')
const filePath2 = join(homeDir, 'excel-files', 'users2.xlsx')
const filePath3 = join(homeDir, 'excel-files', 'QualitiaTestData_New_5k-new.xlsx')
const filePath4 = join(homeDir, 'excel-files', 'users.xlsx')

const readXLSX = (params: {
  filePath: string,
}) => {
  console.log('read start')
  const wb = xlsx.readFile(params.filePath, { cellStyles: true })
  console.log('read complete')
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1 })
  console.log(`ws['A1']`, ws['C1000'].s)
  // console.log(`rows`, rows)
  return rows
}

const streamXLSX = (params: {
  filePath: string,
}): Promise<unknown[]> => {
  return new Promise(resolve => {
    console.log('stream start')

    function process_RS(stream: ReadStream, cb: Function) {
      const buffers: any[] = [];
      stream.on("data", function (data) { buffers.push(data); });
      stream.on("end", function () {
        const buffer = Buffer.concat(buffers);
        const workbook = xlsx.read(buffer);
        cb(workbook);
      });
    }

    const readStream = createReadStream(params.filePath)
    process_RS(readStream, (wb: xlsx.WorkBook) => {
      console.log('stream end')
      const ws = wb.Sheets[wb.SheetNames[0]]
      // const rows = xlsx.utils.sheet_to_json(ws, { header: 1 })
      // @ts-ignore
      const rows = xlsx.utils.sheet_to_row_object_array(ws, { header: 1 })
      resolve(rows)
    })
    // const wbStream = xlsx.stream.to_json(readStream);
    // wbStream.on('data', (row: unknown) => {
    //   console.log(`row`, row)
    // })
    // wbStream.on('end', () => {
    //   console.log('done')
    // })
  })
}

const writeXLSX = (params: {
  filePath: string,
  rows: unknown[][]
}) => {
  const wb = xlsx.utils.book_new()
  const ws = xlsx.utils.aoa_to_sheet([])
  xlsx.utils.book_append_sheet(wb, ws, 'sheet1')
  xlsx.utils.sheet_add_aoa(ws, params.rows)
  xlsx.writeFile(wb, params.filePath)
}

const writeXLSXWithStream = (params: {
  filePath: string,
  rows: unknown[][]
}) => {
  const wb = xlsx.utils.book_new()
  const ws = xlsx.utils.aoa_to_sheet([])
  xlsx.utils.book_append_sheet(wb, ws, 'sheet1')
  xlsx.utils.sheet_add_aoa(ws, params.rows)
  const writeStream = createWriteStream(params.filePath)
  // xlsx.writeFile(wb, params.filePath)
  xlsx.stream.to_csv(ws).pipe(writeStream)
  writeStream.on('finish', () => {
    console.log('written')
  })
}

const rowsCount = 25000
const columnCount = 300
const cellContentPrefix = 'Lorem ipsum dolor sit amet consectetur adipisicing elit.'

const main = async () => {
  // const rows = [...new Array(rowsCount)].map((_, i) => {
  //   return [...new Array(columnCount)].map((_, j) => `${cellContentPrefix}${i}${j}`)
  // })
  // writeXLSXWithStream({ rows, filePath: filePath2 })
  // writeXLSX({ rows, filePath: filePath2 })
  // console.log(readXLSX({ filePath: filePath2 }).length)
  // console.log((await streamXLSX({ filePath: filePath2 })).length)
  // console.log(readXLSX({ filePath: filePath3 }).length)
  // console.log((await streamXLSX({ filePath: filePath3 })).length)
  console.log(readXLSX({ filePath: filePath4 }).length)
  console.log((await streamXLSX({ filePath: filePath4 })).length)
  // console.log(readXLSX({ filePath }).length)
  // console.log((await streamXLSX({ filePath })).length)
}

main()