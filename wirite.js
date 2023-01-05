const Excel = require('exceljs')
const _ = require('lodash')

const chalk = require('chalk')

function gengerate(arr, month, fileName, suffix) {
  const workbook = new Excel.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')

  const alignment = { vertical: 'middle', horizontal: 'center' }

  worksheet.pageSetup.margins = {
    left: 0.2,
    right: 0.2,
  }

  // worksheet.pageSetup.printArea = 'A1:G34';
  worksheet.pageSetup.horizontalCentered = true

  worksheet.pageSetup.fitToPage = true
  worksheet.pageSetup.paperSize = 9

  const font = {
    name: '宋体',
    size: 12,
  }

  const style = {
    font,
    alignment,
  }

  const border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  }

  worksheet.columns = [
    { width: 6, style },
    { width: 20, style },
    { width: 20, style },
    { width: 15, style },
    { width: 17, style },
    { width: 17, style },
    { width: 20, style },
  ]

  const boundfont = {
    name: '宋体',
    size: 14,
    bold: true,
  }

  worksheet.mergeCells('A1:G1')
  worksheet.mergeCells('A2:G2')

  const row = worksheet.getRow(1)

  row.height = 38
  row.font = boundfont
  row.alignment = alignment

  const row2 = worksheet.getRow(2)
  row2.height = 28
  row2.font = { ...boundfont, size: 10 }
  row2.alignment = { ...alignment, horizontal: 'left' }

  row.getCell(1).value = fileName

  row2.getCell(1).value = `记录月份：${month}月`

  const row3 = worksheet.getRow(3)
  row3.height = 22

  for (let i = 3; i <= 34; i++) {
    const rowNew = worksheet.getRow(i)
    rowNew.height = 22
    for (let j = 1; j <= 7; j++) {
      rowNew.getCell(j).border = border
    }
  }

  const rows = _.map(arr, (item, index) => {
    const noIndex = _.findIndex(Object.values(item), (val) => !val)

    if (noIndex > -1) {
      console.log(item, chalk.red.bold('====数据不全===='))
    }
    if (!item.dayNum) {
      return [index + 1, '', item.name, '', item.time, '', item.dep]
    }
    return [
      index + 1,
      `${month}月${item.dayNum}日`,
      item.name,
      item.white ? '白' : '夜',
      item.time,
      '',
      item.dep,
    ]
  })
  worksheet.addTable({
    name: 'table',
    ref: 'A3',
    style: { theme: false },
    columns: [
      { name: '序号' },
      { name: '日期' },
      { name: '姓名' },
      { name: '班次' },
      { name: '工时' },
      { name: '领班签字' },
      { name: '备注' },
    ],
    rows,
  })
  workbook.xlsx
    .writeFile(`./dist/${fileName}-${suffix}.xlsx`)
    .then(() => {})
    .catch(() => {
      console.log(chalk.red.bold('写入文件出错了!!!!!!!'))
    })
}

module.exports = {
  gengerate,
}
