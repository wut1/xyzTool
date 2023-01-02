const Excel = require('exceljs');

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

const alignment = { vertical: 'middle', horizontal: 'center' };

const font = {
  name: '宋体',
  size: 12,
}

const style = {
  font, alignment
}

const border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

worksheet.columns = [
  { width: 6, style, },
  { width: 20, style },
  { width: 20, style },
  { width: 15, style },
  { width: 17, style },
  { width: 17, style },
  { width: 20, style }
];

const boundfont = {
  name: '宋体',
  size: 14,
  bold: true,
}



worksheet.mergeCells('A1:G1');
worksheet.mergeCells('A2:G2');

const row = worksheet.getRow(1);

row.height = 38;
row.font = boundfont
row.alignment = alignment


const row2 = worksheet.getRow(2);
row2.height = 28;
row2.font = { ...boundfont, size: 10 }
row2.alignment = { ...alignment, horizontal: 'left' }

row.getCell(1).value = '天佮劳务宝思利产线计时明细表'

row2.getCell(1).value = `记录月份：月`

const row3 = worksheet.getRow(3);
row3.height = 22;

for (let i = 3; i <= 34; i++) {
  const rowNew = worksheet.getRow(i);
  rowNew.height = 22;
  for (let j = 1; j <= 7; j++) {
    rowNew.getCell(j).border = border
  }
}



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
  rows: [],
});





workbook.xlsx.writeFile('./dist/result.xlsx');