const ExcelJs = require('exceljs')

async function excelTest() {
    let output ={row:-1,column:-1};
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile('C:\\Users\\HP\\Downloads\\ExceldownloadTest.xlsx')
    const worksheet = workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            // console.log(cell.value);
            if (cell.value === 'iPhone') {
               output.row = rowNumber;
               output.column = colNumber;
            }
        })
    })

    const cell = worksheet.getCell(output.row, output.column);
    cell.value = 'jackfruite';
    await workbook.xlsx.writeFile('C:\\Users\\HP\\Downloads\\ExceldownloadTest.xlsx');
}


excelTest();