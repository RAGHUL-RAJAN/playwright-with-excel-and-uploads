const ExcelJs = require('exceljs')

async function excelTest() {
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile('C:\\Users\\HP\\Downloads\\ExceldownloadTest.xlsx')
    const worksheet = workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            // console.log(cell.value);
            if (cell.value === 'Apple') {
                console.log(rowNumber);
                console.log(colNumber);
            }
        })
    })

    const cell = worksheet.getCell(3, 2);
    cell.value = 'iPhone';
    await workbook.xlsx.writeFile('C:\\Users\\HP\\Downloads\\ExceldownloadTest.xlsx');
}


excelTest();