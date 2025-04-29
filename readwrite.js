const ExcelJs = require('exceljs')

async function writeTest(searchText, replaceText, filepath) {

    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile(filepath)
    const worksheet = workbook.getWorksheet('Sheet1');
    const output = await readExcel(worksheet, searchText);
    const cell = worksheet.getCell(output.row, output.column);
    cell.value = replaceText;
    await workbook.xlsx.writeFile(filepath);
}


async function readExcel(worksheet, searchText) {
    let output = { row: -1, column: -1 };
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {

            if (cell.value === searchText) {
                output.row = rowNumber;
                output.column = colNumber;
            }
        })
    })
    return output;

}


writeTest("Apple", "Grapes", "C:\\Users\\HP\\Downloads\\ExceldownloadTest.xlsx");