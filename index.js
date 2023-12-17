const pup = require("puppeteer");
const xlJs = require("exceljs");

const url = "https://www.infomoney.com.br/mercados/";

(async () => {
    const browser = await pup.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto(url);


    await page.waitForSelector("iframe");
    const iframeElement = await page.$("iframe");
    const iframe = await iframeElement.contentFrame();
    const fechar = await iframe.$("#fechar");
    if (fechar != null) {
        await fechar.click();
    };

    await page.waitForSelector("#High");
    const highButton = await page.$('#High');
    await highButton.click();
    
    const wait = (milliseconds) => new Promise(resolve => setTimeout(resolve, milliseconds));

    await wait(5000);

    const tableSelector = '.table-sm';

    await page.waitForSelector(tableSelector);

    const wb = new xlJs.Workbook();
    const ws = wb.addWorksheet('Tabela');


    // Parse do HTML da tabela, incluindo headers
    const rows = await page.evaluate((tableSelector) => {
        const rows = [];
        const table = document.querySelector(tableSelector);
        const rowElements = table.querySelectorAll('tr');

        rowElements.forEach(row => {
            const rowData = [];
            const cellElements = row.querySelectorAll('th, td');

            cellElements.forEach(cell => {
                rowData.push(cell.textContent.trim());
            });

            rows.push(rowData);
        });

        return rows;
    }, tableSelector);

    // Adiciona os dados na planilha
    rows.forEach(row => {
        ws.addRow(row);
    });

    // Formata a tabela no Excel
    ws.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    
        if (rowNumber === 1) {
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } };
            };
        });
    });
    

    // Salva o arquivo Excel
    await wb.xlsx.writeFile('ações em alta.xlsx');


    await browser.close();
})();