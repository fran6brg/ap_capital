'use strict'

// requires
const puppeteer = require('puppeteer');
const config = require('./config.json');
const params = require('./params.json');
const chalk = require('chalk');
const Excel = require('exceljs')

// connect to account with puppeteer
(async () => { 
    let browser = await puppeteer.launch({
        headless: false,
        args: [
        '--no-sandbox',
        // '--headless',
        '--disable-gpu',
        '--window-size=3000x1500'] 
    });
    let page = await browser.newPage();
    await page.goto(params.urls.main, { waitUntil: 'networkidle0' });
    await page.click(params.slt.login.nav_btn);
    await page.type(params.slt.login.mail_input_field, config.mail, { delay: 30 });
    await page.type(params.slt.login.password_input_field, config.password, { delay: 30 });
    // await page.click(params.slt.login.submit_btn);
    await Promise.all([ await page.click(params.slt.login.submit_btn) ]);
    await page.waitForNavigation({ waitUntil: 'networkidle0'});
    await page.waitFor(15000);
    try {
        console.log("trying to login");
        await page.waitFor(params.slt.login.displayred_if_logged);            
        console.log(chalk.green("successfully logged in"));
    } catch (error) {
        console.error("failed to login");
        process.exit(0);
    }
})();

// 
let failed = [];
(async () => {
    var wb = new Excel.Workbook();
    await wb.xlsx.readFile('./linkedin_v2.xlsx')
        .then(function () {
            var worksheet = wb.getWorksheet('links');
            worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
                // console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
                if (row.values[4] == "na" || row.values[4].includes("broke url, can't fetch") || row.values[4].includes("unhandled error")) {
                    let fail = {
                        row: rowNumber,
                        name: row.values[1],
                        site: row.values[2]
                    }
                    failed.push(fail);
                }
            });
        });
    console.table(failed);
})();

async function autoScroll(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;
                if (totalHeight >= scrollHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
}

// get links
(async () => {
    var workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(params.wb.name)
    let worksheet = await workbook.getWorksheet(params.wb.ws.cabinets)
    /* -------------------------------- */
    let index = 1;
    // for (i = 0; i < 26; i++) { console.log((i + 10).toString(36)) };
    for (letter = 0; letter < 26; letter++) {
        try {
            let url = params.urls.cabs + (letter + 10).toString(36) + "/"
            await page.goto(url, { waitUntil: 'domcontentloaded' });
            await page.setViewport({
                width: 1200,
                height: 800
            });
            /* -------------------------------- */
            await autoScroll(page);
            /* -------------------------------- */
            const hrefs = await page.evaluate(
                () => Array.from(
                    document.querySelectorAll('a[href]'),
                    a => a.getAttribute('href')
                )
            );
            /* -------------------------------- */
            let nb_page_for_letter = 1;
            for (var i = 0; i < hrefs.length; i++) {
                if (hrefs[i].includes(params.slt.cabs.link_to_n_letter_page)) {
                    if (parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10) > nb_page_for_letter) {
                        nb_page_for_letter = parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10);
                    }
                }
            }
            /* -------------------------------- */
            let cabs_urls = [];
            for (var i = 1; i <= nb_page_for_letter; i++) {
                if (i > 1) {
                    let link = params.slt.cabs.link_to_n_letter_page + (letter + 10).toString(36) + "/"
                    await page.goto(link_completed, { waitUntil: 'domcontentloaded' });
                    ; // get hrefs
                }
                for (var i = 0; i < hrefs.length; i++) {
                    if (hrefs[i].includes(params.slt.cabs_urls.link_to_cab)) {
                        cabs_urls.push(hrefs[i]);
                    }
                }
            }
            /* -------------------------------- */
            for (var i = 0; i < cabs_urls.length; i++) {
                let row = {
                    id: worksheet.getCell('A' + index + 1),
                    letter: worksheet.getCell('B' + index + 1),
                    page: worksheet.getCell('C' + index + 1),
                    // name: worksheet.getCell('D' + index + 1),
                    // city: worksheet.getCell('E' + index + 1),
                    // address: worksheet.getCell('F' + index + 1),
                    // creation: worksheet.getCell('G' + index + 1),
                    // nb_lawyer: worksheet.getCell('H' + index + 1),
                    // nb_cab_decisions: worksheet.getCell('I' + index + 1),
                    url: worksheet.getCell('J' + index + 1)
                }
                row.id.value = index;
                row.letter.value = (letter + 10).toString(36);
                // row.letter.page = ; ??????????
                row.url.value = cabs_urls[i];
            }
            workbook.xlsx.writeFile(params.wb.name);
            /* -------------------------------- */
            /* -------------------------------- */
        }
        catch (error) {
            console.error("1: " + error)
        }
    }
    /* -------------------------------- */
    await browser.close();
})();
