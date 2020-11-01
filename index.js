'use strict'

// requires
const puppeteer = require('puppeteer');
const config = require('./config.json');
const params = require('./params.json');
const chalk = require('chalk');
// const Excel = require('exceljs');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./credentials.json');

// connect to account with puppeteer
async function login(){ 
    /* -------------------------------- */
    let browser = await puppeteer.launch({
        // headless: false,
        args: [
        '--no-sandbox',
        // '--headless',
        '--disable-gpu',
        '--window-size=3000x1500'] 
    });
    /* -------------------------------- */
    let page = await browser.newPage();
    console.log("page user agent: " + browser.userAgent());
    // await page.setUserAgent('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36');
    /* -------------------------------- */
    await page.goto(params.urls.main, { waitUntil: 'domcontentloaded' });
    await page.click(params.slt.login.nav_btn);
    await page.type(params.slt.login.mail_input_field, config.francis.mail, { delay: 30 });
    await page.type(params.slt.login.password_input_field, config.francis.password, { delay: 30 });
    // await page.click(params.slt.login.submit_btn);
    await Promise.all([ await page.click(params.slt.login.submit_btn) ]);
    await page.waitForNavigation({ waitUntil: 'domcontentloaded'});
    // await page.waitForTimeout(1500);
    /* -------------------------------- */
    try {
        console.log("trying to login");
        await page.waitFor(params.slt.login.displayed_if_logged);            
        console.log(chalk.green("successfully logged in"));
    } catch (error) {
        console.error("failed to login");
        process.exit(0);
    }
}

// auto scroll page
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
async function scrapper() {
    /* -------------------------------- */
    let browser = await puppeteer.launch({
        headless: false,
        args: [
        '--no-sandbox',
        '--headless',
        '--disable-gpu',
        '--window-size=3000x1500'] 
    });
    /* -------------------------------- */
    let page = await browser.newPage();
    console.log("page user agent: " + browser.userAgent());
    // await page.setUserAgent('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36');
    /* -------------------------------- */
    await page.goto(params.urls.main, { waitUntil: 'domcontentloaded' });
    await page.click(params.slt.login.nav_btn);
    await page.type(params.slt.login.mail_input_field, config.francis.mail, { delay: 30 });
    await page.type(params.slt.login.password_input_field, config.francis.password, { delay: 30 });
    // await page.click(params.slt.login.submit_btn);
    await Promise.all([ await page.click(params.slt.login.submit_btn) ]);
    await page.waitForNavigation({ waitUntil: 'domcontentloaded'});
    // await page.waitForTimeout(1500);
    /* -------------------------------- */
    try {
        console.log("trying to login");
        await page.waitFor(params.slt.login.displayed_if_logged);            
        console.log(chalk.green("successfully logged in"));
    } catch (error) {
        console.error("failed to login");
        process.exit(0);
    }
    // /* -------------------------------- */
    // var workbook = new Excel.Workbook();
    // await workbook.xlsx.readFile(params.wb.name)
    // let worksheet = await workbook.getWorksheet(params.wb.ws.cabinets)
    /* -------------------------------- */
    const doc = new GoogleSpreadsheet(params.gsheet.id);
    await doc.useServiceAccountAuth({
        client_email: creds.client_email,
        private_key: creds.private_key,
    });
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    /* -------------------------------- */
    let index = 1;
    let letter = 0;
    let hrefs = [];
    // for (i = 0; i < 26; i++) { console.log((i + 10).toString(36)) };
    for (letter = 0; letter < 26; letter++) {
        try {
            /* -------------------------------- */
            let url = params.urls.cabs + (letter + 10).toString(36) + "/1"
            console.log("go to " + url);
            await page.goto(url, { waitUntil: 'domcontentloaded' });
            await page.setViewport({
                width: 1200,
                height: 800
            });
            await page.waitForTimeout(1500);
            /* -------------------------------- */
            await autoScroll(page);
            /* -------------------------------- */
            hrefs = []
            hrefs = await page.evaluate(
                () => Array.from(
                    document.querySelectorAll('a[href]'),
                    a => a.getAttribute('href')
                )
            );
            console.log("   nb hrefs retrieved: " + hrefs.length);
            /* -------------------------------- */
            let nb_page_for_letter = 1;
            for (var i = 0; i < hrefs.length; i++) {
                if (hrefs[i].includes(params.slt.cabs.link_to_n_letter_page)) {
                    if (parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10) > nb_page_for_letter) {
                        nb_page_for_letter = parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10);
                    }
                }
            }
            console.log("   nb pages (for letter " + (letter + 10).toString(36) + "): " + nb_page_for_letter);
            /* -------------------------------- */
            for (var nb_page = 1; nb_page <= nb_page_for_letter; nb_page++) {
                /* -------------------------------- */
                let cabs_url = [];
                /* -------------------------------- */
                if (nb_page > 1) {
                    /* -------------------------------- */
                    url = params.urls.main.slice(0, -1)
                        + params.slt.cabs.link_to_n_letter_page
                        + (letter + 10).toString(36)
                        + "/"
                        + nb_page
                    console.log("go to " + url);
                    await page.goto(url, { waitUntil: 'domcontentloaded' });
                    await page.setViewport({
                        width: 1200,
                        height: 800
                    });
                    await page.waitForTimeout(1500);
                    /* -------------------------------- */
                    await autoScroll(page);
                    /* -------------------------------- */
                    hrefs = []
                    hrefs = await page.evaluate(
                        () => Array.from(
                            document.querySelectorAll('a[href]'),
                            a => a.getAttribute('href')
                        )
                    );
                    /* -------------------------------- */
                }
                /* -------------------------------- */
                for (var i = 0; i < hrefs.length; i++) {
                    if (hrefs[i].includes(params.slt.cabs.link_to_cab)) {
                        cabs_url.push(hrefs[i]);
                    }
                }
                console.log("   nb cabs url retrieved: " + cabs_url.length);
                /* -------------------------------- */
                let rows = [];
                for (var i = 0; i < cabs_url.length; i++) {
                    let row = {
                        id: '',
                        letter: '',
                        number: '',
                        url: ''
                    }
                    rows.push(row);
                    // console.log("   row: index=" + index + " | letter=" + (letter + 10).toString(36) + " | page_i=" + nb_page + " | url=" + cabs_url[i]);
                    /* -------------------------------- */
                    row.id = 'c' + (letter + 10).toString(36) + nb_page;
                    row.letter = (letter + 10).toString(36);
                    row.number = nb_page;
                    row.url = params.urls.main.slice(0, -1) + cabs_url[i];
                    /* -------------------------------- */
                    // workbook.xlsx.writeFile(params.wb.name);
                    /* -------------------------------- */
                    index++;
                }
                await sheet.addRows(rows);
                console.log(chalk.green("   +" + cabs_url.length + " cabs added to excel"));
            }
        }
        catch (error) {
            console.error("for 1: " + error);
            continue;
        }
    }
    /* -------------------------------- */
    await browser.close();
}

/* -------------------------------- */
async function main() {
    // await login();
    await scrapper();
}
  
main();
