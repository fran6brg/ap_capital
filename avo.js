'use strict'

// requires
const puppeteer = require('puppeteer');
const config = require('./config.json');
const params = require('./params.json');
const chalk = require('chalk');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./credentials.json');

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

// login + scrap the mfs pages
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
            let url = params.urls.avos + (letter + 10).toString(36) + "/1"
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
                if (hrefs[i].includes(params.slt.avos.link_to_n_letter_page)) {
                    if (parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10) > nb_page_for_letter) {
                        nb_page_for_letter = parseInt(hrefs[i].split('/', 5).slice(-1)[0], 10);
                    }
                }
            }
            console.log("   nb pages (for letter " + (letter + 10).toString(36) + "): " + nb_page_for_letter);
            /* -------------------------------- */
            for (var nb_page = 1; nb_page <= nb_page_for_letter; nb_page++) {
                /* -------------------------------- */
                /* go to letter/[nb_page]
                /* -------------------------------- */
                url = params.urls.main.slice(0, -1)
                    + params.slt.avos.link_to_n_letter_page
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
                let rows = [];
                let avo_cards = [];
                avo_cards = await page.$$(params.slt.avos.card);
                console.log("   nb entries: ", avo_cards.length);
                /* -------------------------------- */
                for (let i = 0; i < avo_cards.length; i++)
                {
                    let href = await avo_cards[i].$eval('a', a => a.getAttribute('href'));
                    let name_addr = await (await avo_cards[i].getProperty('innerText')).jsonValue();
                    let tab = name_addr.split(/\r?\n/);
                    let row = {
                        id: 'c' + (letter + 10).toString(36) + ('0' + nb_page).slice(-2),
                        letter: (letter + 10).toString(36),
                        number: nb_page,
                        name: tab[0].trim(),
                        barreau: tab[2].trim(),
                        prestation: tab[4].trim(),
                        address: tab[6].trim(),
                        url: params.urls.main.slice(0, -1) + href
                    }
                    // console.table(row);
                    rows.push(row);
                    index++;
                }
                console.log(chalk.green("   +" + rows.length + " entries added to excel"));
                /* -------------------------------- */
                await sheet.addRows(rows);
            }
        }
        catch (error) {
            console.error("catch: " + error);
            continue;
        }
    }
    /* -------------------------------- */
    await browser.close();
}

/* -------------------------------- */
async function main() {
    await scrapper();
}
  
main();