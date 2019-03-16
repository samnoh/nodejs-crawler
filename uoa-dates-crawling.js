// 16/03/19
// Parse UoA important dates and make an excel file
// You need two packages to run this.

const puppeteer = require('puppeteer');
const excel = require('excel4node');

const uoaDates = async () => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto('https://www.auckland.ac.nz/en/students/academic-information/important-dates.html');

    const semesters = await page.evaluate(() => Array.from(document.querySelectorAll('div.text.section h3'))
        .map(e => ({
            title: e.innerText
        })));

    const dates = await page.evaluate(() => {
        const tbodies = Array.from(document.querySelectorAll('div.table-responsive-wrap'));
        const titles = new Array();
        for (let i = 0; i < tbodies.length; i++) {
            if (i !== 2) { // handle undexpected data
                titles.push(0); // to check where to put semester title
            }
            const trs = tbodies[i].querySelectorAll('table tbody tr td');
            for (let j = 0; j < trs.length; j += 2) {
                if (trs[j] !== null) {
                    titles.push({
                        title: trs[j].innerText.trim(),
                        date: trs[j + 1].innerText.trim(),
                    })
                }
            }
        }
        return titles;
    });

    await browser.close();

    return { dates: dates, semesters: semesters };
};

const excel_file = new excel.Workbook();

uoaDates().then(data => {
    const ws = excel_file.addWorksheet('UoA Dates');
    let count = 0, row = 1;

    for (let i = 0; i < data.dates.length; i++) {
        if (typeof data.dates[i] === 'number') { // semesters titles
            if (i !== 0) {
                ws.cell(row++, 1).string('');
            }
            ws.cell(row++, 1).string(data.semesters[count++].title)
                .style({ font: { bod: true, size: 14, color: '#0000ff' } });
        }
        else { // dates
            ws.cell(row, 1).string(data.dates[i].title);
            ws.cell(row++, 2).string(data.dates[i].date);
        }
    }

    excel_file.write(__dirname + '/uoa_dates.xlsx');
    console.log('Created an excel file at ' + __dirname)
})
