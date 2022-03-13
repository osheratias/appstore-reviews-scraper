const https = require('https')
const bl = require('bl');
const parseString = require('xml2js-parser').parseString;
const excel = require('excel4node');

let workbook = new excel.Workbook({
    dateFormat: 'dd/mm/yyyy hh:mm'
});


// Create a reusable style
const titleStyle = workbook.createStyle({
    font: {
        color: '#FF0800',
        size: 12
    }
});

// These are the apps well scrape ratings format is: name - resulting excel file tab name, url - appstore url of the app
let apps = [
    {name: "בנק הפועלים - ניהול החשבון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=362123205/sortby=mostrecent/xml"},
    {name: "בנק דיסקונט - ניהול החשבון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=342173691/sortby=mostrecent/xml"},
    {name: "בנק לאומי - ניהול החשבון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=337963192/sortby=mostrecent/xml"},
    {name: "בנק מזרחי טפחות - ניהול החשבון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=508767900/sortby=mostrecent/xml"},
    {name: "ביט Bit", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1182007739/sortby=mostrecent/xml"},
    {name: "פייבוקס - PayBox", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=895491053/sortby=mostrecent/xml"},
    {name: "PAY", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1138800563/sortby=mostrecent/xml"},
    {name: "בנק הפועלים - מסחר בשוק ההון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=423213809/sortby=mostrecent/xml"},
    {name: "Pepper Invest", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1331464989/sortby=mostrecent/xml"},
    {name: "מזרחי טפחות - שוק ההון", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=969834635/sortby=mostrecent/xml"},
    {name: "דיסקונט טרייד", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1074658971/sortby=mostrecent/xml"},
    {name: "לאומי טרייד Leumi Trade", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=438890482/sortby=mostrecent/xml"},
    {name: "בנק הפועלים - פועלים לעסקים", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1192185734/sortby=mostrecent/xml"},
    {name: "דיסקונט עסקים+", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=444480867/sortby=mostrecent/xml"},
    {name: "מרכנתיל עסקים+", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=444486794/sortby=mostrecent/xml"},
    {name: "בנק הפועלים – open", url: "https://itunes.apple.com/$location/rss/customerreviews/page=$page/id=1454373103/sortby=mostrecent/xml"}
];

// Geo stores to check
let locationsToCheck = ["us", "il", "ru"];

async function asyncHttpGet(url) {
    return new Promise((resolve, reject) => {
        https.get(url, response => {
            response.setEncoding('utf8');
            response.pipe(bl((err, data) => {
                if (err) {
                    reject(err);
                }
                resolve(data.toString());
            }));
        });
    });
}

async function asyncParseString(xmlData) {
    return new Promise((resolve, reject) => {
        parseString(xmlData, (err, result) => {
            if (err) {
                reject(err);
            }
            resolve(result);
        });
    });
}

async function asyncForEach(array, callback) {
    if (array) {
        for (let index = 0; index < array.length; index++) {
            await callback(array[index], index, array);
        }
    }
}

const start = async () => {

    // Loop through apps
    await asyncForEach(apps, async (app) => {

        console.info("Started working on: %s", app.name);

        // Create worksheet and set title line
        let worksheet = workbook.addWorksheet(app.name);
        worksheet.cell(1, 1).string("appstore geo").style(titleStyle);
        worksheet.cell(1, 2).string("date").style(titleStyle);
        worksheet.column(2).setWidth(20);
        worksheet.cell(1, 3).string("version").style(titleStyle);
        worksheet.cell(1, 4).string("rating").style(titleStyle);
        worksheet.cell(1, 5).string("author").style(titleStyle);
        worksheet.cell(1, 6).string("title").style(titleStyle);
        worksheet.column(6).setWidth(50);
        worksheet.cell(1, 7).string("comment").style(titleStyle);
        worksheet.column(7).setWidth(60);
        let currentLine = 2;

        // Check all locations for current app
        await asyncForEach(locationsToCheck, async (location) => {

            let currUrl = app.url.replace("$location", location);

            // Run over 10 pages (apple limit the oldest review page you can get by rss to 10)
            for (let page=1; page<=10; page++) {
                currUrl = currUrl.replace("$page", page.toString());
                let reviews = null;
                let xmlData = await asyncHttpGet(currUrl);
                reviews = await asyncParseString(xmlData);

                // Loop through reviews
                await asyncForEach(reviews.feed.entry, async (review) => {
                    // Set data in excel
                    worksheet.cell(currentLine, 1).string(location);
                    worksheet.cell(currentLine, 2).date(review.updated[0].slice(0, -6));
                    worksheet.cell(currentLine, 3).string(review["im:version"][0].toString());
                    worksheet.cell(currentLine, 4).string(review["im:rating"][0].toString());
                    worksheet.cell(currentLine, 5).string(review.author[0].name[0].toString());
                    worksheet.cell(currentLine, 6).string(review.title[0].toString());
                    worksheet.cell(currentLine, 7).string(review.content[0]["_"].toString());
                    currentLine++;
                })
            }
        });
    });
    console.info("Writing result file");
    workbook.write("Apple store reviews.xlsx");
}

start();

