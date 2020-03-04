const puppeteer = require('puppeteer');
const config = require('./config');
const Excel = require('exceljs');
const fs = require('fs');
const program = require('commander');

program
  .option('--start <number>', 'Start range number for movies array (optional).')
  .option('--stop <number>', 'Stop range number for movies array (optional).')
  .option('--recreate', 'Delete current excel file and create empty one (optional).')
  .option('--worksheet <name>', 'Add a new worksheet to excel file (optional).')
  .parse(process.argv);

const start = program.start;
const stop = program.stop;
const recreate = program.recreate;
const worksheet_name = program.worksheet;

(async () => {

  await excelHandle(recreate);

  const browser = await puppeteer.launch({
    /*args: [
      '--proxy-server=socks5://198.211.99.227:46437'
    ],*/
    headless: false,
    //slowMo: 200,
    //devtools: true,
    defaultViewport: {
      width: 1920,
      height: 1080
    }
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.0 Safari/537.36');
  page.setDefaultTimeout(0);

  await page.goto(config.url, {
    waitUntil: 'networkidle0'
  });
  console.log('Login page initiated.');
  await page.waitForSelector(config.USERNAME_SELECTOR);
  console.log('Login page username selector exists!');
  //await page.waitFor(config.USERNAME_SELECTOR);

  // Additional step if we start at https://www.disneyplus.com
  //await page.screenshot({path: 'screen.png'});
  //await page.click('button[kind="outline"]');

  //await page.screenshot({ path: 'screenEmail.png' });

  await page.click(config.USERNAME_SELECTOR);
  await page.keyboard.type(config.username);
  await page.click('button[name="dssLoginSubmit"]');
  console.log('Login page username form submitted.');

  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });
  //await page.screenshot({ path: 'screenLogIn.png' });

  await page.click(config.PASSWORD_SELECTOR);
  await page.keyboard.type(config.password);
  await page.click('button[name="dssLoginSubmit"]');
  console.log('Login page password form submitted.');

  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });
  console.log('Main loggedin page initiated.');

  //await page.screenshot({ path: 'screenLoggedIn.png' });

  await page.goto(config.movieUrl, {
    waitUntil: 'networkidle0'
  });
  console.log('Movie page initiated.');
  await page.waitForSelector('.sc-jTzLTM');
  //await page.screenshot({ path: 'screenMovie.png' });

  console.log('Movie page scroll to end, to get all movies visible in DOM.');
  await autoScroll(page);
  await page.waitFor(2000);

  //await page.screenshot({
  //  path: 'scrollCheckMovies.png',
  //  fullPage: true
  //});

  //let linksMovies = await page.$$('a.sc-ckVGcZ'); for featured movies
  let linksMovies = await page.$$('a.sc-dxgOiQ');
  let dataMovies = [];

  let fromNumber = !start ? 0 : (parseInt(start) > (linksMovies.length - 1) ? 0 : parseInt(start));
  let toNumber = !stop ? linksMovies.length : (parseInt(stop) > linksMovies.length ? linksMovies.length : parseInt(stop));
  let startNo = fromNumber > toNumber ? 0 : fromNumber;

  console.log('Movies scraping started.');
  for (let i = startNo; i < toNumber; i++) {
    await page.evaluate((i) => {
      //return ([...document.querySelectorAll('a.sc-ckVGcZ')][i]).click(); for featured movies
      return ([...document.querySelectorAll('a.sc-dxgOiQ')][i]).click();
    }, i);

    //await page.waitForSelector('a.sc-ckVGcZ'); for featured movies
    await page.waitForSelector('div.sc-iSDuPN');

    let urlMovie = await page.url();

    let resultMovie = await page.evaluate((urlMovie) => {
      let res = document.querySelector('div > p[class~="sc-gzVnrw"]').innerText;
      let resArr = res.split('\u2022');

      let genre = resArr[2].trim();

      let release_year = resArr[0].trim();

      let viewable_time_step = resArr[1];
      let viewable_time_step1 = resArr[1].replace(/[\D]/g, ' ');
      let viewable_time_step2 = viewable_time_step1.trim().split(' ');
      let viewable_time = viewable_time_step2.length === 1 ?
        (viewable_time_step.indexOf('h') !== -1 ? viewable_time_step2[0] * 60 * 60 : viewable_time_step2[0] * 60) :
        ((viewable_time_step2[0] * 60 * 60) + (viewable_time_step2[2] * 60));

      let rating_step1 = document.querySelector('p > img');
      let rating = 'N/A';
      if (rating_step1) {
        let rating_step2 = rating_step1.alt.split('_');
        rating = rating_step2[2].toUpperCase();
      }

      let title_step = document.querySelector('div[class~="sc-iujRgT"] > img');
      let title = title_step.alt;

      let source_id_step = urlMovie.split('/');
      let source_id = source_id_step[source_id_step.length - 1];

      // setting date for capture date field
      const dateCapture = new Date();
      var dateOptions = {
        year: "numeric",
        month: "2-digit",
        day: "numeric"
      };

      return {
        'bot_system': 'disneyplus',
        'bot_version': '1.0.0',
        'bot_country': 'us',
        'capture_date': dateCapture.toLocaleString('en', dateOptions),
        'offer_type': 'SVOD',
        'purchase_type': '',
        'picture_quality': '',
        'program_price': '',
        'bundle_price': '',
        'currency': '',
        'addon_name': '',
        'is_movie': 1,
        'season_number': 0,
        'episode_number': 0,
        'title': title,
        'genre': genre,
        'source_id': source_id,
        'program_url': urlMovie,
        'maturity_rating': rating,
        'release_date': '',
        'release_year': parseInt(release_year),
        'viewable_runtime': viewable_time,
        'series_title': '',
        'series_release_year': '',
        'series_source_id': '',
        'series_url': '',
        'series_genre': '',
        'season_source_id': ''
      };
    }, urlMovie);

    dataMovies.push(resultMovie);

    await addDataToExcel(resultMovie, worksheet_name);

    console.log(resultMovie);

    await page.goBack();
  }

  console.log('Movies scraping finished.');

  await browser.close();
})();

async function autoScroll(page) {
  await page.evaluate(async () => {
    await new Promise((resolve, reject) => {
      var totalHeight = 0;
      var distance = 100;
      var timer = setInterval(() => {
        var scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, distance);
        totalHeight += distance;

        if (totalHeight >= scrollHeight) {
          clearInterval(timer);
          resolve();
        }
      }, 500);
    });
  });
};

async function excelHandle(recreate) {
  const filePath = './disneyMovies.xlsx';

  try {
    if (fs.existsSync(filePath)) {
      if (recreate) {
        fs.unlinkSync(filePath);
        await createExcel();
      } else {
        console.log('File already exists, will use existing file!');
      }
    } else {
      await createExcel();
    }
  } catch (e) {
    console.error(e);
  }
};

async function createExcel() {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('FEB 2020');

  worksheet.columns = [
    { header: 'bot_system', key: 'bot_system', width: 10 },
    { header: 'bot_version', key: 'bot_version', width: 10 },
    { header: 'bot_country', key: 'bot_country', width: 10 },
    { header: 'capture_date', key: 'capture_date', width: 10 },
    { header: 'offer_type', key: 'offer_type', width: 10 },
    { header: 'purchase_type', key: 'purchase_type', width: 10 },
    { header: 'picture_quality', key: 'picture_quality', width: 10 },
    { header: 'program_price', key: 'program_price', width: 10 },
    { header: 'bundle_price', key: 'bundle_price', width: 10 },
    { header: 'currency', key: 'currency', width: 10 },
    { header: 'addon_name', key: 'addon_name', width: 10 },
    { header: 'is_movie', key: 'is_movie', width: 10 },
    { header: 'season_number', key: 'season_number', width: 10 },
    { header: 'episode_number', key: 'episode_number', width: 10 },
    { header: 'title', key: 'title', width: 10 },
    { header: 'genre', key: 'genre', width: 10 },
    { header: 'source_id', key: 'source_id', width: 10 },
    { header: 'program_url', key: 'program_url', width: 10 },
    { header: 'maturity_rating', key: 'maturity_rating', width: 10 },
    { header: 'release_date', key: 'release_date', width: 10 },
    { header: 'release_year', key: 'release_year', width: 10 },
    { header: 'viewable_runtime', key: 'viewable_runtime', width: 10 },
    { header: 'series_title', key: 'series_title', width: 10 },
    { header: 'series_release_year', key: 'series_release_year', width: 10 },
    { header: 'series_source_id', key: 'series_source_id', width: 10 },
    { header: 'series_url', key: 'series_url', width: 10 },
    { header: 'series_genre', key: 'series_genre', width: 10 },
    { header: 'season_source_id', key: 'season_source_id', width: 10 }
  ];

  // save under disneyMovies.xlsx
  await workbook.xlsx.writeFile('disneyMovies.xlsx');

  console.log('File is created.');
};

async function addDataToExcel(data, worksheet_name) {
  //load a copy of disneyMovies.xlsx
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('disneyMovies.xlsx');
  let worksheet = null;

  if (!worksheet_name) {
    worksheet = workbook.getWorksheet('FEB 2020');
  } else {
    if (!workbook.getWorksheet(worksheet_name)) {
      worksheet = workbook.addWorksheet(worksheet_name);
    } else {
      worksheet = workbook.getWorksheet(worksheet_name);
    }
  }

  worksheet.columns = [
    { header: 'bot_system', key: 'bot_system', width: 10 },
    { header: 'bot_version', key: 'bot_version', width: 10 },
    { header: 'bot_country', key: 'bot_country', width: 10 },
    { header: 'capture_date', key: 'capture_date', width: 10 },
    { header: 'offer_type', key: 'offer_type', width: 10 },
    { header: 'purchase_type', key: 'purchase_type', width: 10 },
    { header: 'picture_quality', key: 'picture_quality', width: 10 },
    { header: 'program_price', key: 'program_price', width: 10 },
    { header: 'bundle_price', key: 'bundle_price', width: 10 },
    { header: 'currency', key: 'currency', width: 10 },
    { header: 'addon_name', key: 'addon_name', width: 10 },
    { header: 'is_movie', key: 'is_movie', width: 10 },
    { header: 'season_number', key: 'season_number', width: 10 },
    { header: 'episode_number', key: 'episode_number', width: 10 },
    { header: 'title', key: 'title', width: 10 },
    { header: 'genre', key: 'genre', width: 10 },
    { header: 'source_id', key: 'source_id', width: 10 },
    { header: 'program_url', key: 'program_url', width: 10 },
    { header: 'maturity_rating', key: 'maturity_rating', width: 10 },
    { header: 'release_date', key: 'release_date', width: 10 },
    { header: 'release_year', key: 'release_year', width: 10 },
    { header: 'viewable_runtime', key: 'viewable_runtime', width: 10 },
    { header: 'series_title', key: 'series_title', width: 10 },
    { header: 'series_release_year', key: 'series_release_year', width: 10 },
    { header: 'series_source_id', key: 'series_source_id', width: 10 },
    { header: 'series_url', key: 'series_url', width: 10 },
    { header: 'series_genre', key: 'series_genre', width: 10 },
    { header: 'season_source_id', key: 'season_source_id', width: 10 }
  ];

  await worksheet.addRow({
    bot_system: data.bot_system,
    bot_version: data.bot_version,
    bot_country: data.bot_country,
    capture_date: data.capture_date,
    offer_type: data.offer_type,
    purchase_type: data.purchase_type,
    picture_quality: data.picture_quality,
    program_price: data.program_price,
    bundle_price: data.bundle_price,
    currency: data.currency,
    addon_name: data.addon_name,
    is_movie: data.is_movie,
    season_number: data.season_number,
    episode_number: data.episode_number,
    title: data.title,
    genre: data.genre,
    source_id: data.source_id,
    program_url: data.program_url,
    maturity_rating: data.maturity_rating,
    release_date: data.release_date,
    release_year: data.release_year,
    viewable_runtime: data.viewable_runtime,
    series_title: data.series_title,
    series_release_year: data.series_release_year,
    series_source_id: data.series_source_id,
    series_url: data.series_url,
    series_genre: data.series_genre,
    season_source_id: data.season_source_id
  });

  await workbook.xlsx.writeFile('disneyMovies.xlsx');

  console.log("Data is written to file.");
};