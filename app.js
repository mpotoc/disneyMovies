const puppeteer = require('puppeteer');
const config = require('./config');
const json2xls = require('json2xls');
const fs = require('fs');

(async () => {
  const browser = await puppeteer.launch({
    /*args: [
      '--proxy-server=socks5://198.211.99.227:46437'
    ],*/
    //devtools: true,
    defaultViewport: {
      width: 1920,
      height: 1080
    }
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.0 Safari/537.36');
  page.setDefaultTimeout(0);

  //await page.waitForSelector(config.USERNAME_SELECTOR);
  await page.goto(config.url, {
    waitUntil: 'networkidle0'
  });
  await page.waitFor(config.USERNAME_SELECTOR);

  // Additional step if we start at https://www.disneyplus.com
  //await page.screenshot({path: 'screen.png'});
  //await page.click('button[kind="outline"]');

  //await page.screenshot({ path: 'screenEmail.png' });

  await page.click(config.USERNAME_SELECTOR);
  await page.keyboard.type(config.username);
  await page.click('button[name="dssLoginSubmit"]');
  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });
  //await page.screenshot({ path: 'screenLogIn.png' });

  await page.click(config.PASSWORD_SELECTOR);
  await page.keyboard.type(config.password);
  await page.click('button[name="dssLoginSubmit"]');
  await page.waitForNavigation({
    waitUntil: 'networkidle0'
  });
  //await page.screenshot({ path: 'screenLoggedIn.png' });

  /*******************************
  * Disney+ Movies part scrapper
  */
  await page.goto(config.movieUrl, {
    waitUntil: 'networkidle0'
  });
  await page.waitFor('.sc-jTzLTM');
  //await page.screenshot({ path: 'screenMovie.png' });

  await autoScroll(page);
  await page.waitFor(2000);

  //await page.screenshot({
  //  path: 'scrollCheckMovies.png',
  //  fullPage: true
  //});

  let linksMovies = await page.$$('a.sc-ckVGcZ');
  let dataMovies = [];

  for (let i = 0; i < linksMovies.length; i++) {
    await page.evaluate((i) => {
      return ([...document.querySelectorAll('a.sc-ckVGcZ')][i]).click();
    }, i);

    await page.waitFor('a.sc-ckVGcZ');

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
      let rating_step2 = rating_step1.alt.split('_');
      let rating = rating_step2[2].toUpperCase();

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

    console.log(resultMovie);
    
    await page.goBack();
  }

  await browser.close();

  var xlsMovies = json2xls(dataMovies);
  fs.appendFileSync('disneyplusMovies.xlsx', xlsMovies, 'binary');
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
      }, 100);
    });
  });
};