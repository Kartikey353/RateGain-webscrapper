const { Builder, By } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { Workbook } = require('exceljs');
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const progressBar = (percentage) => {
    const progressBarLength = 30;
    const completed = Math.round((progressBarLength * percentage) / 100);
    const remaining = progressBarLength - completed;
    return `[${'='.repeat(completed)}>${' '.repeat(remaining)}] ${percentage}%`;
};


const getArticleDataFromPage = async (driver, pageUrl, pageNumber) => {
    await driver.get(pageUrl);
    let result2 = await driver.findElements(By.css('.blog-item.category-blog'));
    let articleData = [];
    for (const item of result2) {
        let blogTitle, blogDate, blogImageUrl, blogLikesCount;
        let hasImageurl = false;
        let hasblogtitle = false;
        let hasblogdate = false;

        try {
            blogTitle = await item.findElement(By.css('.wrap .content h6')).getText();
            hasblogtitle = !!blogTitle;
        } catch (error) {
            blogTitle = '';
            hasblogtitle = false;
        }

        try {
            blogDate = await item.findElement(By.css('.wrap .content .blog-detail .bd-item')).getText();
            hasblogdate = !!blogDate;
        } catch (error) {
            blogDate = '';
            hasblogdate = false;
        }

        try {
            let blogImageElement = await item.findElement(By.css('.wrap .img a'));
            blogImageUrl = await blogImageElement.getAttribute('data-bg');
            hasImageurl = !!blogImageUrl;
        } catch (error) {
            blogImageUrl = '';
            hasImageurl = false;
        }

        try {
            blogLikesCount = await item.findElement(By.css('.wrap .content .zilla-likes')).getText();
        } catch (error) {
            blogLikesCount = '';
        }

        articleData.push({
            blogTitle: blogTitle || '',
            blogDate: blogDate || '',
            blogImageUrl: blogImageUrl || '',
            blogLikesCount: blogLikesCount || '',
            blogpageNumber: pageNumber,
            hasImageurl: hasImageurl,
            hasblogtitle: hasblogtitle,
            hasblogdate: hasblogdate
        });
    }
    return articleData;
};
const scrapeAllPages = async () => {
    let chromeOptions = new chrome.Options();
    chromeOptions.addArguments('--headless');
    let driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    try {
        await driver.get('https://rategain.com/blog/');

        let totalPagesElement = await driver.findElement(By.css('.wpb_wrapper .pagination a.page-numbers:nth-last-child(2)'));
        let totalPagesText = await totalPagesElement.getText();
        let totalPages = parseInt(totalPagesText, 10);
        // let totalPages = 5;
        let allArticleData = [];
        console.log(`Start fetching the articles from total ${totalPages} pages`)
        let percentageCompleted = 0;
        for (let pageNumber = 1; pageNumber <= totalPages; pageNumber++) {
            let pageUrl = `https://rategain.com/blog/page/${pageNumber}/`;
            let articleData = await getArticleDataFromPage(driver, pageUrl, pageNumber);
            allArticleData = allArticleData.concat(articleData);
            percentageCompleted = Math.round((pageNumber / totalPages) * 100);
            process.stdout.clearLine();
            process.stdout.cursorTo(0);
            process.stdout.write(`Progress: ${progressBar(percentageCompleted)}\r`);
            await sleep(100);
        }

        console.log(`\nData fetching complete! Total articles fetch from ${totalPages} pages are ${allArticleData.length}`);

        const workbook = new Workbook();
        const worksheet1 = workbook.addWorksheet('With-Image-URL-Other-Fields');
        const worksheet2 = workbook.addWorksheet('Without-Image-URL-Other-Fields');
        const worksheet3 = workbook.addWorksheet('With-More-Empty-Values');
        const withImageUrlAndOtherFields = allArticleData.filter(
            article => article.blogImageUrl !== '' && article.hasblogtitle && article.hasblogdate
        );

        const withoutImageUrlButOtherFields = allArticleData.filter(
            article => article.blogImageUrl === '' && article.hasblogtitle && article.hasblogdate
        );

        const withTwoOrMoreEmptyValues = allArticleData.filter(article => {
            const emptyValuesCount = Object.values(article).filter(value => value === '').length;
            return emptyValuesCount >= 2 && article.hasImageurl && article.hasblogtitle && article.hasblogdate;
        });


        const headers = ['Blog Title', 'Blog Date', 'Blog Image URL', 'Blog Likes Count', 'Blog Page Number'];
        worksheet1.addRow(headers);
        worksheet2.addRow(headers);
        worksheet3.addRow(headers);

        withImageUrlAndOtherFields.forEach(article => {
            worksheet1.addRow(Object.values(article));
        });

        withoutImageUrlButOtherFields.forEach(article => {
            worksheet2.addRow(Object.values(article));
        });

        withTwoOrMoreEmptyValues.forEach(article => {
            worksheet3.addRow(Object.values(article));
        });

        await workbook.xlsx.writeFile('categorized_articles.xlsx');
        console.log('Data successfully written to categorized_articles.xlsx');
        console.log(allArticleData);
    } catch (error) {
        console.error('Error occurred:', error);
    } finally {
        await driver.quit();
    }
};

scrapeAllPages();
