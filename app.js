// Require what is needed
const puppeteer = require('puppeteer');
const chalk = require('chalk');
const readline = require('readline-sync');
const os = require('os');
const fs = require('fs');
require('dotenv').config();
const envFile = require('envfile');
const sanitize = require('sanitize-filename');

// Welcoming
console.log(chalk.yellow('Running Node.js with puppeteer. Application PDFTaker ') + chalk.yellowBright('v1.3.3'));
console.log(chalk.yellow('Developed By: [ClearBoth - Abdulla Ashoor]'));

let username = '';
let password = '';
let parsedFile = envFile.parse('.env');
let askCredentials = false;

// Check if there is login saved
if (process.env.EMAIL) {
    if (readline.keyInYN(chalk.cyanBright("\nLoad your saved account?"))) {
        username = process.env.EMAIL;
        password = process.env.PASSWORD;
        console.log(chalk.greenBright('Your Account Loaded Successfully.'));
    } else {
        askCredentials = true;
    }
}

if (!process.env.EMAIL || askCredentials) {
    username = readline.questionEMail(chalk.cyanBright("\nEnter your email account: "));
    password = readline.question(chalk.cyanBright("Enter your password: "), {hideEchoBack: true});
    while (!username || !password) {
        console.error(chalk.red.bold('Please enter your username and password in order to login to your form and get your files.'));
        username = readline.questionEMail(chalk.cyanBright("\nEnter your email account: "));
        password = readline.question(chalk.cyanBright("Enter your password: "), {hideEchoBack: true});
    }

    if (readline.keyInYN(chalk.cyanBright("\nDo you want to save your credentials?"))) {

        parsedFile.EMAIL = username;
        parsedFile.PASSWORD = password;
        fs.writeFileSync('./.env', envFile.stringify(parsedFile));
    }
}

let formURL = readline.question(chalk.cyanBright("\nEnter your form URL: "));

while (!formURL || (!formURL.includes("FormId=") && !formURL.includes("id="))) {
    console.error(chalk.red.bold('Please enter a correct microsoft form url.'));
    formURL = readline.question(chalk.cyanBright("\nEnter your form URL: "));
}

let form_id = '';
let app = 0;
let id_type = 0;

// get the id
if (formURL.includes("FormId=")) {
    app = formURL.indexOf("FormId=");
    id_type = 7;
} else {
    app = formURL.indexOf("id=");
    id_type = 3;
}

form_id = formURL.substring(app);
app = form_id.indexOf("&");
if (app === -1)
    form_id = form_id.substring(id_type);
else
    form_id = form_id.substring(id_type, app);

// build the urls
let formReviewURL = "https://forms.office.com/Pages/DesignPage.aspx?" +
    "auth_pvr=OrgId&lang=en-US&origin=OfficeDotCom&route=Start#Analysis=true&FormId=" +
    form_id + "&Grading=%7B%22View%22%3A%22Student%22%2C%22Index%22%3A0%2C%22GoBackToView%22%3A%22GoBackToAnalysis%22%7D" +
    "&TopView=Grading_ReviewAnswers"

let formStatusURL = "https://forms.office.com/Pages/DesignPage.aspx?" +
    "auth_pvr=OrgId&lang=en-US&origin=OfficeDotCom&route=Start#Analysis=true&FormId=" +
    form_id + "&Grading=%7B%22View%22%3A%22Student%22%2C%22Index%22%3A0%2C%22GoBackToView%22%3A%22GoBackToAnalysis%22%7D";

// Ask user for classes folders
console.log(chalk.cyan("\nMake sure you have student class field as 1st multiple choice question, or simply choose n"));
let studentClass = readline.keyInYN(chalk.cyanBright("Create folder for each class: "));

// Ask user about the file name
console.log(chalk.cyanBright("\nChoose how do you want to display the file name:"));
console.log(chalk.magentaBright("1) use " + chalk.blueBright('respondent number/Account name') + "."));
console.log(chalk.magentaBright("2) use " + chalk.blueBright('serial numbers') + "/don't use form fields."));
console.log(chalk.magentaBright("3) use " + chalk.blueBright('first field') + " (student name) only."));
console.log(chalk.magentaBright("4) use " + chalk.blueBright('second-first field') + " (id-student name) only."));
console.log(chalk.magentaBright("5) use " + chalk.blueBright('second-first-third field') + " (id-student name-class)."));
let userOption = readline.question(chalk.cyanBright("Enter your choice number: "));
while (!userOption || isNaN(parseInt(userOption)) || (userOption != 1 && userOption != 2 && userOption != 3 && userOption != 4 && userOption != 5)) {
    console.error(chalk.red.bold('Please enter an option 1, 2, 3, 4, or 5.'));
    userOption = readline.question(chalk.cyanBright("Enter your choice number: "));
}

console.log(chalk.cyan("\nAdd prefix number in front of the file name, or just skip it with Enter"));
let prefix = readline.question(chalk.cyanBright("Enter a number to be added before the file name: "));
if(parseInt(prefix)) {
    prefix = parseInt(prefix);
    prefix = prefix+'_';
} else {
    prefix = '';
}

let seconds = 0;
let slowSpeed = false;
console.log(chalk.cyan("\nIf you're getting errors, try to increase the time between saving files"));
console.log(chalk.cyan("Just insert the number of seconds you want or put 0"));
seconds = readline.question(chalk.cyanBright("How many seconds to wait: "));
while (!seconds === 0 || isNaN(parseFloat(seconds))) {
    console.error(chalk.red.bold('Total seconds must be a number. Try 3 for example.'));
    seconds = readline.question(chalk.cyanBright("How many seconds to wait: "));
}
if (seconds)
    slowSpeed = true;

function isEven(n) {
    return n % 2 === 0;
}

let name = '';
let studentID = '';
let formTitle = '';
let checkedOption = '';
let respondent = '';
let fileName = '';
let dir = '';
let responses = 0;
let desktop = os.homedir() + "\\Desktop";

(async () => {
    const browser = await puppeteer.launch({
        headless: true,
        executablePath: './node_modules/puppeteer/.local-chromium/win64-818858/chrome-win/chrome.exe'
    });
    const [page] = await browser.pages();

    await page.goto(formStatusURL, {waitUntil: 'networkidle0', timeout: 0})

    await page.waitForTimeout(1000);

    console.log(chalk.yellow('\nLogging in.. Please wait..'));

    await page.type('input[type=email]', username);
    await page.click('#idSIButton9');

    await page.waitForNavigation();

    await page.waitForTimeout(1000);
    await page.type('input[type=password]', password);
    await page.click('#idSIButton9');

    await page.waitForNavigation();

    await page.click('#idSIButton9');

    await page.waitForNavigation();
    await page.waitForTimeout(4000);
    if (slowSpeed)
        await page.waitForTimeout(seconds * 1000);

    // Get total number of responses
    responses = parseInt(await page.$eval('div.simple-data-control-content.office-form-theme-primary-foreground', e => e.innerText));

    if (responses) {
        console.log(chalk.greenBright('Logged In Successfully\n'));
    } else {
        console.log(chalk.redBright('Logging Failed/Form is Incorrect.\n'));
        process.exit(1);
    }

    // Go to review page
    await page.waitForTimeout(1000);
    await page.goto(formReviewURL, {waitUntil: 'networkidle0', timeout: 0})
    await page.waitForTimeout(1000);

    // Get the form name
    formTitle = await page.$eval('div.gradingTitle.office-form-theme-primary-foreground span', e => e.innerText);
    if (formTitle.includes("Review:"))
        formTitle = sanitize(formTitle.substring(7));

    if (formTitle.length > 255)
        formTitle = sanitize(formTitle.substring(0, 200));

    console.log(chalk.yellow('A folder created in your desktop with name:') + chalk.blue.underline.bold(formTitle));
    console.log(chalk.yellow('Start saving files..'));
    // start taking the full page pdf
    for (let i = 0; i < responses; i++) {

        await page.waitForTimeout(100);

        // get the student class name
        try {
            checkedOption = await page.$eval('div.gradingByStudentAnswerContainer.gradingAnswerContainer div', e => e.innerHTML);
            checkedOption = checkedOption.substring(checkedOption.indexOf("checked="));
            checkedOption = checkedOption.substring(62, checkedOption.indexOf("</span>"));
        } catch (error) {
            // just ignore and continue.
        }

        if (checkedOption === 'No answer provided.')
            checkedOption = 'Unsorted';

        // create directory
        if (studentClass) {
            // create the class folder if not exists
            dir = desktop + '\\' + formTitle + '\\' + sanitize(checkedOption);
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, {
                    recursive: true
                });
            }
        } else {
            // create the folder if not exists
            dir = desktop + '\\' + formTitle;
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, {
                    recursive: true
                });
            }
        }

        respondent = await page.$eval('span.select-placeholder-text span span', e => e.innerText);
        name = await page.evaluate(() => document.querySelectorAll("div.gradingByStudentAnswerContainer.gradingAnswerContainer span.gradingPlainTextAnswer span")[0].innerHTML);
        studentID = await page.evaluate(() => document.querySelectorAll("div.gradingByStudentAnswerContainer.gradingAnswerContainer span.gradingPlainTextAnswer span")[1].innerHTML);

        //Set file name as requested choice
        switch (userOption) {
            case '1':
                fileName = sanitize(prefix + respondent);
                break;
            case '2':
                fileName = sanitize(prefix + (i + 1));
                break;
            case '3':
                fileName = sanitize(prefix + name);
                break;
            case '4':
                fileName = sanitize(prefix + studentID + '-' + name);
                break;
            case '5':
                fileName = sanitize(prefix + studentID + '-' + name + '-' + checkedOption);
                break;
        }

        // emulate print page
        await page.click('button.menu-control-trigger.analyze-view-ellipsis-button.button-control.light-background-button');
        await page.keyboard.press('ArrowDown');
        await page.waitForTimeout(50);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(50);


        //take full page pdf
        await page.emulateMediaType('print');
        await page.pdf({
            path: dir + '/' + fileName + '.pdf',
            format: 'A4',
            printBackground: true,
            displayHeaderFooter: false,
        });

        // inform me in console
        if (isEven(i)) {
            console.log(chalk.cyan((i + 1) + ') ' + fileName + ' PDF saved.'));
        } else {
            console.log(chalk.cyanBright((i + 1) + ') ' + fileName + ' PDF saved.'));
        }

        // cancel chromium print button
        //await page.waitForTimeout(100);
        await page.keyboard.press('Escape');
        if (slowSpeed)
            await page.waitForTimeout(seconds * 500);
        await page.waitForTimeout(500);
        await page.keyboard.press('Escape');
        await page.emulateMediaType('screen');
        if (slowSpeed)
            await page.waitForTimeout(seconds * 500);
        await page.waitForTimeout(500);

        // Next student
        await page.click('button.gradingNavBarArrow.nextQuestion.button-control');
    }

    console.log(chalk.greenBright('\n' + responses + ' responses saved successfully.'));
    await browser.close();
})();

process.on('unhandledRejection', error => {
    console.log(chalk.red.bold("Something Went Wrong! Please try again."));
    console.log(chalk.blueBright.bold("Make sure you have the right login information and the correct form URL."));
    console.log('\nShow the following error to the support team so they can help you.')
    console.log(error);
    process.exit(1);
});