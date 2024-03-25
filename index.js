import { GoogleSpreadsheet }  from 'google-spreadsheet'
import { WebClient } from '@slack/web-api';
import schedule from 'node-schedule';
import {JWT} from "google-auth-library";
import 'dotenv/config'


console.log('start')
const serviceAccountAuth = new JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
    ],
});

/**
 * Выбираем рандомную ячейку которая содержит сообщение
 * @return {Promise<number|boolean|string|GoogleSpreadsheetCellErrorValue>}
 */
async function getRandomRow() {
    console.log('start getRandomRow')
    const doc = new GoogleSpreadsheet(process.env.SPREAD_SHEET_ID, serviceAccountAuth);

    await doc.loadInfo(); // загружает свойства документа и рабочие листы
    const sheet = doc.sheetsByIndex[0]; // берем  данные которые находятся на первом листе

    await sheet.loadCells(); // загружаем все ячейки

    const columnNumber = 1; // Номер столбца у которого вычисляем количество заполненых ячеек
    let filledCellCount = null; // заполненые ячейки

    // Перебираем ячейки в указанном столбце и подсчитываем заполненые ячейки
    for (let i = 0; i < sheet.rowCount; i++) {
        const cell = sheet.getCell(i, columnNumber - 1);
        if (cell.value !== null && cell.value !== '') {
            filledCellCount++;
        }
    }

    const randomRowIndex = Math.floor(Math.random() * filledCellCount);
    const cell = await sheet.getCell(randomRowIndex,0);
    console.log('Количество заполенных ячеек',filledCellCount)
    console.log('Выбранная ячейка',randomRowIndex)
    return cell.value;
}


async function postToSlack(message) {
    const slackClient = new WebClient(process.env.SLACK_API_TOKEN);
    try {
        const result = await slackClient.chat.postMessage({
            channel: process.env.CHANNEL_NAME,
            text: message,
        });
        console.log('Message sent successfully:', {result: result.ts, message: message, channel: process.env.CHANNEL_NAME});
    } catch (error) {
        console.error('Error posting message:', error);
    }
}


async function main() {
    const message = await getRandomRow();


    await postToSlack(message);
}

// main().then(() => {}).catch((e) =>{console.log(e)})
/**
 * *    *    *    *    *    *
 * ┬    ┬    ┬    ┬    ┬    ┬
 * │    │    │    │    │    |
 * │    │    │    │    │    └ day of week (0 - 7) (0 or 7 is Sun)
 * │    │    │    │    └───── month (1 - 12)
 * │    │    │    └────────── day of month (1 - 31)
 * │    │    └─────────────── hour (0 - 23)
 * │    └──────────────────── minute (0 - 59)
 * └───────────────────────── second (0 - 59, OPTIONAL)
 */

schedule.scheduleJob("*/1 * * * 1-5", main);