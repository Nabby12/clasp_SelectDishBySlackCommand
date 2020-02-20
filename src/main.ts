const SLACK_VERIFICATIONTOKEN: string = PropertiesService.getScriptProperties().getProperty('SLACK_VERIFICATIONTOKEN');
const SLACK_WEBHOOK_URL: string = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
const SPREADSHEET_ID: string = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEET1NAME: string = PropertiesService.getScriptProperties().getProperty('SHEET1NAME');

function doPost(e: string) {
    let verificationToken: string = e.parameter.token;
    if (verificationToken != SLACK_VERIFICATIONTOKEN) {
        throw new Error('Invalid Token');
    }
    
    let arg: string = e.parameter.text.trim();

    let trgtSpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trgtSh = trgtSpreadSheet.getSheetByName(SHEET1NAME);

    let dataLastRow = trgtSh.getLastRow();
    let trgtRow: number = Math.floor(Math.random() * Math.floor(dataLastRow)) + 1;
    let materials: string = trgtSh.getRange(trgtRow, 2).getValue();

    if ( arg.length > 0 ) {
        let trgtMaterial: string = arg;

        let materialCnt: number = materials.indexOf(trgtMaterial);

        if (materialCnt === -1) {
            do {
                trgtRow = Math.floor(Math.random() * Math.floor(dataLastRow)) + 1;
                materials = trgtSh.getRange(trgtRow, 2).getValue();

                materialCnt = materials.indexOf(trgtMaterial);
            } while (materialCnt === -1);
        }
    }

    let trgtDish: string = trgtSh.getRange(trgtRow, 1).getValue();

    let sendComment: string = 
    `${ trgtDish }

材料：${ materials }`
    
    PostMessageToSlack(sendComment);
    
    return ContentService.createTextOutput();
}

function PostMessageToSlack(sendBody: string) {
    let params: any = {
        method: 'post',
        contentType: 'application/json',
        payload: `{"text":'${ sendBody }'}`
    };
    
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, params);
}