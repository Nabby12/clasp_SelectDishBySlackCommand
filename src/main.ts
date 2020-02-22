const SLACK_VERIFICATIONTOKEN: string = PropertiesService.getScriptProperties().getProperty('SLACK_VERIFICATIONTOKEN');
const SLACK_WEBHOOK_URL: string = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
const SPREADSHEET_ID: string = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEET1NAME: string = PropertiesService.getScriptProperties().getProperty('SHEET1NAME');
const SENDCOMMENTSTR1: string = PropertiesService.getScriptProperties().getProperty('SENDCOMMENTSTR1');

function doPost(e: string) {
    let verificationToken: string = e.parameter.token;
    if (verificationToken != SLACK_VERIFICATIONTOKEN) {
        throw new Error('Invalid Token');
    }
    
    let arg: string = e.parameter.text.trim();

    let trgtSpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trgtSh = trgtSpreadSheet.getSheetByName(SHEET1NAME);

    let dataLastRow = trgtSh.getLastRow();
    let trgtRng = trgtSh.getRange(1, 1, dataLastRow, 2);
    let trgtAry: string[] = trgtRng.getValues();
    let trgtRowIndex: number = Math.floor(Math.random() * Math.floor(dataLastRow));
    
    // 配列のインデックスは「0」から始まるため「-1」
    const dishColIndex: number = 1 - 1;
    const materialColIndex: number = 2 - 1;

    let materials: string = trgtAry[trgtRowIndex][materialColIndex];

    if ( arg.length > 0 ) {
        let trgtMaterial: string = arg;

        let materialCnt: number = materials.indexOf(trgtMaterial);

        if (materialCnt === -1) {
            do {
                trgtRowIndex = Math.floor(Math.random() * Math.floor(dataLastRow));
                materials = trgtAry[trgtRowIndex][materialColIndex];

                materialCnt = materials.indexOf(trgtMaterial);
            } while (materialCnt === -1);
        }
    }

    let trgtDish: string = trgtAry[trgtRowIndex][dishColIndex];

    let sendComment: string = 
    `${ trgtDish }

${ SENDCOMMENTSTR1}：${ materials }`
    
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