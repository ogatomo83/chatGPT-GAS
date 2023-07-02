const apiKey = ScriptProperties.getProperty('chatGPT_API');
const apiUrl = 'https://api.openai.com/v1/chat/completions';
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
var range = sheet.getDataRange();
var values = range.getValues();

function onOpen() {
  var UI = SpreadsheetApp.getUi();
  var MENU = UI.createMenu('chatGPT');
  MENU.addItem('ブログの作成','chatGPT');
  MENU.addToUi();
}

function chatGPT() {
  for (var i = 0; i < values.length; i++) {
    if (values[i][4] === true) {
      Logger.log(values[i]);
      let messages = [{'role': 'system', 'content': 'あなたはブログを作成しています。'}];
      messages.push({'role': 'user', 'content': "これから「キーワード、ブログタイトル、概要、参考URL」を提示しますので、簡単なブログを生成してください。"});
      messages.push({"role": "system", "content": "承知しました。「キーワード、ブログタイトル、概要、参考URL」を共有してください。"}),
      messages.push({'role': 'user', 'content': "キーワード : " + values[i][0]});
      messages.push({"role": "system", "content": "キーワードの内容を把握しました"}),
      messages.push({'role': 'user', 'content': "ブログタイトル : " + values[i][1]});
      messages.push({"role": "system", "content": "ブログタイトルの内容を把握しました"}),
      messages.push({'role': 'user', 'content': "概要 : " + values[i][2]});
      messages.push({"role": "system", "content": "概要の内容を把握しました。"}),
      messages.push({'role': 'user', 'content': "参考URL : " + values[i][3]});
      messages.push({"role": "system", "content": "参考URLを把握しました。"}),
      messages.push({'role': 'user', 'content': "簡単なブログを作成してください。"});

      const requestBody = {
        'model': 'gpt-3.5-turbo',
        'temperature': 0.7,
        'max_tokens': 1000,
        'messages': messages
      }

      const request = {
        method: "POST",
        muteHttpExceptions : true,
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + apiKey,
        },
        payload: JSON.stringify(requestBody),
      }

      try {
        //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
        const response = JSON.parse(UrlFetchApp.fetch(apiUrl, request).getContentText());
        // ChatGPTのAPIレスポンスをセルに記載
          cell = "F" + (i + 1)
          var responseCell = sheet.getRange(cell);
        if (response.choices) {
          Logger.log(response.choices[0].message.content);
          responseCell.setValue(response.choices[0].message.content);
        } else {
          responseCell.setValue('no content');
        }
        checkboxCell = E + (i + 1)
        var responseCheckboxCell = sheet.getRange(cell);
        responseCheckboxCell.setValue('false')
      } catch(e) {
        // 例外エラー処理
        console.log('error')
        console.log(e);
      }
    } else {
      continue;
    }
  }
}
      