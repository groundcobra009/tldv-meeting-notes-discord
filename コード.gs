function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('tldvメール処理')
    .addItem('メールを処理', 'processEmails')
    .addSeparator()
    .addItem('行の高さを24pxに設定', 'setRowHeight')
    .addItem('対象メールアドレスを設定', 'setEmailAddress')
    .addItem('自動実行トリガーを設定', 'setupTrigger')
    .addItem('トリガーを削除', 'removeTriggers')
    .addSeparator()
    .addItem('Discord Webhook URLを設定', 'setDiscordWebhook')
    .addItem('Discordに未送信を通知', 'sendUnsentToDiscord')
    .addItem('Discord自動送信トリガーを設定', 'setupDiscordTrigger')
    .addItem('Discord自動送信トリガーを削除', 'removeDiscordTrigger')
    .addToUi();
}

function processEmails() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const targetEmail = PropertiesService.getScriptProperties().getProperty('TARGET_EMAIL');
    
    if (!targetEmail) {
      SpreadsheetApp.getUi().alert('対象メールアドレスが設定されていません。\nメニューから設定してください。');
      return;
    }
    
    const tldvLabel = GmailApp.getUserLabelByName('tldv');
    if (!tldvLabel) {
      SpreadsheetApp.getUi().alert('「tldv」ラベルが見つかりません。');
      return;
    }
    
    let processedLabel = GmailApp.getUserLabelByName('処理済み');
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel('処理済み');
    }
    
    const threads = tldvLabel.getThreads();
    
    if (threads.length === 0) {
      console.log('処理対象のメールがありません。');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    let currentRow = lastRow + 1;
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      
      messages.forEach(message => {
        const receivedDate = message.getDate();
        const subject = message.getSubject();
        let body = message.getPlainBody();
        
        // 不要な部分を削除
        const removePattern = /機能紹介[\s\S]*?ミーティングの要約の受信を停止するには、こちらから登録を解除.*?してください。/;
        body = body.replace(removePattern, '').trim();
        
        // 冒頭のURLを削除
        body = body.replace(/^\(\s*https:\/\/tldv\.io\/ja\/\s*\)\s*/, '').trim();
        
        sheet.getRange(currentRow, 1).setValue(receivedDate);
        sheet.getRange(currentRow, 2).setValue(subject);
        sheet.getRange(currentRow, 3).setValue(body);
        
        currentRow++;
      });
      
      thread.removeLabel(tldvLabel);
      thread.addLabel(processedLabel);
    });
    
    // 追加した行の高さを24ピクセルに設定
    if (currentRow > lastRow + 1) {
      for (let row = lastRow + 1; row < currentRow; row++) {
        sheet.setRowHeight(row, 24);
      }
    }
    
    SpreadsheetApp.getUi().alert(`${threads.length}件のメールを処理しました。`);
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('処理中にエラーが発生しました。\n' + error.toString());
  }
}

function setEmailAddress() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = PropertiesService.getScriptProperties().getProperty('TARGET_EMAIL') || '';
  
  const result = ui.prompt(
    'メールアドレス設定',
    '対象のメールアドレスを入力してください:\n現在の設定: ' + currentEmail,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const email = result.getResponseText().trim();
    if (email) {
      PropertiesService.getScriptProperties().setProperty('TARGET_EMAIL', email);
      ui.alert('メールアドレスを設定しました: ' + email);
    } else {
      ui.alert('メールアドレスが入力されていません。');
    }
  }
}

function setupTrigger() {
  try {
    removeTriggers();
    
    ScriptApp.newTrigger('processEmails')
      .timeBased()
      .everyHours(1)
      .create();
    
    SpreadsheetApp.getUi().alert('1時間ごとの自動実行トリガーを設定しました。');
  } catch (error) {
    SpreadsheetApp.getUi().alert('トリガーの設定に失敗しました。\n' + error.toString());
  }
}

function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function setRowHeight() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 0) {
      for (let row = 1; row <= lastRow; row++) {
        sheet.setRowHeight(row, 24);
      }
      SpreadsheetApp.getUi().alert(`${lastRow}行の高さを24ピクセルに設定しました。`);
    } else {
      SpreadsheetApp.getUi().alert('シートにデータがありません。');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('行の高さ設定中にエラーが発生しました。\n' + error.toString());
  }
}

function setDiscordWebhook() {
  const ui = SpreadsheetApp.getUi();
  const currentWebhook = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK') || '';
  
  const result = ui.prompt(
    'Discord Webhook URL設定',
    'Discord Webhook URLを入力してください:\n現在の設定: ' + (currentWebhook ? '設定済み' : '未設定'),
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const webhookUrl = result.getResponseText().trim();
    if (webhookUrl && webhookUrl.startsWith('https://discord.com/api/webhooks/')) {
      PropertiesService.getScriptProperties().setProperty('DISCORD_WEBHOOK', webhookUrl);
      ui.alert('Discord Webhook URLを設定しました。');
    } else if (webhookUrl) {
      ui.alert('有効なDiscord Webhook URLを入力してください。\nURLは https://discord.com/api/webhooks/ で始まる必要があります。');
    } else {
      ui.alert('Webhook URLが入力されていません。');
    }
  }
}

function sendUnsentToDiscord() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    
    // シート名の確認を削除して、どのシートでも動作するようにする
    console.log('現在のシート名:', sheetName);
    
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK');
    if (!webhookUrl) {
      SpreadsheetApp.getUi().alert('Discord Webhook URLが設定されていません。\nメニューから設定してください。');
      return;
    }
    
    // A列の最終行を確認（クエリ関数のデータ）
    let lastRow = 1;
    const columnA = sheet.getRange('A:A').getValues();
    for (let i = columnA.length - 1; i >= 0; i--) {
      if (columnA[i][0] !== '') {
        lastRow = i + 1;
        break;
      }
    }
    
    console.log('データの最終行:', lastRow);
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('送信対象のデータがありません。');
      return;
    }
    
    let sentCount = 0;
    
    // データを一括で取得（パフォーマンス向上）
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
    const data = dataRange.getValues();
    console.log('取得したデータ数:', data.length);
    
    let checkedCount = 0;
    let uncheckedCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const dateValue = data[i][0];
      const titleValue = data[i][1];
      const summaryValue = data[i][2];
      const checkboxValue = data[i][3];
      
      console.log(`行${row}: チェックボックス=${checkboxValue}, 日時=${dateValue}, タイトル=${titleValue}`);
      
      if (checkboxValue) {
        checkedCount++;
      } else {
        uncheckedCount++;
      }
      
      // D列がチェックされていない場合のみ送信
      if (!checkboxValue && dateValue && titleValue) {
        const formattedDate = Utilities.formatDate(new Date(dateValue), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        
        const message = {
          content: null,
          thread_name: `${titleValue} - ${formattedDate}`,
          embeds: [{
            title: titleValue,
            description: summaryValue || '概要なし',
            color: 5814783,
            fields: [{
              name: '日時',
              value: formattedDate,
              inline: true
            }],
            footer: {
              text: '東京同窓会 - tldv議事録'
            },
            timestamp: new Date().toISOString()
          }]
        };
        
        const options = {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(message),
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(webhookUrl, options);
        
        if (response.getResponseCode() === 204 || response.getResponseCode() === 200) {
          console.log(`Discord送信成功 (行 ${row}): ${titleValue}`);
          sheet.getRange(row, 4).setValue(true);
          console.log(`チェックボックスを更新しました (行 ${row})`);
          sentCount++;
          
          Utilities.sleep(1000);
        } else {
          console.error(`Discord送信エラー (行 ${row}):`, response.getContentText());
        }
      }
    }
    
    console.log(`チェック済み: ${checkedCount}件, 未チェック: ${uncheckedCount}件`);
    
    if (sentCount > 0) {
      SpreadsheetApp.getUi().alert(`${sentCount}件の通知をDiscordに送信しました。`);
    } else {
      SpreadsheetApp.getUi().alert(`送信対象の項目はありませんでした。\n（既に送信済みか、データが存在しません）\n\nデバッグ情報:\n- データ数: ${data.length}件\n- チェック済み: ${checkedCount}件\n- 未チェック: ${uncheckedCount}件`);
    }
    
  } catch (error) {
    console.error('Discord送信中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('Discord送信中にエラーが発生しました。\n' + error.toString());
  }
}

function setupDiscordTrigger() {
  try {
    removeDiscordTrigger();
    
    // 23時から24時の間のランダムな時間を生成
    const randomMinute = Math.floor(Math.random() * 60);
    
    ScriptApp.newTrigger('sendUnsentToDiscord')
      .timeBased()
      .everyDays(1)
      .atHour(23)
      .nearMinute(randomMinute)
      .create();
    
    SpreadsheetApp.getUi().alert(`Discord自動送信トリガーを設定しました。\n毎日23:${randomMinute.toString().padStart(2, '0')}頃に実行されます。`);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Discord自動送信トリガーの設定に失敗しました。\n' + error.toString());
  }
}

function removeDiscordTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendUnsentToDiscord') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  SpreadsheetApp.getUi().alert('Discord自動送信トリガーを削除しました。');
}