function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('tldvメール処理')
    .addItem('メールを処理', 'processEmails')
    .addSeparator()
    .addItem('行の高さを24pxに設定', 'setRowHeight')
    .addItem('対象メールアドレスを設定', 'setEmailAddress')
    .addItem('自動実行トリガーを設定', 'setupTrigger')
    .addItem('トリガーを削除', 'removeTriggers')
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