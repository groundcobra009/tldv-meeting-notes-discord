function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // デフォルトのWebhook URLを設定（初回のみ）
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('DEFAULT_DISCORD_WEBHOOK')) {
    properties.setProperty('DEFAULT_DISCORD_WEBHOOK', 'https://discord.com/api/webhooks/YOUR_DEFAULT_WEBHOOK_HERE');
  }
  
  ui.createMenu('📋 tldv')
    .addItem('📧 メール処理', 'processEmails')
    .addItem('🔀 キーワード分岐', 'processKeywordBranching')
    .addItem('📤 Discord通知', 'sendUnsentToDiscord')
    .addItem('📢 分岐Discord通知', 'sendBranchDiscordNotifications')
    .addSeparator()
    .addItem('⚙️ 設定', 'showSettingsDialog')
    .addItem('📐 分岐シート書式設定', 'formatBranchSheets')
    .addSeparator()
    .addItem('🔍 設定確認', 'checkSettings')
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
        let subject = message.getSubject();
        
        // 件名から「」と「のミーティングノートが準備できました」を除去
        subject = subject.replace(/^「/, '').replace(/」のミーティングノートが準備できました$/, '');
        
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
    
    // 追加した行の高さを24ピクセルに強制設定
    if (currentRow > lastRow + 1) {
      // setRowHeightsForcedを使用して強制的に設定
      sheet.setRowHeightsForced(lastRow + 1, currentRow - lastRow - 1, 24);
      
      // 念のため個別にも設定
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
      // まず全体の行の高さを自動調整を無効にして設定
      sheet.setRowHeightsForced(1, lastRow, 24);
      
      // 個別に各行を強制的に24ピクセルに設定
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

function forceRowHeightAndDisableAutoResize() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 0) {
      // テキストの折り返しを無効にして、内容による自動リサイズを防ぐ
      const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
      range.setWrap(false);
      
      // 全ての行を強制的に24ピクセルに設定
      sheet.setRowHeightsForced(1, lastRow, 24);
      
      SpreadsheetApp.getUi().alert(`${lastRow}行の高さを24ピクセルに固定しました。\n（テキストの折り返しも無効化しました）`);
    } else {
      SpreadsheetApp.getUi().alert('シートにデータがありません。');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('行の高さ固定中にエラーが発生しました。\n' + error.toString());
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
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = spreadsheet.getSheets()[0];
    
    // デフォルトのWebhook URLを取得
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('DEFAULT_DISCORD_WEBHOOK');
    if (!webhookUrl || webhookUrl.includes('YOUR_DEFAULT_WEBHOOK_HERE')) {
      SpreadsheetApp.getUi().alert('デフォルトのDiscord Webhook URLが設定されていません。\n設定メニューから設定してください。');
      return;
    }
    
    const lastRow = sheet1.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('送信対象のデータがありません。');
      return;
    }
    
    // E列まで確保
    const maxCol = Math.max(sheet1.getLastColumn(), 5);
    const dataRange = sheet1.getRange(2, 1, lastRow - 1, maxCol);
    const data = dataRange.getValues();
    
    let sentCount = 0;
    let logEntries = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const date = data[i][0];
      const title = data[i][1];
      const summary = data[i][2];
      const discordSent = data[i][4]; // E列
      
      // E列がfalseまたは空の場合のみ送信
      if (!discordSent && date && title) {
        const formattedDate = Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        
        const message = {
          content: null,
          thread_name: `${title} - ${formattedDate}`,
          embeds: [{
            title: title,
            description: summary || '概要なし',
            color: 3447003,
            fields: [{
              name: '日時',
              value: formattedDate,
              inline: true
            }, {
              name: 'シート',
              value: 'メイン（全体）',
              inline: true
            }],
            footer: {
              text: 'tldv議事録 - 全体通知'
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
          console.log(`Discord送信成功 (行${row}): ${title}`);
          sheet1.getRange(row, 5).setValue(true); // E列にチェック
          sentCount++;
          
          // ログエントリを追加
          logEntries.push({
            timestamp: new Date(),
            title: title,
            status: '成功',
            webhookUrl: webhookUrl.substring(0, 50) + '...'
          });
          
          Utilities.sleep(1000);
        } else {
          console.error(`Discord送信エラー (行${row}):`, response.getContentText());
          logEntries.push({
            timestamp: new Date(),
            title: title,
            status: 'エラー',
            error: response.getContentText()
          });
        }
      }
    }
    
    // 送信ログを記録
    if (logEntries.length > 0) {
      createSendLog(logEntries);
    }
    
    if (sentCount > 0) {
      SpreadsheetApp.getUi().alert(`${sentCount}件の通知をDiscordに送信しました。\n\n送信ログを確認できます。`);
    } else {
      SpreadsheetApp.getUi().alert('送信対象の項目はありませんでした。\n（既に送信済みか、データが存在しません）');
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

function processKeywordBranching() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = spreadsheet.getSheets()[0];
    const settingSheet = spreadsheet.getSheetByName('設定');
    
    if (!settingSheet) {
      SpreadsheetApp.getUi().alert('「設定」シートが見つかりません。');
      return;
    }
    
    const settingData = settingSheet.getDataRange().getValues();
    const keywordRules = [];
    
    console.log('設定シートデータ:', settingData);
    
    for (let i = 1; i < settingData.length; i++) {
      const keyword = settingData[i][1];  // 列B（キーワード）
      const targetSheet = settingData[i][4];  // 列E（シート名）
      console.log(`行${i}: キーワード="${keyword}", シート名="${targetSheet}"`);
      if (keyword && targetSheet) {
        keywordRules.push({
          regex: new RegExp(keyword, 'i'),
          sheetName: targetSheet,
          keyword: keyword
        });
      }
    }
    
    console.log('作成されたルール:', keywordRules);
    
    if (keywordRules.length === 0) {
      SpreadsheetApp.getUi().alert('設定シートにキーワードルールが設定されていません。');
      return;
    }
    
    const lastRow = sheet1.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('処理対象のデータがありません。');
      return;
    }
    
    const maxCol = Math.max(sheet1.getLastColumn(), 4);
    const dataRange = sheet1.getRange(2, 1, lastRow - 1, maxCol);
    const data = dataRange.getValues();
    let processedCount = 0;
    let debugInfo = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const date = data[i][0];
      const subject = data[i][1];
      const body = data[i][2];
      const isProcessed = data[i][maxCol - 1];
      
      if (isProcessed) {
        debugInfo.push(`行${row}: 既に処理済み`);
        continue;
      }
      
      const searchText = subject + ' ' + body;
      let matched = false;
      
      for (const rule of keywordRules) {
        if (rule.regex.test(searchText)) {
          console.log(`マッチ発見: 行${row}, キーワード="${rule.keyword}", 件名="${subject}"`);
          
          let targetSheet = spreadsheet.getSheetByName(rule.sheetName);
          if (!targetSheet) {
            targetSheet = spreadsheet.insertSheet(rule.sheetName);
            console.log(`新しいシートを作成: ${rule.sheetName}`);
            
            // ヘッダーを設定
            targetSheet.getRange(1, 1).setValue('日時');
            targetSheet.getRange(1, 2).setValue('件名');
            targetSheet.getRange(1, 3).setValue('概要');
            targetSheet.getRange(1, 4).setValue('送信済み');
            
            // ヘッダー行を太字に
            targetSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
            
            // 行の高さを24ピクセルに設定
            targetSheet.setRowHeight(1, 24);
          }
          
          const targetLastRow = targetSheet.getLastRow();
          const targetRow = targetLastRow + 1;
          
          targetSheet.getRange(targetRow, 1).setValue(date);
          targetSheet.getRange(targetRow, 2).setValue(subject);
          targetSheet.getRange(targetRow, 3).setValue(body);
          
          // 新しい行の高さを24ピクセルに設定
          targetSheet.setRowHeight(targetRow, 24);
          
          // 折り返しを無効にする
          targetSheet.getRange(targetRow, 1, 1, 3).setWrap(false);
          
          sheet1.getRange(row, maxCol).setValue(true);
          processedCount++;
          matched = true;
          debugInfo.push(`行${row}: "${rule.keyword}"でマッチ → ${rule.sheetName}に送信`);
          break;
        }
      }
      
      if (!matched) {
        debugInfo.push(`行${row}: マッチなし - "${subject}"`);
      }
    }
    
    console.log('デバッグ情報:', debugInfo);
    
    let alertMessage = `${processedCount}件の行を分岐シートに送信しました。`;
    if (processedCount === 0) {
      alertMessage += '\n\nデバッグ情報:\n' + debugInfo.slice(0, 5).join('\n');
    }
    
    SpreadsheetApp.getUi().alert(alertMessage);
    
  } catch (error) {
    console.error('キーワード分岐処理中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('処理中にエラーが発生しました。\n' + error.toString());
  }
}

function showSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '設定メニュー',
    '設定したい項目を選択してください:\n\n' +
    '1. メールアドレス設定\n' +
    '2. デフォルトDiscord Webhook設定\n' +
    '3. 自動実行トリガー設定\n' +
    '4. 行の高さ設定',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result === ui.Button.OK) {
    const choice = ui.prompt('設定項目選択', '番号を入力してください (1-4):', ui.ButtonSet.OK_CANCEL);
    
    if (choice.getSelectedButton() === ui.Button.OK) {
      const num = choice.getResponseText().trim();
      
      switch(num) {
        case '1':
          setEmailAddress();
          break;
        case '2':
          setDefaultDiscordWebhook();
          break;
        case '3':
          showTriggerMenu();
          break;
        case '4':
          setRowHeight();
          break;
        default:
          ui.alert('無効な番号です。');
      }
    }
  }
}

function showTriggerMenu() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'トリガー設定',
    '実行したい操作を選択してください:\n\n' +
    '1. メール処理トリガー設定\n' +
    '2. Discord自動送信トリガー設定\n' +
    '3. 全トリガー削除',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result === ui.Button.OK) {
    const choice = ui.prompt('トリガー操作選択', '番号を入力してください (1-3):', ui.ButtonSet.OK_CANCEL);
    
    if (choice.getSelectedButton() === ui.Button.OK) {
      const num = choice.getResponseText().trim();
      
      switch(num) {
        case '1':
          setupTrigger();
          break;
        case '2':
          setupDiscordTrigger();
          break;
        case '3':
          removeTriggers();
          removeDiscordTrigger();
          break;
        default:
          ui.alert('無効な番号です。');
      }
    }
  }
}

function checkSettings() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = spreadsheet.getSheetByName('設定');
    
    if (!settingSheet) {
      SpreadsheetApp.getUi().alert('「設定」シートが見つかりません。設定シートを作成してください。');
      return;
    }
    
    const settingData = settingSheet.getDataRange().getValues();
    let settingInfo = '設定シートの内容:\n\n';
    
    if (settingData.length === 0) {
      settingInfo += 'データがありません。';
    } else {
      for (let i = 0; i < settingData.length; i++) {
        const keyword = settingData[i][0] || '(空)';
        const targetSheet = settingData[i][1] || '(空)';
        settingInfo += `行${i + 1}: キーワード="${keyword}" → シート名="${targetSheet}"\n`;
      }
    }
    
    // データの最初の数行も確認
    const sheet1 = spreadsheet.getSheets()[0];
    const lastRow = Math.min(sheet1.getLastRow(), 5);
    settingInfo += '\n\nシート1のサンプルデータ:\n';
    
    if (lastRow > 1) {
      const sampleData = sheet1.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < sampleData.length; i++) {
        const subject = sampleData[i][1] || '(空)';
        settingInfo += `行${i + 2}: "${subject}"\n`;
      }
    } else {
      settingInfo += 'データがありません。';
    }
    
    SpreadsheetApp.getUi().alert(settingInfo);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('設定確認中にエラーが発生しました。\n' + error.toString());
  }
}

function sendBranchDiscordNotifications() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = spreadsheet.getSheetByName('設定');
    
    if (!settingSheet) {
      SpreadsheetApp.getUi().alert('「設定」シートが見つかりません。');
      return;
    }
    
    // 設定シートからWebhook情報を取得
    const settingData = settingSheet.getDataRange().getValues();
    const webhookMap = {};
    
    console.log('Discord通知設定を読み込み中...');
    
    for (let i = 1; i < settingData.length; i++) {
      const keyword = settingData[i][1];  // 列B（キーワード）
      const webhookUrl = settingData[i][2];  // 列C（Webhook URL）
      const sendFormat = settingData[i][3];  // 列D（送信形式）
      const sheetName = settingData[i][4];  // 列E（シート名）
      
      if (keyword && webhookUrl && sheetName && !webhookUrl.includes('YOUR_WEBHOOK_HERE')) {
        webhookMap[sheetName] = {
          webhookUrl: webhookUrl,
          sendFormat: sendFormat,
          keyword: keyword
        };
        console.log(`設定を追加: ${sheetName} → ${webhookUrl.substring(0, 50)}...`);
      }
    }
    
    if (Object.keys(webhookMap).length === 0) {
      SpreadsheetApp.getUi().alert('有効なWebhook URLが設定されていません。\n設定シートの列C（Webhook URL）を確認してください。');
      return;
    }
    
    let totalSent = 0;
    let processedSheets = [];
    
    // 全てのシートを確認して、設定にあるものを処理
    const allSheets = spreadsheet.getSheets();
    
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      
      // 設定シートやシート1は除外
      if (sheetName === '設定' || sheet === spreadsheet.getSheets()[0]) continue;
      
      // このシート用の設定があるか確認
      const webhookInfo = webhookMap[sheetName];
      if (!webhookInfo) {
        console.log(`${sheetName}の設定が見つかりません`);
        continue;
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        console.log(`${sheetName}にデータがありません`);
        continue;
      }
      
      // D列の最大列を確保
      const maxCol = Math.max(sheet.getLastColumn(), 4);
      const dataRange = sheet.getRange(2, 1, lastRow - 1, maxCol);
      const data = dataRange.getValues();
      
      let sheetSentCount = 0;
      
      for (let i = 0; i < data.length; i++) {
        const row = i + 2;
        const date = data[i][0];
        const title = data[i][1];
        const summary = data[i][2];
        const discordSent = data[i][3]; // D列
        
        // D列がfalseまたは空の場合のみ送信
        if (!discordSent && date && title) {
          const formattedDate = Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
          
          let message;
          
          if (webhookInfo.sendFormat === 'フォーラム') {
            message = {
              content: null,
              thread_name: `${title} - ${formattedDate}`,
              embeds: [{
                title: title,
                description: summary || '概要なし',
                color: 5814783,
                fields: [{
                  name: '日時',
                  value: formattedDate,
                  inline: true
                }, {
                  name: 'カテゴリ',
                  value: webhookInfo.keyword,
                  inline: true
                }],
                footer: {
                  text: `${sheetName} - tldv議事録`
                },
                timestamp: new Date().toISOString()
              }]
            };
          } else {
            message = {
              embeds: [{
                title: title,
                description: summary || '概要なし',
                color: 5814783,
                fields: [{
                  name: '日時',
                  value: formattedDate,
                  inline: true
                }, {
                  name: 'カテゴリ',
                  value: webhookInfo.keyword,
                  inline: true
                }],
                footer: {
                  text: `${sheetName} - tldv議事録`
                },
                timestamp: new Date().toISOString()
              }]
            };
          }
          
          const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(message),
            muteHttpExceptions: true
          };
          
          const response = UrlFetchApp.fetch(webhookInfo.webhookUrl, options);
          
          if (response.getResponseCode() === 204 || response.getResponseCode() === 200) {
            console.log(`Discord送信成功 (${sheetName} 行${row}): ${title}`);
            sheet.getRange(row, 4).setValue(true);
            sheetSentCount++;
            totalSent++;
            
            Utilities.sleep(1000);
          } else {
            console.error(`Discord送信エラー (${sheetName} 行${row}):`, response.getContentText());
          }
        }
      }
      
      if (sheetSentCount > 0) {
        processedSheets.push(`${sheetName}: ${sheetSentCount}件`);
      } else if (webhookInfo) {
        console.log(`${sheetName}: 送信対象なし`);
      }
    }
    
    if (totalSent > 0) {
      SpreadsheetApp.getUi().alert(`${totalSent}件の通知をDiscordに送信しました。\n\n詳細:\n${processedSheets.join('\n')}`);
    } else {
      SpreadsheetApp.getUi().alert('送信対象の項目はありませんでした。\n（既に送信済みか、データが存在しません）');
    }
    
  } catch (error) {
    console.error('分岐Discord送信中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('分岐Discord送信中にエラーが発生しました。\n' + error.toString());
  }
}

function formatBranchSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = spreadsheet.getSheetByName('設定');
    
    if (!settingSheet) {
      SpreadsheetApp.getUi().alert('「設定」シートが見つかりません。');
      return;
    }
    
    // 設定シートからシート名を取得
    const settingData = settingSheet.getDataRange().getValues();
    const sheetNames = [];
    
    for (let i = 1; i < settingData.length; i++) {
      const sheetName = settingData[i][4];  // 列E（シート名）
      if (sheetName) {
        sheetNames.push(sheetName);
      }
    }
    
    let formattedCount = 0;
    
    // 各分岐シートを書式設定
    for (const sheetName of sheetNames) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) {
        // 空のシートにヘッダーを追加
        sheet.getRange(1, 1).setValue('日時');
        sheet.getRange(1, 2).setValue('件名');
        sheet.getRange(1, 3).setValue('概要');
        sheet.getRange(1, 4).setValue('送信済み');
        sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        sheet.setRowHeight(1, 24);
      }
      
      if (lastRow > 0) {
        // 全行の高さを24ピクセルに強制設定
        sheet.setRowHeightsForced(1, lastRow, 24);
        
        // 全体の折り返しを無効化
        const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
        range.setWrap(false);
        
        // ヘッダー行を太字に
        sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        
        formattedCount++;
        console.log(`${sheetName}の書式を設定しました`);
      }
    }
    
    SpreadsheetApp.getUi().alert(`${formattedCount}個の分岐シートの書式を設定しました。\n\n設定内容:\n- 行の高さ: 24ピクセル\n- 折り返し: 無効\n- ヘッダー: 日時、件名、概要、送信済み`);
    
  } catch (error) {
    console.error('分岐シート書式設定中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('分岐シート書式設定中にエラーが発生しました。\n' + error.toString());
  }
}

function setDefaultDiscordWebhook() {
  const ui = SpreadsheetApp.getUi();
  const currentWebhook = PropertiesService.getScriptProperties().getProperty('DEFAULT_DISCORD_WEBHOOK') || '';
  
  const result = ui.prompt(
    'デフォルトDiscord Webhook URL設定',
    'シート1全体通知用のDiscord Webhook URLを入力してください:\n現在の設定: ' + (currentWebhook.includes('YOUR_DEFAULT_WEBHOOK_HERE') ? '未設定' : '設定済み'),
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const webhookUrl = result.getResponseText().trim();
    if (webhookUrl && webhookUrl.startsWith('https://discord.com/api/webhooks/')) {
      PropertiesService.getScriptProperties().setProperty('DEFAULT_DISCORD_WEBHOOK', webhookUrl);
      ui.alert('デフォルトDiscord Webhook URLを設定しました。');
    } else if (webhookUrl) {
      ui.alert('有効なDiscord Webhook URLを入力してください。\nURLは https://discord.com/api/webhooks/ で始まる必要があります。');
    } else {
      ui.alert('Webhook URLが入力されていません。');
    }
  }
}

function createSendLog(logEntries) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('送信ログ');
    
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('送信ログ');
      
      // ヘッダーを設定
      logSheet.getRange(1, 1).setValue('送信日時');
      logSheet.getRange(1, 2).setValue('件名');
      logSheet.getRange(1, 3).setValue('ステータス');
      logSheet.getRange(1, 4).setValue('送信先');
      logSheet.getRange(1, 5).setValue('エラー詳細');
      
      // ヘッダー行を太字に
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      logSheet.setRowHeight(1, 24);
    }
    
    // ログエントリを追加
    for (const entry of logEntries) {
      const lastRow = logSheet.getLastRow();
      const newRow = lastRow + 1;
      
      logSheet.getRange(newRow, 1).setValue(Utilities.formatDate(entry.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
      logSheet.getRange(newRow, 2).setValue(entry.title);
      logSheet.getRange(newRow, 3).setValue(entry.status);
      logSheet.getRange(newRow, 4).setValue(entry.webhookUrl || '');
      logSheet.getRange(newRow, 5).setValue(entry.error || '');
      
      // 行の高さを24ピクセルに設定
      logSheet.setRowHeight(newRow, 24);
      
      // 折り返しを無効にする
      logSheet.getRange(newRow, 1, 1, 5).setWrap(false);
      
      // ステータスに応じて背景色を設定
      if (entry.status === '成功') {
        logSheet.getRange(newRow, 3).setBackground('#d4edda');
      } else if (entry.status === 'エラー') {
        logSheet.getRange(newRow, 3).setBackground('#f8d7da');
      }
    }
    
    console.log(`${logEntries.length}件のログを記録しました`);
    
  } catch (error) {
    console.error('送信ログ記録中にエラーが発生しました:', error);
  }
}