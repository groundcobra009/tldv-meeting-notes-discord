# Discord通知実装解説書

## 概要
Google Apps ScriptでスプレッドシートからDiscordのフォーラムチャンネルに自動通知を送信する機能の実装について、発生した問題と解決方法をまとめます。

## 実装した機能
- スプレッドシートのデータをDiscordのフォーラムチャンネルに自動投稿
- D列のチェックボックスによる送信済み管理
- 毎日23時台の自動実行トリガー

## 発生した問題と解決方法

### 問題1: Discordフォーラムチャンネルへの投稿エラー

#### エラー内容
```
{"message": "Webhooks posted to forum channels must have a thread_name or thread_id", "code": 220001}
```

#### 原因
Discordのフォーラム形式チャンネルでは、通常のWebhook投稿と異なり、新しいスレッドを作成する際に`thread_name`または`thread_id`パラメータが必須です。

#### 解決方法
Webhook送信時のJSONペイロードに`thread_name`フィールドを追加：

```javascript
const message = {
  content: null,
  thread_name: `${titleValue} - ${formattedDate}`, // フォーラムチャンネル用
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
```

### 問題2: クエリ関数との相性問題（予想される問題）

#### 問題の背景
スプレッドシートでQUERY関数を使用してA〜C列にデータを表示している場合、以下の問題が発生する可能性があります：

1. **データの動的変更**: クエリ結果が変わると行の順序が変わる
2. **読み取り専用セル**: クエリ関数の結果セルは編集不可
3. **行番号の不一致**: データの並び順が変わるとチェックボックスとの対応が崩れる

#### 採用した解決策
D列のチェックボックスによるシンプルな管理：

```javascript
// D列がチェックされていない場合のみ送信
if (!checkboxValue && dateValue && titleValue) {
  // Discord送信処理
  // ...
  
  // 送信成功後、チェックボックスを更新
  if (response.getResponseCode() === 204 || response.getResponseCode() === 200) {
    sheet.getRange(row, 4).setValue(true);
    sentCount++;
  }
}
```

#### 代替案（複雑だが堅牢）
データ内容ベースの送信履歴管理：

```javascript
// 一意のキーを生成（日時と件名の組み合わせ）
const uniqueKey = `${dateValue.getTime()}_${titleValue}`;

// 送信履歴シートで重複チェック
if (!sentKeys.has(uniqueKey)) {
  // 送信処理
}
```

## 実装のポイント

### 1. データ範囲の取得
クエリ関数のデータに対応するため、実際のデータ範囲を動的に取得：

```javascript
// A列の最終行を確認（クエリ関数のデータ）
let lastRow = 1;
const columnA = sheet.getRange('A:A').getValues();
for (let i = columnA.length - 1; i >= 0; i--) {
  if (columnA[i][0] !== '') {
    lastRow = i + 1;
    break;
  }
}
```

### 2. トリガーの設定
毎日23時台のランダムな時間に実行：

```javascript
function setupDiscordTrigger() {
  // 23時から24時の間のランダムな時間を生成
  const randomMinute = Math.floor(Math.random() * 60);
  
  ScriptApp.newTrigger('sendUnsentToDiscord')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(randomMinute)
    .create();
}
```

### 3. エラーハンドリング
詳細なログ出力でデバッグを容易に：

```javascript
console.log(`行${row}: チェックボックス=${checkboxValue}, 日時=${dateValue}, タイトル=${titleValue}`);

if (response.getResponseCode() === 204 || response.getResponseCode() === 200) {
  console.log(`Discord送信成功 (行 ${row}): ${titleValue}`);
} else {
  console.error(`Discord送信エラー (行 ${row}):`, response.getContentText());
}
```

## 今後の注意点

### Discordフォーラムチャンネル使用時
- 必ず`thread_name`を指定する
- スレッド名は一意性を保つため日時を含める
- 長すぎるスレッド名は避ける（100文字制限）

### クエリ関数使用時
- D列のチェックボックスはクエリ範囲外に配置
- データの並び順が変わる可能性を考慮
- 必要に応じて送信履歴シートでの管理を検討

### パフォーマンス
- データを一括取得してAPI呼び出し回数を最小化
- 送信間隔（1秒）でレート制限を回避
- 大量データの場合はバッチ処理を検討

## まとめ

最も重要だった解決策は**Discordフォーラムチャンネルでの`thread_name`指定**でした。クエリ関数の問題は予想されたものでしたが、実際にはD列のチェックボックス管理で十分対応できており、シンプルな解決策が効果的でした。

今後同様の実装を行う際は、まずDiscordチャンネルの種類（通常 vs フォーラム）を確認し、適切なWebhookペイロード構造を使用することが重要です。