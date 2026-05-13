# Google Apps Script セットアップ手順

## 1. スプレッドシート作成
1. https://sheets.new を開く
2. シート名を「試合記録」にしておく（任意）

## 2. Apps Script を開く
1. メニュー → 「拡張機能」→「Apps Script」
2. 以下のコードを貼り付けて保存（Ctrl+S）

```javascript
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- シート「試合記録」---
    let matchSheet = ss.getSheetByName('試合記録');
    if (!matchSheet) matchSheet = ss.insertSheet('試合記録');
    if (matchSheet.getLastRow() === 0) {
      matchSheet.appendRow(['日時','チームA','チームB','スコアA','スコアB','勝者','イニング数']);
    }
    const winner = data.scoreA > data.scoreB ? data.teamA :
                   data.scoreB > data.scoreA ? data.teamB : '引き分け';
    matchSheet.appendRow([
      new Date().toLocaleString('ja-JP'),
      data.teamA, data.teamB,
      data.scoreA, data.scoreB,
      winner, data.innings
    ]);

    // --- シート「個人成績」---
    let playerSheet = ss.getSheetByName('個人成績');
    if (!playerSheet) playerSheet = ss.insertSheet('個人成績');
    if (playerSheet.getLastRow() === 0) {
      playerSheet.appendRow(['日時','チーム','選手','打席(PA)','打数(AB)','安打(H)',
        '二塁打','三塁打','本塁打','三振(K)','四球(BB)','打点(RBI)','打率','出塁率']);
    }
    const date = new Date().toLocaleString('ja-JP');
    (data.players || []).forEach(p => {
      const avg = p.ab > 0 ? (p.hits / p.ab).toFixed(3) : '---';
      const obp = p.pa > 0 ? ((p.hits + p.walks) / p.pa).toFixed(3) : '---';
      playerSheet.appendRow([
        date, p.team, p.name,
        p.pa, p.ab, p.hits,
        p.doubles, p.triples, p.hrs,
        p.ks, p.walks, p.rbi,
        avg, obp
      ]);
    });

    // --- シート「イニングスコア」---
    let inningSheet = ss.getSheetByName('イニングスコア');
    if (!inningSheet) inningSheet = ss.insertSheet('イニングスコア');
    if (inningSheet.getLastRow() === 0) {
      const headers = ['日時','チーム'];
      for (let i = 1; i <= 9; i++) headers.push(`${i}回`);
      headers.push('合計');
      inningSheet.appendRow(headers);
    }
    const date2 = new Date().toLocaleString('ja-JP');
    (data.inningScores || []).forEach(row => {
      inningSheet.appendRow([date2, row.team, ...row.scores, row.total]);
    });

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

## 3. デプロイ
1. 右上「デプロイ」→「新しいデプロイ」
2. 種類：「ウェブアプリ」
3. 実行するユーザー：「自分」
4. アクセスできるユーザー：「全員」
5. 「デプロイ」ボタン → URLをコピー

## 4. URLをアプリに設定
コピーした URL（https://script.google.com/macros/s/.../exec）を
アプリの「スプレッドシートに送信」ボタンを押したあとの入力欄に貼り付ける
