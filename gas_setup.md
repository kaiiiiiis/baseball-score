# Google Apps Script セットアップ手順

## 1. スプレッドシート作成
1. https://sheets.new を開く

## 2. Apps Script を開く
1. メニュー → 「拡張機能」→「Apps Script」
2. 既存のコードを全部消して、以下を貼り付けて保存（Ctrl+S）

```javascript
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // CORSヘッダー対応
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let data;
    if (e.postData) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else {
      output.setContent(JSON.stringify({ status: 'error', message: 'no data' }));
      return output;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- シート「選手成績」---
    let sheet = ss.getSheetByName('選手成績');
    if (!sheet) {
      sheet = ss.insertSheet('選手成績');
      sheet.appendRow([
        '選手名','試合数','打席(PA)','打数(AB)','安打(H)',
        '二塁打','三塁打','本塁打','四球(BB)','三振(K)','打点(RBI)','打率'
      ]);
      // ヘッダー行の書式設定
      sheet.getRange(1,1,1,12).setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    }

    const players = data.players || [];

    players.forEach(p => {
      const name = p.name;
      const lastRow = sheet.getLastRow();
      let found = false;

      // 既存の選手を検索（2行目以降）
      if (lastRow >= 2) {
        const nameCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < nameCol.length; i++) {
          if (nameCol[i][0] === name) {
            // 既存行を加算更新
            const row = i + 2;
            const existing = sheet.getRange(row, 1, 1, 12).getValues()[0];
            const games   = (existing[1] || 0) + 1;
            const pa      = (existing[2] || 0) + (p.pa || 0);
            const ab      = (existing[3] || 0) + (p.ab || 0);
            const hits    = (existing[4] || 0) + (p.hits || 0);
            const doubles = (existing[5] || 0) + (p.doubles || 0);
            const triples = (existing[6] || 0) + (p.triples || 0);
            const hrs     = (existing[7] || 0) + (p.hrs || 0);
            const walks   = (existing[8] || 0) + (p.walks || 0);
            const ks      = (existing[9] || 0) + (p.ks || 0);
            const rbi     = (existing[10] || 0) + (p.rbi || 0);
            const avg     = ab > 0 ? (hits / ab).toFixed(3) : '.000';

            sheet.getRange(row, 1, 1, 12).setValues([[
              name, games, pa, ab, hits, doubles, triples, hrs, walks, ks, rbi, avg
            ]]);
            found = true;
            break;
          }
        }
      }

      if (!found) {
        // 新規追加
        const ab  = p.ab || 0;
        const hits = p.hits || 0;
        const avg = ab > 0 ? (hits / ab).toFixed(3) : '.000';
        sheet.appendRow([
          name, 1, p.pa||0, ab, hits,
          p.doubles||0, p.triples||0, p.hrs||0,
          p.walks||0, p.ks||0, p.rbi||0, avg
        ]);
      }
    });

    // 打率でソート（降順）
    const lastRow2 = sheet.getLastRow();
    if (lastRow2 >= 3) {
      sheet.getRange(2, 1, lastRow2 - 1, 12).sort({ column: 12, ascending: false });
    }

    output.setContent(JSON.stringify({ status: 'ok', updated: players.length }));
    return output;

  } catch(err) {
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
    return output;
  }
}
```

## 3. デプロイ
1. 右上「デプロイ」→「新しいデプロイ」
2. 種類の歯車アイコン →「ウェブアプリ」を選択
3. 設定：
   - 実行するユーザー：「**自分**」
   - アクセスできるユーザー：「**全員**」
4. 「デプロイ」ボタン → **権限を承認** → URLをコピー

## 4. URLをアプリに設定
コピーした URL（https://script.google.com/macros/s/.../exec）を
アプリの「スプレッドシートに記録する」ボタン → 入力欄に貼り付けて送信

## 注意
- コードを修正したら「新しいデプロイ」ではなく「デプロイを管理」→「編集」→「バージョン：新しいバージョン」で更新
- 同じ選手名であれば試合をまたいで累積加算される
- 打率は毎回（安打÷打数）で再計算
