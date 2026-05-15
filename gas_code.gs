// ============================================================
//  baseball-score GAS コード
//  シート名: 「選手成績」「投手成績」
//  選手成績シート列: A:選手名 B:試合数 C:打席 D:打数 E:安打
//                   F:二塁打 G:三塁打 H:本塁打 I:四球 J:三振
//                   K:打点 L:打率 M:チーム N:異名
// ============================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // ── action=load：スプレッドシートから選手・成績を返す ──
    if (e.parameter && e.parameter.action === 'load') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('選手成績');
      const stats = {};
      const teams = { A: [], B: [], Aname: '', Bname: '' };

      if (sheet && sheet.getLastRow() >= 2) {
        // N列(14列目)まで取得
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
        rows.forEach(r => {
          const name = r[0];
          if (!name || /^選手\d+$/.test(name)) return;
          const teamLabel = (r[12] || '').toString().trim();
          const nickname  = (r[13] || '').toString().trim();

          stats[name] = {
            games:    r[1]  || 0,
            pa:       r[2]  || 0,
            ab:       r[3]  || 0,
            hits:     r[4]  || 0,
            doubles:  r[5]  || 0,
            triples:  r[6]  || 0,
            hrs:      r[7]  || 0,
            walks:    r[8]  || 0,
            ks:       r[9]  || 0,
            rbi:      r[10] || 0,
            teamLabel: teamLabel,
            nickname:  nickname
          };

          // チームA / B の選手リスト（登録順）
          if (teamLabel === 'A') {
            teams.A.push({ name, nickname });
          } else if (teamLabel === 'B') {
            teams.B.push({ name, nickname });
          }
        });
      }

      // チーム名シートがあれば読む（なければデフォルト）
      const configSheet = ss.getSheetByName('設定');
      if (configSheet) {
        teams.Aname = configSheet.getRange('B1').getValue() || 'チームA';
        teams.Bname = configSheet.getRange('B2').getValue() || 'チームB';
      }

      output.setContent(JSON.stringify({ ok: true, stats, teams }));
      return output;
    }

    // ── 通常送信（今試合の成績を累積加算） ──
    let payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      payload = JSON.parse(decodeURIComponent(e.parameter.data));
    } else {
      throw new Error('No data');
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('選手成績') || ss.insertSheet('選手成績');

    // ヘッダー行を確認・作成
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['選手名','試合数','打席','打数','安打','二塁打','三塁打','本塁打','四球','三振','打点','打率','チーム','異名']);
    }

    const players = payload.players || [];

    players.forEach(p => {
      if (!p.name || /^選手\d+$/.test(p.name)) return;

      // 既存行を探す
      const lastRow = sheet.getLastRow();
      let rowIndex  = -1;
      if (lastRow >= 2) {
        const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        rowIndex = names.indexOf(p.name);
        if (rowIndex >= 0) rowIndex += 2; // 1-indexed + ヘッダー行
      }

      if (rowIndex > 0) {
        // 既存行に累積加算
        const row    = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
        const games  = (row[1]  || 0) + 1;
        const pa     = (row[2]  || 0) + (p.pa      || 0);
        const ab     = (row[3]  || 0) + (p.ab      || 0);
        const hits   = (row[4]  || 0) + (p.hits    || 0);
        const doubles= (row[5]  || 0) + (p.doubles || 0);
        const triples= (row[6]  || 0) + (p.triples || 0);
        const hrs    = (row[7]  || 0) + (p.hrs     || 0);
        const walks  = (row[8]  || 0) + (p.walks   || 0);
        const ks     = (row[9]  || 0) + (p.ks      || 0);
        const rbi    = (row[10] || 0) + (p.rbi     || 0);
        const avg    = ab > 0 ? (hits / ab).toFixed(3) : '.000';
        const team   = row[12] || '';
        const nick   = row[13] || ''; // 異名は上書きしない

        sheet.getRange(rowIndex, 1, 1, 14).setValues([[
          p.name, games, pa, ab, hits, doubles, triples, hrs, walks, ks, rbi, avg, team, nick
        ]]);
      } else {
        // 新規行追加
        const ab   = p.ab   || 0;
        const hits = p.hits || 0;
        const avg  = ab > 0 ? (hits / ab).toFixed(3) : '.000';
        sheet.appendRow([
          p.name, 1, p.pa||0, ab, hits,
          p.doubles||0, p.triples||0, p.hrs||0,
          p.walks||0, p.ks||0, p.rbi||0,
          avg, '', ''  // チーム・異名は空欄（後でスプレッドシートで入力）
        ]);
      }
    });

    output.setContent(JSON.stringify({ ok: true }));
    return output;

  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
    return output;
  }
}
