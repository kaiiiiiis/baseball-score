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
    // --- action=load：スプレッドシートから選手・成績を返す ---
    if (e.parameter && e.parameter.action === 'load') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('選手成績');
      const stats = {};
      const teams = { A: [], B: [], Aname: '', Bname: '' };

      if (sheet && sheet.getLastRow() >= 2) {
        // 13列 → 14列（N列:異名を追加）
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
        rows.forEach(r => {
          const name = r[0];
          if (!name || /^選手\d+$/.test(name)) return;
          const teamLabel = r[12] || '';
          const nickname  = r[13] || '';  // N列:異名
          stats[name] = {
            games:   r[1]  || 0,
            pa:      r[2]  || 0,
            ab:      r[3]  || 0,
            hits:    r[4]  || 0,
            doubles: r[5]  || 0,
            triples: r[6]  || 0,
            hrs:     r[7]  || 0,
            walks:   r[8]  || 0,
            ks:      r[9]  || 0,
            rbi:     r[10] || 0,
            avg:     r[11] || '.000',
            nickname: nickname  // ★追加
          };
          if (teamLabel === 'A') teams.A.push({ name, nickname });
          else if (teamLabel === 'B') teams.B.push({ name, nickname });
          else teams.A.push({ name, nickname });
        });
      }

      output.setContent(JSON.stringify({ status: 'ok', stats, teams }));
      return output;
    }

    // --- 通常の記録処理 ---
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
        '二塁打','三塁打','本塁打','四球(BB)','三振(K)','打点(RBI)','打率','チーム','異名'  // ★M・N列追加
      ]);
      sheet.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    }

    const players = data.players || [];

    players.forEach(p => {
      if (!p.name || /^選手\d+$/.test(p.name)) return;

      const name = p.name;
      const lastRow = sheet.getLastRow();
      let found = false;

      if (lastRow >= 2) {
        const nameCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < nameCol.length; i++) {
          if (nameCol[i][0] === name) {
            const row = i + 2;
            const existing = sheet.getRange(row, 1, 1, 14).getValues()[0];
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
            const team    = existing[12] || '';  // チームはそのまま保持
            const nick    = existing[13] || '';  // 異名はそのまま保持（上書きしない）★

            sheet.getRange(row, 1, 1, 14).setValues([[
              name, games, pa, ab, hits, doubles, triples, hrs, walks, ks, rbi, avg, team, nick
            ]]);
            found = true;
            break;
          }
        }
      }

      if (!found) {
        const ab   = p.ab || 0;
        const hits = p.hits || 0;
        const avg  = ab > 0 ? (hits / ab).toFixed(3) : '.000';
        sheet.appendRow([
          name, 1, p.pa||0, ab, hits,
          p.doubles||0, p.triples||0, p.hrs||0,
          p.walks||0, p.ks||0, p.rbi||0, avg, '', ''  // チーム・異名は空欄
        ]);
      }
    });

    // 打率降順ソート
    const lastRow2 = sheet.getLastRow();
    if (lastRow2 >= 3) {
      sheet.getRange(2, 1, lastRow2 - 1, 14).sort({ column: 12, ascending: false });
    }

    // --- シート「投手成績」---
    const pitchers = data.pitchers || [];
    if (pitchers.length > 0) {
      let pSheet = ss.getSheetByName('投手成績');
      if (!pSheet) {
        pSheet = ss.insertSheet('投手成績');
        pSheet.appendRow([
          '選手名','試合数','投球回','打者数(BF)','失点','自責点','被安打','三振','四球','防御率(ERA)','奪三振率(K/IP)'
        ]);
        pSheet.getRange(1,1,1,11).setFontWeight('bold').setBackground('#166534').setFontColor('#ffffff');
      }

      pitchers.forEach(p => {
        if (!p.name || /^投手\d+$/.test(p.name)) return;
        const name = p.name;
        const pLastRow = pSheet.getLastRow();
        let found = false;

        if (pLastRow >= 2) {
          const nameCol = pSheet.getRange(2, 1, pLastRow - 1, 1).getValues();
          for (let i = 0; i < nameCol.length; i++) {
            if (nameCol[i][0] === name) {
              const row = i + 2;
              const ex    = pSheet.getRange(row, 1, 1, 11).getValues()[0];
              const games = (ex[1] || 0) + 1;
              const ip    = (ex[2] || 0) + (p.ip || 0);
              const bf    = (ex[3] || 0) + (p.bf || 0);
              const runs  = (ex[4] || 0) + (p.runs || 0);
              const er    = (ex[5] || 0) + (p.er || 0);
              const hits  = (ex[6] || 0) + (p.hits || 0);
              const ks    = (ex[7] || 0) + (p.ks || 0);
              const walks = (ex[8] || 0) + (p.walks || 0);
              const era   = ip > 0 ? ((er * 3) / ip).toFixed(2) : '---';
              const kip   = ip > 0 ? ((ks * 3) / ip).toFixed(2) : '---';
              pSheet.getRange(row, 1, 1, 11).setValues([[
                name, games, ip, bf, runs, er, hits, ks, walks, era, kip
              ]]);
              found = true;
              break;
            }
          }
        }

        if (!found) {
          const ip  = p.ip || 0;
          const er  = p.er || 0;
          const ks  = p.ks || 0;
          const era = ip > 0 ? ((er * 3) / ip).toFixed(2) : '---';
          const kip = ip > 0 ? ((ks * 3) / ip).toFixed(2) : '---';
          pSheet.appendRow([
            name, 1, ip, p.bf||0, p.runs||0, er,
            p.hits||0, ks, p.walks||0, era, kip
          ]);
        }
      });

      // ERA昇順ソート
      const pLast = pSheet.getLastRow();
      if (pLast >= 3) {
        pSheet.getRange(2, 1, pLast - 1, 11).sort({ column: 10, ascending: true });
      }
    }

    output.setContent(JSON.stringify({ status: 'ok', updated: players.length, pitchers: pitchers.length }));
    return output;

  } catch(err) {
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
    return output;
  }
}
