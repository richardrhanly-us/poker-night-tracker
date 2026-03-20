/**
 * Web app server layer.
 * - doGet() serves mobile.html
 * - mobile_get* functions supply data to the client
 *
 * IMPORTANT:
 * If your scans/medals are not publicly viewable, images may not load for anonymous users.
 * The web app runs as USER_DEPLOYING, so it can *read* Drive as the deployer,
 * but the browser still needs a URL it can fetch.
 */

const SETTINGS = {
  SCANS_FOLDER_ID: '1Ut5E_qls44O6O26GTCATb0BTdE38JzJ7',
  MEDALS_FOLDER_ID: '1Ut5E_qls44O6O26GTCATb0BTdE38JzJ7',

  SCANS_FOLDER_NAME: 'Poker Night Scans',
  MEDALS_FOLDER_NAME: 'Poker Night Scans',

  MEDAL_FILE_NAMES: {
    '1': '1st-Place-Medal.jpg',
    '2': '2nd-place-Medal.jpg',
    '3': '3rd-place-Medal.jpg'
  }
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('mobile')
    .setTitle('Poker');
}

/* =========================
   DATA API: config/master/night/player/top3
========================= */

function mobile_getConfig() {
  const ss = getSS_();
  const sessionsSh = ss.getSheetByName(SHEET_NAMES.SESSIONS);

  const players = (PLAYERS || []).slice();

  let nights = [];
  if (sessionsSh && sessionsSh.getLastRow() >= 2) {
    const values = sessionsSh.getRange(2, 1, sessionsSh.getLastRow() - 1, 2).getDisplayValues();
    nights = values
      .map(r => String(r[1] || '').trim())
      .filter(Boolean)
      .sort();
  }

  return {
    players,
    nights
  };
}

function mobile_getMaster() {
  const ss = getSS_();
  const sh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);
  if (!sh || sh.getLastRow() < 2) {
    return { headers: [], rows: [] };
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getDisplayValues();

  const byPlayer = {};
  (PLAYERS || []).forEach(p => {
    byPlayer[p] = {
      gamesPlayed: 0,
      totalBuyIn: 0,
      totalCashOut: 0,
      netTotal: 0
    };
  });

  values.forEach(row => {
    const player = String(row[2] || '').trim();
    if (!player) return;

    if (!byPlayer[player]) {
      byPlayer[player] = {
        gamesPlayed: 0,
        totalBuyIn: 0,
        totalCashOut: 0,
        netTotal: 0
      };
    }

    byPlayer[player].gamesPlayed += 1;
    byPlayer[player].totalBuyIn += toNumber_(row[7]);
    byPlayer[player].totalCashOut += toNumber_(row[8]);
    byPlayer[player].netTotal += toNumber_(row[9]);
  });

  const headers = [
    'Player',
    'Games Played',
    'Avg Buy-in',
    'Avg Cash-Out',
    'Net Avg',
    'Total Buy-in',
    'Total Cash-Out',
    'Net Total'
  ];

  const rows = Object.keys(byPlayer)
    .map(player => {
      const d = byPlayer[player];
      const games = d.gamesPlayed || 0;
      const avgBuy = games ? d.totalBuyIn / games : 0;
      const avgCash = games ? d.totalCashOut / games : 0;
      const avgNet = games ? d.netTotal / games : 0;

      return [
        player,
        games,
        avgBuy,
        avgCash,
        avgNet,
        d.totalBuyIn,
        d.totalCashOut,
        d.netTotal
      ];
    })
    .filter(r => (r[1] || 0) >= 4);

  rows.sort((a, b) => (b[7] || 0) - (a[7] || 0) || String(a[0]).localeCompare(String(b[0])));

  return { headers, rows };
}

function mobile_getNightTable(nightName) {
  const ss = getSS_();
  const sessionPlayersSh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);
  const sessionsSh = ss.getSheetByName(SHEET_NAMES.SESSIONS);

  if (!sessionPlayersSh || sessionPlayersSh.getLastRow() < 2) {
    return { rows: [], images: [] };
  }

  const targetDate = String(nightName || '').trim();
  if (!targetDate) return { rows: [], images: [] };

  const values = sessionPlayersSh.getRange(2, 1, sessionPlayersSh.getLastRow() - 1, 13).getDisplayValues();

  const rows = values
    .filter(r => String(r[1] || '').trim() === targetDate) // session_date
    .map(r => [
      r[2],               // player_name
      toNumber_(r[3]),    // buy_in
      toNumber_(r[4]),    // rebuy_1
      toNumber_(r[5]),    // rebuy_2
      toNumber_(r[6]),    // rebuy_3
      toNumber_(r[7]),    // total_buy_in
      toNumber_(r[8]),    // cash_out
      toNumber_(r[9])     // net
    ]);

  let images = [];
  if (sessionsSh && sessionsSh.getLastRow() >= 2) {
    const sessionValues = sessionsSh.getRange(2, 1, sessionsSh.getLastRow() - 1, 8).getDisplayValues();
    const sessionRow = sessionValues.find(r => String(r[1] || '').trim() === targetDate);

    if (sessionRow) {
      const imageLink = String(sessionRow[5] || '').trim(); // image_link
      const sourceSheet = String(sessionRow[2] || '').trim();

      if (imageLink) {
        images = [{
          id: '',
          name: 'Session image',
          thumb: imageLink,
          view: imageLink
        }];
      } else if (sourceSheet) {
        images = listNightImages_(sourceSheet);
      }
    }
  }

  return { rows, images };
}

function mobile_getPlayerTable(playerName) {
  const ss = getSS_();
  const sh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);

  const player = String(playerName || '').trim();
  if (!player || !sh || sh.getLastRow() < 2) {
    return { rows: [], totals: summarizeRows_([]), yearReviews: [] };
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getDisplayValues();

  const rows = values
    .filter(r => String(r[2] || '').trim() === player)
    .map(r => [
      r[1],               // session_date
      toNumber_(r[3]),    // buy_in
      toNumber_(r[4]),    // rebuy_1
      toNumber_(r[5]),    // rebuy_2
      toNumber_(r[6]),    // rebuy_3
      toNumber_(r[7]),    // total_buy_in
      toNumber_(r[8]),    // cash_out
      toNumber_(r[9])     // net
    ]);

  rows.sort((a, b) => {
    const da = new Date(a[0]).getTime() || 0;
    const db = new Date(b[0]).getTime() || 0;
    return da - db;
  });

  const totals = summarizeRows_(rows);

  const byYear = {};
  rows.forEach(r => {
    const y = String(r[0] || '').slice(0, 4) || 'Unknown';
    if (!byYear[y]) byYear[y] = [];
    byYear[y].push(r);
  });

  const yearReviews = Object.keys(byYear)
    .sort()
    .map(y => {
      const yrRows = byYear[y];
      const t = summarizeRows_(yrRows);
      const chart = runningChart_(yrRows);
      return { year: y, rows: yrRows, totals: t, chart };
    });

  return { rows, totals, yearReviews };
}

function mobile_getLeaderboardTop3() {
  const master = mobile_getMaster();
  const rows = master.rows || [];

  const top3 = rows.slice(0, 3).map((r, i) => ({
    player: r[0],
    net: toNumber_(r[7]),
    rank: i + 1
  }));

  return {
    top3,
    medals: resolveMedalUrls_()
  };
}

/* =========================
   Helpers: totals/chart/drive urls
========================= */

function toNumber_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  if (s[0] === '(' && s[s.length - 1] === ')') s = '-' + s.slice(1, -1);
  s = s.replace(/[^0-9.\-]/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function driveThumbnailUrl_(fileId, size) {
  const sz = size || 400;
  return `https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w${sz}`;
}

/**
 * rows are: [Date, Buy-in, Rebuy1, Rebuy2, Rebuy3, TotalBuy-in, Cash-out, Net]
 * Net index in player rows is 7.
 * Total buy-in index is 5.
 * Cash-out index is 6.
 */
function summarizeRows_(rows) {
  let buy = 0, cash = 0, net = 0;
  (rows || []).forEach(r => {
    buy  += toNumber_(r[5]);
    cash += toNumber_(r[6]);
    net  += toNumber_(r[7]);
  });
  const games = (rows || []).length;
  return {
    totalBuyIn: buy,
    cashOut: cash,
    net: net,
    gamesPlayed: games,
    avgBuyIn: games ? buy / games : 0,
    avgCashOut: games ? cash / games : 0,
    avgNet: games ? net / games : 0
  };
}

function runningChart_(rows) {
  let acc = 0;
  const dates = [];
  const running = [];
  (rows || []).forEach(r => {
    acc += toNumber_(r[7]);
    dates.push(r[0]);
    running.push(acc);
  });
  return { dates, running };
}

function debugMedals_() {
  const result = mobile_getLeaderboardTop3();
  Logger.log(JSON.stringify(result, null, 2));
}

/* ---------- Drive: folders ---------- */

function getFolderByIdOrName_(id, name) {
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  if (name) {
    const it = DriveApp.getFoldersByName(name);
    if (it.hasNext()) return it.next();
  }
  return null;
}

function getScansFolder_() {
  return getFolderByIdOrName_(SETTINGS.SCANS_FOLDER_ID, SETTINGS.SCANS_FOLDER_NAME);
}

function getMedalsFolder_() {
  return getFolderByIdOrName_(SETTINGS.MEDALS_FOLDER_ID, SETTINGS.MEDALS_FOLDER_NAME);
}

/* ---------- Drive: images for a given night ---------- */

function listNightImages_(nightName) {
  const scansRoot = getScansFolder_();
  if (!scansRoot) return [];

  const out = [];

  // Prefer: subfolder named exactly like the night (YYYY-MM-DD)
  const subFolders = scansRoot.getFoldersByName(nightName);
  if (subFolders.hasNext()) {
    const f = subFolders.next();
    return listImageFilesInFolder_(f);
  }

  // Fallback: files in root whose names contain the nightName
  const files = scansRoot.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const nm = (file.getName() || '').toLowerCase();
    if (!nm.includes(nightName.toLowerCase())) continue;
    if (!isImage_(file)) continue;
    out.push(fileToClientImage_(file));
  }

  return out;
}

function listImageFilesInFolder_(folder) {
  const out = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!isImage_(file)) continue;
    out.push(fileToClientImage_(file));
  }
  // Stable ordering: by name
  out.sort((a,b) => String(a.name || '').localeCompare(String(b.name || '')));
  return out;
}

function isImage_(file) {
  const mt = (file.getMimeType() || '').toLowerCase();
  return mt.startsWith('image/');
}

function fileToClientImage_(file) {
  const id = file.getId();
  const name = file.getName();
  return {
    id,
    name,
    // These are browser-fetchable URLs when sharing allows it
    thumb: directViewUrl_(id), // reliable as img src if file is shared
    view: file.getUrl()
  };
}

function directViewUrl_(fileId) {
  // This is the most common "works in <img>" URL format
  return `https://drive.google.com/uc?export=view&id=${encodeURIComponent(fileId)}`;
}

/* ---------- Drive: medals ---------- */

function resolveMedalUrls_() {
  const folder = getMedalsFolder_();
  const urls = { '1': null, '2': null, '3': null };
  if (!folder) return urls;

  const filesByName = {};
  const it = folder.getFiles();

  while (it.hasNext()) {
    const f = it.next();
    if (!isImage_(f)) continue;
    filesByName[String(f.getName() || '').trim().toLowerCase()] = f;
  }

  Object.keys(SETTINGS.MEDAL_FILE_NAMES).forEach(rank => {
    const expectedName = String(SETTINGS.MEDAL_FILE_NAMES[rank] || '').trim().toLowerCase();
    const hit = filesByName[expectedName];
    if (hit) {
      urls[rank] = driveThumbnailUrl_(hit.getId(), 300);
    }
  });

  return urls;
}