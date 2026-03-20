/** Poker backend + legacy support **/

const SPREADSHEET_ID = '';

const PLAYERS = [
  'Ray','Justin','Marisa','Windham','Meagan','Danny','Claudia','Pelon',
  'James','Adam','Benny','Marshall','Curtis','Scales',
];

const SHEET_NAMES = {
  MASTER: 'Master',
  SESSION_ENTRY: 'SessionEntry',
  SESSION_PLAYERS: 'SessionPlayers',
  SESSIONS: 'Sessions',
};

const NIGHT_LAYOUT = {
  DATE_LABEL: 'A1',
  DATE_VALUE: 'B1',
  NOTES_LABEL: 'A2',
  NOTES_VALUE: 'B2',
  IMAGE_LABEL: 'A3',
  IMAGE_VALUE: 'B3',
  HEADER_ROW: 3,
  START_ROW: 4,
  COL_PLAYER: 1,
  COL_BUYIN: 2,
  COL_REBUY1: 3,
  COL_REBUY2: 4,
  COL_REBUY3: 5,
  COL_TOTAL_BUYIN: 6,
  COL_CASHOUT: 7,
  COL_NET: 8,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Poker')
    .addItem('Format Legacy Night Sheets + Master', 'formatDatesAndMaster')
    .addItem('Refresh Legacy Master Only', 'refreshMasterOnly')
    .addSeparator()
    .addItem('Set Up Session Entry Sheet', 'setupSessionEntrySheet')
    .addItem('Migrate Old Night Sheets to New Backend', 'migrateOldNightSheetsToNewBackend')
    .addItem('Finalize Session Entry to New Backend', 'finalizeSessionEntry')
    .addItem('Delete Session by session_id', 'deleteSessionBySessionId')
    .addToUi();
}

/* =========================
   Shared helpers
========================= */

function validateSessionEntry_() {
  const ss = getSS_();
  const entrySh = ss.getSheetByName(SHEET_NAMES.SESSION_ENTRY);
  if (!entrySh) throw new Error('Missing SessionEntry sheet.');

  const rawDate = entrySh.getRange('B1').getValue();
  if (!rawDate) throw new Error('Session date is required in B1.');

  const dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
  if (isNaN(dateObj.getTime())) throw new Error('Session date in B1 is not valid.');

  const sessionDate = formatDateYmd_(dateObj);
  const sessionId = buildSessionId_(sessionDate);

  const sessionsSh = getSS_().getSheetByName(SHEET_NAMES.SESSIONS);
  if (sessionsSh) {
    const existingSessionIds = getExistingSessionIds_(sessionsSh);
    if (existingSessionIds.has(sessionId)) {
      throw new Error(`Session already exists: ${sessionId}`);
    }
  }

  const startRow = 7;
  const lastRow = entrySh.getLastRow();
  if (lastRow < startRow) throw new Error('No player rows found in SessionEntry.');

  const values = entrySh.getRange(startRow, 1, lastRow - startRow + 1, 8).getDisplayValues();

  const errors = [];
  const seenPlayers = new Set();
  let completedRows = 0;

  values.forEach((row, idx) => {
    const sheetRow = startRow + idx;

    const playerName = String(row[0] || '').trim();
    const buyIn = toNumber_(row[1]);
    const rebuy1 = toNumber_(row[2]);
    const rebuy2 = toNumber_(row[3]);
    const rebuy3 = toNumber_(row[4]);
    const totalBuyIn = toNumber_(row[5]);
    const cashOut = toNumber_(row[6]);
    const net = toNumber_(row[7]);

    const hasAnyMoney =
      buyIn !== 0 || rebuy1 !== 0 || rebuy2 !== 0 || rebuy3 !== 0 || totalBuyIn !== 0 || cashOut !== 0 || net !== 0;

    if (!playerName && !hasAnyMoney) return;

    if (!playerName && hasAnyMoney) {
      errors.push(`Row ${sheetRow}: has amounts but no player name.`);
      return;
    }

    completedRows += 1;

    if (seenPlayers.has(playerName)) {
      errors.push(`Row ${sheetRow}: duplicate player "${playerName}".`);
    }
    seenPlayers.add(playerName);

    const nums = [buyIn, rebuy1, rebuy2, rebuy3, totalBuyIn, cashOut];
    if (nums.some(n => n < 0)) {
      errors.push(`Row ${sheetRow}: negative values are not allowed.`);
    }

    const expectedTotalBuyIn = buyIn + rebuy1 + rebuy2 + rebuy3;
    if (Math.abs(totalBuyIn - expectedTotalBuyIn) > 0.001) {
      errors.push(
        `Row ${sheetRow}: total_buy_in (${totalBuyIn}) does not match buy_in + rebuys (${expectedTotalBuyIn}).`
      );
    }

    const expectedNet = cashOut - totalBuyIn;
    if (Math.abs(net - expectedNet) > 0.001) {
      errors.push(
        `Row ${sheetRow}: net (${net}) does not match cash_out - total_buy_in (${expectedNet}).`
      );
    }
  });

  if (completedRows === 0) {
    errors.push('No completed player rows found in SessionEntry.');
  }

  if (errors.length) {
    throw new Error(
      'Validation failed:\n\n• ' + errors.join('\n• ') 
    );
  }

  return {
    sessionId,
    sessionDate
  };
}

function clearSessionEntryForm_() {
  const ss = getSS_();
  const entrySh = ss.getSheetByName(SHEET_NAMES.SESSION_ENTRY);
  if (!entrySh) return;

  entrySh.getRange('B1:B3').clearContent();
  entrySh.getRange(7, 2, PLAYERS.length, 5).clearContent(); // buy_in through rebuy_3
  entrySh.getRange(7, 7, PLAYERS.length, 1).clearContent(); // cash_out
}

function getSS_() {
  if (SPREADSHEET_ID) {
    try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch (e) {}
  }
  return SpreadsheetApp.getActive();
}

function parseSheetDate_(name) {
  const s = String(name || '').trim();

  let m = s.match(/^(\d{1,2})[\/\-_](\d{1,2})[\/\-_]((?:19|20)\d{2})$/);
  if (m) return new Date(`${m[3]}-${m[1]}-${m[2]}`);

  m = s.match(/^((?:19|20)\d{2})[\/\-_](\d{1,2})[\/\-_](\d{1,2})$/);
  if (m) return new Date(`${m[1]}-${m[2]}-${m[3]}`);

  return null;
}

function formatDateYmd_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function buildSessionId_(sessionDateYmd) {
  return 'S_' + String(sessionDateYmd).replace(/-/g, '_');
}

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

function listDateSheets_(ss = getSS_()) {
  return ss.getSheets()
    .map(sh => sh.getName())
    .filter(n =>
      n !== SHEET_NAMES.MASTER &&
      n !== SHEET_NAMES.SESSION_ENTRY &&
      n !== SHEET_NAMES.SESSION_PLAYERS &&
      n !== SHEET_NAMES.SESSIONS &&
      parseSheetDate_(n)
    )
    .sort((a, b) => parseSheetDate_(a) - parseSheetDate_(b));
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getExistingSessionPlayerKeys_(sh) {
  const keys = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return keys;

  const values = sh.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  values.forEach(row => {
    const sessionId = String(row[0] || '').trim();
    const playerName = String(row[2] || '').trim();
    if (sessionId && playerName) keys.add(`${sessionId}||${playerName}`);
  });

  return keys;
}

function getExistingSessionIds_(sh) {
  const ids = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return ids;

  const values = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  values.forEach(row => {
    const sessionId = String(row[0] || '').trim();
    if (sessionId) ids.add(sessionId);
  });

  return ids;
}

/* =========================
   New backend sheets
========================= */

function setupSessionEntrySheet() {
  const ss = getSS_();
  const sh = getOrCreateSheet_(ss, SHEET_NAMES.SESSION_ENTRY);

  sh.clear();

  sh.getRange('A1').setValue('session_date').setFontWeight('bold');
  sh.getRange('A2').setValue('notes').setFontWeight('bold');
  sh.getRange('A3').setValue('image_link').setFontWeight('bold');

  const headers = [
    'player_name','buy_in','rebuy_1','rebuy_2','rebuy_3',
    'total_buy_in','cash_out','net'
  ];
  sh.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E7F2F8');

  sh.getRange(7, 1, PLAYERS.length, 1).setValues(PLAYERS.map(p => [p]));
  sh.getRange(7, 1, PLAYERS.length, 1).setFontWeight('bold');

  const startRow = 7;
  sh.getRange(startRow, 6).setFormula(`=SUM(B${startRow}:E${startRow})`);
  sh.getRange(startRow, 8).setFormula(`=G${startRow}-F${startRow}`);
  sh.getRange(startRow, 6, 1, 1).autoFill(sh.getRange(startRow, 6, PLAYERS.length, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sh.getRange(startRow, 8, 1, 1).autoFill(sh.getRange(startRow, 8, PLAYERS.length, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  sh.getRange(startRow, 2, PLAYERS.length, 7)
    .setNumberFormat('#,##0.00')
    .setHorizontalAlignment('right');

  sh.setFrozenRows(6);
  [170,90,90,90,90,120,110,100].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  SpreadsheetApp.getUi().alert('SessionEntry set up.');
}

function ensureBackendSheets_() {
  const ss = getSS_();

  const sessionPlayers = getOrCreateSheet_(ss, SHEET_NAMES.SESSION_PLAYERS);
  if (sessionPlayers.getLastRow() === 0) {
    sessionPlayers.getRange(1, 1, 1, 13).setValues([[
      'session_id','session_date','player_name',
      'buy_in','rebuy_1','rebuy_2','rebuy_3',
      'total_buy_in','cash_out','net',
      'source_sheet','imported_at','notes'
    ]]);
  }

  const sessions = getOrCreateSheet_(ss, SHEET_NAMES.SESSIONS);
  if (sessions.getLastRow() === 0) {
    sessions.getRange(1, 1, 1, 8).setValues([[
      'session_id','session_date','source_sheet','imported_at',
      'notes','image_link','total_buy_in','total_cash_out'
    ]]);
  }
}

/* =========================
   Session Management - Delete session
========================= */

function deleteSessionBySessionId() {
  const ui = SpreadsheetApp.getUi();
  const ss = getSS_();

  const sessionsSh = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const sessionPlayersSh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);

  if (!sessionsSh || !sessionPlayersSh) {
    ui.alert('Missing Sessions or SessionPlayers sheet.');
    return;
  }

  const lastRow = sessionsSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No sessions found.');
    return;
  }

  const sessionValues = sessionsSh.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  const sessionList = sessionValues
    .map(r => `${r[0]} | ${r[1]}`)
    .join('\n');

  const res = ui.prompt(
    'Delete Session',
    'Enter the session_id to delete.\n\nAvailable sessions:\n' + sessionList,
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const sessionId = String(res.getResponseText() || '').trim();
  if (!sessionId) {
    ui.alert('No session_id entered.');
    return;
  }

  const confirm = ui.alert(
    'Confirm Delete',
    `Delete session ${sessionId} from Sessions and SessionPlayers?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let deletedSessionPlayers = 0;
  let deletedSessions = 0;

  deletedSessionPlayers = deleteRowsByValue_(sessionPlayersSh, 1, sessionId);
  deletedSessions = deleteRowsByValue_(sessionsSh, 1, sessionId);

  ui.alert(
    'Delete complete.\n' +
    `SessionPlayers rows deleted: ${deletedSessionPlayers}\n` +
    `Sessions rows deleted: ${deletedSessions}`
  );
}

function deleteRowsByValue_(sh, columnIndex, targetValue) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const values = sh.getRange(2, columnIndex, lastRow - 1, 1).getDisplayValues();
  const rowsToDelete = [];

  values.forEach((row, i) => {
    if (String(row[0] || '').trim() === String(targetValue).trim()) {
      rowsToDelete.push(i + 2);
    }
  });

  rowsToDelete.reverse().forEach(r => sh.deleteRow(r));
  return rowsToDelete.length;
}

/* =========================
   Migration: old night sheets -> new backend
========================= */

function migrateOldNightSheetsToNewBackend() {
  const ss = getSS_();
  ensureBackendSheets_();

  const sessionPlayersSh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);
  const sessionsSh = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const dateSheets = listDateSheets_(ss);

  if (!dateSheets.length) {
    SpreadsheetApp.getUi().alert('No date-named sheets found to migrate.');
    return;
  }

  const existingSessionPlayerKeys = getExistingSessionPlayerKeys_(sessionPlayersSh);
  const existingSessionIds = getExistingSessionIds_(sessionsSh);

  const sessionPlayerRowsToAppend = [];
  const sessionRowsToAppend = [];

  dateSheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const dateObj = parseSheetDate_(sheetName);
    if (!dateObj) return;

    const sessionDate = formatDateYmd_(dateObj);
    const sessionId = buildSessionId_(sessionDate);
    const importedAt = new Date();

    let sessionTotalBuyIn = 0;
    let sessionTotalCashOut = 0;
    let hasAnyPlayerRows = false;

    const startRow = NIGHT_LAYOUT.START_ROW;
    const lastRow = sh.getLastRow();
    if (lastRow < startRow) return;

    const values = sh.getRange(startRow, 1, lastRow - startRow + 1, 8).getDisplayValues();

    values.forEach(row => {
      const playerName = String(row[0] || '').trim();
      if (!playerName) return;
      if (/^totals?$/i.test(playerName)) return;

      const buyIn = toNumber_(row[1]);
      const rebuy1 = toNumber_(row[2]);
      const rebuy2 = toNumber_(row[3]);
      const rebuy3 = toNumber_(row[4]);
      const totalBuyIn = toNumber_(row[5]);
      const cashOut = toNumber_(row[6]);
      const net = toNumber_(row[7]);

      const uniqueKey = `${sessionId}||${playerName}`;
      if (existingSessionPlayerKeys.has(uniqueKey)) return;

      sessionPlayerRowsToAppend.push([
        sessionId,
        sessionDate,
        playerName,
        buyIn,
        rebuy1,
        rebuy2,
        rebuy3,
        totalBuyIn,
        cashOut,
        net,
        sheetName,
        importedAt,
        ''
      ]);

      existingSessionPlayerKeys.add(uniqueKey);
      sessionTotalBuyIn += totalBuyIn;
      sessionTotalCashOut += cashOut;
      hasAnyPlayerRows = true;
    });

    if (hasAnyPlayerRows && !existingSessionIds.has(sessionId)) {
      sessionRowsToAppend.push([
        sessionId,
        sessionDate,
        sheetName,
        importedAt,
        '',
        '',
        sessionTotalBuyIn,
        sessionTotalCashOut
      ]);
      existingSessionIds.add(sessionId);
    }
  });

  if (sessionPlayerRowsToAppend.length) {
    sessionPlayersSh
      .getRange(sessionPlayersSh.getLastRow() + 1, 1, sessionPlayerRowsToAppend.length, 13)
      .setValues(sessionPlayerRowsToAppend);
  }

  if (sessionRowsToAppend.length) {
    sessionsSh
      .getRange(sessionsSh.getLastRow() + 1, 1, sessionRowsToAppend.length, 8)
      .setValues(sessionRowsToAppend);
  }

  SpreadsheetApp.getUi().alert(
    'Migration complete.\n' +
    'SessionPlayers rows added: ' + sessionPlayerRowsToAppend.length + '\n' +
    'Sessions rows added: ' + sessionRowsToAppend.length
  );
}

/* =========================
   SessionEntry
========================= */

function finalizeSessionEntry() {
  const ss = getSS_();
  ensureBackendSheets_();

  const entrySh = ss.getSheetByName(SHEET_NAMES.SESSION_ENTRY);
  const sessionPlayersSh = ss.getSheetByName(SHEET_NAMES.SESSION_PLAYERS);
  const sessionsSh = ss.getSheetByName(SHEET_NAMES.SESSIONS);

  if (!entrySh) throw new Error('Missing SessionEntry sheet.');

  const validated = validateSessionEntry_();
  const sessionId = validated.sessionId;
  const sessionDate = validated.sessionDate;

  const notes = String(entrySh.getRange('B2').getDisplayValue() || '').trim();
  const imageLink = String(entrySh.getRange('B3').getDisplayValue() || '').trim();
  const importedAt = new Date();

  const existingSessionPlayerKeys = getExistingSessionPlayerKeys_(sessionPlayersSh);

  const startRow = 7;
  const lastRow = entrySh.getLastRow();
  const values = entrySh.getRange(startRow, 1, lastRow - startRow + 1, 8).getDisplayValues();

  const playerRows = [];
  let sessionTotalBuyIn = 0;
  let sessionTotalCashOut = 0;

  values.forEach(row => {
    const playerName = String(row[0] || '').trim();
    if (!playerName) return;

    const buyIn = toNumber_(row[1]);
    const rebuy1 = toNumber_(row[2]);
    const rebuy2 = toNumber_(row[3]);
    const rebuy3 = toNumber_(row[4]);
    const totalBuyIn = toNumber_(row[5]);
    const cashOut = toNumber_(row[6]);
    const net = toNumber_(row[7]);

    const hasAnyMoney =
      buyIn !== 0 || rebuy1 !== 0 || rebuy2 !== 0 || rebuy3 !== 0 || totalBuyIn !== 0 || cashOut !== 0 || net !== 0;

    if (!hasAnyMoney) return;

    const uniqueKey = `${sessionId}||${playerName}`;
    if (existingSessionPlayerKeys.has(uniqueKey)) {
      throw new Error(`Duplicate player/session detected: ${playerName} / ${sessionId}`);
    }

    playerRows.push([
      sessionId,
      sessionDate,
      playerName,
      buyIn,
      rebuy1,
      rebuy2,
      rebuy3,
      totalBuyIn,
      cashOut,
      net,
      SHEET_NAMES.SESSION_ENTRY,
      importedAt,
      notes
    ]);

    sessionTotalBuyIn += totalBuyIn;
    sessionTotalCashOut += cashOut;
  });

  if (!playerRows.length) {
    throw new Error('No completed player rows found in SessionEntry.');
  }

  sessionPlayersSh
    .getRange(sessionPlayersSh.getLastRow() + 1, 1, playerRows.length, 13)
    .setValues(playerRows);

  sessionsSh
    .getRange(sessionsSh.getLastRow() + 1, 1, 1, 8)
    .setValues([[
      sessionId,
      sessionDate,
      SHEET_NAMES.SESSION_ENTRY,
      importedAt,
      notes,
      imageLink,
      sessionTotalBuyIn,
      sessionTotalCashOut
    ]]);

  clearSessionEntryForm_();

  SpreadsheetApp.getUi().alert(
    'Session finalized.\n' +
    'Session ID: ' + sessionId + '\n' +
    'Players added: ' + playerRows.length
  );
}

/* =========================
   Legacy support: old night sheets + Master
========================= */

function formatDatesAndMaster() {
  const ss = getSS_();
  const dateSheets = listDateSheets_(ss);

  dateSheets.forEach(name => {
    ensureNightSheet_(ss, name);
    styleNight_(ss.getSheetByName(name));
  });

  makeMaster_(ss, dateSheets);
  styleMaster_(ss.getSheetByName(SHEET_NAMES.MASTER));
}

function refreshMasterOnly() {
  const ss = getSS_();
  const dateSheets = listDateSheets_(ss);
  makeMaster_(ss, dateSheets);
  styleMaster_(ss.getSheetByName(SHEET_NAMES.MASTER));
}

function addNightByDate() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Add Night', 'Enter date as YYYY-MM-DD (e.g., 2025-01-03):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const name = String(res.getResponseText() || '').trim();
  if (!/^(?:19|20)\d{2}-\d{1,2}-\d{1,2}$/.test(name)) {
    ui.alert('Please use YYYY-MM-DD.');
    return;
  }

  const ss = getSS_();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureNightSheet_(ss, name);
  styleNight_(sh);
  ss.setActiveSheet(sh);
}

function ensureNightSheet_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const startRow = NIGHT_LAYOUT.START_ROW;
  const endRow = startRow + PLAYERS.length - 1;

  sh.getRange('A1').setValue('Date:').setFontWeight('bold');
  if (!sh.getRange('B1').getDisplayValue()) {
    const d = parseSheetDate_(name);
    if (d) sh.getRange('B1').setValue(d);
  }

  const headers = ['Player','Buy-in','Rebuy 1','Rebuy 2','Rebuy 3','Total Buy-in','Cash-out','Net'];
  sh.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E7F2F8');

  if (!sh.getRange(startRow, 1).getDisplayValue()) {
    sh.getRange(startRow, 1, PLAYERS.length, 1).setValues(PLAYERS.map(p => [p]));
  }

  sh.getRange(startRow, 6).setFormula(`=SUM(B${startRow}:E${startRow})`);
  sh.getRange(startRow, 8).setFormula(`=G${startRow}-F${startRow}`);
  sh.getRange(startRow, 6, 1, 1).autoFill(sh.getRange(startRow, 6, PLAYERS.length, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sh.getRange(startRow, 8, 1, 1).autoFill(sh.getRange(startRow, 8, PLAYERS.length, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  sh.getRange(startRow, 2, PLAYERS.length, 7)
    .setNumberFormat('#,##0.00')
    .setHorizontalAlignment('right');

  const totalsRow = endRow + 1;
  sh.getRange(totalsRow, 1).setValue('Totals').setFontWeight('bold');
  sh.getRange(totalsRow, 6).setFormula(`=SUM(F${startRow}:F${endRow})`).setFontWeight('bold');
  sh.getRange(totalsRow, 7).setFormula(`=SUM(G${startRow}:G${endRow})`).setFontWeight('bold');
  sh.getRange(totalsRow, 8).setFormula(`=SUM(H${startRow}:H${endRow})`).setFontWeight('bold');

  sh.setFrozenRows(3);
}

function styleNight_(sh) {
  if (!sh) return;

  const startRow = NIGHT_LAYOUT.START_ROW;
  const endRow = startRow + PLAYERS.length - 1;
  const totalsRow = endRow + 1;

  const rules = [];
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground('#FBFDFF')
      .setRanges([sh.getRange(startRow, 1, PLAYERS.length, 8)])
      .build()
  );

  const netRange = sh.getRange(startRow, 8, (totalsRow - startRow + 1), 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#0B8F3A')
      .setRanges([netRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#9C0006')
      .setRanges([netRange])
      .build()
  );

  sh.setConditionalFormatRules(rules);
}

function makeMaster_(ss, sheetNames) {
  const sh = ss.getSheetByName(SHEET_NAMES.MASTER) || ss.insertSheet(SHEET_NAMES.MASTER);
  sh.clear();

  const headers = [
    'Player','Games Played','Average Buy-in','Average Cash-Out',
    'Net Average','Sum of Total Buy-in','Sum of Total Cash-out','Net Total'
  ];
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#E7F2F8');

  sh.getRange(2, 1, PLAYERS.length, 1).setValues(PLAYERS.map(p => [p]));

  const qName = n => n.replace(/'/g, "''");
  const gamesExpr   = row => sheetNames.map(n => `COUNTIF('${qName(n)}'!A4:A, $A${row})`).join(' + ');
  const sumBuyExpr  = row => sheetNames.map(n => `SUMIF('${qName(n)}'!A4:A, $A${row}, '${qName(n)}'!F4:F)`).join(' + ');
  const sumCashExpr = row => sheetNames.map(n => `SUMIF('${qName(n)}'!A4:A, $A${row}, '${qName(n)}'!G4:G)`).join(' + ');
  const sumNetExpr  = row => sheetNames.map(n => `SUMIF('${qName(n)}'!A4:A, $A${row}, '${qName(n)}'!H4:H)`).join(' + ');

  for (let i = 0; i < PLAYERS.length; i++) {
    const r = 2 + i;
    const games = `(${gamesExpr(r)})`;
    const buySum = `(${sumBuyExpr(r)})`;
    const cashSum = `(${sumCashExpr(r)})`;
    const netSum = `(${sumNetExpr(r)})`;

    sh.getRange(r, 2).setFormula(`=${games}`);
    sh.getRange(r, 3).setFormula(`=IFERROR(${buySum}/${games},0)`);
    sh.getRange(r, 4).setFormula(`=IFERROR(${cashSum}/${games},0)`);
    sh.getRange(r, 5).setFormula(`=IFERROR(${netSum}/${games},0)`);
    sh.getRange(r, 6).setFormula(`=${buySum}`);
    sh.getRange(r, 7).setFormula(`=${cashSum}`);
    sh.getRange(r, 8).setFormula(`=${netSum}`);
  }

  sortMasterByNet_(sh);
}

function sortMasterByNet_(sh) {
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= 2) return;

  sh.getRange(2, 1, lastRow - 1, 8).sort([
    { column: 8, ascending: false },
    { column: 1, ascending: true }
  ]);
}

function styleMaster_(sh) {
  if (!sh) return;
  const lastRow = Math.max(2, sh.getLastRow());
  try { sh.getFilter() && sh.getFilter().remove(); } catch (e) {}
  sh.getRange(1, 1, lastRow, 8).createFilter();
}
