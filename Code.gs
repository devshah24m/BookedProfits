// BookedProfits - Google Apps Script Backend
// 1. Go to your Google Sheet
// 2. Extensions -> Apps Script -> paste this entire file as Code.gs
// 3. Deploy -> New Deployment -> Web App -> Execute as: Me -> Who has access: Anyone -> Deploy
// 4. Copy the Web App URL and paste it into the dashboard

const SHEET_ID = '1DP9SIv2JX8ZJYg__lz-qmE-D1C0EzpCz-7OOYLclI9I';

function doGet(e) {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const eq  = readSheet(ss, 'Equity',  readEquityRow);
    const fno = readSheet(ss, 'FNO',     readFnoRow);
    const out = JSON.stringify({ ok: true, equity: eq, fno: fno });
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function readSheet(ss, name, rowFn) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var hdrs = data[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  });
  return data.slice(1)
    .filter(function(row) { return row.some(function(c) { return c !== '' && c !== null; }); })
    .map(function(row) {
      var r = {};
      hdrs.forEach(function(h, i) { r[h] = row[i]; });
      return rowFn(r);
    })
    .filter(function(r) { return r !== null; });
}

function readEquityRow(r) {
  var sym   = String(r['symbol'] || r['stock'] || r['scrip'] || '').trim().toUpperCase();
  var qty   = parseFloat(r['qty'] || r['quantity'] || 0) || 0;
  var buyP  = parseFloat(r['buyprice'] || r['buy'] || r['avgbuy'] || r['entryprice'] || 0) || 0;
  var sellP = parseFloat(r['sellprice'] || r['sell'] || r['avgsell'] || r['exitprice'] || 0) || 0;
  var pnl   = parseFloat(r['pnl'] || r['profit'] || r['profitloss'] || 0) || 0;
  if (!pnl && buyP && sellP && qty) pnl = (sellP - buyP) * qty;
  var buyDate  = fmtDate(r['buydate']  || r['purchasedate'] || r['entrydate'] || '');
  var sellDate = fmtDate(r['selldate'] || r['exitdate']     || r['saledate']  || '');
  if (!sym || qty <= 0 || buyP <= 0 || sellP <= 0) return null;
  return { sym: sym, qty: qty, buyDate: buyDate, buyP: buyP, sellDate: sellDate, sellP: sellP, pnl: Math.round(pnl * 100) / 100 };
}

function readFnoRow(r) {
  var sym   = String(r['symbol'] || r['stock'] || r['scrip'] || '').trim().toUpperCase();
  var inst  = String(r['instrument'] || r['type'] || r['contract'] || r['option'] || '').trim();
  var qty   = parseFloat(r['qty'] || r['lots'] || r['quantity'] || 0) || 0;
  var buyP  = parseFloat(r['buyprice'] || r['buy'] || r['entry'] || 0) || 0;
  var sellP = parseFloat(r['sellprice'] || r['sell'] || r['exit'] || 0) || 0;
  var pnl   = parseFloat(r['pnl'] || r['profit'] || r['profitloss'] || r['gain'] || r['loss'] || 0) || 0;
  if (!pnl && buyP && sellP && qty) pnl = (sellP - buyP) * qty;
  if (!sym) return null;
  return { sym: sym, inst: inst, qty: qty, buyP: buyP, sellP: sellP, pnl: Math.round(pnl * 100) / 100 };
}

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' +
      String(val.getMonth()+1).padStart(2,'0') + '-' +
      String(val.getDate()).padStart(2,'0');
  }
  return String(val).trim();
}
