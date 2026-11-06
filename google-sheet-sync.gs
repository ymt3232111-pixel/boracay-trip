// 長灘島記帳：四人共用帳本（貼到 Google 試算表的 Apps Script）
const HEAD = ['id', 'name', 'amount', 'cur', 'payer', 'split', 'day', 'cat', 'ts'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('記帳');
  if (!s) { s = ss.insertSheet('記帳'); s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}

function state_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const expenses = rows.filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), amount: Number(r[2]), cur: String(r[3]),
    payer: String(r[4]), split: String(r[5]).split(',').filter(Boolean),
    day: String(r[6]), cat: String(r[7]), ts: Number(r[8])
  }));
  const saved = PropertiesService.getScriptProperties().getProperty('rates');
  return { expenses, rates: saved ? JSON.parse(saved) : { PHP: 0.55, USD: 33 } };
}

function text_(v) {
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(state_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const s = sheet_();
    const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
    if (req.op === 'upsert') {
      const x = req.exp;
      const row = [text_(x.id), text_(x.name), Number(x.amount), text_(x.cur), text_(x.payer),
                   text_((x.split || []).join(',')), text_(x.day), text_(x.cat), Number(x.ts)];
      const i = ids.indexOf(String(x.id));
      if (i > 0) s.getRange(i + 1, 1, 1, row.length).setValues([row]);
      else s.appendRow(row);
    } else if (req.op === 'delete') {
      const i = ids.indexOf(String(req.id));
      if (i > 0) s.deleteRow(i + 1);
    } else if (req.op === 'rates') {
      PropertiesService.getScriptProperties().setProperty('rates',
        JSON.stringify({ PHP: Number(req.rates.PHP), USD: Number(req.rates.USD) }));
    }
    return json_(state_());
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
