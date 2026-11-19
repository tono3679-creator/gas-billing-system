/**
 * シートの読み書き。列は COLUMNS の定義順で固定し、
 * 数式で計算する列（calc: true）には書き込まない。
 */

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」が見つかりません。メニュー「' + APP.menuName + ' › 管理 › 初期セットアップ」を実行してください。');
  return sh;
}

function sheetKeyOf_(name) {
  for (const k in SHEET) if (SHEET[k] === name) return k;
  return null;
}

/**
 * 1 列目（キー列）に値がある最後の行。ARRAYFORMULA の空文字で
 * getLastRow() が膨らむのを避けるため、キー列を下から探す。
 */
function lastDataRow_(sh, keyCol) {
  const max = sh.getLastRow();
  if (max < 2) return 1;
  const vals = sh.getRange(2, keyCol || 1, max - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] !== null) return i + 2;
  }
  return 1;
}

function normalizeCell_(col, v) {
  if (col.type === 'date') return toYmd(v);
  if (col.type === 'yen' || col.type === 'number' || col.type === 'percent') return v === '' || v === null ? '' : Number(v);
  if (col.type === 'checkbox') return v === true;
  return v === null || v === undefined ? '' : (typeof v === 'string' ? v.trim() : v);
}

/** 表を読み、{ 列key: 値, _row: 行番号 } の配列で返す。キー列が空の行は飛ばす */
function readTable_(tableKey, keyCol) {
  const sh = getSheet_(SHEET[tableKey]);
  const cols = COLUMNS[tableKey];
  const last = lastDataRow_(sh, keyCol || 1);
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (row[(keyCol || 1) - 1] === '' || row[(keyCol || 1) - 1] === null) return;
    const obj = { _row: i + 2 };
    cols.forEach(function (c, j) { obj[c.key] = normalizeCell_(c, row[j]); });
    out.push(obj);
  });
  return out;
}

function readRow_(tableKey, row) {
  const sh = getSheet_(SHEET[tableKey]);
  const cols = COLUMNS[tableKey];
  const values = sh.getRange(row, 1, 1, cols.length).getValues()[0];
  const obj = { _row: row };
  cols.forEach(function (c, j) { obj[c.key] = normalizeCell_(c, values[j]); });
  return obj;
}

function toCellValue_(col, v) {
  if (v === undefined || v === null) return '';
  if (col.type === 'date' && typeof v === 'string' && v) return ymdToDate(v);
  return v;
}

/**
 * 複数行をまとめて書く。計算列をまたがないよう、連続した入力列ごとに setValues する。
 * @param {number} startRow 書き始める行
 */
function writeRows_(tableKey, startRow, objs) {
  if (!objs.length) return;
  const sh = getSheet_(SHEET[tableKey]);
  const cols = COLUMNS[tableKey];
  let i = 0;
  while (i < cols.length) {
    if (cols[i].calc) { i++; continue; }
    let j = i;
    while (j + 1 < cols.length && !cols[j + 1].calc) j++;
    const block = objs.map(function (o) {
      const r = [];
      for (let k = i; k <= j; k++) r.push(toCellValue_(cols[k], o[cols[k].key]));
      return r;
    });
    sh.getRange(startRow, i + 1, objs.length, j - i + 1).setValues(block);
    i = j + 1;
  }
}

function appendRows_(tableKey, objs, keyCol) {
  const sh = getSheet_(SHEET[tableKey]);
  const start = lastDataRow_(sh, keyCol || 1) + 1;
  if (start + objs.length - 1 > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), start + objs.length - 1 - sh.getMaxRows() + 50);
  }
  writeRows_(tableKey, start, objs);
  return start;
}

/** 1 行の一部の列だけ書き換える。{ key: value } で渡す */
function updateCells_(tableKey, row, patch) {
  const sh = getSheet_(SHEET[tableKey]);
  const cols = COLUMNS[tableKey];
  Object.keys(patch).forEach(function (key) {
    const idx = COL[tableKey][key];
    if (!idx) throw new Error('列 ' + key + ' がありません');
    if (cols[idx - 1].calc) throw new Error('計算列 ' + key + ' には書き込めません');
    sh.getRange(row, idx).setValue(toCellValue_(cols[idx - 1], patch[key]));
  });
}

/** セルにリンク（=HYPERLINK）を入れる */
function setLink_(tableKey, row, key, url, text) {
  const sh = getSheet_(SHEET[tableKey]);
  const safe = String(text).replace(/"/g, '""');
  sh.getRange(row, COL[tableKey][key]).setFormula('=HYPERLINK("' + url + '","' + safe + '")');
}

/** =HYPERLINK のセルから Drive のファイル ID を取り出す（なければ ''） */
function getLinkedFileId_(tableKey, row, key) {
  const f = getSheet_(SHEET[tableKey]).getRange(row, COL[tableKey][key]).getFormula();
  const m = String(f || '').match(/\/d\/([\w-]{20,})/);
  return m ? m[1] : '';
}

// ---------- 設定 ----------

function getSettings_() {
  const sh = getSheet_(SHEET.settings);
  const last = sh.getLastRow();
  const values = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  const byLabel = {};
  values.forEach(function (r) { byLabel[String(r[0]).trim()] = r[1]; });
  const s = {};
  SETTINGS_SPEC.forEach(function (spec) {
    if (!spec.key) return;
    const v = byLabel[spec.label];
    s[spec.key] = v === undefined || v === '' ? spec.value : v;
    if (typeof spec.value === 'number') s[spec.key] = Number(s[spec.key]) || spec.value;
    if (typeof spec.value === 'boolean') s[spec.key] = s[spec.key] === true || s[spec.key] === 'TRUE';
  });
  return s;
}

function setSetting_(key, value) {
  const spec = SETTINGS_SPEC.filter(function (x) { return x.key === key; })[0];
  const sh = getSheet_(SHEET.settings);
  const last = sh.getLastRow();
  const labels = sh.getRange(2, 1, Math.max(last - 1, 1), 1).getValues();
  for (let i = 0; i < labels.length; i++) {
    if (String(labels[i][0]).trim() === spec.label) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  throw new Error('設定項目「' + spec.label + '」が見つかりません');
}

// ---------- ログ ----------

function writeLog_(action, target, ok, detail) {
  try {
    const sh = getSs_().getSheetByName(SHEET.log);
    if (!sh) return;
    let user = '';
    try { user = Session.getActiveUser().getEmail(); } catch (e) { user = ''; }
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 6).setValues([[new Date(), action, target || '', ok ? '成功' : '失敗', detail || '', user]]);
  } catch (e) {
    console.error('ログの書き込みに失敗: ' + e);
  }
}

// ---------- まとめて読む ----------

/** 業務処理でよく使う一式を読む */
function loadBook_() {
  const customers = readTable_('customers');
  const products = readTable_('products');
  const deals = readTable_('deals');
  const lines = readTable_('lines');
  const payments = readTable_('payments', 2);
  const customerByName = {};
  customers.forEach(function (c) { customerByName[c.name] = c; });
  const productByName = {};
  products.forEach(function (p) { productByName[p.name] = p; });
  const linesByDeal = {};
  lines.forEach(function (l) { (linesByDeal[l.dealId] = linesByDeal[l.dealId] || []).push(l); });
  const paidByInvoice = {};
  payments.forEach(function (p) { paidByInvoice[p.invoiceNo] = (paidByInvoice[p.invoiceNo] || 0) + (Number(p.amount) || 0); });
  // シートの数式とは独立に、JS 側でも金額と入金状況を計算しておく
  deals.forEach(function (d) {
    const t = calcTotals(linesByDeal[d.id] || []);
    d.calc = t;
    d.calcPaid = d.invoiceNo ? paidByInvoice[d.invoiceNo] || 0 : 0;
  });
  return { customers: customers, products: products, deals: deals, lines: lines, payments: payments, customerByName: customerByName, productByName: productByName, linesByDeal: linesByDeal, paidByInvoice: paidByInvoice };
}

/** 支払状況の判定に渡す形へ */
function dealForStatus_(d) {
  return { orderStatus: d.orderStatus, invoiceNo: d.invoiceNo, total: d.calc ? d.calc.total : d.total, paid: d.calcPaid !== undefined ? d.calcPaid : d.paid, dueOn: d.dueOn };
}

function todayYmd_() {
  return Utilities.formatDate(new Date(), APP.timeZone, 'yyyy-MM-dd');
}
