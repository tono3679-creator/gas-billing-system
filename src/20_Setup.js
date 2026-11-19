/**
 * 初期セットアップ：シート作成・書式・入力規則・数式・ダッシュボード・デモデータ。
 * 何度実行しても同じ状態に作り直せる（既存データは消える）。
 */

function q_(sheetName) {
  return "'" + sheetName + "'!";
}

/** 各表の計算列の数式（1 行目に見出し付きで置き、下の行へ自動で広がる） */
function columnFormulas_() {
  const D = q_(SHEET.deals);
  const L = q_(SHEET.lines);
  const P = q_(SHEET.payments);
  const C = q_(SHEET.customers);
  return {
    deals: {
      customerId: '={"顧客ID";ARRAYFORMULA(IF(C2:C="","",IFERROR(VLOOKUP(C2:C,{' + C + 'B2:B,' + C + 'A2:A},2,FALSE),"未登録")))}',
      subtotal: '={"税抜金額";MAP(A2:A,LAMBDA(id,IF(id="","",SUMIFS(' + L + 'I2:I,' + L + 'A2:A,id))))}',
      // 税率ごとに合計してから 1 回だけ切り捨て（適格請求書の端数処理）
      tax: '={"消費税";MAP(A2:A,LAMBDA(id,IF(id="","",ROUNDDOWN(SUMIFS(' + L + 'I2:I,' + L + 'A2:A,id,' + L + 'H2:H,0.1)*10/100)+ROUNDDOWN(SUMIFS(' + L + 'I2:I,' + L + 'A2:A,id,' + L + 'H2:H,0.08)*8/100))))}',
      total: '={"税込合計";ARRAYFORMULA(IF(A2:A="","",G2:G+H2:H))}',
      paid: '={"入金済額";MAP(A2:A,M2:M,LAMBDA(id,no,IF(OR(id="",no=""),"",SUMIFS(' + P + 'D2:D,' + P + 'B2:B,no))))}',
      balance: '={"未入金残高";ARRAYFORMULA(IF((A2:A="")+(M2:M=""),"",I2:I-R2:R))}',
      paymentStatus: '={"入金ステータス";ARRAYFORMULA(IF(A2:A="","",IF(F2:F="失注","失注",IF(M2:M="","未請求",IF(S2:S<=0,"入金済",IF((O2:O<>"")*(O2:O<TODAY()),"期限超過",IF(R2:R>0,"一部入金","未入金")))))))}',
      overdueDays: '={"期限まで/超過";ARRAYFORMULA(IF(((T2:T="未入金")+(T2:T="一部入金")+(T2:T="期限超過"))*(O2:O<>""),TODAY()-O2:O,""))}',
    },
    lines: {
      customer: '={"顧客名";ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,{' + D + 'A2:A,' + D + 'C2:C},2,FALSE),"案件が見つかりません")))}',
      amount: '={"金額（税抜）";ARRAYFORMULA(IF((A2:A="")+(E2:E=""),"",ROUND(E2:E*G2:G)))}',
    },
    payments: {
      customer: '={"顧客名";ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{' + D + 'M2:M,' + D + 'C2:C},2,FALSE),"請求番号が見つかりません")))}',
    },
  };
}

const NUMBER_FORMAT = {
  date: 'yyyy/mm/dd',
  datetime: 'yyyy/mm/dd hh:mm',
  yen: '"¥"#,##0',
  number: '#,##0.##',
  percent: '0%',
};

// ---------- 入口 ----------

/** メニュー：デモデータ入りで作り直す */
function menuSetupWithDemo() {
  const ui = SpreadsheetApp.getUi();
  if (hasUserData_()) {
    const ok = ui.alert('初期セットアップ', 'すべてのシートを作り直し、デモデータを入れます。\n今入っている案件・明細・入金・マスタは消えます。続けますか？', ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;
  }
  const result = setupSpreadsheet_({ demo: true, samplePdfs: true });
  ui.alert('セットアップ完了', result.message, ui.ButtonSet.OK);
}

/** メニュー：空の状態で作り直す（本番運用を始めるとき） */
function menuSetupEmpty() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('空の状態で作り直す', 'すべてのシートを作り直し、データを空にします（設定シートは初期値に戻ります）。続けますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const result = setupSpreadsheet_({ demo: false, samplePdfs: false });
  ui.alert('セットアップ完了', result.message, ui.ButtonSet.OK);
}

function hasUserData_() {
  const sh = getSs_().getSheetByName(SHEET.deals);
  return !!sh && lastDataRow_(sh, 1) >= 2;
}

/**
 * セットアップ本体。
 * @param {{demo: boolean, samplePdfs: boolean, today?: string}} opt
 */
function setupSpreadsheet_(opt) {
  const ss = getSs_();
  const today = opt.today || todayYmd_();
  ss.setSpreadsheetTimeZone(APP.timeZone);
  try { ss.setSpreadsheetLocale('ja_JP'); } catch (e) { /* ロケール変更不可の環境では無視 */ }

  // 既存の設定（フォルダ ID など）は引き継ぐ
  let keepFolderId = '';
  if (ss.getSheetByName(SHEET.settings)) {
    try { keepFolderId = getSettings_().driveFolderId || ''; } catch (e) { keepFolderId = ''; }
  }

  const sheets = {};
  SHEET_ORDER.forEach(function (name, i) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, i);
    resetSheet_(sh);
    sheets[name] = sh;
  });
  // 並び順をそろえ、既定の「シート1」などを消す
  SHEET_ORDER.forEach(function (name, i) {
    ss.setActiveSheet(sheets[name]);
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheets().forEach(function (sh) {
    if (SHEET_ORDER.indexOf(sh.getName()) < 0 && sh.getName() !== SHEET.selfTest) ss.deleteSheet(sh);
  });

  const formulas = columnFormulas_();
  buildTable_('deals', formulas.deals, { freezeCols: 1, rows: 600 });
  buildTable_('lines', formulas.lines, { freezeCols: 1, rows: 1500 });
  buildTable_('payments', formulas.payments, { freezeCols: 1, rows: 600 });
  buildTable_('customers', {}, { freezeCols: 2, rows: 200 });
  buildTable_('products', {}, { freezeCols: 2, rows: 200 });
  buildTable_('log', {}, { freezeCols: 0, rows: 300 });
  applyValidations_();
  applyConditionalFormats_();
  buildSettings_(keepFolderId);
  buildGuide_();

  let demo = null;
  if (opt.demo) demo = insertDemoData_(today);
  buildDashboard_();
  SpreadsheetApp.flush();

  const messages = [];
  if (demo) {
    messages.push('デモデータ：顧客' + demo.customers + '社・商品' + demo.products + '件・案件' + demo.deals + '件・明細' + demo.lines + '行・入金' + demo.payments + '件');
  }
  if (opt.samplePdfs && demo) {
    try {
      const made = createSamplePdfs_();
      messages.push('見本PDF：' + made.join('、') + ' を作成し Google ドライブに保存しました');
    } catch (e) {
      messages.push('見本PDFの作成はスキップしました（' + e.message + '）');
    }
  }
  messages.push('メニュー「' + APP.menuName + '」から見積書・請求書の作成やメール下書きを試せます。');
  writeLog_('初期セットアップ', opt.demo ? 'デモデータ' : '空の状態', true, messages.join(' / '));
  ss.setActiveSheet(sheets[SHEET.dashboard]);
  return { ok: true, message: messages.join('\n'), demo: demo };
}

// ---------- シートの土台 ----------

function resetSheet_(sh) {
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  sh.clearConditionalFormatRules();
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sh.clear();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.setTabColor(THEME.sheetTabs[sh.getName()] || null);
  sh.setHiddenGridlines(true);
}

function fitSize_(sh, rows, cols) {
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

function buildTable_(tableKey, formulas, opt) {
  const sh = getSheet_(SHEET[tableKey]);
  const cols = COLUMNS[tableKey];
  fitSize_(sh, opt.rows, cols.length);
  const rows = sh.getMaxRows();

  const all = sh.getRange(1, 1, rows, cols.length);
  all.setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.ink).setVerticalAlignment('middle');
  sh.setRowHeights(2, rows - 1, 26);
  sh.setRowHeight(1, 34);

  const header = sh.getRange(1, 1, 1, cols.length);
  header.setBackground(THEME.header).setFontColor(THEME.headerText).setFontWeight('bold').setHorizontalAlignment('center');

  cols.forEach(function (c, i) {
    const col = i + 1;
    sh.setColumnWidth(col, c.width);
    const body = sh.getRange(2, col, rows - 1, 1);
    if (c.calc) {
      sh.getRange(1, col).setFormula(formulas[c.key]).setBackground(THEME.calcHeader);
      body.setFontColor('#334155');
    } else {
      sh.getRange(1, col).setValue(c.label);
    }
    if (NUMBER_FORMAT[c.type]) body.setNumberFormat(NUMBER_FORMAT[c.type]);
    if (c.type === 'yen' || c.type === 'number' || c.type === 'percent') body.setHorizontalAlignment('right');
    if (c.type === 'date' || c.type === 'checkbox') body.setHorizontalAlignment('center');
    if (c.type === 'checkbox') body.insertCheckboxes();
  });

  if (tableKey === 'deals') {
    sh.getRange(2, COL.deals.overdueDays, rows - 1, 1).setNumberFormat('"超過 "0"日";"あと "0"日";"当日"').setHorizontalAlignment('center');
    [COL.deals.orderStatus, COL.deals.paymentStatus, COL.deals.quotePdf, COL.deals.invoicePdf].forEach(function (c) {
      sh.getRange(2, c, rows - 1, 1).setHorizontalAlignment('center');
    });
  }

  const banding = all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banding.setHeaderRowColor(THEME.header).setFirstRowColor('#ffffff').setSecondRowColor(THEME.band);
  // 計算列の見出しは少し明るい色にして「入力不要」を示す
  cols.forEach(function (c, i) {
    if (c.calc) sh.getRange(1, i + 1).setBackground(THEME.calcHeader).setNote('自動計算の列です（入力不要）');
  });

  sh.setFrozenRows(1);
  if (opt.freezeCols) sh.setFrozenColumns(opt.freezeCols);

  // 計算列は誤って上書きしないよう警告付きで保護
  cols.forEach(function (c, i) {
    if (!c.calc) return;
    const p = sh.getRange(1, i + 1, rows, 1).protect();
    p.setDescription('自動計算列：' + c.label);
    p.setWarningOnly(true);
  });
}

function applyValidations_() {
  const ss = getSs_();
  const deals = getSheet_(SHEET.deals);
  const lines = getSheet_(SHEET.lines);
  const pays = getSheet_(SHEET.payments);
  const customers = getSheet_(SHEET.customers);
  const products = getSheet_(SHEET.products);

  const rule = function () { return SpreadsheetApp.newDataValidation(); };
  const body = function (sh, col) { return sh.getRange(2, col, sh.getMaxRows() - 1, 1); };

  body(deals, COL.deals.customer).setDataValidation(rule().requireValueInRange(customers.getRange('B2:B'), true).setAllowInvalid(false).setHelpText('顧客マスタの会社名から選んでください').build());
  body(deals, COL.deals.orderStatus).setDataValidation(rule().requireValueInList(ORDER_STATUS, true).setAllowInvalid(false).build());
  ['receivedOn', 'quoteOn', 'invoiceOn', 'dueOn'].forEach(function (k) {
    body(deals, COL.deals[k]).setDataValidation(rule().requireDate().setAllowInvalid(false).setHelpText('日付を入力してください').build());
  });

  body(lines, COL.lines.dealId).setDataValidation(rule().requireValueInRange(deals.getRange('A2:A'), true).setAllowInvalid(false).setHelpText('案件・請求シートの案件番号から選んでください').build());
  body(lines, COL.lines.product).setDataValidation(rule().requireValueInRange(products.getRange('B2:B'), true).setAllowInvalid(true).setHelpText('商品マスタから選ぶと単位・単価・税率が自動で入ります（マスタにない品目も入力できます）').build());
  body(lines, COL.lines.qty).setDataValidation(rule().requireNumberGreaterThan(0).setAllowInvalid(false).setHelpText('0 より大きい数値').build());
  body(lines, COL.lines.unitPrice).setDataValidation(rule().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  body(lines, COL.lines.taxRate).setDataValidation(rule().requireFormulaSatisfied('=OR(H2=0.1,H2=0.08)').setAllowInvalid(false).setHelpText('10% または 8%（軽減税率）').build());

  body(pays, COL.payments.invoiceNo).setDataValidation(rule().requireValueInRange(deals.getRange('M2:M'), true).setAllowInvalid(false).setHelpText('案件・請求シートの請求番号から選んでください').build());
  body(pays, COL.payments.amount).setDataValidation(rule().requireNumberGreaterThan(0).setAllowInvalid(false).build());
  body(pays, COL.payments.paidOn).setDataValidation(rule().requireDate().setAllowInvalid(false).build());
  body(pays, COL.payments.method).setDataValidation(rule().requireValueInList(PAYMENT_METHODS, true).setAllowInvalid(false).build());

  body(customers, COL.customers.terms).setDataValidation(rule().requireValueInList(PAYMENT_TERMS.map(function (t) { return t.label; }), true).setAllowInvalid(false).build());
  body(customers, COL.customers.name).setDataValidation(rule().requireFormulaSatisfied('=COUNTIF($B$2:$B,B2)=1').setAllowInvalid(false).setHelpText('会社名は重複できません（案件と結びつけるキーです）').build());
  body(customers, COL.customers.email).setDataValidation(rule().requireTextIsEmail().setAllowInvalid(false).build());
  body(products, COL.products.taxRate).setDataValidation(rule().requireFormulaSatisfied('=OR(F2=0.1,F2=0.08)').setAllowInvalid(false).setHelpText('10% または 8%（軽減税率）').build());
  body(products, COL.products.name).setDataValidation(rule().requireFormulaSatisfied('=COUNTIF($B$2:$B,B2)=1').setAllowInvalid(false).setHelpText('商品名は重複できません').build());
  return ss;
}

function statusRule_(range, text) {
  const c = THEME.status[text];
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(c.bg).setFontColor(c.fg).setBold(true).setRanges([range]).build();
}

function applyConditionalFormats_() {
  const deals = getSheet_(SHEET.deals);
  const n = deals.getMaxRows() - 1;
  const lastCol = COLUMNS.deals.length;
  const statusRange = deals.getRange(2, COL.deals.paymentStatus, n, 1);
  const orderRange = deals.getRange(2, COL.deals.orderStatus, n, 1);
  const daysRange = deals.getRange(2, COL.deals.overdueDays, n, 1);
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="失注"').setFontColor('#9ca3af').setRanges([deals.getRange(2, 1, n, lastCol)]).build());
  PAYMENT_STATUS.forEach(function (s) { rules.push(statusRule_(statusRange, s)); });
  ORDER_STATUS.forEach(function (s) { rules.push(statusRule_(orderRange, s)); });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(U2),U2>0)').setFontColor('#b42318').setBold(true).setRanges([daysRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(U2),U2<=0,U2>=-7)').setFontColor('#b54708').setRanges([daysRange]).build());
  deals.setConditionalFormatRules(rules);

  const lines = getSheet_(SHEET.lines);
  lines.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('案件が見つかりません').setFontColor('#b42318').setRanges([lines.getRange(2, COL.lines.customer, lines.getMaxRows() - 1, 1)]).build(),
  ]);
  const pays = getSheet_(SHEET.payments);
  pays.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('請求番号が見つかりません').setFontColor('#b42318').setRanges([pays.getRange(2, COL.payments.customer, pays.getMaxRows() - 1, 1)]).build(),
  ]);
}

// ---------- 設定シート ----------

/** 設定シートで、その項目が何行目にあるか（ダッシュボードの数式から参照する） */
function settingsRow_(key) {
  let row = 2;
  for (let i = 0; i < SETTINGS_SPEC.length; i++) {
    if (SETTINGS_SPEC[i].key === key) return row;
    row++;
  }
  throw new Error('設定項目 ' + key + ' がありません');
}

function buildSettings_(keepFolderId) {
  const sh = getSheet_(SHEET.settings);
  fitSize_(sh, SETTINGS_SPEC.length + 3, 3);
  sh.getRange(1, 1, sh.getMaxRows(), 3).setFontFamily(THEME.font).setFontSize(10).setVerticalAlignment('middle').setFontColor(THEME.ink);
  sh.getRange(1, 1, 1, 3).setValues([['項目', '値', '説明']]).setBackground(THEME.header).setFontColor(THEME.headerText).setFontWeight('bold');
  sh.setRowHeight(1, 34);
  const values = SETTINGS_SPEC.map(function (s) {
    if (s.section) return [s.section, '', ''];
    let v = s.value;
    if (s.key === 'driveFolderId' && keepFolderId) v = keepFolderId;
    return [s.label, v, s.help || ''];
  });
  sh.getRange(2, 1, values.length, 3).setValues(values);
  sh.setRowHeights(2, values.length, 28);
  SETTINGS_SPEC.forEach(function (s, i) {
    const r = i + 2;
    if (s.section) {
      sh.getRange(r, 1, 1, 3).setBackground('#eef2f7').setFontWeight('bold').setFontColor('#334155');
    } else {
      sh.getRange(r, 2).setBackground('#fffdf5').setBorder(null, null, true, null, null, null, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(r, 3).setFontColor(THEME.muted).setFontSize(9);
      if (typeof s.value === 'boolean') sh.getRange(r, 2).insertCheckboxes().setValue(s.value);
      if (typeof s.value === 'number') sh.getRange(r, 2).setHorizontalAlignment('left');
      if (s.key === 'mailSignature') { sh.setRowHeight(r, 90); sh.getRange(r, 2).setWrap(true); }
    }
  });
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 460);
  sh.getRange(1, 3, sh.getMaxRows(), 1).setWrap(true);
  sh.setFrozenRows(1);
}

// ---------- 使い方シート ----------

function buildGuide_() {
  const sh = getSheet_(SHEET.guide);
  const rows = [
    ['見積・請求・入金管理システム　使い方', ''],
    ['見積から入金確認までを1つのスプレッドシートで管理します。上部メニュー「' + APP.menuName + '」から操作します。', ''],
    ['', ''],
    ['手順', '内容'],
    ['① 案件を登録', '「案件・請求」で顧客名を選ぶと、案件番号・受付日・受注状況（見積中）が自動で入ります。件名を入力します。'],
    ['② 明細を入力', '「売上明細」で案件番号と商品を選ぶと、単位・単価・税率が自動で入ります。数量を入れると金額が計算され、案件の税抜・消費税・合計に反映されます。'],
    ['③ 見積書PDF', '案件の行を選び「' + APP.menuName + ' › 見積書PDFを作成」。見積番号を採番し、Googleドライブに保存してリンクを貼ります。'],
    ['④ 請求書PDF', '同じく「請求書PDFを発行」。請求番号を採番し、顧客の支払条件から支払期限を自動計算します。'],
    ['⑤ メール下書き', '「請求メールの下書きを作成」で、PDFを添付した Gmail の下書きを作ります。自動送信はしません。Gmail で内容を確認してから送ってください。'],
    ['⑥ 入金を記録', '「入金記録」で請求番号を選ぶと、入金日と残額が自動で入ります（操作パネルからも記録できます）。入金ステータスが自動で切り替わります。'],
    ['⑦ 未入金アラート', '支払期限を過ぎた案件は赤で表示され、ダッシュボードに一覧が出ます。期限超過分の督促メール下書きもまとめて作れます。毎朝の自動チェックを有効にすると操作ログに記録されます。'],
    ['⑧ ダッシュボード', '月別の売上・入金、顧客別ランキング、未入金残高を自動集計します。'],
    ['', ''],
    ['見出しの色', '濃い紺の列は入力する列、少し明るい紺の列は自動計算の列です（入力不要・上書きすると警告が出ます）。'],
    ['操作パネル', '「' + APP.menuName + ' › 操作パネルを開く」で、選択中の案件の状況確認と操作をボタンで行えます。'],
    ['デモデータ', '会社・顧客・人名・住所はすべて架空です。顧客メールは予約ドメイン（.example）なので、誤って送信しても届きません。'],
    ['本番で使うとき', '「' + APP.menuName + ' › 管理 › 空の状態で作り直す」を実行し、「設定」シートに自社情報と振込先を入れてください。'],
  ];
  fitSize_(sh, rows.length + 2, 2);
  sh.getRange(1, 1, sh.getMaxRows(), 2).setFontFamily(THEME.font).setFontSize(10).setVerticalAlignment('middle').setWrap(true).setFontColor(THEME.ink);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).merge().setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1, 1, 2).merge().setFontColor(THEME.muted);
  sh.getRange(4, 1, 1, 2).setBackground(THEME.header).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(5, 1, rows.length - 4, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 760);
  sh.setRowHeight(1, 44);
  sh.setRowHeights(5, rows.length - 4, 40);
}

// ---------- デモデータ ----------

function insertDemoData_(today) {
  const data = generateDemoData(today);
  writeRows_('customers', 2, data.customers);
  writeRows_('products', 2, data.products);
  writeRows_('deals', 2, data.deals.map(function (d) {
    return {
      id: d.id, receivedOn: d.receivedOn, customer: d.customer, subject: d.subject, orderStatus: d.orderStatus,
      quoteNo: d.quoteNo || '', quoteOn: d.quoteOn || '', invoiceNo: d.invoiceNo || '', invoiceOn: d.invoiceOn || '', dueOn: d.dueOn || '', note: d.note || '',
    };
  }));
  const lines = [];
  data.deals.forEach(function (d) {
    d.lines.forEach(function (l) {
      lines.push({ dealId: d.id, product: l.product, detail: l.detail, qty: l.qty, unit: l.unit, unitPrice: l.unitPrice, taxRate: l.taxRate, note: '' });
    });
  });
  writeRows_('lines', 2, lines);
  writeRows_('payments', 2, data.payments);
  return { customers: data.customers.length, products: data.products.length, deals: data.deals.length, lines: lines.length, payments: data.payments.length };
}

/** セットアップ直後に見本の PDF を 2 件作る（直近の請求と、回答待ちの見積） */
function createSamplePdfs_() {
  const deals = readTable_('deals');
  const made = [];
  const invoiced = deals.filter(function (d) { return d.invoiceNo; }).sort(function (a, b) { return a.invoiceOn < b.invoiceOn ? 1 : -1; })[0];
  if (invoiced) made.push(createDocument_('invoice', invoiced._row, { force: true }).number);
  const quoting = deals.filter(function (d) { return d.orderStatus === '見積中'; }).sort(function (a, b) { return a.quoteOn < b.quoteOn ? 1 : -1; })[0];
  if (quoting) made.push(createDocument_('quote', quoting._row, { force: true }).number);
  return made;
}

// ---------- ダッシュボード ----------

function buildDashboard_() {
  const sh = getSheet_(SHEET.dashboard);
  const D = q_(SHEET.deals);
  const P = q_(SHEET.payments);
  const S = q_(SHEET.settings);
  fitSize_(sh, 70, 14);
  const widths = [16, 120, 170, 170, 100, 110, 92, 92, 92, 92, 92, 92, 92, 16];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, 70, 14).setFontFamily(THEME.font).setFontSize(10).setFontColor(THEME.ink).setVerticalAlignment('middle');
  sh.setRowHeights(1, 70, 24);
  sh.setRowHeight(1, 10);

  // 見出し
  sh.getRange('B2:H2').merge().setValue('見積・請求・入金ダッシュボード').setFontSize(18).setFontWeight('bold');
  sh.setRowHeight(2, 40);
  sh.getRange('I2:M2').merge().setFormula('=' + S + 'B' + settingsRow_('companyName') + '&"　｜　"&TEXT(TODAY(),"yyyy/mm/dd")&" 時点"').setHorizontalAlignment('right').setFontColor(THEME.muted);
  sh.getRange('B3:M3').merge().setValue('案件・明細・入金のシートに入力すると自動で集計されます。期限超過は赤で表示します。').setFontColor(THEME.muted).setFontSize(9);

  // KPI カード（4 枚）
  const monthStart = '(EOMONTH(TODAY(),-1)+1)';
  const monthEnd = 'EOMONTH(TODAY(),0)';
  const prevStart = '(EOMONTH(TODAY(),-2)+1)';
  const prevEnd = 'EOMONTH(TODAY(),-1)';
  const salesBetween = function (a, b) { return 'SUMIFS(' + D + 'G:G,' + D + 'N:N,">="&' + a + ',' + D + 'N:N,"<="&' + b + ')'; };
  // 残高が正の請求 = 未入金・一部入金・期限超過（未請求・失注は空欄、入金済は 0 以下）。
  // ※ SUMIFS に配列の条件 {"未入金",…} を渡すと、Sheets では ARRAYFORMULA なしだと先頭の条件しか使われない
  const cards = [
    { range: ['B', 'C'], label: '今月の売上（税抜・請求ベース）', value: '=' + salesBetween(monthStart, monthEnd), sub: '="前月 "&TEXT(' + salesBetween(prevStart, prevEnd) + ',"¥#,##0")&"　／　今月の請求 "&COUNTIFS(' + D + 'N:N,">="&' + monthStart + ',' + D + 'N:N,"<="&' + monthEnd + ')&"件"', color: THEME.ink },
    { range: ['D', 'E'], label: '見積中（回答待ち・税込）', value: '=SUMIFS(' + D + 'I:I,' + D + 'F:F,"見積中")', sub: '=COUNTIF(' + D + 'F:F,"見積中")&"件の見積が回答待ち"', color: THEME.ink },
    { range: ['F', 'H'], label: '未入金残高（税込）', value: '=SUMIFS(' + D + 'S:S,' + D + 'S:S,">0")', sub: '=COUNTIF(' + D + 'S:S,">0")&"件の請求が入金待ち"', color: THEME.ink },
    { range: ['I', 'K'], label: '期限超過（税込）', value: '=SUMIFS(' + D + 'S:S,' + D + 'T:T,"期限超過")', sub: '=IF(COUNTIF(' + D + 'T:T,"期限超過")=0,"期限超過はありません",COUNTIF(' + D + 'T:T,"期限超過")&"件　／　最長 "&MAXIFS(' + D + 'U:U,' + D + 'T:T,"期限超過")&"日超過")', color: '#b42318' },
  ];
  cards.forEach(function (c) {
    const a = c.range[0];
    const b = c.range[1];
    sh.getRange(a + '5:' + b + '5').merge().setValue(c.label).setFontColor(THEME.muted).setFontSize(9);
    sh.getRange(a + '6:' + b + '6').merge().setFormula(c.value).setFontSize(20).setFontWeight('bold').setNumberFormat(NUMBER_FORMAT.yen).setHorizontalAlignment('left').setFontColor(c.color);
    sh.getRange(a + '7:' + b + '7').merge().setFormula(c.sub).setFontColor(THEME.muted).setFontSize(9);
    const card = sh.getRange(a + '5:' + b + '7');
    card.setBackground('#f8fafc').setBorder(true, true, true, true, null, null, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(a + '5:' + a + '7').setBorder(null, true, null, null, null, null, c.color === THEME.ink ? THEME.accent : c.color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.setRowHeight(5, 26);
  sh.setRowHeight(6, 40);
  sh.setRowHeight(7, 24);

  const section = function (a1, text) {
    sh.getRange(a1).setValue(text).setFontSize(12).setFontWeight('bold');
    const row = sh.getRange(a1).getRow();
    sh.getRange(row, 2, 1, 12).setBorder(null, null, true, null, null, null, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 32);
  };
  const tableHeader = function (range) {
    range.setBackground('#eef2f7').setFontWeight('bold').setFontColor('#334155').setHorizontalAlignment('center');
  };

  // 月別（直近 12 か月）
  section('B9', '月別の売上と入金（直近12か月）');
  sh.getRange('B10:E10').setValues([['月', '売上（税抜）', '入金額', '請求件数']]);
  tableHeader(sh.getRange('B10:E10'));
  const monthRows = [];
  for (let i = 0; i < 12; i++) {
    const r = 11 + i;
    const start = 'DATEVALUE($B' + r + '&"/01")';
    monthRows.push([
      '=TEXT(EDATE(EOMONTH(TODAY(),-1)+1,' + (i - 11) + '),"yyyy/mm")',
      '=SUMIFS(' + D + '$G:$G,' + D + '$N:$N,">="&' + start + ',' + D + '$N:$N,"<"&EDATE(' + start + ',1))',
      '=SUMIFS(' + P + '$D:$D,' + P + '$A:$A,">="&' + start + ',' + P + '$A:$A,"<"&EDATE(' + start + ',1))',
      '=COUNTIFS(' + D + '$N:$N,">="&' + start + ',' + D + '$N:$N,"<"&EDATE(' + start + ',1))',
    ]);
  }
  sh.getRange('B11:E22').setFormulas(monthRows);
  sh.getRange('B23:E23').setFormulas([['="合計"', '=SUM(C11:C22)', '=SUM(D11:D22)', '=SUM(E11:E22)']]).setFontWeight('bold').setBorder(true, null, null, null, null, null, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('C11:D23').setNumberFormat(NUMBER_FORMAT.yen).setHorizontalAlignment('right');
  sh.getRange('E11:E23').setNumberFormat('0"件"').setHorizontalAlignment('right');
  sh.getRange('B11:B23').setHorizontalAlignment('center');
  sh.getRange('B11:E22').setBorder(null, null, null, null, null, true, '#f1f5f9', SpreadsheetApp.BorderStyle.SOLID);

  // 顧客別ランキング
  section('B25', '顧客別の売上ランキング（直近12か月・請求ベース・税抜）');
  sh.getRange('B26:E26').setValues([['顧客名', '', '売上（税抜）', '構成比']]);
  sh.getRange('B26:C26').merge();
  tableHeader(sh.getRange('B26:E26'));
  const from12 = 'TEXT(EDATE(EOMONTH(TODAY(),-1)+1,-11),"yyyy-mm-dd")';
  // 顧客名を B:C の幅で見せるため、QUERY の結果（顧客名・売上）を B と D に分けて置く
  sh.getRange('F27').setFormula('=IFERROR(QUERY(' + D + 'A2:V,"select C, sum(G) where M <> \'\' and N >= date \'"&' + from12 + '&"\' group by C order by sum(G) desc limit 10 label sum(G) \'\'",0),)');
  const rankRows = [];
  for (let i = 0; i < 10; i++) {
    const r = 27 + i;
    rankRows.push(['=IF($F' + r + '="","",$F' + r + ')', '', '=IF($F' + r + '="","",$G' + r + ')', '=IF($F' + r + '="","",$G' + r + '/SUM($G$27:$G$36))']);
  }
  sh.getRange('B27:E36').setFormulas(rankRows);
  for (let i = 0; i < 10; i++) sh.getRange(27 + i, 2, 1, 2).merge();
  sh.getRange('D27:D36').setNumberFormat(NUMBER_FORMAT.yen).setHorizontalAlignment('right');
  sh.getRange('E27:E36').setNumberFormat('0.0%').setHorizontalAlignment('right');
  sh.getRange('B27:E36').setBorder(null, null, null, null, null, true, '#f1f5f9', SpreadsheetApp.BorderStyle.SOLID);
  // QUERY の作業列は見えないよう文字色を背景と同じにして小さくする（グラフの元データ）
  sh.getRange('F27:G36').setFontColor('#ffffff').setFontSize(6).setNumberFormat('0');

  // 未入金アラート
  const soonRow = settingsRow_('alertSoonDays');
  section('B39', '未入金アラート（期限超過・期限まで' + '設定日数以内）');
  sh.getRange('B39').setFormula('="未入金アラート（期限超過と、期限まで"&' + S + 'B' + soonRow + '&"日以内の未入金）"');
  sh.getRange('B40:H40').setValues([['請求番号', '顧客名', '件名', '支払期限', '未入金残高', 'ステータス', '期限まで/超過']]);
  tableHeader(sh.getRange('B40:H40'));
  const t = function (col) { return D + col + '2:' + col; };
  sh.getRange('B41').setFormula(
    '=IFERROR(SORT(FILTER({' + [t('M'), t('C'), t('E'), t('O'), t('S'), t('T'), t('U')].join(',') + '},' +
    '(' + t('T') + '="期限超過")+((' + t('T') + '="未入金")+(' + t('T') + '="一部入金"))*(' + t('O') + '<>"")*(' + t('O') + '<=TODAY()+' + S + 'B' + soonRow + ')),7,FALSE),"該当はありません")'
  );
  sh.getRange('E41:E65').setNumberFormat(NUMBER_FORMAT.date).setHorizontalAlignment('center');
  sh.getRange('F41:F65').setNumberFormat(NUMBER_FORMAT.yen).setHorizontalAlignment('right');
  sh.getRange('G41:G65').setHorizontalAlignment('center');
  sh.getRange('H41:H65').setNumberFormat('"超過 "0"日";"あと "0"日";"当日"').setHorizontalAlignment('center');
  sh.getRange('B41:H65').setBorder(null, null, null, null, null, true, '#f1f5f9', SpreadsheetApp.BorderStyle.SOLID);
  const alertStatus = sh.getRange('G41:G65');
  sh.setConditionalFormatRules([
    statusRule_(alertStatus, '期限超過'),
    statusRule_(alertStatus, '未入金'),
    statusRule_(alertStatus, '一部入金'),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(H41),H41>0)').setFontColor('#b42318').setBold(true).setRanges([sh.getRange('H41:H65')]).build(),
  ]);

  // グラフ（1 グラフ 1 軸。金額の 2 系列は同じ単位なので同じ軸に並べる）
  const monthly = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('B10:D22'))
    .setNumHeaders(1)
    .setPosition(9, 6, 8, 4)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'end', textStyle: { color: '#374151', fontSize: 11 } })
    .setOption('colors', [THEME.chart.sales, THEME.chart.paid])
    .setOption('vAxis', { format: 'short', gridlines: { color: '#eef2f7' }, minValue: 0, textStyle: { color: '#6b7280', fontSize: 10 } })
    .setOption('hAxis', { textStyle: { color: '#6b7280', fontSize: 10 } })
    .setOption('bar', { groupWidth: '62%' })
    .setOption('chartArea', { left: 64, top: 36, width: '84%', height: '72%' })
    .setOption('width', 660)
    .setOption('height', 330)
    .build();
  sh.insertChart(monthly);

  const ranking = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('F27:G36'))
    .setNumHeaders(0)
    .setPosition(25, 6, 8, 4)
    .setOption('title', '')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [THEME.chart.sales])
    .setOption('hAxis', { format: 'short', gridlines: { color: '#eef2f7' }, minValue: 0, textStyle: { color: '#6b7280', fontSize: 10 } })
    .setOption('vAxis', { textStyle: { color: '#374151', fontSize: 10 } })
    .setOption('chartArea', { left: 180, top: 16, width: '68%', height: '86%' })
    .setOption('width', 660)
    .setOption('height', 300)
    .build();
  sh.insertChart(ranking);

  sh.setFrozenRows(0);
}
