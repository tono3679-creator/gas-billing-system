/**
 * 模擬スプレッドシートでの通し動作：
 * セットアップ → 新規案件の入力（自動入力）→ 見積 PDF → 請求 PDF → メール下書き → 入金 → 督促 → 毎朝チェック
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGas } = require('./gas-mock');

function select(ctx, sheetName, row, col) {
  const sh = ctx.ss.getSheetByName(sheetName);
  ctx.ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(row, col || 1));
}

/** 人がセルを編集したときと同じく値を入れて onEdit を呼ぶ */
function edit(ctx, sheetName, row, col, value) {
  const sh = ctx.ss.getSheetByName(sheetName);
  const range = sh.getRange(row, col);
  range.setValue(value);
  ctx.g.onEdit({ range, value });
}

/** 模擬シートでは数式が計算されないので、計算列の値を JS で埋める（本物では Sheets が計算する） */
function fillCalcColumns(ctx) {
  const { g } = ctx;
  const book = g.loadBook_();
  const dealsSh = ctx.ss.getSheetByName(g.SHEET.deals);
  const today = g.todayYmd_();
  for (const d of book.deals) {
    const st = g.computePaymentStatus(g.dealForStatus_(d), today);
    const c = g.COL.deals;
    const set = (col, v) => dealsSh._set(d._row, col, v);
    set(c.customerId, (book.customerByName[d.customer] || {}).id || '未登録');
    set(c.subtotal, d.calc.subtotal); set(c.tax, d.calc.tax); set(c.total, d.calc.total);
    set(c.paid, d.invoiceNo ? d.calcPaid : ''); set(c.balance, d.invoiceNo ? d.calc.total - d.calcPaid : '');
    set(c.paymentStatus, st.status); set(c.overdueDays, st.daysPastDue);
  }
}

test('通し動作：セットアップから入金・督促まで', () => {
  const ctx = loadGas();
  const { g, ss, drive, gmail, ui } = ctx;

  // 1) セットアップ
  g.onOpen();
  assert.ok(ui.menuAdded, 'メニューが追加される');
  assert.match(ss.toasts.at(-1).msg, /初期セットアップ/, '未セットアップなら案内のトースト');
  const res = g.setupSpreadsheet_({ demo: true, samplePdfs: true });
  assert.ok(res.ok);
  assert.deepEqual(ss.getSheets().map((s) => s.getName()), [...g.SHEET_ORDER]);
  assert.match(res.message, /見本PDF：INV-\d{6}-\d{3}、Q-\d{6}-\d{3}/);
  const dash = ss.getSheetByName(g.SHEET.dashboard);
  assert.equal(dash.charts.length, 2, 'グラフ 2 つ');
  // 埋め込みグラフのオプションはドット区切りのキーと単純な値で渡す（オブジェクトで渡すと実機で空白になった）
  for (const c of dash.charts) {
    for (const [k, v] of Object.entries(c.opts)) {
      assert.ok(Array.isArray(v) || typeof v !== 'object', `グラフのオプション ${k} がオブジェクト`);
      assert.ok(!/%$/.test(String(v)), `グラフのオプション ${k} に割合指定`);
    }
  }
  // Sheets の SUMIFS/COUNTIF は ARRAYFORMULA なしで配列の条件 {..} を渡すと先頭しか使わない（実機で発覚）
  for (const f of dash.formulas.values()) {
    if (/(SUMIFS|COUNTIFS?|MAXIFS)\([^)]*\{"/.test(f)) assert.match(f, /ARRAYFORMULA/, f);
  }
  const pdfs = [...drive.files.values()];
  assert.equal(pdfs.length, 2);
  assert.match(pdfs[0].name, /^請求書_INV-\d{6}-\d{3}_.+\.pdf$/);
  const folderId = g.getSettings_().driveFolderId;
  assert.ok(folderId, '保存先フォルダ ID が設定に記録される');
  assert.ok(drive.folders.get(folderId).name.startsWith('見積・請求管理'));

  const deals0 = g.readTable_('deals');
  assert.ok(deals0.length >= 50);
  assert.equal(deals0.filter((d) => d.invoicePdf || d.quotePdf).length, 0, 'PDF 列は =HYPERLINK なので値は読み取らない形');
  // =HYPERLINK が入っていること
  const linked = deals0.filter((d) => g.getLinkedFileId_('deals', d._row, 'invoicePdf'));
  assert.equal(linked.length, 1);

  // 2) 新しい案件：顧客を選ぶと案件番号・受付日・受注状況が入る
  const dealsSh = ss.getSheetByName(g.SHEET.deals);
  const newRow = g.lastDataRow_(dealsSh, 1) + 1;
  edit(ctx, g.SHEET.deals, newRow, g.COL.deals.customer, '株式会社北斗フーズ');
  dealsSh.getRange(newRow, g.COL.deals.subject).setValue('業務改善ダッシュボード構築');
  const created = g.readRow_('deals', newRow);
  const year = g.todayYmd_().slice(0, 4);
  assert.match(created.id, new RegExp(`^PJ-${year}-\\d{4}$`));
  assert.equal(created.receivedOn, g.todayYmd_());
  assert.equal(created.orderStatus, '見積中');

  // 3) 明細：商品を選ぶと単位・単価・税率が入る
  const linesSh = ss.getSheetByName(g.SHEET.lines);
  let lr = g.lastDataRow_(linesSh, 1) + 1;
  linesSh.getRange(lr, g.COL.lines.dealId).setValue(created.id);
  edit(ctx, g.SHEET.lines, lr, g.COL.lines.product, 'DX・業務改善コンサルティング');
  linesSh.getRange(lr, g.COL.lines.qty).setValue(12);
  let line = g.readRow_('lines', lr);
  assert.equal(line.unit, '時間'); assert.equal(line.unitPrice, 15000); assert.equal(line.taxRate, 0.1);
  lr += 1;
  linesSh.getRange(lr, g.COL.lines.dealId).setValue(created.id);
  edit(ctx, g.SHEET.lines, lr, g.COL.lines.product, '打合せ用 飲料・軽食');
  line = g.readRow_('lines', lr);
  assert.equal(line.qty, 1, '数量が空なら 1'); assert.equal(line.taxRate, 0.08);

  // 4) 見積書 PDF（操作パネル経由）
  select(ctx, g.SHEET.lines, lr, 3); // 明細の行を選んでいても案件を特定できる
  const state = g.panelGetState();
  assert.equal(state.deal.id, created.id);
  assert.equal(state.deal.total, 180000 + 18000 + 3000 + 240);
  let r = g.panelAction('quote', { row: newRow });
  assert.ok(r.ok, r.error);
  assert.match(r.number, /^Q-\d{6}-\d{3}$/);
  const allQuotes = g.readTable_('deals').map((d) => d.quoteNo).filter(Boolean);
  assert.equal(new Set(allQuotes).size, allQuotes.length, '見積番号は重複しない');
  const quoteFile = drive.files.get(r.fileId);
  assert.match(quoteFile.blob.content, /御見積書/);
  assert.match(quoteFile.blob.content, /業務改善ダッシュボード構築/);
  assert.match(quoteFile.blob.content, /¥201,240/);
  assert.match(quoteFile.blob.content, /有効期限/);
  // 2 回目は確認が必要、force で作り直すと古い PDF はゴミ箱へ
  r = g.panelAction('quote', { row: newRow });
  assert.equal(r.needsConfirm, true);
  r = g.panelAction('quote', { row: newRow, force: true });
  assert.ok(r.ok);
  assert.equal(quoteFile.trashed, true);
  assert.equal(g.readRow_('deals', newRow).quoteNo, allQuotes.find((q) => q === r.number), '再作成でも番号は同じ');

  // 5) 請求書 PDF（メニュー経由）：受注になり、期限は顧客の支払条件（月末締め翌月末払い）
  select(ctx, g.SHEET.deals, newRow, 5);
  g.menuCreateInvoice();
  const inv = g.readRow_('deals', newRow);
  assert.match(inv.invoiceNo, /^INV-\d{6}-\d{3}$/);
  assert.equal(inv.orderStatus, '受注');
  assert.equal(inv.invoiceOn, g.todayYmd_());
  assert.equal(inv.dueOn, g.calcDueDate(g.todayYmd_(), '月末締め翌月末払い'));
  assert.equal(ui.dialogs.at(-1).title, '請求書を作成しました');
  const invFile = [...drive.files.values()].find((f) => f.name.includes(inv.invoiceNo));
  assert.match(invFile.blob.content, /御請求書/);
  assert.match(invFile.blob.content, /お振込先/);
  assert.match(invFile.blob.content, /8%対象（軽減税率）/);
  assert.ok(drive.folders.get(invFile.parent.id).name === g.todayYmd_().slice(0, 7), '月別フォルダに保存');

  // 6) 請求メールの下書き（送信はしない）
  g.menuDraftInvoice();
  assert.equal(gmail.drafts.length, 1);
  const draft = gmail.drafts[0];
  assert.equal(draft.to, 'yamamoto@hokuto-foods.example');
  assert.match(draft.subject, /【ご請求書送付】業務改善ダッシュボード構築/);
  assert.match(draft.body, /株式会社北斗フーズ\n営業企画部 山本 健一 様/);
  assert.match(draft.body, /¥201,240/);
  assert.equal(draft.options.attachments.length, 1);
  assert.match(draft.options.attachments[0].getName(), /^請求書_/);
  assert.match(g.readRow_('deals', newRow).mailDraft, /^請求書 下書き /);
  assert.equal(gmail.sent.length, 0, 'メールは送信しない');

  // 7) 入金：一部 → 残り
  r = g.panelAction('payment', { row: newRow, amount: 100000, paidOn: g.todayYmd_(), method: '銀行振込' });
  assert.ok(r.ok, r.error);
  assert.equal(r.statusAfter, '一部入金');
  assert.equal(r.balanceAfter, 101240);
  r = g.panelAction('payment', { row: newRow, paidOn: g.todayYmd_() });
  assert.equal(r.amount, 101240, '金額を空にすると残額すべて');
  assert.equal(r.statusAfter, '入金済');
  const paysForInv = g.readTable_('payments', 2).filter((p) => p.invoiceNo === inv.invoiceNo);
  assert.equal(paysForInv.length, 2);

  // 入金記録シートで請求番号を選ぶと、入金日と残額が入る（別の未入金の請求で確認）
  const book = g.loadBook_();
  const unpaid = book.deals.find((d) => d.invoiceNo && d.calcPaid === 0 && g.computePaymentStatus(g.dealForStatus_(d), g.todayYmd_()).status === '期限超過');
  const paySh = ss.getSheetByName(g.SHEET.payments);
  const pr = g.lastDataRow_(paySh, 2) + 1;
  edit(ctx, g.SHEET.payments, pr, g.COL.payments.invoiceNo, unpaid.invoiceNo);
  const payRow = g.readRow_('payments', pr);
  assert.equal(payRow.paidOn, g.todayYmd_());
  assert.equal(payRow.amount, unpaid.calc.total);
  paySh.getRange(pr, 1, 1, 6).clearContent(); // 試しの行は消す

  // 8) 未入金アラートと督促メール下書き
  const s = g.summarizeAlerts_();
  assert.ok(s.overdueCount >= 3);
  assert.ok(s.items[0].days >= s.items[s.items.length - 1].days, '超過日数の大きい順');
  g.menuReminderDrafts();
  const reminders = gmail.drafts.slice(1);
  assert.equal(reminders.length, s.overdueCount);
  for (const d of reminders) {
    assert.match(d.subject, /【ご入金のお願い】INV-\d{6}-\d{3}/);
    assert.match(d.body, /未入金残高：¥[\d,]+/);
    assert.match(d.to, /\.example$/);
  }
  assert.equal(gmail.sent.length, 0, '督促も下書きだけ');

  // 9) 毎朝のチェック：既定では通知しない。オンにしても宛先は本人だけ
  g.setDailyTrigger_(true);
  assert.equal(ctx.triggers.length, 1);
  assert.equal(ctx.triggers[0].getHandlerFunction(), 'dailyCheck');
  g.setDailyTrigger_(true);
  assert.equal(ctx.triggers.length, 1, '二重に登録しない');
  g.dailyCheck();
  assert.equal(gmail.sent.length, 0);
  g.setSetting_('alertNotifySelf', true);
  g.dailyCheck();
  assert.equal(gmail.sent.length, 1);
  assert.equal(gmail.sent[0].to, 'owner@example.com');
  assert.match(gmail.sent[0].subject, /未入金アラート/);
  g.setDailyTrigger_(false);
  assert.equal(ctx.triggers.length, 0);

  // ダッシュボードだけの作り直しではデータが変わらない
  const before = JSON.stringify(g.readTable_('deals'));
  g.menuRebuildDashboard();
  assert.equal(JSON.stringify(g.readTable_('deals')), before);
  assert.equal(ss.getSheetByName(g.SHEET.dashboard).charts.length, 2);

  // 10) 開いたときのトースト
  g.onOpen();
  assert.match(ss.toasts.at(-1).msg, /期限超過 \d+件/);

  // 11) 操作ログ
  const log = g.readTable_('log');
  const actions = log.map((l) => l.action);
  for (const a of ['初期セットアップ', '見積書PDF作成', '請求書PDF作成', '請求書メール下書き', '入金記録', '督促メール下書き', '毎朝の未入金チェック', 'アラート通知']) {
    assert.ok(actions.includes(a), `ログに ${a}`);
  }
  assert.ok(log.every((l) => l.result === '成功'));

  // 12) セルフテスト（模擬シートでは数式が計算されないので、計算列を JS で埋めてから実行）
  fillCalcColumns(ctx);
  const st = g.runSelfTest();
  const failed = st.results.filter((x) => x[1] === '失敗');
  // ダッシュボードの数式は模擬環境では計算されないため、その 2 項目以外は成功すること
  assert.deepEqual([...failed.map((x) => x[0]).filter((n) => !n.startsWith('ダッシュボード'))], [], JSON.stringify(st.results));
});

test('入力ミスは分かりやすいエラーにする', () => {
  const ctx = loadGas();
  const { g, ss, ui } = ctx;
  g.setupSpreadsheet_({ demo: true, samplePdfs: false });
  const dealsSh = ss.getSheetByName(g.SHEET.deals);
  const row = g.lastDataRow_(dealsSh, 1) + 1;
  edit(ctx, g.SHEET.deals, row, g.COL.deals.customer, '青葉デンタルクリニック');
  // 明細なし
  let r = g.panelAction('invoice', { row });
  assert.equal(r.ok, false);
  assert.match(r.error, /明細がありません/);
  // 失注には請求書を出せない
  dealsSh.getRange(row, g.COL.deals.orderStatus).setValue('失注');
  const linesSh = ss.getSheetByName(g.SHEET.lines);
  const lr = g.lastDataRow_(linesSh, 1) + 1;
  linesSh.getRange(lr, 1).setValue(g.readRow_('deals', row).id);
  edit(ctx, g.SHEET.lines, lr, g.COL.lines.product, 'ページ追加・更新作業');
  r = g.panelAction('invoice', { row });
  assert.match(r.error, /失注/);
  // 請求前の入金
  dealsSh.getRange(row, g.COL.deals.orderStatus).setValue('受注');
  r = g.panelAction('payment', { row, amount: 1000 });
  assert.match(r.error, /請求書を発行していません/);
  // 案件以外のシートで実行
  select(ctx, g.SHEET.settings, 3, 2);
  g.menuCreateQuote();
  assert.match(ui.alerts.at(-1).msg, /案件の行/);
  // 税率が不正
  linesSh.getRange(lr, g.COL.lines.taxRate).setValue(0.05);
  r = g.panelAction('quote', { row });
  assert.match(r.error, /税率/);
  // エラーはログに残る
  assert.ok(g.readTable_('log').some((l) => l.result === '失敗'));
});

test('メニューのキャンセルでは何も作らない', () => {
  const ctx = loadGas({ alertAnswers: ['CANCEL', 'CANCEL'] });
  const { g, ss, gmail } = ctx;
  g.setupSpreadsheet_({ demo: true, samplePdfs: false });
  const row = g.readTable_('deals').find((d) => d.invoiceNo)._row;
  select(ctx, g.SHEET.deals, row);
  g.menuDraftInvoice();
  g.menuReminderDrafts();
  assert.equal(gmail.drafts.length, 0);
  assert.equal(ss.toasts.length, 0);
});

test('計算列の数式：括弧の対応とシート参照', () => {
  const { g } = loadGas();
  const formulas = g.columnFormulas_();
  for (const table of Object.keys(formulas)) {
    for (const [key, f] of Object.entries(formulas[table])) {
      let depth = 0;
      for (const ch of f.replace(/"[^"]*"/g, '')) {
        if (ch === '(' || ch === '{') depth++;
        if (ch === ')' || ch === '}') depth--;
        assert.ok(depth >= 0, `${table}.${key}`);
      }
      assert.equal(depth, 0, `${table}.${key} の括弧`);
      const label = g.COLUMNS[table][g.COL[table][key] - 1].label;
      assert.ok(f.startsWith(`={"${label}";`), `${table}.${key} の見出し`);
      for (const ref of f.match(/'[^']+'!/g) || []) assert.ok(Object.values(g.SHEET).includes(ref.slice(1, -2)), ref);
    }
  }
  // 計算列は COLUMNS の calc と一致
  for (const table of Object.keys(g.COLUMNS)) {
    const calc = [...g.COLUMNS[table].filter((c) => c.calc).map((c) => c.key)].sort();
    assert.deepEqual(calc, Object.keys(formulas[table] || {}).sort(), table);
  }
});
