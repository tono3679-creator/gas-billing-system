/**
 * メニュー・編集時の自動入力・操作パネル（サイドバー）。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(APP.menuName)
    .addItem('操作パネルを開く', 'menuOpenPanel')
    .addSeparator()
    .addItem('見積書PDFを作成（選択中の案件）', 'menuCreateQuote')
    .addItem('請求書PDFを発行（選択中の案件）', 'menuCreateInvoice')
    .addSeparator()
    .addItem('見積メールの下書きを作成', 'menuDraftQuote')
    .addItem('請求メールの下書きを作成', 'menuDraftInvoice')
    .addSeparator()
    .addItem('入金を記録（操作パネル）', 'menuOpenPanel')
    .addItem('未入金チェックを実行', 'menuAlertCheck')
    .addItem('期限超過分の督促メール下書きを作成', 'menuReminderDrafts')
    .addSeparator()
    .addSubMenu(ui.createMenu('管理')
      .addItem('毎朝の自動チェックをオンにする', 'menuTriggerOn')
      .addItem('毎朝の自動チェックをオフにする', 'menuTriggerOff')
      .addSeparator()
      .addItem('セルフテストを実行', 'menuSelfTest')
      .addItem('ダッシュボードだけ作り直す（データはそのまま）', 'menuRebuildDashboard')
      .addSeparator()
      .addItem('初期セットアップ（デモデータで作り直す）', 'menuSetupWithDemo')
      .addItem('空の状態で作り直す（本番運用の開始）', 'menuSetupEmpty'))
    .addToUi();

  // 開いたときに未入金の状況を右下に表示
  try {
    if (!getSs_().getSheetByName(SHEET.deals)) {
      getSs_().toast('メニュー「' + APP.menuName + ' › 管理 › 初期セットアップ」から始めてください。', 'はじめに', 10);
      return;
    }
    const s = summarizeAlerts_();
    if (s.overdueCount > 0) {
      getSs_().toast('期限超過 ' + s.overdueCount + '件（' + formatYen(s.overdueAmount) + '）／期限間近 ' + s.soonCount + '件。ダッシュボードの「未入金アラート」で確認できます。', '未入金アラート', 12);
    }
  } catch (e) {
    console.warn('onOpen のアラート表示をスキップ: ' + e);
  }
}

// ---------- 選択中の案件 ----------

/** 選択中のセルから案件の行を特定する（案件・売上明細・入金記録・ダッシュボードのアラート表に対応） */
function getSelectedDealRow_() {
  const ss = getSs_();
  const sh = ss.getActiveSheet();
  const row = ss.getActiveRange() ? ss.getActiveRange().getRow() : 0;
  const name = sh.getName();
  const hint = '「' + SHEET.deals + '」シートで案件の行（どのセルでも可）を選んでから実行してください。';
  if (name === SHEET.deals) {
    if (row < 2 || !sh.getRange(row, COL.deals.id).getValue()) throw new Error(hint);
    return row;
  }
  let key = '';
  let field = '';
  if (name === SHEET.lines && row >= 2) { key = sh.getRange(row, COL.lines.dealId).getValue(); field = 'id'; }
  if (name === SHEET.payments && row >= 2) { key = sh.getRange(row, COL.payments.invoiceNo).getValue(); field = 'invoiceNo'; }
  if (name === SHEET.dashboard && row >= 41) { key = sh.getRange(row, 2).getValue(); field = 'invoiceNo'; }
  if (key) {
    const hit = readTable_('deals').filter(function (d) { return d[field] === key; })[0];
    if (hit) return hit._row;
  }
  throw new Error(hint);
}

function withUi_(fn) {
  try {
    return fn(SpreadsheetApp.getUi());
  } catch (e) {
    SpreadsheetApp.getUi().alert('実行できませんでした', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog_('エラー', '', false, e.message);
    return null;
  }
}

function showLinkDialog_(title, message, links) {
  const html = '<div style="font-family:sans-serif;font-size:13px;line-height:1.7;color:#1f2937">' +
    '<p style="margin:0 0 12px 0;white-space:pre-wrap">' + escapeHtml(message) + '</p>' +
    links.map(function (l) {
      return '<a href="' + escapeHtml(l.url) + '" target="_blank" style="display:inline-block;margin:0 8px 8px 0;padding:6px 14px;border-radius:6px;background:#1f6feb;color:#fff;text-decoration:none">' + escapeHtml(l.label) + '</a>';
    }).join('') +
    '</div>';
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html).setWidth(440).setHeight(190), title);
}

// ---------- メニューの処理 ----------

function menuCreateQuote() { menuCreateDoc_('quote'); }
function menuCreateInvoice() { menuCreateDoc_('invoice'); }

function menuCreateDoc_(type) {
  withUi_(function (ui) {
    const row = getSelectedDealRow_();
    let res = createDocument_(type, row, {});
    if (res.needsConfirm) {
      const ans = ui.alert(DOC_TYPES[type].name + 'の再発行', res.number + ' の PDF は作成済みです。同じ番号で作り直しますか？（古いPDFはゴミ箱に移動します）', ui.ButtonSet.YES_NO);
      if (ans !== ui.Button.YES) return;
      res = createDocument_(type, row, { force: true });
    }
    showLinkDialog_(DOC_TYPES[type].name + 'を作成しました', res.number + '（' + formatYen(res.total) + '）を Google ドライブに保存しました。', [{ label: 'PDFを開く', url: res.url }]);
  });
}

function menuDraftQuote() { menuDraft_('quote'); }
function menuDraftInvoice() { menuDraft_('invoice'); }

function menuDraft_(type) {
  withUi_(function (ui) {
    const row = getSelectedDealRow_();
    const deal = readRow_('deals', row);
    const customer = readTable_('customers').filter(function (c) { return c.name === deal.customer; })[0];
    const to = customer ? customer.email : '（未登録）';
    const ans = ui.alert(DOC_TYPES[type].name + 'メールの下書き', '宛先 ' + to + ' 宛の下書きを Gmail に作成します。\nメールは送信しません（Gmail で確認してから送ってください）。', ui.ButtonSet.OK_CANCEL);
    if (ans !== ui.Button.OK) return;
    const res = createMailDraft_(type, row);
    showLinkDialog_('下書きを作成しました', '宛先：' + res.to + '\n件名：' + res.subject + '\nPDF を添付済みです。送信はしていません。', [{ label: 'Gmail の下書きを開く', url: 'https://mail.google.com/mail/u/0/#drafts' }]);
  });
}

function menuAlertCheck() {
  withUi_(function (ui) {
    const s = summarizeAlerts_();
    writeLog_('未入金チェック', '期限超過 ' + s.overdueCount + '件', true, '期限超過 ' + formatYen(s.overdueAmount) + '／期限間近 ' + s.soonCount + '件');
    getSs_().setActiveSheet(getSheet_(SHEET.dashboard));
    getSheet_(SHEET.dashboard).setActiveRange(getSheet_(SHEET.dashboard).getRange('B39'));
    ui.alert('未入金チェック', s.items.length ? alertText_(s) : '期限超過・期限間近の未入金はありません。', ui.ButtonSet.OK);
  });
}

function menuReminderDrafts() {
  withUi_(function (ui) {
    const s = summarizeAlerts_();
    if (!s.overdueCount) { ui.alert('督促メール', '期限超過の請求はありません。', ui.ButtonSet.OK); return; }
    const ans = ui.alert('督促メールの下書き', '期限超過 ' + s.overdueCount + '件（' + formatYen(s.overdueAmount) + '）の督促メールを Gmail の下書きに作成します。\nメールは送信しません。続けますか？', ui.ButtonSet.OK_CANCEL);
    if (ans !== ui.Button.OK) return;
    const res = createReminderDrafts_();
    const ok = res.filter(function (r) { return r.ok; }).length;
    showLinkDialog_('督促メールの下書きを作成しました', ok + '件の下書きを作成しました（送信はしていません）。', [{ label: 'Gmail の下書きを開く', url: 'https://mail.google.com/mail/u/0/#drafts' }]);
  });
}

function menuTriggerOn() {
  withUi_(function (ui) {
    setDailyTrigger_(true);
    ui.alert('毎朝の自動チェック', '毎朝 8 時台に未入金チェックを実行し、操作ログに記録します。\n自分宛の通知メールは「設定」シートでオンにしたときだけ送ります（顧客には送りません）。', ui.ButtonSet.OK);
  });
}

function menuTriggerOff() {
  withUi_(function (ui) {
    setDailyTrigger_(false);
    ui.alert('毎朝の自動チェック', '自動チェックを止めました。', ui.ButtonSet.OK);
  });
}

function menuSelfTest() {
  withUi_(function (ui) {
    const r = runSelfTest();
    getSs_().setActiveSheet(getSheet_(SHEET.selfTest));
    ui.alert('セルフテスト', r.passed + '件成功 / ' + r.failed + '件失敗 / ' + r.warned + '件注意\n詳しくは「' + SHEET.selfTest + '」シートを確認してください。', ui.ButtonSet.OK);
  });
}

function menuRebuildDashboard() {
  withUi_(function (ui) {
    buildDashboard_();
    SpreadsheetApp.flush();
    writeLog_('ダッシュボード再作成', '', true, '数式とグラフを作り直しました（データは変更なし）');
    getSs_().setActiveSheet(getSheet_(SHEET.dashboard));
    ui.alert('ダッシュボード', 'ダッシュボードを作り直しました。案件・明細・入金のデータは変更していません。', ui.ButtonSet.OK);
  });
}

function menuOpenPanel() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('操作パネル');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------- 操作パネル（google.script.run から呼ぶ） ----------

function panelGetState() {
  const out = { deal: null, hint: '', alerts: null, triggerOn: false };
  try {
    out.alerts = summarizeAlerts_();
    out.alerts.items = out.alerts.items.slice(0, 6);
    out.triggerOn = isDailyTriggerOn_();
  } catch (e) {
    out.hint = e.message;
    return out;
  }
  try {
    const row = getSelectedDealRow_();
    const book = loadBook_();
    const d = book.deals.filter(function (x) { return x._row === row; })[0];
    const st = computePaymentStatus(dealForStatus_(d), todayYmd_());
    out.deal = {
      row: row, id: d.id, customer: d.customer, subject: d.subject, orderStatus: d.orderStatus,
      subtotal: d.calc.subtotal, tax: d.calc.tax, total: d.calc.total, lineCount: (book.linesByDeal[d.id] || []).length,
      quoteNo: d.quoteNo, quoteOn: d.quoteOn, hasQuotePdf: !!getLinkedFileId_('deals', row, 'quotePdf'),
      invoiceNo: d.invoiceNo, invoiceOn: d.invoiceOn, dueOn: d.dueOn, hasInvoicePdf: !!getLinkedFileId_('deals', row, 'invoicePdf'),
      paid: d.calcPaid, balance: st.balance, status: st.status, days: st.daysPastDue, mailDraft: d.mailDraft,
    };
  } catch (e) {
    out.hint = e.message;
  }
  return out;
}

function panelAction(action, payload) {
  payload = payload || {};
  try {
    const row = payload.row;
    switch (action) {
      case 'quote': return createDocument_('quote', row, { force: !!payload.force });
      case 'invoice': return createDocument_('invoice', row, { force: !!payload.force });
      case 'draftQuote': return createMailDraft_('quote', row);
      case 'draftInvoice': return createMailDraft_('invoice', row);
      case 'payment': return recordPayment_(row, payload);
      case 'reminders': return { ok: true, results: createReminderDrafts_() };
      case 'goto': {
        const sh = getSheet_(SHEET.deals);
        getSs_().setActiveSheet(sh);
        sh.setActiveRange(sh.getRange(row, 1));
        return { ok: true };
      }
      default: throw new Error('不明な操作: ' + action);
    }
  } catch (e) {
    writeLog_('エラー', action, false, e.message);
    return { ok: false, error: e.message };
  }
}

// ---------- 編集時の自動入力（シンプルトリガー） ----------

function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sh = e.range.getSheet();
    const name = sh.getName();
    const r0 = e.range.getRow();
    const c0 = e.range.getColumn();
    const nRows = e.range.getNumRows();
    const nCols = e.range.getNumColumns();
    const touches = function (col) { return c0 <= col && col < c0 + nCols; };
    if (r0 < 2 && r0 + nRows - 1 < 2) return;
    const rowsFrom = Math.max(r0, 2);
    const rowsTo = r0 + nRows - 1;

    if (name === SHEET.lines) {
      if (touches(COL.lines.product)) autofillLines_(sh, rowsFrom, rowsTo);
      warnIfInvoiced_(sh, rowsFrom, rowsTo);
    } else if (name === SHEET.deals && touches(COL.deals.customer)) {
      autofillDeals_(sh, rowsFrom, rowsTo);
    } else if (name === SHEET.payments && touches(COL.payments.invoiceNo)) {
      autofillPayments_(sh, rowsFrom, rowsTo);
    }
  } catch (err) {
    console.warn('onEdit: ' + err);
  }
}

/** 商品を選んだら単位・単価・税率を入れ、数量が空なら 1 にする */
function autofillLines_(sh, from, to) {
  const products = readTable_('products');
  const byName = {};
  products.forEach(function (p) { byName[p.name] = p; });
  for (let r = from; r <= to; r++) {
    const name = sh.getRange(r, COL.lines.product).getValue();
    const p = byName[name];
    if (!p) continue;
    sh.getRange(r, COL.lines.unit).setValue(p.unit);
    sh.getRange(r, COL.lines.unitPrice).setValue(p.unitPrice);
    sh.getRange(r, COL.lines.taxRate).setValue(p.taxRate);
    if (sh.getRange(r, COL.lines.qty).getValue() === '') sh.getRange(r, COL.lines.qty).setValue(1);
  }
}

/** 請求済みの案件の明細を触ったら注意を出す */
function warnIfInvoiced_(sh, from, to) {
  const ids = sh.getRange(from, COL.lines.dealId, to - from + 1, 1).getValues().map(function (v) { return v[0]; }).filter(String);
  if (!ids.length) return;
  const invoiced = readTable_('deals').filter(function (d) { return d.invoiceNo && ids.indexOf(d.id) >= 0; });
  if (invoiced.length) {
    getSs_().toast(invoiced.map(function (d) { return d.id; }).join('、') + ' は請求書を発行済みです。明細を直した場合は請求書を再発行してください。', '請求済みの案件です', 8);
  }
}

/** 顧客を選んだら案件番号・受付日・受注状況を入れる */
function autofillDeals_(sh, from, to) {
  const settings = getSettings_();
  const today = todayYmd_();
  const last = lastDataRow_(sh, COL.deals.id);
  const ids = last >= 2 ? sh.getRange(2, COL.deals.id, last - 1, 1).getValues().map(function (v) { return v[0]; }) : [];
  for (let r = from; r <= to; r++) {
    if (!sh.getRange(r, COL.deals.customer).getValue()) continue;
    if (!sh.getRange(r, COL.deals.id).getValue()) {
      const id = nextDealNumber(ids, settings.dealPrefix, today);
      ids.push(id);
      sh.getRange(r, COL.deals.id).setValue(id);
    }
    if (!sh.getRange(r, COL.deals.receivedOn).getValue()) sh.getRange(r, COL.deals.receivedOn).setValue(ymdToDate(today));
    if (!sh.getRange(r, COL.deals.orderStatus).getValue()) sh.getRange(r, COL.deals.orderStatus).setValue('見積中');
  }
}

/** 請求番号を選んだら入金日（今日）と残額を入れる */
function autofillPayments_(sh, from, to) {
  const book = loadBook_();
  const today = todayYmd_();
  for (let r = from; r <= to; r++) {
    const no = sh.getRange(r, COL.payments.invoiceNo).getValue();
    if (!no) continue;
    const d = book.deals.filter(function (x) { return x.invoiceNo === no; })[0];
    if (!d) continue;
    if (!sh.getRange(r, COL.payments.paidOn).getValue()) sh.getRange(r, COL.payments.paidOn).setValue(ymdToDate(today));
    if (sh.getRange(r, COL.payments.amount).getValue() === '') {
      // この行の入金額はまだ空なので、残額 = 合計 - これまでの入金
      const balance = d.calc.total - d.calcPaid;
      if (balance > 0) sh.getRange(r, COL.payments.amount).setValue(balance);
    }
    if (!sh.getRange(r, COL.payments.method).getValue()) sh.getRange(r, COL.payments.method).setValue('銀行振込');
  }
}
