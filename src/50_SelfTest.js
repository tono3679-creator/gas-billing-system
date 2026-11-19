/**
 * セルフテスト：実際のスプレッドシート上で、
 * シートの数式の計算結果と JS 側の計算（PDF に使う値）が一致するかなどを確かめる。
 * 結果は「セルフテスト」シートに書き出す。データは書き換えない。
 */

function runSelfTest() {
  const results = [];
  const add = function (name, level, detail) { results.push([name, level, detail || '']); };
  const check = function (name, fn) {
    try {
      const r = fn();
      if (r === true || r === undefined) add(name, '成功', '');
      else if (r && r.warn) add(name, '注意', r.warn);
      else add(name, '成功', String(r));
    } catch (e) {
      add(name, '失敗', e.message);
    }
  };
  const ss = getSs_();
  SpreadsheetApp.flush();

  check('必要なシートがそろっている', function () {
    const missing = SHEET_ORDER.filter(function (n) { return !ss.getSheetByName(n); });
    if (missing.length) throw new Error('見つからないシート: ' + missing.join('、'));
    return SHEET_ORDER.length + 'シート';
  });

  check('計算列の数式が入っている', function () {
    const formulas = columnFormulas_();
    const broken = [];
    Object.keys(formulas).forEach(function (tableKey) {
      const sh = getSheet_(SHEET[tableKey]);
      Object.keys(formulas[tableKey]).forEach(function (key) {
        const f = sh.getRange(1, COL[tableKey][key]).getFormula();
        if (f !== formulas[tableKey][key]) broken.push(SHEET[tableKey] + '・' + COLUMNS[tableKey][COL[tableKey][key] - 1].label);
      });
    });
    if (broken.length) throw new Error('数式が消えている・書き換わっている列: ' + broken.join('、') + '（初期セットアップで復旧できます）');
    return '計算列 ' + Object.keys(formulas).reduce(function (s, k) { return s + Object.keys(formulas[k]).length; }, 0) + '列';
  });

  check('計算列にエラー値がない', function () {
    const errors = [];
    ['deals', 'lines', 'payments'].forEach(function (tableKey) {
      const sh = getSheet_(SHEET[tableKey]);
      const last = lastDataRow_(sh, tableKey === 'payments' ? 2 : 1);
      if (last < 2) return;
      const vals = sh.getRange(1, 1, last, COLUMNS[tableKey].length).getDisplayValues();
      vals.forEach(function (row, i) {
        row.forEach(function (v, j) {
          if (/^#(REF|N\/A|VALUE|ERROR|DIV\/0|NAME)/.test(String(v))) errors.push(SHEET[tableKey] + ' ' + (i + 1) + '行' + (j + 1) + '列: ' + v);
        });
      });
    });
    if (errors.length) throw new Error(errors.slice(0, 5).join(' / ') + (errors.length > 5 ? ' ほか' + (errors.length - 5) + '件' : ''));
  });

  const book = loadBook_();
  const today = todayYmd_();

  check('案件の金額：シートの数式とプログラムの計算が一致', function () {
    const diffs = [];
    book.deals.forEach(function (d) {
      if (d.subtotal !== d.calc.subtotal || d.tax !== d.calc.tax || d.total !== d.calc.total) {
        diffs.push(d.id + '（シート ' + d.subtotal + '/' + d.tax + '/' + d.total + '、計算 ' + d.calc.subtotal + '/' + d.calc.tax + '/' + d.calc.total + '）');
      }
    });
    if (diffs.length) throw new Error(diffs.slice(0, 5).join(' / '));
    return book.deals.length + '件すべて一致';
  });

  check('入金ステータス：シートの数式とプログラムの判定が一致', function () {
    const diffs = [];
    book.deals.forEach(function (d) {
      const s = computePaymentStatus(dealForStatus_(d), today);
      const sheetBalance = d.balance === '' ? '' : Number(d.balance);
      const sheetDays = d.overdueDays === '' ? '' : Number(d.overdueDays);
      if (d.paymentStatus !== s.status || (s.status !== '入金済' && sheetBalance !== s.balance) || sheetDays !== s.daysPastDue) {
        diffs.push(d.id + '（シート ' + d.paymentStatus + '/' + sheetBalance + '/' + sheetDays + '、計算 ' + s.status + '/' + s.balance + '/' + s.daysPastDue + '）');
      }
    });
    if (diffs.length) throw new Error(diffs.slice(0, 5).join(' / '));
    return book.deals.length + '件すべて一致';
  });

  check('見積番号・請求番号・案件番号に重複がない', function () {
    const dup = [];
    ['id', 'quoteNo', 'invoiceNo'].forEach(function (k) {
      const seen = {};
      book.deals.forEach(function (d) {
        if (!d[k]) return;
        if (seen[d[k]]) dup.push(d[k]);
        seen[d[k]] = true;
      });
    });
    if (dup.length) throw new Error('重複: ' + dup.join('、'));
  });

  check('案件・明細・入金のつながり', function () {
    const ids = {};
    const invs = {};
    book.deals.forEach(function (d) { ids[d.id] = true; if (d.invoiceNo) invs[d.invoiceNo] = true; });
    const orphanLines = book.lines.filter(function (l) { return !ids[l.dealId]; }).length;
    const orphanPays = book.payments.filter(function (p) { return !invs[p.invoiceNo]; }).length;
    const unknownCustomers = book.deals.filter(function (d) { return !book.customerByName[d.customer]; }).length;
    const noLines = book.deals.filter(function (d) { return !(book.linesByDeal[d.id] || []).length; }).map(function (d) { return d.id; });
    if (orphanLines || orphanPays || unknownCustomers) throw new Error('案件のない明細 ' + orphanLines + '件、請求のない入金 ' + orphanPays + '件、マスタにない顧客 ' + unknownCustomers + '件');
    if (noLines.length) return { warn: '明細のない案件: ' + noLines.join('、') };
  });

  check('支払期限が支払条件どおり', function () {
    const off = [];
    book.deals.forEach(function (d) {
      if (!d.invoiceNo || !d.invoiceOn || !d.dueOn) return;
      const c = book.customerByName[d.customer];
      if (c && calcDueDate(d.invoiceOn, c.terms) !== d.dueOn) off.push(d.invoiceNo + '（' + d.dueOn + '、条件どおりなら ' + calcDueDate(d.invoiceOn, c.terms) + '）');
    });
    if (off.length) return { warn: '手で変更された支払期限: ' + off.join('、') };
  });

  check('ダッシュボードの未入金残高が明細からの計算と一致', function () {
    const s = summarizeAlerts_(today);
    const expectedUnpaid = book.deals.reduce(function (sum, d) {
      const st = computePaymentStatus(dealForStatus_(d), today);
      return sum + (['未入金', '一部入金', '期限超過'].indexOf(st.status) >= 0 ? st.balance : 0);
    }, 0);
    const dash = getSheet_(SHEET.dashboard);
    const shown = Number(dash.getRange('F6').getValue());
    const overdueShown = Number(dash.getRange('I6').getValue());
    if (shown !== expectedUnpaid) throw new Error('ダッシュボード ' + shown + '、計算 ' + expectedUnpaid);
    if (overdueShown !== s.overdueAmount) throw new Error('期限超過 ダッシュボード ' + overdueShown + '、計算 ' + s.overdueAmount);
    return '未入金 ' + formatYen(expectedUnpaid) + '／期限超過 ' + formatYen(s.overdueAmount) + '（' + s.overdueCount + '件）';
  });

  check('ダッシュボードの月別売上の合計が案件と一致', function () {
    const dash = getSheet_(SHEET.dashboard);
    const months = dash.getRange('B11:C22').getValues();
    const byMonth = {};
    book.deals.forEach(function (d) { if (d.invoiceOn) byMonth[d.invoiceOn.slice(0, 7).replace('-', '/')] = (byMonth[d.invoiceOn.slice(0, 7).replace('-', '/')] || 0) + d.calc.subtotal; });
    const diffs = months.filter(function (m) { return (byMonth[m[0]] || 0) !== Number(m[1]); }).map(function (m) { return m[0] + '（' + m[1] + ' / ' + (byMonth[m[0]] || 0) + '）'; });
    if (diffs.length) throw new Error(diffs.join(' / '));
    return '12か月分一致';
  });

  check('設定：帳票に必要な項目', function () {
    const s = getSettings_();
    const missing = ['companyName', 'companyAddress', 'companyTel', 'bankName', 'bankAccount', 'bankHolder'].filter(function (k) { return !s[k]; });
    if (missing.length) throw new Error('未入力: ' + missing.join('、'));
    if (s.registrationNo && !/^T\d{13}$/.test(String(s.registrationNo))) return { warn: '登録番号の形式が T＋13桁 ではありません' };
  });

  check('PDF 変換（保存はしない）', function () {
    const d = book.deals.filter(function (x) { return (book.linesByDeal[x.id] || []).length; })[0];
    if (!d) return { warn: '明細のある案件がないため省略' };
    const html = renderDocumentHtml({
      type: 'invoice', number: 'TEST', issuedOn: today, limitOn: today, subject: d.subject, customer: book.customerByName[d.customer],
      lines: book.linesByDeal[d.id].map(function (l) { return { name: l.product, detail: l.detail, qty: l.qty, unit: l.unit, unitPrice: l.unitPrice, taxRate: l.taxRate }; }), note: '',
    }, getSettings_());
    const pdf = Utilities.newBlob(html, 'text/html', 'test.html').getAs(MimeType.PDF);
    const size = pdf.getBytes().length;
    if (size < 5000) throw new Error('PDF が小さすぎます（' + size + ' バイト）');
    return 'PDF ' + Math.round(size / 1024) + 'KB';
  });

  check('保存先フォルダにアクセスできる', function () {
    const id = getSettings_().driveFolderId;
    if (!id) return { warn: 'まだ PDF を作成していないため未作成（初回の PDF 作成時に自動作成）' };
    const f = DriveApp.getFolderById(id);
    return f.getName();
  });

  check('顧客へメールを自動送信する処理がない', function () {
    // 時間主導トリガーは dailyCheck だけ。dailyCheck の通知先は所有者本人に固定
    const others = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() !== DAILY_HANDLER; });
    if (others.length) return { warn: '想定外のトリガー: ' + others.map(function (t) { return t.getHandlerFunction(); }).join('、') };
    return '毎朝の自動チェック: ' + (isDailyTriggerOn_() ? 'オン' : 'オフ');
  });

  // 書き出し
  let sh = ss.getSheetByName(SHEET.selfTest);
  if (!sh) sh = ss.insertSheet(SHEET.selfTest);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.setHiddenGridlines(true);
  const header = [['テスト項目', '結果', '詳細']];
  sh.getRange(1, 1, 1, 3).setValues(header).setBackground(THEME.header).setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, results.length, 3).setValues(results);
  sh.getRange(results.length + 3, 1).setValue('実行日時: ' + Utilities.formatDate(new Date(), APP.timeZone, 'yyyy/MM/dd HH:mm:ss')).setFontColor(THEME.muted);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 700);
  sh.getRange(1, 1, results.length + 3, 3).setFontFamily(THEME.font).setVerticalAlignment('middle').setWrap(true);
  const res = sh.getRange(2, 2, results.length, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('成功').setBackground(THEME.status['入金済'].bg).setFontColor(THEME.status['入金済'].fg).setRanges([res]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('失敗').setBackground(THEME.status['期限超過'].bg).setFontColor(THEME.status['期限超過'].fg).setRanges([res]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('注意').setBackground(THEME.status['未入金'].bg).setFontColor(THEME.status['未入金'].fg).setRanges([res]).build(),
  ]);
  sh.setTabColor('#9ca3af');
  const summary = {
    passed: results.filter(function (r) { return r[1] === '成功'; }).length,
    failed: results.filter(function (r) { return r[1] === '失敗'; }).length,
    warned: results.filter(function (r) { return r[1] === '注意'; }).length,
    results: results,
  };
  writeLog_('セルフテスト', summary.passed + '/' + results.length, summary.failed === 0, summary.failed ? results.filter(function (r) { return r[1] === '失敗'; }).map(function (r) { return r[0]; }).join('、') : '失敗なし');
  return summary;
}
