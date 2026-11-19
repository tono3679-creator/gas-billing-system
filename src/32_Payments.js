/**
 * 入金の記録と未入金アラート。
 */

/**
 * 入金を 1 件記録する。
 * @param {number} row 案件・請求シートの行
 * @param {{paidOn?:string, amount?:number, method?:string, note?:string}} input
 */
function recordPayment_(row, input) {
  const book = loadBook_();
  const deal = book.deals.filter(function (d) { return d._row === row; })[0];
  if (!deal) throw new Error('案件の行を選んでください。');
  if (!deal.invoiceNo) throw new Error('案件 ' + deal.id + ' はまだ請求書を発行していません。');
  const status = computePaymentStatus(dealForStatus_(deal), todayYmd_());
  const amount = input.amount === undefined || input.amount === '' ? status.balance : Number(input.amount);
  if (!(amount > 0)) throw new Error('入金額は 0 より大きい数値で入力してください。');
  const paidOn = toYmd(input.paidOn) || todayYmd_();
  appendRows_('payments', [{ paidOn: paidOn, invoiceNo: deal.invoiceNo, amount: amount, method: input.method || '銀行振込', note: input.note || '' }], 2);
  const after = computePaymentStatus(Object.assign(dealForStatus_(deal), { paid: deal.calcPaid + amount }), todayYmd_());
  writeLog_('入金記録', deal.invoiceNo, true, deal.customer + '／' + formatYen(amount) + '／' + paidOn + '／入金後: ' + after.status);
  return { ok: true, invoiceNo: deal.invoiceNo, amount: amount, statusAfter: after.status, balanceAfter: after.balance };
}

/** 未入金の状況をまとめる（ダッシュボード・トースト・毎朝のチェックで共用） */
function summarizeAlerts_(todayOpt) {
  const today = todayOpt || todayYmd_();
  const book = loadBook_();
  const settings = getSettings_();
  const items = buildAlerts(book.deals.map(function (d) { return Object.assign(dealForStatus_(d), { _deal: d }); }), today, Number(settings.alertSoonDays) || 7);
  const overdue = items.filter(function (a) { return a.level === 'overdue'; });
  const soon = items.filter(function (a) { return a.level === 'soon'; });
  const sum = function (list) { return list.reduce(function (s, a) { return s + a.status.balance; }, 0); };
  return {
    today: today,
    overdueCount: overdue.length,
    overdueAmount: sum(overdue),
    soonCount: soon.length,
    soonAmount: sum(soon),
    items: items.map(function (a) {
      const d = a.deal._deal;
      return { level: a.level, invoiceNo: d.invoiceNo, customer: d.customer, subject: d.subject, dueOn: d.dueOn, balance: a.status.balance, days: a.status.daysPastDue, row: d._row };
    }),
  };
}

function alertText_(s) {
  const lines = ['未入金アラート（' + s.today + ' 時点）', ''];
  lines.push('期限超過：' + s.overdueCount + '件　' + formatYen(s.overdueAmount));
  lines.push('期限間近：' + s.soonCount + '件　' + formatYen(s.soonAmount));
  lines.push('');
  s.items.forEach(function (i) {
    lines.push((i.level === 'overdue' ? '［超過 ' + i.days + '日］' : '［あと ' + -i.days + '日］') + i.invoiceNo + '　' + i.customer + '　' + formatYen(i.balance) + '（期限 ' + i.dueOn + '）');
  });
  return lines.join('\n');
}

// ---------- 毎朝の自動チェック ----------

const DAILY_HANDLER = 'dailyCheck';

/** 時間主導トリガーから呼ばれる。ログに記録し、設定がオンなら所有者本人にだけ通知 */
function dailyCheck() {
  const s = summarizeAlerts_();
  writeLog_('毎朝の未入金チェック', '期限超過 ' + s.overdueCount + '件', true, '期限超過 ' + formatYen(s.overdueAmount) + '／期限間近 ' + s.soonCount + '件 ' + formatYen(s.soonAmount));
  const settings = getSettings_();
  if (settings.alertNotifySelf && s.overdueCount > 0) {
    // 宛先はスクリプトを実行している本人のアドレスに固定（顧客には送らない）
    const me = Session.getEffectiveUser().getEmail();
    if (me) {
      GmailApp.sendEmail(me, '【未入金アラート】期限超過 ' + s.overdueCount + '件（' + formatYen(s.overdueAmount) + '）', alertText_(s) + '\n\n' + getSs_().getUrl());
      writeLog_('アラート通知', me, true, '自分宛に通知しました');
    }
  }
  return s;
}

function isDailyTriggerOn_() {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === DAILY_HANDLER; });
}

function setDailyTrigger_(on) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === DAILY_HANDLER) ScriptApp.deleteTrigger(t);
  });
  if (on) ScriptApp.newTrigger(DAILY_HANDLER).timeBased().everyDays(1).atHour(8).inTimezone(APP.timeZone).create();
  writeLog_('毎朝の自動チェック', on ? 'オン' : 'オフ', true, on ? '毎朝 8 時台に未入金チェックを実行します' : '自動チェックを止めました');
  return on;
}
