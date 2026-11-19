/**
 * Gmail の下書き作成。
 * このシステムは顧客宛のメールを自動送信しない（GmailApp.createDraft のみ）。
 * 送信は利用者が Gmail で内容を確認してから手動で行う。
 */

function mailVars_(settings, deal, customer, extra) {
  const to = buildAddressee(customer);
  const vars = {
    '自社名': settings.companyName,
    '担当者': settings.representative,
    '電話番号': settings.companyTel,
    'メール': settings.companyEmail,
    '顧客名': customer.name,
    '宛名': to.honorific ? customer.name + ' ' + to.honorific : customer.name + '\n' + to.person,
    '件名': deal.subject,
  };
  Object.keys(extra || {}).forEach(function (k) { vars[k] = extra[k]; });
  vars['署名'] = fillTemplate(settings.mailSignature, vars);
  return vars;
}

function buildMailBody_(kind, vars) {
  const bodies = {
    quote: [
      '{宛名}',
      '',
      'いつもお世話になっております。{自社名}の{担当者}です。',
      '',
      '「{件名}」のお見積書をお送りいたします。',
      '添付のPDFをご確認くださいますよう、よろしくお願い申し上げます。',
      '',
      '　見積番号　：{番号}',
      '　お見積金額：{金額}（税込）',
      '　有効期限　：{期限}',
      '',
      'ご不明な点やご要望がございましたら、お気軽にお申し付けください。',
      '何卒よろしくお願いいたします。',
      '',
      '{署名}',
    ],
    invoice: [
      '{宛名}',
      '',
      'いつもお世話になっております。{自社名}の{担当者}です。',
      '',
      '「{件名}」のご請求書をお送りいたします。',
      '添付のPDFをご確認のうえ、お支払期限までにお振込みくださいますようお願い申し上げます。',
      '',
      '　請求番号　：{番号}',
      '　ご請求金額：{金額}（税込）',
      '　お支払期限：{期限}',
      '',
      '今後ともよろしくお願いいたします。',
      '',
      '{署名}',
    ],
    reminder: [
      '{宛名}',
      '',
      'いつもお世話になっております。{自社名}の{担当者}です。',
      '',
      '{請求日}付でお送りしました下記のご請求につきまして、',
      'お支払期限を過ぎておりますが、本日時点でご入金の確認ができておりません。',
      '',
      '　請求番号　：{番号}',
      '　件名　　　：{件名}',
      '　ご請求金額：{金額}（税込）',
      '　未入金残高：{残高}',
      '　お支払期限：{期限}',
      '',
      'お手数ですが、ご確認のうえお手続きをお願いできますでしょうか。',
      '行き違いでお支払い済みの場合は、失礼をお詫びするとともに本メールを破棄してください。',
      '',
      '{署名}',
    ],
  };
  return fillTemplate(bodies[kind].join('\n'), vars);
}

/**
 * 見積・請求メールの下書きを作る。PDF がなければ先に作る。
 * @return {{ok:boolean, draftId:string, to:string, subject:string}}
 */
function createMailDraft_(type, row) {
  let file = getDocFile_(type, row);
  if (!file) {
    createDocument_(type, row, { force: true });
    file = getDocFile_(type, row);
    if (!file) throw new Error('PDF の作成に失敗しました。');
  }
  const book = loadBook_();
  const deal = book.deals.filter(function (d) { return d._row === row; })[0];
  const customer = book.customerByName[deal.customer];
  if (!customer.email) throw new Error('顧客「' + customer.name + '」のメールアドレスが顧客マスタにありません。');
  const settings = getSettings_();
  const t = DOC_TYPES[type];
  const limit = type === 'quote' ? addDays(deal.quoteOn, Number(settings.quoteValidDays) || 30) : deal.dueOn;
  const vars = mailVars_(settings, deal, customer, {
    '番号': deal[t.numberKey],
    '金額': formatYen(deal.calc.total),
    '期限': formatJpDate(limit),
  });
  const subject = fillTemplate(type === 'quote' ? settings.mailSubjectQuote : settings.mailSubjectInvoice, vars);
  const body = buildMailBody_(type, vars);
  const draft = GmailApp.createDraft(customer.email, subject, body, { attachments: [file.getBlob()], name: settings.companyName });
  const stamp = Utilities.formatDate(new Date(), APP.timeZone, 'M/d HH:mm');
  updateCells_('deals', row, { mailDraft: t.name + ' 下書き ' + stamp });
  writeLog_(t.name + 'メール下書き', deal[t.numberKey], true, '宛先 ' + customer.email + '／件名 ' + subject + '（送信はしていません）');
  return { ok: true, draftId: draft.getId(), to: customer.email, subject: subject };
}

/** 期限超過の請求すべてに督促メールの下書きを作る */
function createReminderDrafts_(todayOpt) {
  const today = todayOpt || todayYmd_();
  const book = loadBook_();
  const settings = getSettings_();
  const overdue = buildAlerts(book.deals.map(function (d) { return Object.assign(dealForStatus_(d), { _deal: d }); }), today, 0)
    .filter(function (a) { return a.level === 'overdue'; });
  const results = [];
  overdue.forEach(function (a) {
    const deal = a.deal._deal;
    const customer = book.customerByName[deal.customer];
    if (!customer || !customer.email) {
      results.push({ ok: false, number: deal.invoiceNo, reason: 'メールアドレスなし' });
      return;
    }
    const vars = mailVars_(settings, deal, customer, {
      '番号': deal.invoiceNo,
      '金額': formatYen(deal.calc.total),
      '残高': formatYen(a.status.balance),
      '期限': formatJpDate(deal.dueOn),
      '請求日': formatJpDate(deal.invoiceOn),
    });
    const subject = fillTemplate(settings.mailSubjectReminder, vars);
    const file = getDocFile_('invoice', deal._row);
    const options = { name: settings.companyName };
    if (file) options.attachments = [file.getBlob()];
    GmailApp.createDraft(customer.email, subject, buildMailBody_('reminder', vars), options);
    const stamp = Utilities.formatDate(new Date(), APP.timeZone, 'M/d HH:mm');
    updateCells_('deals', deal._row, { mailDraft: '督促 下書き ' + stamp });
    results.push({ ok: true, number: deal.invoiceNo, to: customer.email, days: a.status.daysPastDue });
  });
  writeLog_('督促メール下書き', results.length + '件', true, results.map(function (r) { return r.number + (r.ok ? '（' + r.days + '日超過）' : '（' + r.reason + '）'); }).join('、') + '（送信はしていません）');
  return results;
}
