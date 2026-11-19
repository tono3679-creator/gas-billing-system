/**
 * 見積書・請求書の PDF 作成と Google ドライブへの保存。
 */

const DOC_TYPES = {
  quote: { name: '見積書', folder: '見積書', numberKey: 'quoteNo', dateKey: 'quoteOn', pdfKey: 'quotePdf', prefixKey: 'quotePrefix' },
  invoice: { name: '請求書', folder: '請求書', numberKey: 'invoiceNo', dateKey: 'invoiceOn', pdfKey: 'invoicePdf', prefixKey: 'invoicePrefix' },
};

/**
 * 案件の行から書類（PDF）を作る。番号が未採番なら採番し、日付・期限も埋める。
 * 既に PDF があるときは opt.force が無ければ { needsConfirm: true } を返す（再発行の確認用）。
 * @return {{ok:boolean, number:string, url:string, fileId:string, total:number, needsConfirm?:boolean}}
 */
function createDocument_(type, row, opt) {
  opt = opt || {};
  const t = DOC_TYPES[type];
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const book = loadBook_();
    const deal = book.deals.filter(function (d) { return d._row === row; })[0];
    if (!deal) throw new Error(row + '行目に案件がありません。案件・請求シートで案件の行を選んでください。');
    const settings = getSettings_();
    const customer = book.customerByName[deal.customer];
    const lines = book.linesByDeal[deal.id] || [];
    if (!customer) throw new Error('顧客「' + deal.customer + '」が顧客マスタにありません。');
    if (!lines.length) throw new Error('案件 ' + deal.id + ' の明細がありません。売上明細シートに明細を入力してください。');
    if (type === 'invoice' && deal.orderStatus === '失注') throw new Error('失注の案件には請求書を発行できません。');
    lines.forEach(function (l, i) {
      if (!l.product) throw new Error('明細 ' + (i + 1) + ' 行目の商品・サービスが空です。');
      if (!(Number(l.qty) > 0)) throw new Error('明細「' + l.product + '」の数量を入力してください。');
      if ([10, 8].indexOf(ratePercent(l.taxRate)) < 0) throw new Error('明細「' + l.product + '」の税率は 10% か 8% にしてください。');
    });
    if (!opt.force && getLinkedFileId_('deals', row, t.pdfKey)) return { ok: false, needsConfirm: true, number: deal[t.numberKey] };

    const today = opt.today || todayYmd_();
    const patch = {};
    let number = deal[t.numberKey];
    const issuedOn = deal[t.dateKey] || today;
    if (!number) {
      const existing = book.deals.map(function (d) { return d[t.numberKey]; });
      number = nextDocNumber(existing, settings[t.prefixKey], issuedOn);
      patch[t.numberKey] = number;
    }
    if (!deal[t.dateKey]) patch[t.dateKey] = issuedOn;
    let limitOn;
    if (type === 'quote') {
      limitOn = addDays(issuedOn, Number(settings.quoteValidDays) || 30);
      if (!deal.orderStatus) patch.orderStatus = '見積中';
    } else {
      limitOn = deal.dueOn || calcDueDate(issuedOn, customer.terms);
      if (!deal.dueOn) patch.dueOn = limitOn;
      if (deal.orderStatus !== '受注') patch.orderStatus = '受注';
    }

    const doc = {
      type: type,
      number: number,
      issuedOn: issuedOn,
      limitOn: limitOn,
      subject: deal.subject,
      customer: customer,
      lines: lines.map(function (l) {
        return { name: l.product, detail: l.detail, qty: l.qty, unit: l.unit, unitPrice: l.unitPrice, taxRate: l.taxRate, amount: calcLineAmount(l.qty, l.unitPrice) };
      }),
      note: type === 'quote' ? '本見積書の有効期限は発行日から' + (Number(settings.quoteValidDays) || 30) + '日です。' : '',
    };
    const html = renderDocumentHtml(doc, settings);
    const fileName = safeFileName(t.name + '_' + number + '_' + customer.name) + '.pdf';
    const pdf = Utilities.newBlob(html, 'text/html', fileName).getAs(MimeType.PDF).setName(fileName);
    const folder = getDocFolder_(settings, t.folder, issuedOn);
    // 同じ名前の古い版はゴミ箱へ（再発行時）
    const olds = folder.getFilesByName(fileName);
    while (olds.hasNext()) olds.next().setTrashed(true);
    const file = folder.createFile(pdf);

    if (Object.keys(patch).length) updateCells_('deals', row, patch);
    setLink_('deals', row, t.pdfKey, file.getUrl(), '開く');
    writeLog_(t.name + 'PDF作成', number, true, customer.name + '／' + deal.subject + '／' + formatYen(calcTotals(doc.lines).total) + '／' + file.getName());
    return { ok: true, number: number, url: file.getUrl(), fileId: file.getId(), total: calcTotals(doc.lines).total };
  } finally {
    lock.releaseLock();
  }
}

/** 保存先：<ルート>/<見積書|請求書>/<YYYY-MM> */
function getDocFolder_(settings, kind, ymd) {
  const root = getRootFolder_(settings);
  return getOrCreateChild_(getOrCreateChild_(root, kind), ymd.slice(0, 7));
}

function getRootFolder_(settings) {
  if (settings.driveFolderId) {
    try {
      const f = DriveApp.getFolderById(settings.driveFolderId);
      if (!f.isTrashed()) return f;
    } catch (e) {
      // 見つからない・権限がない場合は作り直す
    }
  }
  const name = '見積・請求管理（' + getSs_().getName() + '）';
  const folder = DriveApp.createFolder(name);
  setSetting_('driveFolderId', folder.getId());
  writeLog_('保存先フォルダ作成', name, true, folder.getUrl());
  return folder;
}

function getOrCreateChild_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 案件行の PDF ファイル（なければ null） */
function getDocFile_(type, row) {
  const id = getLinkedFileId_('deals', row, DOC_TYPES[type].pdfKey);
  if (!id) return null;
  try {
    const f = DriveApp.getFileById(id);
    return f.isTrashed() ? null : f;
  } catch (e) {
    return null;
  }
}
