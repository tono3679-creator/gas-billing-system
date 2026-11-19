/**
 * 見積書・請求書の HTML（PDF 化の元）を組み立てる純粋関数。
 * Apps Script の HTML→PDF 変換はフレックスボックスなどに弱いため、
 * レイアウトは table と固定幅だけで組んでいる。
 */

const DOC_LABEL = {
  quote: { title: '御見積書', lead: '下記の通りお見積り申し上げます。', amountLabel: 'お見積金額（税込）', dateLabel: '有効期限' },
  invoice: { title: '御請求書', lead: '下記の通りご請求申し上げます。', amountLabel: 'ご請求金額（税込）', dateLabel: 'お支払期限' },
};

/**
 * @param {Object} doc  { type, number, issuedOn, limitOn, subject, customer, lines, note }
 * @param {Object} co   設定シートの値（getSettings() の戻り値）
 */
function renderDocumentHtml(doc, co) {
  const label = DOC_LABEL[doc.type];
  const totals = calcTotals(doc.lines);
  const accent = /^#[0-9a-fA-F]{6}$/.test(String(co.accentColor || '')) ? co.accentColor : '#1f6feb';
  const to = buildAddressee(doc.customer);
  const hasReduced = totals.byRate.some(function (r) { return r.pct === 8; });

  const rowsHtml = doc.lines.map(function (l, i) {
    const reduced = ratePercent(l.taxRate) === 8;
    const amount = l.amount !== undefined && l.amount !== '' ? l.amount : calcLineAmount(l.qty, l.unitPrice);
    return (
      '<tr>' +
      '<td class="c no">' + (i + 1) + '</td>' +
      '<td class="name">' + escapeHtml(l.name) + (reduced ? ' <span class="mark">※</span>' : '') +
      (l.detail ? '<div class="detail">' + escapeHtml(l.detail) + '</div>' : '') + '</td>' +
      '<td class="r">' + escapeHtml(formatNumber(l.qty)) + '</td>' +
      '<td class="c">' + escapeHtml(l.unit) + '</td>' +
      '<td class="r">' + escapeHtml(formatYen(l.unitPrice)) + '</td>' +
      '<td class="c">' + ratePercent(l.taxRate) + '%</td>' +
      '<td class="r">' + escapeHtml(formatYen(amount)) + '</td>' +
      '</tr>'
    );
  });
  // 明細が少ないときも表の高さが極端に変わらないよう空行で埋める
  for (let i = doc.lines.length; i < 8; i++) {
    rowsHtml.push('<tr class="blank"><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>');
  }

  const rateRows = totals.byRate.map(function (r) {
    return (
      '<tr><th>' + r.pct + '%対象' + (r.pct === 8 ? '（軽減税率）' : '') + '</th>' +
      '<td>' + escapeHtml(formatYen(r.base)) + '</td>' +
      '<th>消費税</th><td>' + escapeHtml(formatYen(r.tax)) + '</td></tr>'
    );
  }).join('');

  // 社印：文字を縦書き 2 列（右の列から読む）に並べた角印
  let seal = '';
  if (co.sealText) {
    const chars = Array.from(String(co.sealText).replace(/\s/g, '')).slice(0, 12);
    const perCol = chars.length <= 2 ? chars.length : Math.ceil(chars.length / 2);
    const right = chars.slice(0, perCol);
    const left = chars.slice(perCol);
    const col = function (cs) { return '<td>' + cs.map(escapeHtml).join('<br>') + '</td>'; };
    seal = '<table class="seal"><tr>' + (left.length ? col(left) : '') + col(right) + '</tr></table>';
  }

  const bank = doc.type === 'invoice'
    ? '<table class="box"><tr><th>お振込先</th></tr><tr><td>' +
      escapeHtml(co.bankName) + '<br>' + escapeHtml(co.bankAccount) + '<br>口座名義：' + escapeHtml(co.bankHolder) +
      '<div class="small">※ 恐れ入りますが、振込手数料は貴社にてご負担ください。</div></td></tr></table>'
    : '';

  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>',
    '@page { size: A4; margin: 14mm 14mm 14mm 14mm; }',
    'body { font-family: "Noto Sans JP", "Hiragino Sans", "Yu Gothic", sans-serif; color: #1f2937; font-size: 9.5pt; margin: 0; }',
    'table { border-collapse: collapse; }',
    '.bar { height: 5px; background: ' + accent + '; margin-bottom: 14px; }',
    '.top { width: 100%; }',
    '.top td { vertical-align: top; }',
    'h1 { font-size: 20pt; letter-spacing: 8px; margin: 0 0 4px 0; font-weight: bold; }',
    '.meta { font-size: 9pt; color: #4b5563; }',
    '.meta th { text-align: left; font-weight: normal; color: #6b7280; padding: 1px 10px 1px 0; }',
    '.meta td { text-align: right; padding: 1px 0; }',
    '.to { font-size: 13pt; font-weight: bold; border-bottom: 1.5px solid #1f2937; padding-bottom: 4px; margin-top: 16px; }',
    '.to .hon { font-size: 11pt; margin-left: 8px; }',
    '.toperson { font-size: 10.5pt; margin-top: 4px; }',
    '.toaddr { font-size: 8.5pt; color: #6b7280; margin-top: 2px; }',
    '.lead { margin: 14px 0 8px 0; }',
    '.issuer { font-size: 9pt; line-height: 1.6; }',
    '.issuer .cname { font-size: 11.5pt; font-weight: bold; }',
    '.seal { border: 2px solid #c0392b; border-radius: 4px; margin-top: 4px; }',
    '.seal td { color: #c0392b; font-size: 9pt; font-weight: bold; line-height: 1.12; text-align: center; padding: 4px 3px; vertical-align: middle; }',
    '.amount { width: 100%; margin: 6px 0 14px 0; }',
    '.amount th { background: #f3f4f6; text-align: left; padding: 9px 12px; font-size: 10pt; width: 34%; border-left: 4px solid ' + accent + '; }',
    '.amount td { padding: 9px 12px; font-size: 18pt; font-weight: bold; border-bottom: 1.5px solid #1f2937; }',
    '.subject { margin-bottom: 10px; }',
    '.subject th { text-align: left; color: #6b7280; font-weight: normal; padding: 2px 12px 2px 0; }',
    '.lines { width: 100%; }',
    '.lines th { background: #1f2a44; color: #fff; font-weight: normal; padding: 6px 6px; font-size: 8.5pt; }',
    '.lines td { border-bottom: 1px solid #e5e7eb; padding: 6px 6px; vertical-align: top; }',
    '.lines tr.blank td { height: 16px; }',
    '.lines .no { width: 22px; color: #9ca3af; }',
    '.detail { font-size: 8pt; color: #6b7280; margin-top: 2px; }',
    '.mark { color: #b42318; }',
    '.r { text-align: right; } .c { text-align: center; }',
    '.sum { width: 100%; margin-top: 10px; }',
    '.sum td { vertical-align: top; }',
    '.totals { width: 100%; }',
    '.totals th { text-align: left; font-weight: normal; color: #4b5563; padding: 4px 8px; background: #f9fafb; border-bottom: 1px solid #e5e7eb; }',
    '.totals td { text-align: right; padding: 4px 8px; border-bottom: 1px solid #e5e7eb; }',
    '.totals tr.grand th, .totals tr.grand td { font-weight: bold; font-size: 11pt; color: #111827; background: #eef2ff; border-bottom: 1.5px solid #1f2937; }',
    '.rates { width: 100%; margin-top: 6px; font-size: 8.5pt; }',
    '.rates th { text-align: left; font-weight: normal; color: #6b7280; padding: 2px 6px; }',
    '.rates td { text-align: right; padding: 2px 6px; }',
    '.box { width: 100%; margin-top: 12px; }',
    '.box th { text-align: left; background: #f3f4f6; padding: 5px 8px; font-size: 8.5pt; color: #374151; }',
    '.box td { padding: 6px 8px; border: 1px solid #e5e7eb; line-height: 1.6; }',
    '.small { font-size: 7.5pt; color: #6b7280; margin-top: 4px; }',
    '.foot { margin-top: 14px; font-size: 7.5pt; color: #9ca3af; text-align: right; }',
    '</style></head><body>',
    '<div class="bar"></div>',
    '<table class="top"><tr>',
    '<td style="width:58%">',
    '<h1>' + label.title + '</h1>',
    '<div class="to">' + escapeHtml(to.company) + (to.honorific ? '<span class="hon">' + to.honorific + '</span>' : '') + '</div>',
    to.person ? '<div class="toperson">' + escapeHtml(to.person) + '</div>' : '',
    doc.customer.address ? '<div class="toaddr">〒' + escapeHtml(doc.customer.zip) + ' ' + escapeHtml(doc.customer.address) + '</div>' : '',
    '<div class="lead">' + label.lead + '</div>',
    '</td>',
    '<td style="width:42%; text-align:right">',
    '<table class="meta" align="right">',
    '<tr><th>番号</th><td>' + escapeHtml(doc.number) + '</td></tr>',
    '<tr><th>発行日</th><td>' + formatJpDate(doc.issuedOn) + '</td></tr>',
    '<tr><th>' + label.dateLabel + '</th><td>' + formatJpDate(doc.limitOn) + '</td></tr>',
    '</table>',
    '<div style="clear:both"></div>',
    '<table align="right" style="margin-top:12px"><tr>',
    '<td class="issuer" style="text-align:left">',
    '<div class="cname">' + escapeHtml(co.companyName) + '</div>',
    '〒' + escapeHtml(co.companyZip) + '<br>' + escapeHtml(co.companyAddress) + '<br>',
    'TEL ' + escapeHtml(co.companyTel) + '<br>' + escapeHtml(co.companyEmail) + '<br>',
    co.registrationNo ? '登録番号 ' + escapeHtml(co.registrationNo) + '<br>' : '',
    co.representative ? '担当：' + escapeHtml(co.representative) : '',
    '</td><td style="vertical-align:top; padding-left:8px">' + seal + '</td>',
    '</tr></table>',
    '</td></tr></table>',
    '<table class="amount"><tr><th>' + label.amountLabel + '</th><td>' + escapeHtml(formatYen(totals.total)) + ' -</td></tr></table>',
    '<table class="subject"><tr><th>件名</th><td>' + escapeHtml(doc.subject) + '</td></tr></table>',
    '<table class="lines"><thead><tr>',
    '<th>No.</th><th style="text-align:left">品名・内容</th><th style="width:52px">数量</th><th style="width:40px">単位</th>',
    '<th style="width:80px">単価</th><th style="width:40px">税率</th><th style="width:90px">金額</th>',
    '</tr></thead><tbody>',
    rowsHtml.join(''),
    '</tbody></table>',
    '<table class="sum"><tr>',
    '<td style="width:52%; padding-right:16px">',
    hasReduced ? '<div class="small">※ は軽減税率（8%）対象品目です。</div>' : '',
    bank,
    doc.note ? '<table class="box"><tr><th>備考</th></tr><tr><td>' + escapeHtml(doc.note).replace(/\n/g, '<br>') + '</td></tr></table>' : '',
    '</td>',
    '<td style="width:48%">',
    '<table class="totals">',
    '<tr><th>小計（税抜）</th><td>' + escapeHtml(formatYen(totals.subtotal)) + '</td></tr>',
    '<tr><th>消費税</th><td>' + escapeHtml(formatYen(totals.tax)) + '</td></tr>',
    '<tr class="grand"><th>合計（税込）</th><td>' + escapeHtml(formatYen(totals.total)) + '</td></tr>',
    '</table>',
    '<table class="rates">' + rateRows + '</table>',
    '</td></tr></table>',
    '<div class="foot">' + escapeHtml(co.companyName) + ' ・ ' + escapeHtml(doc.number) + '</div>',
    '</body></html>',
  ].join('\n');
}
