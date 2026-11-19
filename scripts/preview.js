/**
 * 帳票（見積書・請求書）の HTML をデモデータで作り、Chromium で PNG/PDF にして見た目を確認する。
 * 使い方: npm run preview -- <出力先ディレクトリ>
 * ※ 本番の PDF は Apps Script の変換機能で作るため、細部（フォントなど）は少し異なる。
 */
const fs = require('fs');
const path = require('path');
const { loadGas } = require('../tests/gas-mock');

(async () => {
  const out = path.resolve(process.argv[2] || 'preview');
  fs.mkdirSync(out, { recursive: true });
  const { g } = loadGas();
  const settings = {};
  g.SETTINGS_SPEC.forEach((s) => { if (s.key) settings[s.key] = s.value; });
  const data = g.generateDemoData('2026-11-19');
  const customers = new Map(data.customers.map((c) => [c.name, c]));
  const pick = (pred) => data.deals.find(pred);
  const docs = [
    ['invoice', pick((d) => d.invoiceNo && d.lines.some((l) => l.taxRate === 0.08)) || pick((d) => d.invoiceNo)],
    ['quote', pick((d) => d.quoteNo && d.lines.length >= 3)],
  ];
  const { chromium } = require('playwright');
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 2 });
  for (const [type, d] of docs) {
    const html = g.renderDocumentHtml({
      type,
      number: type === 'quote' ? d.quoteNo : d.invoiceNo,
      issuedOn: type === 'quote' ? d.quoteOn : d.invoiceOn,
      limitOn: type === 'quote' ? g.addDays(d.quoteOn, 30) : d.dueOn,
      subject: d.subject,
      customer: customers.get(d.customer),
      lines: d.lines.map((l) => ({ name: l.product, detail: l.detail, qty: l.qty, unit: l.unit, unitPrice: l.unitPrice, taxRate: l.taxRate })),
      note: type === 'quote' ? '本見積書の有効期限は発行日から30日です。' : '',
    }, settings);
    fs.writeFileSync(path.join(out, `${type}.html`), html);
    await page.setContent(html, { waitUntil: 'load' });
    await page.pdf({ path: path.join(out, `${type}.pdf`), format: 'A4', printBackground: true, preferCSSPageSize: true });
    await page.emulateMedia({ media: 'print' });
    await page.screenshot({ path: path.join(out, `${type}.png`), fullPage: true });
    console.log(type, d.subject, '→', path.join(out, `${type}.png`));
  }
  await browser.close();
})();
