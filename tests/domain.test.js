const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGas } = require('./gas-mock');

const { g } = loadGas();

test('支払期限：締め日と支払月の組み合わせ', () => {
  const cases = [
    ['2026-11-19', '月末締め翌月末払い', '2026-12-31'],
    ['2026-11-30', '月末締め翌月末払い', '2026-12-31'],
    ['2026-01-31', '月末締め翌月末払い', '2026-02-28'],
    ['2027-12-15', '月末締め翌月末払い', '2028-01-31'],
    ['2026-11-19', '月末締め翌々月末払い', '2027-01-31'],
    ['2026-12-01', '月末締め翌々月末払い', '2027-02-28'],
    ['2026-11-20', '20日締め翌月10日払い', '2026-12-10'],
    ['2026-11-21', '20日締め翌月10日払い', '2027-01-10'],
    ['2026-12-25', '20日締め翌月10日払い', '2027-02-10'],
    ['2026-11-15', '15日締め当月末払い', '2026-11-30'],
    ['2026-11-16', '15日締め当月末払い', '2026-12-31'],
    ['2026-02-10', '15日締め当月末払い', '2026-02-28'],
    ['2026-11-19', '請求日から30日以内', '2026-12-19'],
    ['2026-12-20', '請求日から14日以内', '2027-01-03'],
    ['2026-11-19', '存在しない条件', '2026-12-19'],
    ['2028-01-31', '月末締め翌月末払い', '2028-02-29'],
  ];
  for (const [inv, terms, due] of cases) assert.equal(g.calcDueDate(inv, terms), due, `${inv} ${terms}`);
});

test('日付の補助関数', () => {
  assert.equal(g.addDays('2026-12-31', 1), '2027-01-01');
  assert.equal(g.addDays('2028-03-01', -1), '2028-02-29');
  assert.equal(g.diffDays('2026-11-01', '2026-11-19'), 18);
  assert.equal(g.diffDays('2026-03-08', '2026-03-09'), 1);
  assert.equal(g.toYmd(new Date(2026, 10, 19, 23, 59)), '2026-11-19');
  assert.equal(g.toYmd('2026/1/5'), '2026-01-05');
  assert.equal(g.toYmd(''), '');
  assert.equal(g.toYmd(new Date('invalid')), '');
  assert.equal(g.toYmd(g.ymdToDate('2026-02-28')), '2026-02-28');
});

test('消費税：税率ごとに合計してから 1 回切り捨て', () => {
  const t = g.calcTotals([
    { qty: 1, unitPrice: 999, taxRate: 0.1 },
    { qty: 1, unitPrice: 999, taxRate: 0.1 },
    { qty: 3, unitPrice: 333, taxRate: 0.08 },
  ]);
  // 10%: 1998 → 199.8 → 199 / 8%: 999 → 79.92 → 79（行ごとに切り捨てると 99+99=198 になり違う）
  assert.equal(t.subtotal, 2997);
  assert.equal(t.tax, 199 + 79);
  assert.equal(t.total, 2997 + 278);
  assert.deepEqual(JSON.parse(JSON.stringify(t.byRate.map((r) => [r.pct, r.base, r.tax]))), [[10, 1998, 199], [8, 999, 79]]);
});

test('消費税：浮動小数の誤差で 1 円ずれない', () => {
  // 0.07 * 100 のような誤差が出やすい値でも整数計算で正しく切り捨てる
  for (const base of [100, 1100, 2900, 12345, 99999, 1234567]) {
    const t = g.calcTotals([{ qty: 1, unitPrice: base, taxRate: 0.08 }]);
    assert.equal(t.tax, Math.floor((base * 8) / 100), String(base));
  }
  assert.equal(g.ratePercent(0.1), 10);
  assert.equal(g.ratePercent(0.08), 8);
  assert.equal(g.ratePercent(10), 10);
  assert.equal(g.calcLineAmount(1.5, 15000), 22500);
  assert.equal(g.calcLineAmount(0.3333, 1000), 333);
});

test('入金ステータスの判定', () => {
  const today = '2026-11-19';
  const base = { orderStatus: '受注', invoiceNo: 'INV-202610-001', total: 110000, paid: 0, dueOn: '2026-11-30' };
  assert.equal(g.computePaymentStatus(Object.assign({}, base, { orderStatus: '失注' }), today).status, '失注');
  assert.equal(g.computePaymentStatus(Object.assign({}, base, { invoiceNo: '' }), today).status, '未請求');
  let s = g.computePaymentStatus(base, today);
  assert.equal(s.status, '未入金'); assert.equal(s.balance, 110000); assert.equal(s.daysPastDue, -11);
  s = g.computePaymentStatus(Object.assign({}, base, { paid: 50000 }), today);
  assert.equal(s.status, '一部入金'); assert.equal(s.balance, 60000);
  s = g.computePaymentStatus(Object.assign({}, base, { dueOn: '2026-11-19' }), today);
  assert.equal(s.status, '未入金', '期限当日はまだ超過ではない'); assert.equal(s.daysPastDue, 0);
  s = g.computePaymentStatus(Object.assign({}, base, { dueOn: '2026-11-10', paid: 10000 }), today);
  assert.equal(s.status, '期限超過'); assert.equal(s.daysPastDue, 9); assert.equal(s.balance, 100000);
  assert.equal(g.computePaymentStatus(Object.assign({}, base, { paid: 110000, dueOn: '2026-10-01' }), today).status, '入金済');
  assert.equal(g.computePaymentStatus(Object.assign({}, base, { paid: 120000 }), today).status, '入金済');
});

test('未入金アラート：期限超過と期限間近だけを、超過日数の大きい順に', () => {
  const today = '2026-11-19';
  const deals = [
    { invoiceNo: 'A', total: 100, paid: 0, dueOn: '2026-11-01' },
    { invoiceNo: 'B', total: 100, paid: 0, dueOn: '2026-11-25' },
    { invoiceNo: 'C', total: 100, paid: 0, dueOn: '2026-12-31' },
    { invoiceNo: 'D', total: 100, paid: 100, dueOn: '2026-10-01' },
    { invoiceNo: 'E', total: 100, paid: 30, dueOn: '2026-10-15' },
    { invoiceNo: '', total: 100, paid: 0, dueOn: '' },
  ];
  const a = g.buildAlerts(deals, today, 7);
  assert.deepEqual([...a.map((x) => x.deal.invoiceNo)], ['E', 'A', 'B']);
  assert.deepEqual([...a.map((x) => x.level)], ['overdue', 'overdue', 'soon']);
});

test('採番：月ごと・年ごとの連番', () => {
  assert.equal(g.nextDocNumber([], 'INV', '2026-11-19'), 'INV-202611-001');
  assert.equal(g.nextDocNumber(['INV-202611-001', 'INV-202611-009', 'INV-202610-020', 'Q-202611-050'], 'INV', '2026-11-30'), 'INV-202611-010');
  assert.equal(g.nextDocNumber(['INV-202611-099'], 'INV', '2026-11-30'), 'INV-202611-100');
  assert.equal(g.nextDocNumber(['INV-202611-005'], 'INV', '2026-12-01'), 'INV-202612-001');
  assert.equal(g.nextDealNumber(['PJ-2026-0041', 'PJ-2025-0100'], 'PJ', '2026-11-19'), 'PJ-2026-0042');
  assert.equal(g.nextDealNumber(['PJ-2026-0041'], 'PJ', '2027-01-02'), 'PJ-2027-0001');
});

test('表示・差し込み・宛名', () => {
  assert.equal(g.formatYen(1234567), '¥1,234,567');
  assert.equal(g.formatYen(-500), '-¥500');
  assert.equal(g.formatNumber(1.5), '1.5');
  assert.equal(g.formatJpDate('2026-01-05'), '2026年1月5日');
  assert.equal(g.fillTemplate('{件名}（{番号}）{未知}', { '件名': 'サイト制作', '番号': 'Q-1' }), 'サイト制作（Q-1）{未知}');
  assert.equal(g.escapeHtml('<a href="x">&\'</a>'), '&lt;a href=&quot;x&quot;&gt;&amp;&#39;&lt;/a&gt;');
  assert.deepEqual(Object.assign({}, g.buildAddressee({ name: 'A社', department: '総務部', contact: '山田' })), { company: 'A社', person: '総務部 山田 様', honorific: '' });
  assert.deepEqual(Object.assign({}, g.buildAddressee({ name: 'B社', department: '', contact: '' })), { company: 'B社', person: '', honorific: '御中' });
  assert.equal(g.safeFileName('請求書_INV/1:株式会社"A"'), '請求書_INV_1_株式会社_A_');
});

test('帳票 HTML：金額・税率別の内訳・エスケープ', () => {
  const html = g.renderDocumentHtml({
    type: 'invoice', number: 'INV-202611-001', issuedOn: '2026-11-19', limitOn: '2026-12-31', subject: '<script>件名</script>',
    customer: { name: '株式会社テスト', department: '経理部', contact: '田中', zip: '100-0000', address: '東京都' },
    lines: [
      { name: 'Web制作', detail: 'トップ', qty: 1, unit: '式', unitPrice: 100000, taxRate: 0.1 },
      { name: '軽食', detail: '', qty: 2, unit: '式', unitPrice: 1500, taxRate: 0.08 },
    ],
    note: '',
  }, { companyName: '株式会社ソラノ制作所', companyZip: '100-0000', companyAddress: '住所', companyTel: '03', companyEmail: 'a@example.com', registrationNo: 'T0000000000000', sealText: '印', bankName: 'デモ銀行', bankAccount: '普通 1', bankHolder: 'カ）ソラノ', accentColor: '#123456' });
  assert.match(html, /御請求書/);
  assert.match(html, /INV-202611-001/);
  assert.match(html, /2026年12月31日/);
  assert.match(html, /¥113,240/); // 100000+3000 + 10000 + 240
  assert.match(html, /8%対象（軽減税率）/);
  assert.match(html, /※ は軽減税率/);
  assert.match(html, /お振込先/);
  assert.match(html, /登録番号 T0000000000000/);
  assert.doesNotMatch(html, /<script>件名/);
  assert.match(html, /#123456/);
  const quote = g.renderDocumentHtml({ type: 'quote', number: 'Q-1', issuedOn: '2026-11-19', limitOn: '2026-12-19', subject: 's', customer: { name: 'X' }, lines: [{ name: 'a', qty: 1, unit: '式', unitPrice: 1, taxRate: 0.1 }], note: '' }, { accentColor: 'red' });
  assert.match(quote, /御見積書/);
  assert.doesNotMatch(quote, /お振込先/);
  assert.match(quote, /#1f6feb/, '不正な色は既定色に');
});
