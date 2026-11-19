const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGas } = require('./gas-mock');

const { g } = loadGas();

// どの日にセットアップしても、見どころがそろったデモデータになることを確かめる
const DAYS = ['2026-11-19', '2026-11-01', '2026-11-30', '2027-01-01', '2027-01-02', '2028-02-29', '2026-03-31', '2026-07-15'];

for (const today of DAYS) {
  test(`デモデータの整合性（今日 = ${today}）`, () => {
    const data = g.generateDemoData(today);
    const customers = new Map(data.customers.map((c) => [c.name, c]));
    const yesterday = g.addDays(today, -1);

    assert.equal(data.customers.length, 10);
    assert.equal(data.products.length, 11);
    assert.ok(data.deals.length >= 50, '案件は 50 件以上');

    // 番号の重複なし・形式
    for (const key of ['id', 'quoteNo', 'invoiceNo']) {
      const list = data.deals.map((d) => d[key]).filter(Boolean);
      assert.equal(new Set(list).size, list.length, `${key} が重複`);
    }
    for (const d of data.deals) {
      assert.match(d.id, /^PJ-\d{4}-\d{4}$/);
      assert.ok(customers.has(d.customer), '顧客がマスタにある');
      assert.ok(d.lines.length >= 1, '明細がある');
      assert.ok(d.receivedOn <= today, '受付日は今日以前');
      if (d.quoteNo) {
        assert.match(d.quoteNo, /^Q-\d{6}-\d{3}$/);
        assert.equal(d.quoteNo.slice(2, 8), d.quoteOn.slice(0, 4) + d.quoteOn.slice(5, 7), '見積番号の年月 = 見積日');
        assert.ok(d.quoteOn <= today);
      }
      if (d.invoiceNo) {
        assert.match(d.invoiceNo, /^INV-\d{6}-\d{3}$/);
        assert.ok(d.invoiceOn <= yesterday, '請求日は昨日以前');
        assert.ok(d.invoiceOn >= d.receivedOn);
        assert.equal(d.dueOn, g.calcDueDate(d.invoiceOn, customers.get(d.customer).terms), '期限は支払条件どおり');
        assert.notEqual(d.orderStatus, '失注');
      }
      for (const l of d.lines) {
        assert.ok(l.qty > 0 && l.unitPrice > 0);
        assert.ok([0.1, 0.08].includes(l.taxRate));
      }
    }

    // 入金：請求番号が存在し、合計を超えず、日付が今日以前
    const byInv = new Map(data.deals.filter((d) => d.invoiceNo).map((d) => [d.invoiceNo, d]));
    const paid = new Map();
    for (const p of data.payments) {
      assert.ok(byInv.has(p.invoiceNo), '入金の請求番号が存在');
      assert.ok(p.paidOn <= today, '入金日は今日以前');
      assert.ok(p.paidOn >= byInv.get(p.invoiceNo).invoiceOn, '入金日は請求日以降');
      paid.set(p.invoiceNo, (paid.get(p.invoiceNo) || 0) + p.amount);
    }
    for (const [inv, amount] of paid) assert.ok(amount <= g.calcTotals(byInv.get(inv).lines).total, `${inv} の入金が請求額を超えない`);

    // 見どころ：期限超過（全額未入金と一部入金）、見積中、失注、未入金（期限前）
    const statuses = data.deals.map((d) => g.computePaymentStatus({ orderStatus: d.orderStatus, invoiceNo: d.invoiceNo, total: g.calcTotals(d.lines).total, paid: paid.get(d.invoiceNo) || 0, dueOn: d.dueOn }, today));
    const count = (s) => statuses.filter((x) => x.status === s).length;
    assert.ok(count('期限超過') >= 3, `期限超過が 3 件以上（${count('期限超過')}）`);
    assert.ok(statuses.some((x, i) => x.status === '期限超過' && (paid.get(data.deals[i].invoiceNo) || 0) > 0), '一部入金のまま期限超過の案件がある');
    assert.ok(count('入金済') >= 20, '入金済が十分ある');
    assert.ok(count('失注') === 2);
    assert.ok(count('未入金') + count('一部入金') >= 1, '期限前の未入金がある');
    assert.ok(data.deals.some((d) => d.orderStatus === '見積中'), '見積中の案件がある');
  });
}

test('同じ日付なら同じデータ（乱数の種を固定）', () => {
  assert.equal(JSON.stringify(g.generateDemoData('2026-11-19')), JSON.stringify(g.generateDemoData('2026-11-19')));
});

test('デモの連絡先は予約ドメインだけ（実在のアドレスに届かない）', () => {
  for (const c of g.generateDemoData('2026-11-19').customers) assert.match(c.email, /\.example$/);
});
