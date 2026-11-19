/**
 * デモ用の架空データを作る純粋関数。
 * 「今日」を基準に直近 12 か月分を作るので、いつセットアップしても
 * ダッシュボードに当月・前月の数字や期限超過の案件が並ぶ。
 * 乱数は種を固定しているため、同じ日付なら毎回同じデータになる。
 * 会社名・人名・住所・メールはすべて架空（メールは予約ドメイン .example）。
 */

function seededRandom_(seed) {
  let a = seed >>> 0;
  return function () {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const DEMO_CUSTOMERS = [
  ['C-001', '株式会社北斗フーズ', '営業企画部', '山本 健一', 'yamamoto@hokuto-foods.example', '03-0000-1101', '105-0000', '東京都港区芝公園0-1-1', '月末締め翌月末払い', '月額保守契約あり'],
  ['C-002', '青葉デンタルクリニック', '', '青木 由美', 'info@aoba-dental.example', '045-000-2202', '220-0000', '神奈川県横浜市西区みなとみらい0-2-2', '請求日から30日以内', ''],
  ['C-003', '有限会社みなと工務店', '総務部', '港 一郎', 'minato@minato-koumuten.example', '044-000-3303', '210-0000', '神奈川県川崎市川崎区港町0-3-3', '20日締め翌月10日払い', ''],
  ['C-004', '株式会社さくら不動産', '販売推進課', '桜井 美咲', 'sakurai@sakura-fudosan.example', '03-0000-4404', '160-0000', '東京都新宿区西新宿0-4-4', '月末締め翌月末払い', 'SNS運用代行の月額契約あり'],
  ['C-005', '合同会社ヒカリ商事', '', '', 'contact@hikari-shoji.example', '048-000-5505', '330-0000', '埼玉県さいたま市大宮区桜木町0-5-5', '月末締め翌々月末払い', '担当者未定のため御中宛'],
  ['C-006', '株式会社東雲ロジスティクス', '管理本部', '東 拓也', 'azuma@shinonome-logi.example', '03-0000-6606', '135-0000', '東京都江東区東雲0-6-6', '月末締め翌月末払い', ''],
  ['C-007', '医療法人やまびこ会', '事務局', '山田 花子', 'jimu@yamabiko-kai.example', '042-000-7707', '190-0000', '東京都立川市曙町0-7-7', '15日締め当月末払い', '口座振替'],
  ['C-008', '株式会社こだま印刷', '制作部', '児玉 誠', 'kodama@kodama-print.example', '03-0000-8808', '113-0000', '東京都文京区本郷0-8-8', '20日締め翌月10日払い', ''],
  ['C-009', 'NPO法人つばさの会', '事務局', '高橋 翼', 'office@tsubasa-npo.example', '03-0000-9909', '150-0000', '東京都渋谷区神宮前0-9-9', '請求日から14日以内', ''],
  ['C-010', '株式会社ミナミ精機', '総務部', '南 和也', 'minami@minami-seiki.example', '0466-00-1010', '251-0000', '神奈川県藤沢市辻堂0-10-10', '月末締め翌月末払い', ''],
];

const DEMO_PRODUCTS = [
  ['P-001', 'Webサイト制作（コーポレート）', '制作', '式', 600000, 0.1, 'トップ＋下層10ページ程度、レスポンシブ対応'],
  ['P-002', 'ランディングページ制作', '制作', '式', 250000, 0.1, '1ページ完結型、フォーム設置込み'],
  ['P-003', 'ページ追加・更新作業', '制作', 'ページ', 30000, 0.1, '既存デザインに沿った追加・更新'],
  ['P-004', 'サイト保守・運用（月額）', '保守', '月', 30000, 0.1, 'サーバー監視・更新代行（月5件まで）'],
  ['P-005', 'SNS運用代行（月額）', '運用', '月', 80000, 0.1, '投稿企画・作成 月12本'],
  ['P-006', '撮影ディレクション・撮影', '撮影', '日', 120000, 0.1, 'カメラマン1名・レタッチ込み'],
  ['P-007', 'チラシ・パンフレットデザイン', 'デザイン', '式', 150000, 0.1, 'A4 両面〜8ページ'],
  ['P-008', 'DX・業務改善コンサルティング', 'コンサル', '時間', 15000, 0.1, '業務フロー整理・ツール導入支援'],
  ['P-009', '打合せ用 飲料・軽食', '実費', '式', 3000, 0.08, '軽減税率対象（飲食料品）'],
  ['P-010', '交通費・出張費', '実費', '式', 10000, 0.1, '実費精算'],
  ['P-011', '印刷費（A4チラシ 5,000部）', '実費', '式', 45000, 0.1, '提携印刷会社への外注費'],
];

/** 案件のひな形。lines は [商品ID, 数量, 摘要] */
const DEMO_PROJECTS = [
  { subject: 'コーポレートサイト リニューアル', days: 45, lines: [['P-001', 1, 'デザイン刷新・CMS導入'], ['P-003', 5, '事業紹介ページ'], ['P-006', 1, '社屋・スタッフ撮影'], ['P-010', 1, '']] },
  { subject: '新商品ランディングページ制作', days: 21, lines: [['P-002', 1, ''], ['P-006', 1, '商品撮影']] },
  { subject: '採用サイト制作', days: 35, lines: [['P-001', 1, '採用特設サイト'], ['P-003', 3, '社員インタビュー']] },
  { subject: '会社案内パンフレット制作', days: 28, lines: [['P-007', 1, 'A4 8ページ'], ['P-011', 1, ''], ['P-006', 1, '']] },
  { subject: '受発注業務のDX支援', days: 30, lines: [['P-008', 24, '現状分析・スプレッドシート自動化'], ['P-009', 3, '定例会議']] },
  { subject: '社内向けDX研修', days: 14, lines: [['P-008', 8, '研修講師・資料作成'], ['P-009', 2, '研修時']] },
  { subject: 'Webサイト ページ追加', days: 12, lines: [['P-003', 4, 'お知らせ・実績ページ']] },
  { subject: 'イベント告知チラシ制作', days: 15, lines: [['P-007', 1, 'A4 両面'], ['P-011', 1, '']] },
  { subject: '店舗・メニュー撮影', days: 10, lines: [['P-006', 1, ''], ['P-010', 1, '']] },
];

/** 各案件の顧客（ひな形と月の組み合わせを決め打ちにして毎回同じ並びにする） */
const DEMO_SCHEDULE = [
  // [何か月前, 顧客ID, ひな形番号]
  [11, 'C-002', 1], [11, 'C-006', 6],
  [10, 'C-003', 3], [10, 'C-009', 5],
  [9, 'C-010', 4], [9, 'C-008', 7],
  [8, 'C-005', 0], [8, 'C-002', 8],
  [7, 'C-007', 2], [7, 'C-003', 6],
  [6, 'C-006', 4], [6, 'C-009', 7], [6, 'C-008', 1],
  [5, 'C-010', 5], [5, 'C-002', 6],
  [4, 'C-007', 3], [4, 'C-005', 8], [4, 'C-006', 1],
  [3, 'C-003', 0], [3, 'C-008', 4],
  [2, 'C-009', 2], [2, 'C-010', 7], [2, 'C-007', 6],
  [1, 'C-002', 4], [1, 'C-005', 1], [1, 'C-008', 8],
  [0, 'C-006', 5], [0, 'C-003', 7], [0, 'C-010', 1], [0, 'C-009', 3],
];

/** 失注にする案件（スケジュールの番号） */
const DEMO_LOST = [3, 16];

function generateDemoData(todayYmd) {
  const rnd = seededRandom_(20261119);
  const customers = DEMO_CUSTOMERS.map(function (c) {
    return { id: c[0], name: c[1], department: c[2], contact: c[3], email: c[4], phone: c[5], zip: c[6], address: c[7], terms: c[8], note: c[9] };
  });
  const products = DEMO_PRODUCTS.map(function (p) {
    return { id: p[0], name: p[1], category: p[2], unit: p[3], unitPrice: p[4], taxRate: p[5], description: p[6], active: true };
  });
  const productById = {};
  products.forEach(function (p) { productById[p.id] = p; });
  const customerById = {};
  customers.forEach(function (c) { customerById[c.id] = c; });

  const yesterday = addDays(todayYmd, -1);
  const deals = [];

  // 1) 単発の案件
  DEMO_SCHEDULE.forEach(function (s, idx) {
    const monthsAgo = s[0];
    const tpl = DEMO_PROJECTS[s[2]];
    const firstOfMonth = monthDay_(todayYmd, -monthsAgo, 1);
    const maxDay = monthsAgo === 0 ? Math.max(1, Number(todayYmd.slice(8, 10)) - 1) : 26;
    let receivedOn = addDays(firstOfMonth, Math.floor(rnd() * maxDay));
    if (receivedOn > todayYmd) receivedOn = todayYmd;
    let quoteOn = addDays(receivedOn, 1 + Math.floor(rnd() * 3));
    if (quoteOn > todayYmd) quoteOn = todayYmd;
    const finishOn = addDays(quoteOn, tpl.days + Math.floor(rnd() * 6));
    const lines = tpl.lines.map(function (l) {
      const p = productById[l[0]];
      // 案件ごとに単価を少し変える（値引き・規模の違い）
      const adjust = p.category === '実費' ? 1 : [1, 1, 0.9, 1.2, 1.1][Math.floor(rnd() * 5)];
      return { product: p.name, detail: l[2], qty: l[1], unit: p.unit, unitPrice: Math.round((p.unitPrice * adjust) / 1000) * 1000 || p.unitPrice, taxRate: p.taxRate };
    });
    const deal = {
      receivedOn: receivedOn,
      customerId: s[1],
      customer: customerById[s[1]].name,
      subject: tpl.subject,
      quoteOn: quoteOn,
      lines: lines,
      orderStatus: '受注',
      invoiceOn: '',
      note: '',
    };
    if (DEMO_LOST.indexOf(idx) >= 0) {
      deal.orderStatus = '失注';
      deal.note = '他社に決定';
    } else if (finishOn <= yesterday) {
      deal.invoiceOn = finishOn;
    } else if (diffDays(quoteOn, todayYmd) <= 12) {
      deal.orderStatus = '見積中';
    }
    deals.push(deal);
  });

  // 2) 月額契約（保守・SNS運用）。前月分までを月末に請求
  [['C-001', 'P-004', 'サイト保守・運用'], ['C-004', 'P-005', 'SNS運用代行']].forEach(function (r) {
    const p = productById[r[1]];
    for (let m = 11; m >= 0; m--) {
      const first = monthDay_(todayYmd, -m, 1);
      const last = monthDay_(todayYmd, -m, 31);
      const label = Number(first.slice(0, 4)) + '年' + Number(first.slice(5, 7)) + '月分';
      deals.push({
        receivedOn: first,
        customerId: r[0],
        customer: customerById[r[0]].name,
        subject: r[2] + '（' + label + '）',
        quoteOn: '',
        lines: [{ product: p.name, detail: label, qty: 1, unit: p.unit, unitPrice: p.unitPrice, taxRate: p.taxRate }],
        orderStatus: '受注',
        invoiceOn: last <= yesterday ? last : '',
        note: '月額契約',
      });
    }
  });

  // 番号を日付順に振る
  deals.sort(function (a, b) { return a.receivedOn < b.receivedOn ? -1 : a.receivedOn > b.receivedOn ? 1 : 0; });
  const dealNos = [];
  deals.forEach(function (d) {
    d.id = nextDealNumber(dealNos, 'PJ', d.receivedOn);
    dealNos.push(d.id);
  });
  const quoteNos = [];
  deals.filter(function (d) { return d.quoteOn; })
    .sort(function (a, b) { return a.quoteOn < b.quoteOn ? -1 : a.quoteOn > b.quoteOn ? 1 : 0; })
    .forEach(function (d) {
      d.quoteNo = nextDocNumber(quoteNos, 'Q', d.quoteOn);
      quoteNos.push(d.quoteNo);
    });
  const invoiceNos = [];
  deals.filter(function (d) { return d.invoiceOn; })
    .sort(function (a, b) { return a.invoiceOn < b.invoiceOn ? -1 : a.invoiceOn > b.invoiceOn ? 1 : 0; })
    .forEach(function (d) {
      d.invoiceNo = nextDocNumber(invoiceNos, 'INV', d.invoiceOn);
      invoiceNos.push(d.invoiceNo);
      d.dueOn = calcDueDate(d.invoiceOn, customerById[d.customerId].terms);
    });

  // 3) 入金。期限を過ぎたものは基本的に入金済み。見どころとして数件だけ未入金・一部入金にする
  const payments = [];
  const pastDue = deals
    .filter(function (d) { return d.invoiceNo && d.dueOn < todayYmd; })
    .sort(function (a, b) { return a.dueOn > b.dueOn ? -1 : a.dueOn < b.dueOn ? 1 : 0; });
  const unpaidIdx = [0, 3, 7]; // 直近・中くらい・かなり前に期限が来たもの
  const partialIdx = [4];
  pastDue.forEach(function (d, i) {
    const total = calcTotals(d.lines).total;
    const method = d.customerId === 'C-007' ? '口座振替' : '銀行振込';
    if (unpaidIdx.indexOf(i) >= 0) {
      d.note = d.note ? d.note + '／' : '';
      d.note += '入金確認待ち';
      return;
    }
    if (partialIdx.indexOf(i) >= 0) {
      const half = Math.round(total / 2 / 1000) * 1000;
      payments.push({ paidOn: addDays(d.dueOn, -2), invoiceNo: d.invoiceNo, customer: d.customer, amount: half, method: method, note: '分割1回目（残額は翌月入金予定と連絡あり）' });
      return;
    }
    let paidOn = addDays(d.dueOn, -Math.floor(rnd() * 4));
    if (paidOn < d.invoiceOn) paidOn = d.invoiceOn;
    payments.push({ paidOn: paidOn, invoiceNo: d.invoiceNo, customer: d.customer, amount: total, method: method, note: '' });
  });
  // 期限前に入金された例を 1 件
  const notYetDue = deals.filter(function (d) { return d.invoiceNo && d.dueOn >= todayYmd; });
  if (notYetDue.length > 1) {
    const d = notYetDue[0];
    let paidOn = addDays(d.invoiceOn, 5);
    if (paidOn > todayYmd) paidOn = todayYmd;
    payments.push({ paidOn: paidOn, invoiceNo: d.invoiceNo, customer: d.customer, amount: calcTotals(d.lines).total, method: '銀行振込', note: '期日前入金' });
  }
  payments.sort(function (a, b) { return a.paidOn < b.paidOn ? -1 : a.paidOn > b.paidOn ? 1 : 0; });

  return { customers: customers, products: products, deals: deals, payments: payments };
}
