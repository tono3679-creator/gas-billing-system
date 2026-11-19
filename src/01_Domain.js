/**
 * 業務ロジック（純粋関数）。
 * スプレッドシートや Google のサービスに触れないので、Node でも単体テストできる。
 * 日付は 'YYYY-MM-DD' 文字列で扱い、タイムゾーンの揺れを持ち込まない。
 */

// ---------- 日付 ----------

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

/** Date（スクリプトのタイムゾーン）→ 'YYYY-MM-DD'。空や不正値は '' */
function toYmd(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'string') {
    const m = value.trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    return m ? m[1] + '-' + pad2_(Number(m[2])) + '-' + pad2_(Number(m[3])) : '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getFullYear() + '-' + pad2_(value.getMonth() + 1) + '-' + pad2_(value.getDate());
  }
  return '';
}

/** 'YYYY-MM-DD' → その日の 0 時の Date（スクリプトのタイムゾーン） */
function ymdToDate(ymd) {
  const p = ymd.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function daysInMonth_(year, month1) {
  return new Date(year, month1, 0).getDate();
}

function addDays(ymd, days) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + days));
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

/** 月を足して、指定日（31 なら月末）に合わせる */
function monthDay_(ymd, monthsLater, day) {
  const p = ymd.split('-').map(Number);
  const idx = p[0] * 12 + (p[1] - 1) + monthsLater;
  const y = Math.floor(idx / 12);
  const m = (idx % 12) + 1;
  const d = Math.min(day, daysInMonth_(y, m));
  return y + '-' + pad2_(m) + '-' + pad2_(d);
}

/** 2 つの日付の差（b - a、日数） */
function diffDays(aYmd, bYmd) {
  const a = aYmd.split('-').map(Number);
  const b = bYmd.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

function monthKey(ymd) {
  return ymd.slice(0, 7);
}

// ---------- 支払期限 ----------

function findPaymentTerms(label) {
  for (let i = 0; i < PAYMENT_TERMS.length; i++) {
    if (PAYMENT_TERMS[i].label === label) return PAYMENT_TERMS[i];
  }
  return null;
}

/**
 * 請求日と支払条件から支払期限を求める。
 * 例: 月末締め翌月末払い、請求日 2026-11-19 → 2026-12-31
 *     20日締め翌月10日払い、請求日 2026-11-25 → 締めは 12/20 → 2027-01-10
 * 支払条件が見つからないときは「請求日から30日」とみなす。
 */
function calcDueDate(invoiceYmd, termsLabel) {
  const t = findPaymentTerms(termsLabel) || { kind: 'net', days: 30 };
  if (t.kind === 'net') return addDays(invoiceYmd, t.days);
  const day = Number(invoiceYmd.slice(8, 10));
  const closingThisMonth = monthDay_(invoiceYmd, 0, t.closingDay);
  const closingMonthOffset = day <= Number(closingThisMonth.slice(8, 10)) ? 0 : 1;
  return monthDay_(invoiceYmd, closingMonthOffset + t.monthsLater, t.payDay);
}

// ---------- 金額 ----------

/** 税率を整数のパーセントに（0.1 → 10）。浮動小数の誤差を避けるため */
function ratePercent(rate) {
  const n = Number(rate);
  if (!isFinite(n)) return 0;
  return n < 1 ? Math.round(n * 100) : Math.round(n);
}

function calcLineAmount(qty, unitPrice) {
  const q = Number(qty) || 0;
  const p = Number(unitPrice) || 0;
  return Math.round(q * p);
}

/**
 * 明細から合計を出す。消費税は税率ごとに 1 回だけ端数を切り捨てる
 * （適格請求書の「税率ごとに区分した合計に対して1回の端数処理」）。
 * シート側の数式（案件・請求の消費税列）も同じ計算をする。
 */
function calcTotals(lines) {
  const byRate = {};
  let subtotal = 0;
  lines.forEach(function (l) {
    const amount = l.amount !== undefined && l.amount !== '' ? Number(l.amount) : calcLineAmount(l.qty, l.unitPrice);
    const pct = ratePercent(l.taxRate);
    subtotal += amount;
    byRate[pct] = (byRate[pct] || 0) + amount;
  });
  const rates = Object.keys(byRate)
    .map(Number)
    .sort(function (a, b) { return b - a; })
    .map(function (pct) {
      return { pct: pct, base: byRate[pct], tax: Math.floor((byRate[pct] * pct) / 100) };
    });
  const tax = rates.reduce(function (s, r) { return s + r.tax; }, 0);
  return { subtotal: subtotal, tax: tax, total: subtotal + tax, byRate: rates };
}

// ---------- 入金ステータス ----------

/**
 * 案件の入金状況を判定する。シートの「入金ステータス」列の数式と同じ規則。
 *  失注 → 失注 / 請求番号なし → 未請求 / 残高 0 以下 → 入金済
 *  期限 < 今日 → 期限超過 / 一部でも入金あり → 一部入金 / それ以外 → 未入金
 * daysPastDue は 今日 - 支払期限（正なら超過日数、負なら期限までの日数）。
 */
function computePaymentStatus(deal, todayYmd) {
  const total = Number(deal.total) || 0;
  const paid = Number(deal.paid) || 0;
  if (deal.orderStatus === '失注') return { status: '失注', balance: '', daysPastDue: '' };
  if (!deal.invoiceNo) return { status: '未請求', balance: '', daysPastDue: '' };
  const balance = total - paid;
  if (balance <= 0) return { status: '入金済', balance: balance, daysPastDue: '' };
  const due = deal.dueOn || '';
  const days = due ? diffDays(due, todayYmd) : '';
  let status;
  if (due && days > 0) status = '期限超過';
  else if (paid > 0) status = '一部入金';
  else status = '未入金';
  return { status: status, balance: balance, daysPastDue: days };
}

/**
 * 未入金アラートの一覧。期限超過と、期限まで soonDays 日以内の未入金を出す。
 * 超過日数の大きい順 → 期限の早い順。
 */
function buildAlerts(deals, todayYmd, soonDays) {
  const list = [];
  deals.forEach(function (d) {
    const s = computePaymentStatus(d, todayYmd);
    if (s.status === '期限超過') {
      list.push({ level: 'overdue', deal: d, status: s });
    } else if ((s.status === '未入金' || s.status === '一部入金') && s.daysPastDue !== '' && -s.daysPastDue <= soonDays) {
      list.push({ level: 'soon', deal: d, status: s });
    }
  });
  list.sort(function (a, b) {
    return b.status.daysPastDue - a.status.daysPastDue || (a.deal.dueOn < b.deal.dueOn ? -1 : 1);
  });
  return list;
}

// ---------- 採番 ----------

/**
 * 書類番号の次の番号。形式は 接頭辞-YYYYMM-連番3桁（月ごとに 001 から）。
 * 例: nextDocNumber(['INV-202611-001','INV-202611-004'], 'INV', '2026-11-19') → 'INV-202611-005'
 */
function nextDocNumber(existing, prefix, ymd) {
  const head = prefix + '-' + ymd.slice(0, 4) + ymd.slice(5, 7) + '-';
  let max = 0;
  existing.forEach(function (no) {
    const s = String(no || '');
    if (s.indexOf(head) === 0) {
      const n = Number(s.slice(head.length));
      if (isFinite(n) && n > max) max = n;
    }
  });
  const next = max + 1;
  return head + (next < 10 ? '00' : next < 100 ? '0' : '') + next;
}

/** 案件番号の次の番号。形式は 接頭辞-YYYY-連番4桁（年ごと） */
function nextDealNumber(existing, prefix, ymd) {
  const head = prefix + '-' + ymd.slice(0, 4) + '-';
  let max = 0;
  existing.forEach(function (no) {
    const s = String(no || '');
    if (s.indexOf(head) === 0) {
      const n = Number(s.slice(head.length));
      if (isFinite(n) && n > max) max = n;
    }
  });
  return head + ('000' + (max + 1)).slice(-4);
}

// ---------- 表示 ----------

function formatYen(n) {
  const v = Math.round(Number(n) || 0);
  const s = String(Math.abs(v)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (v < 0 ? '-' : '') + '¥' + s;
}

function formatNumber(n) {
  const v = Number(n) || 0;
  const parts = String(v).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function formatJpDate(ymd) {
  if (!ymd) return '';
  const p = ymd.split('-').map(Number);
  return p[0] + '年' + p[1] + '月' + p[2] + '日';
}

function escapeHtml(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** '{件名}' のような差し込みを置き換える。未定義のキーは残す */
function fillTemplate(template, vars) {
  return String(template || '').replace(/\{([^{}]+)\}/g, function (m, key) {
    return Object.prototype.hasOwnProperty.call(vars, key) ? String(vars[key]) : m;
  });
}

/** 宛名。担当者がいれば「会社名\n部署 担当者 様」、いなければ「会社名 御中」 */
function buildAddressee(customer) {
  const person = [customer.department, customer.contact].filter(Boolean).join(' ');
  if (customer.contact) return { company: customer.name, person: person + ' 様', honorific: '' };
  return { company: customer.name, person: customer.department || '', honorific: '御中' };
}

/** ファイル名に使えない文字を除く */
function safeFileName(s) {
  return String(s || '').replace(/[\\/:*?"<>|\r\n]+/g, '_').slice(0, 120);
}
