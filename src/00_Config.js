/**
 * 設定値・シート定義。
 * Apps Script は全ファイルが同じグローバル空間で読み込まれるため、
 * ここで定義した定数は他のファイルからそのまま参照できる。
 */

const APP = {
  name: '見積・請求・入金管理',
  version: '1.0.0',
  timeZone: 'Asia/Tokyo',
  menuName: '請求管理',
};

const SHEET = {
  dashboard: 'ダッシュボード',
  deals: '案件・請求',
  lines: '売上明細',
  payments: '入金記録',
  customers: '顧客マスタ',
  products: '商品マスタ',
  settings: '設定',
  log: '操作ログ',
  selfTest: 'セルフテスト',
  guide: '使い方',
};

/** 表示順（セットアップ時にこの順に並べる） */
const SHEET_ORDER = [
  SHEET.dashboard,
  SHEET.deals,
  SHEET.lines,
  SHEET.payments,
  SHEET.customers,
  SHEET.products,
  SHEET.settings,
  SHEET.log,
  SHEET.guide,
];

/**
 * 列定義。key は 1 始まりの列番号に変換して COL.* で使う。
 * calc: true の列は数式で自動計算される（手入力しない列）。
 */
const COLUMNS = {
  deals: [
    { key: 'id', label: '案件番号', width: 120 },
    { key: 'receivedOn', label: '受付日', width: 96, type: 'date' },
    { key: 'customer', label: '顧客名', width: 210 },
    { key: 'customerId', label: '顧客ID', width: 72, calc: true },
    { key: 'subject', label: '件名', width: 240 },
    { key: 'orderStatus', label: '受注状況', width: 90 },
    { key: 'subtotal', label: '税抜金額', width: 110, type: 'yen', calc: true },
    { key: 'tax', label: '消費税', width: 96, type: 'yen', calc: true },
    { key: 'total', label: '税込合計', width: 116, type: 'yen', calc: true },
    { key: 'quoteNo', label: '見積番号', width: 128 },
    { key: 'quoteOn', label: '見積日', width: 96, type: 'date' },
    { key: 'quotePdf', label: '見積書PDF', width: 96 },
    { key: 'invoiceNo', label: '請求番号', width: 136 },
    { key: 'invoiceOn', label: '請求日', width: 96, type: 'date' },
    { key: 'dueOn', label: '支払期限', width: 96, type: 'date' },
    { key: 'invoicePdf', label: '請求書PDF', width: 96 },
    { key: 'mailDraft', label: 'メール下書き', width: 150 },
    { key: 'paid', label: '入金済額', width: 110, type: 'yen', calc: true },
    { key: 'balance', label: '未入金残高', width: 110, type: 'yen', calc: true },
    { key: 'paymentStatus', label: '入金ステータス', width: 110, calc: true },
    { key: 'overdueDays', label: '期限まで/超過', width: 104, calc: true },
    { key: 'note', label: '備考', width: 200 },
  ],
  lines: [
    { key: 'dealId', label: '案件番号', width: 120 },
    { key: 'customer', label: '顧客名', width: 200, calc: true },
    { key: 'product', label: '商品・サービス', width: 220 },
    { key: 'detail', label: '明細摘要（任意）', width: 220 },
    { key: 'qty', label: '数量', width: 64, type: 'number' },
    { key: 'unit', label: '単位', width: 56 },
    { key: 'unitPrice', label: '単価', width: 100, type: 'yen' },
    { key: 'taxRate', label: '税率', width: 64, type: 'percent' },
    { key: 'amount', label: '金額（税抜）', width: 116, type: 'yen', calc: true },
    { key: 'note', label: '備考', width: 180 },
  ],
  payments: [
    { key: 'paidOn', label: '入金日', width: 100, type: 'date' },
    { key: 'invoiceNo', label: '請求番号', width: 140 },
    { key: 'customer', label: '顧客名', width: 210, calc: true },
    { key: 'amount', label: '入金額', width: 120, type: 'yen' },
    { key: 'method', label: '入金方法', width: 100 },
    { key: 'note', label: '備考', width: 240 },
  ],
  customers: [
    { key: 'id', label: '顧客ID', width: 72 },
    { key: 'name', label: '会社名', width: 220 },
    { key: 'department', label: '部署', width: 120 },
    { key: 'contact', label: '担当者名', width: 110 },
    { key: 'email', label: 'メール', width: 220 },
    { key: 'phone', label: '電話', width: 120 },
    { key: 'zip', label: '郵便番号', width: 90 },
    { key: 'address', label: '住所', width: 260 },
    { key: 'terms', label: '支払条件', width: 190 },
    { key: 'note', label: '備考', width: 200 },
  ],
  products: [
    { key: 'id', label: '商品ID', width: 72 },
    { key: 'name', label: '商品・サービス名', width: 240 },
    { key: 'category', label: '区分', width: 100 },
    { key: 'unit', label: '単位', width: 56 },
    { key: 'unitPrice', label: '標準単価', width: 110, type: 'yen' },
    { key: 'taxRate', label: '税率', width: 64, type: 'percent' },
    { key: 'description', label: '説明', width: 300 },
    { key: 'active', label: '有効', width: 56, type: 'checkbox' },
  ],
  log: [
    { key: 'at', label: '日時', width: 150, type: 'datetime' },
    { key: 'action', label: '操作', width: 160 },
    { key: 'target', label: '対象', width: 160 },
    { key: 'result', label: '結果', width: 70 },
    { key: 'detail', label: '詳細', width: 520 },
    { key: 'user', label: '実行者', width: 200 },
  ],
};

/** COL.deals.invoiceNo → 13 のように 1 始まりの列番号を引く */
const COL = Object.keys(COLUMNS).reduce(function (acc, sheetKey) {
  acc[sheetKey] = {};
  COLUMNS[sheetKey].forEach(function (c, i) {
    acc[sheetKey][c.key] = i + 1;
  });
  return acc;
}, {});

const ORDER_STATUS = ['見積中', '受注', '失注'];
const PAYMENT_STATUS = ['未請求', '未入金', '一部入金', '期限超過', '入金済', '失注'];
const PAYMENT_METHODS = ['銀行振込', '口座振替', '現金', '相殺', 'その他'];
const TAX_RATES = [0.1, 0.08];

/** 支払条件（顧客マスタのプルダウンと期限計算で共用） */
const PAYMENT_TERMS = [
  { label: '月末締め翌月末払い', kind: 'closing', closingDay: 31, monthsLater: 1, payDay: 31 },
  { label: '月末締め翌々月末払い', kind: 'closing', closingDay: 31, monthsLater: 2, payDay: 31 },
  { label: '20日締め翌月10日払い', kind: 'closing', closingDay: 20, monthsLater: 1, payDay: 10 },
  { label: '15日締め当月末払い', kind: 'closing', closingDay: 15, monthsLater: 0, payDay: 31 },
  { label: '請求日から30日以内', kind: 'net', days: 30 },
  { label: '請求日から14日以内', kind: 'net', days: 14 },
];

/**
 * 設定シートの項目。value は初期値（デモ用の架空企業）。
 * 設定シートは「項目 | 値 | 説明」の3列で、プログラムは項目名で値を引く。
 */
const SETTINGS_SPEC = [
  { section: '自社情報（帳票に印字）' },
  { key: 'companyName', label: '会社名', value: '株式会社ソラノ制作所', help: '架空のデモ企業です。自社名に書き換えてください。' },
  { key: 'companyZip', label: '郵便番号', value: '100-0000' },
  { key: 'companyAddress', label: '住所', value: '東京都千代田区丸の内0-0-0 ソラノビル5F' },
  { key: 'companyTel', label: '電話番号', value: '03-0000-0000' },
  { key: 'companyEmail', label: 'メール', value: 'billing@example.com' },
  { key: 'representative', label: '担当者（帳票・メール署名）', value: '経理担当 佐藤' },
  { key: 'registrationNo', label: '適格請求書発行事業者 登録番号', value: 'T0000000000000', help: 'インボイス制度の登録番号（T＋13桁）。デモ値です。' },
  { key: 'sealText', label: '社印の文字（空欄で非表示）', value: 'ソラノ制作所之印' },
  { section: '振込先（請求書に印字）' },
  { key: 'bankName', label: '銀行名・支店名', value: 'デモ銀行 本店営業部' },
  { key: 'bankAccount', label: '口座種別・番号', value: '普通 0000000' },
  { key: 'bankHolder', label: '口座名義', value: 'カ）ソラノセイサクショ' },
  { section: '帳票・採番' },
  { key: 'quotePrefix', label: '見積番号の接頭辞', value: 'Q' },
  { key: 'invoicePrefix', label: '請求番号の接頭辞', value: 'INV' },
  { key: 'dealPrefix', label: '案件番号の接頭辞', value: 'PJ' },
  { key: 'quoteValidDays', label: '見積の有効期限（日数）', value: 30 },
  { key: 'accentColor', label: '帳票のアクセント色', value: '#1f6feb', help: '#RRGGBB 形式' },
  { section: '保存・メール' },
  { key: 'driveFolderId', label: 'PDF保存先フォルダID', value: '', help: '空欄なら初回のPDF作成時にマイドライブへ自動作成します。' },
  { key: 'mailSubjectQuote', label: '見積メールの件名', value: '【お見積書送付】{件名}（{番号}）' },
  { key: 'mailSubjectInvoice', label: '請求メールの件名', value: '【ご請求書送付】{件名}（{番号}）' },
  { key: 'mailSubjectReminder', label: '督促メールの件名', value: '【ご入金のお願い】{番号} お支払期限経過のご案内' },
  { key: 'mailSignature', label: 'メール署名', value: '――――――――――――――\n{自社名}\n{担当者}\n{電話番号} / {メール}\n――――――――――――――' },
  { section: '未入金アラート' },
  { key: 'alertSoonDays', label: '期限間近として扱う日数', value: 7, help: '支払期限までこの日数以内の未入金も一覧に出します。' },
  { key: 'alertNotifySelf', label: '毎朝のアラートを自分宛にメール通知', value: false, help: 'オンにすると、毎朝の自動チェックで期限超過があるときだけ、このスプレッドシートの所有者本人にだけ通知メールを送ります（顧客には送りません）。' },
];

/** デザイン（業務システム風のフラットな配色） */
const THEME = {
  font: 'Noto Sans JP',
  ink: '#1f2937',
  muted: '#6b7280',
  line: '#e5e7eb',
  header: '#1f2a44',
  headerText: '#ffffff',
  calcHeader: '#3b4a6b',
  band: '#f8fafc',
  accent: '#1f6feb',
  sheetTabs: {
    'ダッシュボード': '#1f6feb',
    '案件・請求': '#0f766e',
    '売上明細': '#0f766e',
    '入金記録': '#0f766e',
    '顧客マスタ': '#6b7280',
    '商品マスタ': '#6b7280',
    '設定': '#6b7280',
    '操作ログ': '#9ca3af',
    '使い方': '#9ca3af',
  },
  status: {
    '期限超過': { bg: '#fde8e8', fg: '#b42318' },
    '未入金': { bg: '#fff4e5', fg: '#b54708' },
    '一部入金': { bg: '#e8f0fe', fg: '#1d4ed8' },
    '入金済': { bg: '#e7f6ec', fg: '#067647' },
    '未請求': { bg: '#f3f4f6', fg: '#4b5563' },
    '失注': { bg: '#f3f4f6', fg: '#9ca3af' },
    '見積中': { bg: '#eef2ff', fg: '#4338ca' },
    '受注': { bg: '#e7f6ec', fg: '#067647' },
  },
  chart: {
    sales: '#2a78d6',
    paid: '#1baf7a',
  },
};
