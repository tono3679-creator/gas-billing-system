/**
 * Apps Script のサービスを Node 上で模擬する最小限のモック。
 * データの読み書き（値・数式・行の挿入）は本物に近く動かし、
 * 書式や見た目だけの呼び出しは何もしない「連鎖可能な空オブジェクト」で受ける。
 * src/*.js を 1 つの vm コンテキストに読み込み、Apps Script と同じく全関数をグローバルに置く。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

/** どんなメソッド呼び出しも受けて自分を返す（書式設定などの no-op 用） */
function chain(overrides) {
  const target = Object.assign({}, overrides || {});
  const p = new Proxy(target, {
    get(t, prop) {
      if (prop in t) return t[prop];
      if (prop === 'then' || typeof prop === 'symbol') return undefined;
      return function () { return p; };
    },
  });
  return p;
}

function colToNum(s) {
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    if (row < 1 || col < 1 || numRows < 1 || numCols < 1) throw new Error(`invalid range ${row},${col},${numRows},${numCols}`);
    if (row + numRows - 1 > sheet.maxRows || col + numCols - 1 > sheet.maxCols) {
      throw new Error(`Range out of bounds on ${sheet.name}: r${row}+${numRows} c${col}+${numCols} (max ${sheet.maxRows}x${sheet.maxCols})`);
    }
    this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols;
    const proxy = new Proxy(this, {
      get(t, prop) {
        if (prop in t) return t[prop];
        if (typeof prop === 'symbol' || prop === 'then') return undefined;
        return function () { return proxy; };
      },
    });
    return proxy;
  }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getSheet() { return this.sheet; }
  getA1Notation() { return `R${this.row}C${this.col}:R${this.row + this.numRows - 1}C${this.col + this.numCols - 1}`; }
  _each(fn) {
    for (let i = 0; i < this.numRows; i++) for (let j = 0; j < this.numCols; j++) fn(this.row + i, this.col + j, i, j);
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.numRows; i++) {
      const r = [];
      for (let j = 0; j < this.numCols; j++) r.push(this.sheet._get(this.row + i, this.col + j));
      out.push(r);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map((r) => r.map((v) => (v instanceof Date ? v.toISOString() : String(v)))); }
  getValue() { return this.sheet._get(this.row, this.col); }
  getFormula() { return this.sheet.formulas.get(`${this.row},${this.col}`) || ''; }
  getFormulas() {
    const out = [];
    for (let i = 0; i < this.numRows; i++) {
      const r = [];
      for (let j = 0; j < this.numCols; j++) r.push(this.sheet.formulas.get(`${this.row + i},${this.col + j}`) || '');
      out.push(r);
    }
    return out;
  }
  setValues(values) {
    if (values.length !== this.numRows || values.some((r) => r.length !== this.numCols)) {
      throw new Error(`setValues size mismatch on ${this.sheet.name}: range ${this.numRows}x${this.numCols}, data ${values.length}x${values[0] && values[0].length}`);
    }
    this._each((r, c, i, j) => this.sheet._set(r, c, values[i][j]));
    return this;
  }
  setValue(v) { this._each((r, c) => this.sheet._set(r, c, v)); return this; }
  setFormula(f) { this._each((r, c) => this.sheet._setFormula(r, c, f)); return this; }
  setFormulas(fs2) { this._each((r, c, i, j) => this.sheet._setFormula(r, c, fs2[i][j])); return this; }
  clearContent() { this._each((r, c) => this.sheet._set(r, c, '')); return this; }
  insertCheckboxes() { this._each((r, c) => this.sheet._set(r, c, false)); return this; }
  setDataValidation(rule) { this.sheet.validations.push({ range: this.getA1Notation(), rule }); return this; }
  protect() { const p = chain({ remove() {} }); this.sheet.protections.push(p); return p; }
  applyRowBanding() { const b = chain({ remove: () => { this.sheet.bandings = []; } }); this.sheet.bandings.push(b); return b; }
  merge() { this.sheet.merges.push(this.getA1Notation()); return this; }
}

class FakeSheet {
  constructor(ss, name) {
    this.ss = ss; this.name = name;
    this.maxRows = 1000; this.maxCols = 26;
    this.data = new Map(); // "r,c" -> value
    this.formulas = new Map();
    this.charts = []; this.bandings = []; this.protections = []; this.validations = []; this.merges = [];
    this.cfRules = [];
    this.activeRange = null;
  }
  _get(r, c) { const v = this.data.get(`${r},${c}`); return v === undefined ? '' : v; }
  _set(r, c, v) {
    this.formulas.delete(`${r},${c}`);
    if (v === '' || v === null || v === undefined) this.data.delete(`${r},${c}`);
    else this.data.set(`${r},${c}`, v);
  }
  _setFormula(r, c, f) {
    if (!f) { this._set(r, c, ''); return; }
    this.data.delete(`${r},${c}`);
    this.formulas.set(`${r},${c}`, f);
    // 見出しの配列数式は見出しの文字列を値として持つ（本物も 1 行目に見出しが表示される）
    const m = String(f).match(/^=\{"([^"]+)";/);
    if (m) this.data.set(`${r},${c}`, m[1]);
  }
  getName() { return this.name; }
  getParent() { return this.ss; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/);
      if (!m) throw new Error('bad a1 ' + a);
      const c1 = colToNum(m[1]);
      const r1 = m[2] ? Number(m[2]) : 1;
      const c2 = m[3] ? colToNum(m[3]) : c1;
      const r2 = m[3] ? (m[4] ? Number(m[4]) : this.maxRows) : r1;
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  getLastRow() {
    let max = 0;
    for (const k of [...this.data.keys(), ...this.formulas.keys()]) max = Math.max(max, Number(k.split(',')[0]));
    return max;
  }
  getLastColumn() {
    let max = 0;
    for (const k of [...this.data.keys(), ...this.formulas.keys()]) max = Math.max(max, Number(k.split(',')[1]));
    return max;
  }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  _shiftRows(from, delta) {
    for (const map of [this.data, this.formulas]) {
      const entries = [...map.entries()];
      map.clear();
      for (const [k, v] of entries) {
        const [r, c] = k.split(',').map(Number);
        if (r >= from) {
          if (delta < 0 && r < from - delta) continue; // 削除された行
          map.set(`${r + delta},${c}`, v);
        } else map.set(k, v);
      }
    }
  }
  insertRowAfter(r) { this._shiftRows(r + 1, 1); this.maxRows += 1; return this; }
  insertRowsAfter(r, n) { this._shiftRows(r + 1, n); this.maxRows += n; return this; }
  deleteRows(r, n) { this._shiftRows(r, -n); this.maxRows -= n; return this; }
  insertColumnsAfter(c, n) { this.maxCols += n; return this; }
  deleteColumns(c, n) {
    for (const map of [this.data, this.formulas]) for (const k of [...map.keys()]) if (Number(k.split(',')[1]) >= c) map.delete(k);
    this.maxCols -= n; return this;
  }
  clear() { this.data.clear(); this.formulas.clear(); this.merges = []; return this; }
  getCharts() { return this.charts.slice(); }
  removeChart(c) { this.charts = this.charts.filter((x) => x !== c); }
  newChart() {
    const opts = {};
    const ranges = [];
    const b = chain({
      addRange: (r) => { ranges.push(r.getA1Notation()); return b; },
      setOption: (k, v) => { opts[k] = v; return b; },
      build: () => ({ opts, ranges }),
    });
    return b;
  }
  insertChart(c) { this.charts.push(c); }
  getBandings() { return this.bandings.slice(); }
  getProtections() { return this.protections.slice(); }
  setConditionalFormatRules(rules) { this.cfRules = rules; }
  getConditionalFormatRules() { return this.cfRules; }
  clearConditionalFormatRules() { this.cfRules = []; }
  setActiveRange(r) { this.activeRange = r; this.ss.active = this; return r; }
  getActiveRange() { return this.activeRange; }
}
// 書式系は何もしない
['setFrozenRows', 'setFrozenColumns', 'setTabColor', 'setHiddenGridlines', 'setColumnWidth', 'setColumnWidths', 'setRowHeight', 'setRowHeights', 'hideSheet', 'showSheet', 'activate'].forEach((m) => {
  FakeSheet.prototype[m] = function () { return this; };
});

class FakeSpreadsheet {
  constructor() {
    this.sheets = [new FakeSheet(this, 'シート1')];
    this.active = this.sheets[0];
    this.toasts = [];
  }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(name, idx) {
    if (this.getSheetByName(name)) throw new Error('duplicate sheet ' + name);
    const s = new FakeSheet(this, name);
    this.sheets.splice(idx === undefined ? this.sheets.length : idx, 0, s);
    return s;
  }
  deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
  setActiveSheet(s) { this.active = s; return s; }
  getActiveSheet() { return this.active; }
  getActiveRange() { return this.active.activeRange; }
  moveActiveSheet(pos) {
    this.sheets = this.sheets.filter((x) => x !== this.active);
    this.sheets.splice(pos - 1, 0, this.active);
  }
  toast(msg, title) { this.toasts.push({ msg, title }); }
  getName() { return 'デモ_見積請求管理'; }
  getId() { return 'ss-1'; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/ss-1/edit'; }
  setSpreadsheetTimeZone() {}
  setSpreadsheetLocale() {}
}

function makeIterator(list) {
  let i = 0;
  return { hasNext: () => i < list.length, next: () => list[i++] };
}

class FakeDrive {
  constructor() { this.folders = new Map(); this.files = new Map(); this.seq = 0; }
  _id(prefix) { this.seq += 1; return `${prefix}${String(this.seq).padStart(24, '0')}`; }
  _folder(name, parent) {
    const id = this._id('fold');
    const drive = this;
    const f = {
      id, name, parent, trashed: false,
      getId: () => id, getName: () => name, getUrl: () => `https://drive.google.com/drive/folders/${id}`, isTrashed: () => f.trashed,
      createFolder: (n) => drive._folder(n, f),
      getFoldersByName: (n) => makeIterator([...drive.folders.values()].filter((x) => x.parent === f && x.name === n && !x.trashed)),
      getFilesByName: (n) => makeIterator([...drive.files.values()].filter((x) => x.parent === f && x.name === n && !x.trashed)),
      createFile: (blob) => drive._file(blob, f),
    };
    this.folders.set(id, f);
    return f;
  }
  _file(blob, parent) {
    const id = this._id('file');
    const f = {
      id, name: blob.getName(), parent, blob, trashed: false,
      getId: () => id, getName: () => f.name, getUrl: () => `https://drive.google.com/file/d/${id}/view`,
      setTrashed: (v) => { f.trashed = v; return f; }, isTrashed: () => f.trashed, getBlob: () => blob,
    };
    this.files.set(id, f);
    return f;
  }
  service() {
    return {
      createFolder: (n) => this._folder(n, null),
      getFolderById: (id) => { const f = this.folders.get(id); if (!f) throw new Error('not found'); return f; },
      getFileById: (id) => { const f = this.files.get(id); if (!f) throw new Error('not found'); return f; },
    };
  }
}

function makeBlob(content, type, name) {
  let n = name;
  const blob = {
    content, type,
    getName: () => n,
    setName: (x) => { n = x; return blob; },
    getContentType: () => type,
    getBytes: () => Buffer.from(String(content)),
    getDataAsString: () => String(content),
    getAs: (mime) => makeBlob(`%PDF-mock\n${content}`, mime, n),
  };
  return blob;
}

function formatDate(date, tz, pattern) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date)
    .reduce((acc, p) => { acc[p.type] = p.value; return acc; }, {});
  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('HH', parts.hour)
    .replace('mm', parts.minute)
    .replace('ss', parts.second)
    .replace(/\bM\b/, String(Number(parts.month)))
    .replace(/\bd\b/, String(Number(parts.day)));
}

/**
 * src を読み込んだコンテキストを作る。
 * @param {{alertAnswers?: string[]}} opt  ui.alert が返すボタン（既定は OK/YES）
 */
function loadGas(opt) {
  opt = opt || {};
  const ss = new FakeSpreadsheet();
  const drive = new FakeDrive();
  const gmail = { drafts: [], sent: [] };
  const triggers = [];
  const ui = {
    alerts: [], dialogs: [], sidebars: [],
    Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' },
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
  };
  const answers = (opt.alertAnswers || []).slice();
  ui.alert = (title, msg) => { ui.alerts.push({ title, msg }); return answers.length ? answers.shift() : (msg === undefined ? 'OK' : 'OK'); };
  ui.createMenu = () => { const m = chain({ addToUi: () => { ui.menuAdded = true; } }); return m; };
  ui.showModelessDialog = (html, title) => ui.dialogs.push({ title, html: html.content });
  ui.showSidebar = (html) => ui.sidebars.push(html);

  const sandbox = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ui,
      flush: () => {},
      newDataValidation: () => { const r = chain({ build: () => ({ kind: 'validation' }) }); return r; },
      newConditionalFormatRule: () => { const r = chain({ build: () => ({ kind: 'cf' }) }); return r; },
      BandingTheme: { LIGHT_GREY: 'LIGHT_GREY' },
      ProtectionType: { RANGE: 'RANGE' },
      BorderStyle: { SOLID: 'SOLID', SOLID_THICK: 'SOLID_THICK' },
    },
    Charts: { ChartType: { COLUMN: 'COLUMN', BAR: 'BAR', LINE: 'LINE' } },
    Utilities: {
      formatDate,
      newBlob: (content, type, name) => makeBlob(content, type, name),
    },
    MimeType: { PDF: 'application/pdf' },
    DriveApp: drive.service(),
    GmailApp: {
      createDraft: (to, subject, body, options) => { const d = { to, subject, body, options, getId: () => 'draft-' + gmail.drafts.length }; gmail.drafts.push(d); return d; },
      sendEmail: (to, subject, body) => { gmail.sent.push({ to, subject, body }); },
    },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: {
      getActiveUser: () => ({ getEmail: () => 'owner@example.com' }),
      getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }),
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
      newTrigger: (fn) => {
        const b = chain({ create: () => { const t = { getHandlerFunction: () => fn }; triggers.push(t); return t; } });
        return b;
      },
    },
    HtmlService: {
      createHtmlOutput: (content) => chain({ content, getContent: () => content }),
      createHtmlOutputFromFile: (name) => chain({ content: fs.readFileSync(path.join(__dirname, '..', 'src', name + '.html'), 'utf8') }),
    },
  };
  vm.createContext(sandbox);
  const dir = path.join(__dirname, '..', 'src');
  const code = fs.readdirSync(dir).filter((f) => f.endsWith('.js')).sort().map((f) => `// ---- ${f}\n` + fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  // const/let の宣言もテストから参照できるよう、最後にまとめて公開する
  const exportNames = [...code.matchAll(/^(?:const|let|function)\s+([A-Za-z_$][\w$]*)/gm)].map((m) => m[1]);
  vm.runInContext(code + `\n;globalThis.__exports = { ${exportNames.join(', ')} };`, sandbox, { filename: 'gas-bundle.js' });
  const g = Object.assign({}, sandbox.__exports);
  return { g, ss, drive, gmail, triggers, ui, sandbox };
}

module.exports = { loadGas, FakeSheet };
