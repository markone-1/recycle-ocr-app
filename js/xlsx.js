"use strict";
// 台帳 xlsx の読み書き。ファイル全体を作り直さず、変更する行のセルとテーブル範囲だけを XML 上で書き換える
// （書式・印刷設定・プリンタ設定などはそのまま残る）。JSZip を使う。
// 列: A交付日 B排出者 C住所 D電話 E製造業者 Fコード G管理番号 H備考 I引取日（3行目が見出し、4行目から）

const XNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const FIRST_ROW = 4;
const COL = { date: 1, name: 2, addr: 3, tel: 4, maker: 5, code: 6, number: 7, note: 8, pickup: 9 };
const DATE_COLS = [1, 9];
const EPOCH = Date.UTC(1899, 11, 30);

const colLetter = (n) => String.fromCharCode(64 + n);
const colIndex = (ref) => ref.replace(/\d/g, "").split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const serialToIso = (n) => new Date(EPOCH + Math.round(n) * 86400000).toISOString().slice(0, 10);
const isoToSerial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - EPOCH) / 86400000; };
const resolve = (base, target) => {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const p of target.split("/")) p === ".." ? parts.pop() : parts.push(p);
  return parts.join("/");
};
const parseXml = (s) => new DOMParser().parseFromString(s, "application/xml");
const kids = (el, name) => Array.from(el.children).filter((c) => c.localName === name);

class Conflict extends Error {}

class Ledger {
  static async open(buffer) {
    const lg = new Ledger();
    lg.zip = await JSZip.loadAsync(buffer);
    const wb = parseXml(await lg.zip.file("xl/workbook.xml").async("string"));
    const rels = parseXml(await lg.zip.file("xl/_rels/workbook.xml.rels").async("string"));
    const target = {};
    for (const r of rels.getElementsByTagName("Relationship")) target[r.getAttribute("Id")] = r.getAttribute("Target");
    const ssFile = lg.zip.file("xl/sharedStrings.xml");
    lg.strings = [];
    if (ssFile) {
      const ss = parseXml(await ssFile.async("string"));
      for (const si of ss.getElementsByTagNameNS(XNS, "si")) {
        // ふりがな（rPh）は除き、本文の t だけをつなぐ
        lg.strings.push(Array.from(si.getElementsByTagNameNS(XNS, "t")).filter((t) => t.parentNode.localName !== "rPh").map((t) => t.textContent).join(""));
      }
    }
    lg.sheets = {};
    lg.dirty = new Set();
    for (const s of wb.getElementsByTagNameNS(XNS, "sheet")) {
      const name = s.getAttribute("name");
      if (!(ITEMS.includes(name.slice(0, -1)) && /[AB]$/.test(name))) continue;
      const rid = s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      const path = resolve("xl/workbook.xml", target[rid]);
      lg.sheets[name] = { path, doc: parseXml(await lg.zip.file(path).async("string")), tables: [] };
      const relPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
      const rf = lg.zip.file(relPath);
      if (rf) {
        for (const r of parseXml(await rf.async("string")).getElementsByTagName("Relationship")) {
          if (r.getAttribute("Type").endsWith("/table")) {
            const tp = resolve(path, r.getAttribute("Target"));
            lg.sheets[name].tables.push({ path: tp, doc: parseXml(await lg.zip.file(tp).async("string")) });
          }
        }
      }
    }
    lg.rows = lg.readRows();
    return lg;
  }

  cellValue(c) {
    const t = c.getAttribute("t");
    const v = kids(c, "v")[0];
    if (t === "s") return v ? this.strings[+v.textContent] : null;
    if (t === "inlineStr") return Array.from(c.getElementsByTagNameNS(XNS, "t")).map((x) => x.textContent).join("");
    if (!v) return null;
    if (t === "str" || t === "e") return v.textContent;
    if (t === "b") return v.textContent === "1" ? "TRUE" : "FALSE";
    return Number(v.textContent);
  }

  readRows() {
    const rows = [];
    for (const [name, sh] of Object.entries(this.sheets)) {
      const data = sh.doc.getElementsByTagNameNS(XNS, "sheetData")[0];
      for (const row of kids(data, "row")) {
        const r = +row.getAttribute("r");
        if (r < FIRST_ROW) continue;
        const vals = {};
        for (const c of kids(row, "c")) {
          const ci = colIndex(c.getAttribute("r"));
          if (ci <= 9) vals[ci] = this.cellValue(c);
        }
        if (vals[COL.name] == null && vals[COL.number] == null) continue;
        const text = (ci) => {
          const v = vals[ci];
          if (v == null) return "";
          if (typeof v === "number") return DATE_COLS.includes(ci) ? serialToIso(v) : String(v);
          return String(v).trim();
        };
        rows.push({ sheet: name, row: r, date: text(1), name: text(2), addr: text(3), tel: text(4), maker: text(5),
          code: digits(text(6)), number: text(7), note: text(8), pickup: text(9) });
      }
    }
    const seen = {};
    for (const x of rows) (seen[digits(x.number)] ||= []).push(x);
    for (const x of rows) {
      const d = digits(x.number);
      x.issue = d && seen[d].length > 1 ? "管理番号が重複" : x.number && !numberOk(x.number) ? "管理番号の形式または検算が合いません" : "";
    }
    return rows;
  }

  rowEl(sheet, r, create) {
    const data = this.sheets[sheet].doc.getElementsByTagNameNS(XNS, "sheetData")[0];
    let before = null;
    for (const row of kids(data, "row")) {
      const n = +row.getAttribute("r");
      if (n === r) return row;
      if (n > r) { before = row; break; }
    }
    if (!create) return null;
    const el = data.ownerDocument.createElementNS(XNS, "row");
    el.setAttribute("r", String(r));
    data.insertBefore(el, before);
    return el;
  }

  cellEl(row, r, ci) {
    const ref = colLetter(ci) + r;
    let before = null;
    for (const c of kids(row, "c")) {
      const n = colIndex(c.getAttribute("r"));
      if (n === ci) return c;
      if (n > ci) { before = c; break; }
    }
    const el = row.ownerDocument.createElementNS(XNS, "c");
    el.setAttribute("r", ref);
    row.insertBefore(el, before);
    return el;
  }

  /** セルに値を入れる。value: null=空 / number / string（インライン文字列）。書式（s）は既存のものを使い、無ければ上の行から写す。 */
  setCell(sheet, r, ci, value) {
    const row = this.rowEl(sheet, r, true);
    const c = this.cellEl(row, r, ci);
    if (!c.hasAttribute("s")) {
      const above = this.rowEl(sheet, r - 1, false);
      const src = above && kids(above, "c").find((x) => colIndex(x.getAttribute("r")) === ci);
      if (src && src.hasAttribute("s")) c.setAttribute("s", src.getAttribute("s"));
    }
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    const doc = c.ownerDocument;
    if (typeof value === "number") {
      const v = doc.createElementNS(XNS, "v"); v.textContent = String(value); c.appendChild(v);
    } else if (value != null && value !== "") {
      c.setAttribute("t", "inlineStr");
      const is = doc.createElementNS(XNS, "is"), t = doc.createElementNS(XNS, "t");
      t.textContent = String(value);
      if (/^\s|\s$/.test(t.textContent)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      is.appendChild(t); c.appendChild(is);
    }
    this.dirty.add(this.sheets[sheet].path);
  }

  setDate(sheet, r, ci, text) {
    const iso = parseDate(text);
    this.setCell(sheet, r, ci, iso ? isoToSerial(iso) : (text || null));
  }

  lastRow(sheet) {
    return Math.max(FIRST_ROW - 1, ...this.rows.filter((x) => x.sheet === sheet).map((x) => x.row));
  }

  /** 新しい行が足りなければ、行の高さ等を上の行から写し、テーブルと dimension の範囲を広げる。 */
  prepareRow(sheet, r) {
    const sh = this.sheets[sheet];
    const exists = this.rowEl(sheet, r, false);
    const row = this.rowEl(sheet, r, true);
    if (!exists) {
      const above = this.rowEl(sheet, r - 1, false);
      if (above) for (const a of above.attributes) if (a.name !== "r") row.setAttributeNS(a.namespaceURI, a.name, a.value);
    }
    for (const t of sh.tables) {
      const tbl = t.doc.documentElement;
      const [a, b] = tbl.getAttribute("ref").split(":");
      if (r > +b.replace(/\D/g, "")) {
        const ref = `${a}:${b.replace(/\d/g, "")}${r}`;
        tbl.setAttribute("ref", ref);
        const af = kids(tbl, "autoFilter")[0];
        if (af) af.setAttribute("ref", ref);
        this.dirty.add(t.path);
      }
    }
    const dim = sh.doc.getElementsByTagNameNS(XNS, "dimension")[0];
    if (dim) {
      const [a, b] = dim.getAttribute("ref").split(":");
      if (b && r > +b.replace(/\D/g, "")) dim.setAttribute("ref", `${a}:${b.replace(/\d/g, "")}${r}`);
    }
  }

  codeValue(sheet, refRow, code) {
    // コード列はシートにより数値/文字列が混在するので、基準の行の型に合わせる
    const row = this.rowEl(sheet, refRow, false);
    const c = row && kids(row, "c").find((x) => colIndex(x.getAttribute("r")) === COL.code);
    const isNum = c && kids(c, "v")[0] && !c.getAttribute("t");
    return isNum ? Number(code) : String(code);
  }

  applyOne(it) {
    const f = it.fields;
    const target = digits(f.number);
    const index = {};
    for (const x of this.rows) { const d = digits(x.number); if (d) (index[d] ||= []).push(x); }
    if (it.action === "add") {
      if (index[target]) throw new Conflict(`管理番号が ${index[target][0].sheet} ${index[target][0].row}行に登録済みです`);
      if (!this.sheets[f.sheet]) throw new Conflict(`シート ${f.sheet} がありません`);
      if (f.note && this.rows.some((x) => x.sheet === f.sheet && x.note === f.note)) throw new Conflict(`備考 ${f.note} は ${f.sheet} で使用済みです`);
      const last = this.lastRow(f.sheet), r = last + 1;
      this.prepareRow(f.sheet, r);
      this.setDate(f.sheet, r, COL.date, f.date);
      for (const k of ["name", "addr", "tel", "maker", "number", "note"]) this.setCell(f.sheet, r, COL[k], f[k] || null);
      this.setCell(f.sheet, r, COL.code, this.codeValue(f.sheet, last, f.code));
      if (f.pickup) this.setDate(f.sheet, r, COL.pickup, f.pickup);
      this.rows.push({ sheet: f.sheet, row: r, date: parseDate(f.date), name: f.name, addr: f.addr, tel: f.tel, maker: f.maker,
        code: digits(f.code), number: f.number, note: f.note, pickup: parseDate(f.pickup), issue: "" });
      return [f.sheet, r];
    }
    // 紫：既存の行に引取日を入れる（番号違いの修正を含む）
    const orig = digits(it.orig_number || f.number);
    let hits = (index[orig] || []).filter((x) => !it.sheet || x.sheet === it.sheet);
    if (it.row) hits = hits.filter((x) => x.row === it.row).length ? hits.filter((x) => x.row === it.row) : hits;
    if (!hits.length) throw new Conflict("対象の行が台帳に見つかりません");
    if (hits.length > 1) throw new Conflict("同じ管理番号の行が複数あります");
    const x = hits[0];
    const nw = parseDate(f.pickup);
    if (x.pickup && x.pickup !== nw && !it.overwrite) throw new Conflict(`引取日がすでに ${x.pickup} で入っています`);
    if (target !== orig && index[target]) throw new Conflict("修正後の管理番号が別の行で使われています");
    this.setDate(x.sheet, x.row, COL.pickup, f.pickup);
    if (target !== orig) { this.setCell(x.sheet, x.row, COL.number, f.number); x.number = f.number; }
    x.pickup = nw;
    return [x.sheet, x.row];
  }

  /** 複数件を反映する。戻り値は [{it, where, err}]。問題のある件は err に理由が入り、書き込まれない。 */
  apply(list) {
    return list.map((it) => {
      try { return { it, where: this.applyOne(it), err: "" }; }
      catch (e) { if (e instanceof Conflict) return { it, where: null, err: e.message }; throw e; }
    });
  }

  async save() {
    const ser = new XMLSerializer();
    const docs = {};
    for (const sh of Object.values(this.sheets)) { docs[sh.path] = sh.doc; for (const t of sh.tables) docs[t.path] = t.doc; }
    for (const path of this.dirty) {
      // ブラウザにより XML 宣言の有無が違うので、取り除いてから Excel と同じ宣言を付け直す
      const body = ser.serializeToString(docs[path]).replace(/^<\?xml[^>]*\?>\s*/, "");
      // createFolders:false … Excel の元ファイルに無いフォルダ項目を足さない
      this.zip.file(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body, { createFolders: false });
    }
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
}
