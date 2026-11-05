"use strict";
// 台帳との照合ルール（管理番号の検算・登録済みチェック・住所補完・備考の連番・読取結果の確認用項目化）

const ITEMS = ["エアコン", "テレビ", "冷蔵庫", "洗濯機", "乾燥機"];
const NUMBER_RE = /^\d{4}-\d{8}-\d$/;

// 製造業者マスタの初期値。台帳の表記揺れを 1 コード 1 表記にまとめたもの（コード, 名称, グループ）
const MAKERS_SEED = [
  [100, "パナソニック", "A"], [101, "三洋電機", "A"], [110, "東芝ライフスタイル", "A"], [114, "TVS REGZA", "A"],
  [120, "ダイキン工業", "A"], [130, "JVCケンウッド", "A"], [151, "ハイセンスジャパン", "A"], [180, "LGエレクトロニクス", "A"],
  [185, "コロナ", "A"], [191, "サムスン電子ジャパン", "A"], [199, "ヤマダHD", "A"],
  [300, "日立グローバルライフ", "B"], [303, "日立ジョンソンコントロールズ空調", "B"], [310, "シャープ", "B"],
  [320, "三菱電機", "B"], [321, "三菱電機エンジニア", "B"], [340, "ソニー", "B"], [350, "富士通ゼネラル", "B"],
  [360, "ソニー（アイワ）", "B"], [382, "トヨトミ", "B"], [387, "良品計画", "B"], [392, "ハイアールジャパン", "B"],
  [393, "アクア", "B"], [394, "吉井電気", "B"], [395, "アイリスオーヤマ", "B"], [511, "ツインバード", "B"],
  [524, "三ツ星貿易", "B"], [532, "正和", "B"], [535, "日本ゼネラル", "B"], [543, "アズマ", "B"], [550, "カンデラ", "B"],
  [559, "Freedom", "B"], [566, "ユニテク", "B"], [570, "ピクセラ", "B"], [573, "エスケイジャパン", "B"],
  [582, "イー・エム・エー", "B"], [586, "テクタイト", "B"], [596, "ジーマックス", "B"], [701, "マクスゼン", "B"],
  [702, "アントビー", "B"], [704, "山善", "B"], [706, "ベルソス", "B"], [708, "ノジマ", "B"], [713, "ドウシシャ", "B"],
  [717, "WIS", "B"], [718, "アイリスオーヤマ(718)", "B"], [719, "A-stage", "B"], [721, "ティーズネットワーク", "B"],
  [727, "TCLジャパン", "B"], [732, "ベステックグループ", "B"], [735, "ニトリ", "B"], [739, "ジョワイユ", "B"],
  [752, "グリーンハウス", "B"], [758, "都光", "B"], [797, "ワールプールジャパン", "B"], [999, "指定法人", "B"],
].map(([code, name, group]) => ({ code: String(code), name, group }));

const digits = (v) => String(v ?? "").replace(/\D/g, "");

/** 形式と検算桁（先頭12桁÷7の余り＝末尾1桁）を確認する。 */
function numberOk(number) {
  const d = digits(number);
  return NUMBER_RE.test(number || "") && Number(BigInt(d.slice(0, 12)) % 7n) === Number(d[12]);
}

const telKey = (v) => { const d = digits(v); return d.startsWith("0558") ? d.slice(4) : d; };

/** 2026-10-30 / 2026/10/30 / 10/30 を "YYYY-MM-DD" へ。解釈できなければ ""。 */
function parseDate(text, today = new Date()) {
  const t = String(text || "").trim().replace(/\//g, "-");
  let y, m, d, x;
  if ((x = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t))) [y, m, d] = [+x[1], +x[2], +x[3]];
  else if ((x = /^(\d{1,2})-(\d{1,2})$/.exec(t))) {
    [y, m, d] = [today.getFullYear(), +x[1], +x[2]];
    if (new Date(y, m - 1, d) > today) y -= 1;
  } else return "";
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return "";
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

const findNumber = (rows, number) => { const d = digits(number); return d ? rows.filter((r) => digits(r.number) === d) : []; };

/** 同じ電話番号→同じ名前の順に、最も新しい行の住所を返す。 */
function findAddress(rows, tel, name) {
  // 名前は全角/半角の（株）等の違いを吸収して比べる
  const norm = (s) => String(s || "").normalize("NFKC").replace(/\s/g, "");
  const key = telKey(tel), n = norm(name);
  const tests = [(r) => key && telKey(r.tel) === key, (r) => n && norm(r.name) === n];
  for (const test of tests) {
    const hits = rows.filter((r) => r.addr && test(r));
    if (hits.length) return hits.reduce((a, b) => ((/^\d/.test(b.date) ? b.date : "") >= (/^\d/.test(a.date) ? a.date : "") ? b : a)).addr;
  }
  return "";
}

/** シートの備考の最大連番＋1。接頭辞は最終行に合わせる（洗濯機Bは全角Ｂ）。 */
function nextNote(rows, sheet) {
  const mine = rows.filter((r) => r.sheet === sheet);
  const nums = mine.map((r) => /(\d+)$/.exec(r.note)).filter(Boolean).map((m) => +m[1]);
  let prefix = sheet.slice(-1) + "-";
  for (let i = mine.length - 1; i >= 0; i--) { const m = /^(.*?[-‐－])\d+$/.exec(mine[i].note); if (m) { prefix = m[1]; break; } }
  return prefix + ((nums.length ? Math.max(...nums) : 0) + 1);
}

function matchMaker(makers, name, code) {
  const c = digits(code);
  return makers.find((m) => m.code === c) || makers.find((m) => name && (m.name.includes(name) || name.includes(m.name))) || null;
}

/** 読取結果1枚を、台帳と突き合わせた確認用の項目にする。 */
function buildItem(file, r, rows, makers) {
  const number = String(r.number || "").trim();
  const kind = r.kind === "purple" ? "purple" : "green";
  const mk = matchMaker(makers, r.maker || "", r.code || "");
  const item = ITEMS.includes(r.item) ? r.item : "";
  const f = { date: r.date || "", name: String(r.name || "").trim(), tel: String(r.tel || "").trim(), addr: "",
    maker: mk ? mk.name : "", code: mk ? mk.code : digits(r.code), number, note: String(r.note || "").trim(),
    pickup: r.pickup || "", item, sheet: item && mk ? item + mk.group : "" };
  f.addr = findAddress(rows, f.tel, f.name);
  const it = { id: Math.random().toString(36).slice(2, 12), file, kind, action: "add", fields: f, excluded: false,
    created: new Date().toISOString() };
  const msgs = [];
  if (r.error) msgs.push("読み取りに失敗しました。写真を見て入力してください。");
  else if (!numberOk(number)) msgs.push("管理番号の検算が合いません。写真と見比べてください。");
  if (!mk && !r.error) msgs.push(`製造業者「${r.maker || "不明"}」が未登録です。選択するか新規登録してください。`);
  const hits = findNumber(rows, number);
  if (kind === "green") {
    if (hits.length) msgs.push(`この番号は ${hits[0].sheet} ${hits[0].row}行に登録済みです。除外してください。`);
  } else {
    if (!f.pickup) msgs.push("受領印がありません。印が戻ってから送るか、引取日を入力してください。");
    if (hits.length) {
      const h = hits[0];
      Object.assign(it, { action: "pickup", sheet: h.sheet, row: h.row, orig_number: h.number });
      Object.assign(f, { addr: h.addr, sheet: h.sheet, note: h.note, name: h.name, tel: h.tel, item: h.sheet.slice(0, -1) });
      if (h.pickup && h.pickup !== f.pickup) {
        msgs.push(`引取日がすでに ${h.pickup} で入っています。上書きする場合はチェックしてください。`);
        it.ask_overwrite = true;
      } else if (h.pickup) msgs.push("同じ引取日が入力済みです。除外してください。");
    } else {
      const near = rows.filter((x) => f.tel && telKey(x.tel) === telKey(f.tel) && x.date === f.date && x.code === f.code);
      if (near.length) {
        it.candidate = { sheet: near[0].sheet, row: near[0].row, number: near[0].number };
        msgs.push(`台帳に番号がありません。${near[0].sheet} ${near[0].row}行（${near[0].number}）の番号違いの可能性があります。`);
      } else msgs.push("台帳に番号がありません（緑が未登録）。行を追加して引取日も入れます。");
    }
  }
  if (it.action === "add" && !f.note && f.sheet) f.note = nextNote(rows, f.sheet);
  if (it.action === "add" && !f.addr) msgs.push("住所は伝票にないため入力してください。");
  it.msgs = msgs;
  return it;
}

/** 保存前の入力チェック。問題があればその理由、なければ ""。 */
function validate(it) {
  const f = it.fields;
  if (!numberOk(f.number)) return "管理番号の形式または検算が合いません";
  if (it.action === "pickup") return parseDate(f.pickup) ? "" : "引取日を入力してください";
  const need = [["date", "交付日"], ["name", "排出者"], ["addr", "住所"], ["sheet", "品目"], ["code", "製造業者"]];
  for (const [k, label] of need) if (!f[k] || (k === "date" && !parseDate(f[k]))) return `${label}を入力してください`;
  return "";
}
