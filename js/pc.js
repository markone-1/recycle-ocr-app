"use strict";
// PC：読み取り結果を取り込む → 1枚ずつ確認・編集 → 「全データ保存」で台帳 Excel に反映。
// 台帳のあるフォルダを一度選ぶと、ブラウザがそのフォルダの台帳を直接読み書きする（Chrome / Edge）。
// 取り込んだ伝票・写真・製造業者一覧・バックアップは同じフォルダの recycle_data に置く。

const DATA_DIR = "recycle_data";
let folder = null;          // 台帳フォルダ（FsFolder。検査時は差し替える）
let list = [], makers = [], pos = 0, rot = 0, timer = null;
const NEW = "__new";
const FIELDS = ["date", "pickup", "name", "tel", "addr", "number", "note"];

// ── フォルダの記憶（IndexedDB にフォルダの参照だけ保存） ──
const idb = (mode, fn) => new Promise((res, rej) => {
  const o = indexedDB.open("recycle", 1);
  o.onupgradeneeded = () => o.result.createObjectStore("kv");
  o.onsuccess = () => { const tx = o.result.transaction("kv", mode); const r = fn(tx.objectStore("kv")); tx.oncomplete = () => res(r && r.result); tx.onerror = () => rej(tx.error); };
  o.onerror = () => rej(o.error);
});

class FsFolder {
  constructor(dir) { this.dir = dir; }
  async init() {
    const names = [];
    for await (const [n, h] of this.dir.entries()) if (h.kind === "file" && /\.xlsx$/i.test(n) && !n.startsWith("~$")) names.push(n);
    this.ledgerName = names.find((n) => n.includes("台帳")) || names[0];
    if (!this.ledgerName) throw new Error("このフォルダに台帳の Excel がありません");
    this.data = await this.dir.getDirectoryHandle(DATA_DIR, { create: true });
  }
  async readLedger() { return (await (await this.dir.getFileHandle(this.ledgerName)).getFile()).arrayBuffer(); }
  async writeLedger(bytes) {
    const w = await (await this.dir.getFileHandle(this.ledgerName)).createWritable();
    await w.write(bytes);
    await w.close();   // Excel が開いているとここで失敗し、台帳は元のまま
  }
  async excelOpen() {
    try { await this.dir.getFileHandle("~$" + this.ledgerName); return true; } catch { return false; }
  }
  async sub(path, create) {
    const parts = path.split("/"); let d = this.data;
    for (const p of parts.slice(0, -1)) d = await d.getDirectoryHandle(p, { create: true });
    return [d, parts[parts.length - 1]];
  }
  async read(path) {
    try { const [d, n] = await this.sub(path); return await (await d.getFileHandle(n)).getFile(); } catch { return null; }
  }
  async write(path, data) {
    const [d, n] = await this.sub(path);
    const w = await (await d.getFileHandle(n, { create: true })).createWritable();
    await w.write(data); await w.close();
  }
}

const readJson = async (path, def) => { const f = await folder.read(path); return f ? JSON.parse(await f.text()) : def; };
const writeJson = (path, obj) => folder.write(path, JSON.stringify(obj, null, 1));
const saveItems = () => writeJson("items.json", list);

function bar(t, cls) { $("bar").textContent = t; $("bar").className = t ? cls : ""; }

// ── フォルダを開く ──
async function openFolder(pick) {
  if (!window.showDirectoryPicker) return bar("このブラウザでは台帳を扱えません。PC の Chrome か Edge で開いてください。", "e");
  try {
    let dir = await idb("readonly", (s) => s.get("dir"));
    if (pick || !dir) { dir = await window.showDirectoryPicker({ id: "recycle", mode: "readwrite" }); await idb("readwrite", (s) => s.put(dir, "dir")); }
    if ((await dir.queryPermission({ mode: "readwrite" })) !== "granted") {
      // 開いた直後は許可の確認を出せない（ボタン操作が必要）ので、許可ボタンを表示する
      if (!pick || (await dir.requestPermission({ mode: "readwrite" })) !== "granted") return showLocked(dir);
    }
    const f = new FsFolder(dir); await f.init();
    await useFolder(f);
  } catch (e) {
    if (e.name === "AbortError") return;
    if (e.name === "SecurityError" || e.name === "NotAllowedError") return showLocked(null);
    bar(e.message, "e");
  }
}
function showLocked(dir) {
  // ブラウザの決まりで、許可の確認はボタンを押したときにしか出せない
  $("folderName").textContent = dir ? `${dir.name}（許可が必要）` : "未選択";
  $("openFolder").textContent = dir ? "台帳フォルダを許可" : "台帳フォルダを選ぶ";
  $("openFolder").onclick = async () => {
    if (dir && (await dir.requestPermission({ mode: "readwrite" })) === "granted") { const f = new FsFolder(dir); await f.init(); await useFolder(f); }
    else openFolder(true);
  };
  showCard();
}

async function useFolder(f) {
  folder = f;
  $("folderName").textContent = f.ledgerName;
  $("openFolder").textContent = "フォルダを変更";
  $("openFolder").onclick = () => openFolder(true);
  makers = await readJson("makers.json", null);
  if (!makers) { makers = MAKERS_SEED.slice(); await writeJson("makers.json", makers); }
  list = await readJson("items.json", []);
  pos = 0;
  await refreshExcel(); showCard(); remote();
}
async function refreshExcel() {
  if (!folder) return;
  const open = await folder.excelOpen();
  $("xl").textContent = open ? "Excel：開いています" : "Excel：閉じています";
  $("xl").className = "tag " + (open ? "w" : "n");
}

// ── 1枚ずつの表示と編集 ──
function fillMakers(name, code) {
  const sel = $("maker"); sel.innerHTML = "";
  sel.add(new Option("選択してください", ""));
  makers.forEach((m) => sel.add(new Option(`${m.name}（${m.code}）`, m.code + "|" + m.name)));
  sel.add(new Option("＋ 新規登録", NEW));
  if (!code) return;
  const val = code + "|" + name;
  if (![...sel.options].some((o) => o.value === val)) sel.add(new Option(`${name}（${code}）`, val), 1);
  sel.value = val;
}
const groupOf = (code) => (makers.find((m) => m.code === code) || {}).group || "";

async function showCard() {
  const empty = !folder || !list.length;
  $("work").style.display = empty ? "none" : "";
  $("empty").style.display = empty ? "" : "none";
  if (!folder) { $("empty").innerHTML = "最初に「台帳フォルダを選ぶ」で、台帳の Excel があるフォルダを選んでください。"; return; }
  if (!list.length) { $("empty").innerHTML = "確認する伝票はありません。<br>スマホから送信した後、「読み取り結果を取り込む」を押してください。"; return; }
  pos = Math.min(pos, list.length - 1);
  const it = list[pos], f = it.fields, pick = it.action === "pickup";
  $("pos").textContent = `${pos + 1} / ${list.length}`;
  $("kind").innerHTML = `<span class="tag ${it.kind === "green" ? "g" : "p"}">${it.kind === "green" ? "緑" : "紫"}</span>`
    + (it.excluded ? ' <span class="tag n">除外</span>' : "");
  $("title").textContent = pick ? `引取日を入力（${it.sheet} ${it.row}行）` : "新規行を追加";
  const img = await folder.read("images/" + it.file);
  if ($("pic").dataset.url) URL.revokeObjectURL($("pic").dataset.url);
  $("pic").dataset.url = img ? URL.createObjectURL(img) : "";
  $("pic").src = $("pic").dataset.url; $("pic").style.transform = `rotate(${rot}deg)`;
  FIELDS.forEach((k) => { $(k).value = f[k] || ""; $(k).readOnly = pick && !["pickup", "number"].includes(k); });
  $("item").value = f.item || ""; $("item").disabled = pick; $("sheet").value = f.sheet || "";
  fillMakers(f.maker || "", f.code || ""); $("maker").disabled = pick;
  $("np").style.display = "none";
  $("candL").style.display = it.candidate ? "flex" : "none";
  if (it.candidate) { $("candT").textContent = `${it.candidate.sheet} ${it.candidate.row}行の管理番号を直して引取日を入れる（新規追加しない）`; $("cand").checked = !!it.use_candidate; }
  $("overL").style.display = it.ask_overwrite ? "flex" : "none"; $("over").checked = !!it.overwrite;
  $("msgs").innerHTML = (it.msgs || []).map((t) => `<p>${esc(t)}</p>`).join("");
  $("exclude").textContent = it.excluded ? "除外を取り消す" : "この伝票を除外";
  $("form").classList.toggle("ex", !!it.excluded);
}

function collect() {
  const it = list[pos], f = it.fields;
  FIELDS.forEach((k) => (f[k] = $(k).value.trim()));
  f.item = $("item").value; f.sheet = $("sheet").value;
  const v = $("maker").value;
  if (v && v !== NEW) { const [code, ...n] = v.split("|"); f.code = code; f.maker = n.join("|"); }
  it.use_candidate = $("cand").checked; it.overwrite = $("over").checked;
}
async function persist() {
  clearTimeout(timer);
  if (!folder || !list.length) return;
  collect();
  try { await saveItems(); } catch (e) { bar("入力内容を保存できませんでした: " + e.message, "e"); }
}
const later = () => { clearTimeout(timer); timer = setTimeout(persist, 600); };
function updateSheet() {
  const g = groupOf($("maker").value.split("|")[0]);
  if ($("item").value && g) $("sheet").value = $("item").value + g;
}

// ── 取り込み・保存 ──
async function remote() {
  if (!cfg().token) { $("remote").textContent = "「設定」でトークンを登録してください"; $("remote").className = "tag d"; return; }
  try {
    const br = (await gh("/branches?per_page=100")).filter((b) => b.name.startsWith("job/"));
    let done = 0;
    for (const b of br) if (await gh(`/contents/out/result.json?ref=${encodeURIComponent(b.name)}`, { allow404: true })) done++;
    $("remote").textContent = br.length ? `未取込：読み取り済み ${done}件・読み取り中 ${br.length - done}件` : "未取込の送信はありません";
    $("remote").className = "tag " + (done ? "o" : "n");
  } catch (e) { $("remote").textContent = e.message; $("remote").className = "tag d"; }
}

async function importAll() {
  if (!folder) return bar("先に台帳フォルダを選んでください。", "e");
  if (!cfg().token) return bar("先に「設定」でトークンを登録してください。", "e");
  $("import").disabled = true; bar("取り込み中…", "k");
  try {
    await persist();
    const rows = (await Ledger.open(await folder.readLedger())).rows;
    const branches = (await gh("/branches?per_page=100")).filter((b) => b.name.startsWith("job/"));
    let count = 0, waiting = 0; const errors = [];
    for (const b of branches) {
      const tree = (await gh(`/git/trees/${b.commit.sha}?recursive=1`)).tree;
      const res = tree.find((t) => t.path === "out/result.json");
      if (!res) { waiting++; continue; }
      const result = JSON.parse(new TextDecoder().decode(unb64((await gh(`/git/blobs/${res.sha}`)).content)));
      const by = {}; (result.items || []).forEach((r) => (by[r.file] = r));
      const jobId = b.name.split("/")[1];
      const imgs = tree.filter((t) => t.path.startsWith("in/") && t.path.endsWith(".jpg")).sort((a, c) => a.path.localeCompare(c.path));
      const fresh = [];
      for (const t of imgs) {
        const base = t.path.slice(3), name = `${jobId}_${base}`;
        await folder.write("images/" + name, unb64((await gh(`/git/blobs/${t.sha}`)).content));
        fresh.push(buildItem(name, by[base] || { error: true }, rows, makers));
      }
      list.push(...fresh); await saveItems();
      await gh(`/git/refs/heads/${b.name}`, { method: "DELETE" });
      count += fresh.length; errors.push(...(result.errors || []));
    }
    showCard(); remote();
    bar(count ? `${count}枚を取り込みました。${errors.join(" ")}` : `取り込める伝票はありません。${waiting ? "読み取り中の送信があります。数分待ってください。" : ""}`, errors.length ? "e" : "k");
  } catch (e) { bar(e.message, "e"); }
  $("import").disabled = false;
}

async function saveAll() {
  if (!folder || !list.length) return bar("保存する伝票がありません。", "e");
  await persist();
  if (await folder.excelOpen()) { await refreshExcel(); return bar("台帳の Excel が開いています。閉じてからもう一度押してください。", "e"); }
  const n = list.filter((i) => !i.excluded).length;
  if (!confirm(`${n}枚を台帳に反映します。除外した${list.length - n}枚は破棄されます。よろしいですか？`)) return;
  $("saveAll").disabled = true;
  try {
    const original = await folder.readLedger();
    const lg = await Ledger.open(original.slice(0));
    const todo = [], remain = [];
    for (const it of list) {
      if (it.excluded) continue;
      const job = JSON.parse(JSON.stringify(it));
      if (it.use_candidate && it.candidate) Object.assign(job, { action: "pickup", sheet: it.candidate.sheet, row: it.candidate.row, orig_number: it.candidate.number });
      const err = validate(job);
      if (err) { it.msgs = [err]; remain.push(it); } else todo.push([it, job]);
    }
    const results = lg.apply(todo.map(([, j]) => j));
    const done = [];
    results.forEach((r, i) => { const it = todo[i][0]; if (r.err) { it.msgs = [r.err]; remain.push(it); } else done.push({ ...it, saved_to: r.where, saved: new Date().toISOString() }); });
    if (done.length) {
      const d = new Date(), p2 = (n) => String(n).padStart(2, "0");
      const stamp = `${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}_${p2(d.getHours())}${p2(d.getMinutes())}${p2(d.getSeconds())}`;
      await folder.write(`backup/${folder.ledgerName.replace(/\.xlsx$/i, "")}_${stamp}.xlsx`, original);
      try { await folder.writeLedger(await lg.save()); }
      catch { await refreshExcel(); $("saveAll").disabled = false; return bar("台帳に書き込めませんでした。Excel を閉じてからもう一度押してください（台帳は変更していません）。", "e"); }
      const hist = await readJson("history.json", []);
      await writeJson("history.json", hist.concat(done));
    }
    list = remain; pos = 0; await saveItems();
    showCard(); refreshExcel();
    bar(`台帳に${done.length}枚を反映しました。` + (remain.length ? `${remain.length}枚は問題があるため残しています。内容を確認してください。` : ""), remain.length ? "e" : "k");
  } catch (e) { bar(e.message, "e"); }
  $("saveAll").disabled = false;
}

function initPc() {
  FIELDS.forEach((k) => ($(k).oninput = later));
  $("item").onchange = () => { updateSheet(); later(); };
  $("maker").onchange = () => {
    const isNew = $("maker").value === NEW;
    $("np").style.display = isNew ? "block" : "none"; $("ne").textContent = "";
    if (!isNew) { updateSheet(); later(); }
  };
  $("cand").onchange = $("over").onchange = later;
  $("na").onclick = async () => {
    const name = $("nn").value.trim(), code = $("nc").value.trim(), group = $("ng").value;
    if (!name || !/^\d{3}$/.test(code)) return ($("ne").textContent = "製造業者名と3桁のコードを入力してください。");
    if (makers.some((m) => m.code === code || m.name === name)) return ($("ne").textContent = "同じ名前かコードが登録済みです。");
    makers.push({ code, name, group }); makers.sort((a, b) => a.code.localeCompare(b.code));
    await writeJson("makers.json", makers);
    fillMakers(name, code); $("np").style.display = "none"; $("nn").value = $("nc").value = "";
    updateSheet(); persist();
  };
  $("prev").onclick = async () => { await persist(); pos = (pos - 1 + list.length) % list.length; rot = 0; showCard(); };
  $("next").onclick = async () => { await persist(); pos = (pos + 1) % list.length; rot = 0; showCard(); };
  $("exclude").onclick = async () => { list[pos].excluded = !list[pos].excluded; await persist(); showCard(); };
  $("rot").onclick = () => { rot = (rot + 90) % 360; $("pic").style.transform = `rotate(${rot}deg)`; };
  $("zoom").onclick = $("pic").onclick = () => $("pic").classList.toggle("z");
  $("import").onclick = importAll;
  $("saveAll").onclick = saveAll;
  $("openFolder").onclick = () => openFolder(true);
  setInterval(() => { if (!document.hidden && folder && $("tabPc").classList.contains("on")) { remote(); refreshExcel(); } }, 30000);
}
