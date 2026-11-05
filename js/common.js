"use strict";
// 設定（リポジトリ・トークン）と GitHub API。設定はこの端末のブラウザにだけ保存する。

const API = "https://api.github.com", DEFAULT_REPO = "markone-1/recycle-ocr";
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

const mem = {};
let storageOk = true;
const store = {
  get(k) { try { const v = localStorage.getItem(k); if (v !== null) return v; } catch { storageOk = false; } return mem[k] || ""; },
  set(k, v) { mem[k] = v; try { localStorage.setItem(k, v); } catch { storageOk = false; } },
};
const cfg = () => ({ repo: store.get("repo") || DEFAULT_REPO, token: store.get("token") });

async function gh(path, opt = {}) {
  const c = cfg();
  const r = await fetch(`${API}/repos/${c.repo}${path}`, {
    method: opt.method || "GET",
    headers: { Authorization: "Bearer " + c.token, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28" },
    body: opt.body ? JSON.stringify(opt.body) : undefined,
  });
  if (opt.allow404 && r.status === 404) return null;
  if (!r.ok) throw new Error(r.status === 401 ? "トークンが無効です（設定を確認）" : `GitHub との通信に失敗しました（${r.status}）`);
  return r.status === 204 ? null : r.json();
}

function unb64(s) {
  const bin = atob(s.replace(/\s/g, "")), u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}

// 処理中ダイアログ（閉じられない）。pct（0〜1）を渡すとプログレスバー、省略するとくるくるを出す
function busy(text, pct) {
  const d = $("busy");
  $("busyText").textContent = text;
  $("busySpin").classList.toggle("hidden", pct !== undefined);
  $("busyBar").classList.toggle("hidden", pct === undefined);
  if (pct !== undefined) $("busyFill").style.width = Math.round(Math.min(1, Math.max(0, pct)) * 100) + "%";
  if (!d.open) d.showModal();
}
const busyEnd = () => { if ($("busy").open) $("busy").close(); };

function initSettings(onSaved) {
  $("busy").oncancel = (e) => e.preventDefault();   // Esc で閉じない

  $("openCfg").onclick = () => { const c = cfg(); $("repo").value = c.repo; $("token").value = c.token; $("cfg").showModal(); };
  $("cancelCfg").onclick = () => $("cfg").close();
  $("saveCfg").onclick = () => {
    const repo = $("repo").value.trim(), token = $("token").value.trim();
    if (!/^[\w.-]+\/[\w.-]+$/.test(repo)) return alert("処理用リポジトリは「ユーザー名/リポジトリ名」の形で入力してください");
    if (!token) return alert("GitHub アクセストークンを入力してください");
    store.set("repo", repo); store.set("token", token); $("cfg").close();
    if (!storageOk) alert("このブラウザでは設定を保存できません。画面を開いている間だけ有効です。");
    onSaved();
  };
}
