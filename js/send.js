"use strict";
// 送信：撮影/フォルダ選択 → 向きを直す → 長辺1280px・JPEG75に圧縮 → job/<ID> ブランチに in/NN.jpg として送る
// （ブランチができると GitHub Actions が読み取りを始める）

const MAX_EDGE = 1280, QUALITY = 0.75, PREVIEW = 640;
let photos = [];

function drawPhoto(canvas, p, maxEdge) {
  const s = Math.min(1, maxEdge / Math.max(p.img.naturalWidth, p.img.naturalHeight));
  const w = Math.round(p.img.naturalWidth * s), h = Math.round(p.img.naturalHeight * s);
  const side = p.rot % 180 !== 0;
  canvas.width = side ? h : w; canvas.height = side ? w : h;
  const c = canvas.getContext("2d");
  c.translate(canvas.width / 2, canvas.height / 2);
  c.rotate(p.rot * Math.PI / 180);
  c.drawImage(p.img, -w / 2, -h / 2, w, h);
}

function renderPhotos() {
  const box = $("photos"); box.innerHTML = "";
  photos.forEach((p, i) => {
    const card = document.createElement("div"); card.className = "card";
    const cv = document.createElement("canvas"); drawPhoto(cv, p, PREVIEW); card.append(cv);
    const tools = document.createElement("div"); tools.className = "tools3";
    [["左へ回転", () => turn(p, -90)], ["右へ回転", () => turn(p, 90)], ["削除", () => { photos.splice(i, 1); renderPhotos(); }]]
      .forEach(([t, fn]) => { const b = document.createElement("button"); b.textContent = t; b.onclick = fn; if (t === "削除") b.className = "del"; tools.append(b); });
    card.append(tools); box.append(card);
  });
  $("send").textContent = photos.length ? `送信して読み取り（${photos.length}枚）` : "送信して読み取り";
  $("send").disabled = !photos.length;   // 写真が無いうちは押せない（薄い色）
}
function turn(p, deg) { p.rot = (p.rot + deg + 360) % 360; renderPhotos(); }
function sendNote(t, cls) { $("sendNote").textContent = t; $("sendNote").className = "note " + (cls || ""); }

function addFiles(files) {
  Promise.all([...files].map((f) => new Promise((res) => {
    const img = new Image();
    img.onload = () => res({ img, rot: 0 });
    img.onerror = () => res(null);
    img.src = URL.createObjectURL(f);
  }))).then((list) => {
    const ok = list.filter(Boolean);
    if (ok.length < list.length) sendNote("読み込めない画像がありました。", "err");
    photos.push(...ok); renderPhotos();
  });
}

async function sendPhotos() {
  if (!photos.length) return sendNote("先に撮影または選択をしてください。", "err");
  if (!cfg().token) { sendNote("先に「設定」でトークンを登録してください。", "err"); return $("openCfg").click(); }
  $("send").disabled = true;
  busy("画像を圧縮しています…");
  await new Promise((r) => setTimeout(r, 50));   // 圧縮で画面が固まる前にダイアログを描かせる
  try {
    const data = photos.map((p) => { const cv = document.createElement("canvas"); drawPhoto(cv, p, MAX_EDGE); return cv.toDataURL("image/jpeg", QUALITY).split(",")[1]; });
    const kb = Math.round(data.reduce((a, d) => a + d.length * 0.75, 0) / 1024);
    const now = new Date(), p2 = (n) => String(n).padStart(2, "0");
    const jobId = `${now.getFullYear()}${p2(now.getMonth() + 1)}${p2(now.getDate())}${p2(now.getHours())}${p2(now.getMinutes())}${p2(now.getSeconds())}-`
      + Math.random().toString(36).slice(2, 6).padEnd(4, "0");
    const main = await gh("/git/ref/heads/main");
    const base = await gh(`/git/commits/${main.object.sha}`);
    const tree = [];
    for (let i = 0; i < data.length; i++) {
      busy(`送信中 ${i + 1} / ${data.length}枚（圧縮後 ${kb}KB）`);
      const b = await gh("/git/blobs", { method: "POST", body: { content: data[i], encoding: "base64" } });
      tree.push({ path: `in/${p2(i + 1)}.jpg`, mode: "100644", type: "blob", sha: b.sha });
    }
    busy("送信を仕上げています…");
    const t = await gh("/git/trees", { method: "POST", body: { base_tree: base.tree.sha, tree } });
    const c = await gh("/git/commits", { method: "POST", body: { message: `job ${jobId}`, tree: t.sha, parents: [main.object.sha] } });
    await gh("/git/refs", { method: "POST", body: { ref: `refs/heads/job/${jobId}`, sha: c.sha } });
    photos = []; renderPhotos();
    sendNote(`${data.length}枚を送信しました（圧縮後 ${kb}KB）。数分で読み取りが終わります。PCで取り込んでください。`, "ok");
  } catch (e) {
    sendNote(`送信できませんでした（${e.message}）。電波と設定を確認してください。`, "err");
  }
  busyEnd();
  $("send").disabled = !photos.length;
}

function initSend() {
  $("cam").onchange = (e) => { addFiles(e.target.files); e.target.value = ""; };
  $("lib").onchange = (e) => { addFiles(e.target.files); e.target.value = ""; };
  $("send").onclick = sendPhotos;
}
