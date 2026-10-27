/*
 * app.js — UI task pane + lapisan data.
 *   - Di dalam Excel  : baca/tulis tabel tblMaster, tblDept, tblJob, tblLog lewat Office.js
 *   - Di browser biasa: Demo Mode dengan data fiktif (demo-data.js), tidak menyentuh Excel
 * Semua aturan bisnis ada di core.js (HRCore).
 */
"use strict";

const PWD = "SSOT2026"; // password proteksi sheet MASTER & CHANGE_LOG (POC). Samakan dengan di workbook.
const H = HRCore, C = H.C, s = H.s;
const PROTECT_OPTS = { allowAutoFilter: true, allowSort: true, allowFormatColumns: true };

let layer = null;        // DataExcel atau DataDemo
let data = null;         // {master, dept, job, log}
let st = null;           // state hasil HRCore.buildState
let currentView = "home";

/* ======================= DATA LAYER: EXCEL ======================= */
const DataExcel = {
  name: "Excel",
  async load() {
    return Excel.run(async (ctx) => {
      const t = (n) => { const r = ctx.workbook.tables.getItem(n).getDataBodyRange(); r.load("values"); return r; };
      const m = t("tblMaster"), d = t("tblDept"), j = t("tblJob"), l = t("tblLog");
      await ctx.sync();
      return { master: m.values, dept: d.values, job: j.values, log: l.values.filter((r) => s(r[1])) };
    });
  },
  /** Terapkan hasil HRCore (updates + newRows + logs) ke workbook */
  async commit(res) {
    return Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const shM = wb.worksheets.getItem("MASTER"), shL = wb.worksheets.getItem("CHANGE_LOG");
      shM.protection.load("protected"); shL.protection.load("protected");
      await ctx.sync();
      if (shM.protection.protected) shM.protection.unprotect(PWD);
      if (shL.protection.protected) shL.protection.unprotect(PWD);
      await ctx.sync();
      try {
        const tM = wb.tables.getItem("tblMaster");
        const body = tM.getDataBodyRange();
        res.updates.forEach((u) => { body.getCell(u.index, 0).getResizedRange(0, 13).values = [u.row]; });
        await appendRows(ctx, tM, res.newRows, H.MASTER_FMT);
        await appendRows(ctx, wb.tables.getItem("tblLog"), res.logs, H.LOG_FMT);
        await ctx.sync();
      } finally {
        shM.protection.protect(PROTECT_OPTS, PWD);
        shL.protection.protect(PROTECT_OPTS, PWD);
        await ctx.sync();
      }
    });
  },
};

/** Tambah baris ke tabel. Format kolom di-set dulu supaya ID tetap teks (tidak jadi angka). */
async function appendRows(ctx, table, rows, fmts) {
  if (!rows.length) return;
  const body = table.getDataBodyRange();
  body.load("values,rowCount");
  await ctx.sync();
  const blank = () => fmts.map(() => "");
  let start = body.rowCount;
  if (start === 1 && body.values[0].every((v) => s(v) === "")) {   // tabel kosong: isi baris bawaan dulu
    start = 0;
    if (rows.length > 1) table.rows.add(null, rows.slice(1).map(blank));
  } else {
    table.rows.add(null, rows.map(blank));
  }
  const rng = table.getDataBodyRange().getCell(start, 0).getResizedRange(rows.length - 1, fmts.length - 1);
  rng.numberFormat = rows.map(() => fmts);
  rng.values = rows;
}

/* ======================= DATA LAYER: DEMO ======================= */
const DataDemo = {
  name: "Demo",
  async load() { return JSON.parse(JSON.stringify(HR_DEMO)); },
  async commit(res) {
    res.updates.forEach((u) => { HR_DEMO.master[u.index] = u.row; });
    res.newRows.forEach((r) => HR_DEMO.master.push(r));
    res.logs.forEach((r) => HR_DEMO.log.push(r));
  },
};

/* ======================= HELPERS ======================= */
const $ = (sel, el = document) => el.querySelector(sel);
const esc = (v) => s(v).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const empLabel = (r) => `${s(r[C.NM]) || "(nama kosong)"} · ${s(r[C.JOBN]) || "-"}`;
const todayIso = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
function getBy() { try { return localStorage.getItem("hrm.by") || ""; } catch (e) { return ""; } }
function setBy(v) { try { localStorage.setItem("hrm.by", v); } catch (e) { /* abaikan */ } }

async function refresh() {
  data = await layer.load();
  st = H.buildState(data.master, data.dept, data.job);
  const opt = (v, l) => `<option value="${esc(v)}">${esc(l)}</option>`;
  const withId = st.master.filter((r) => s(r[C.ID]));
  $("#dl-emp").innerHTML = withId.map((r) => opt(r[C.ID], `${empLabel(r)} · ${s(r[C.ST])}`)).join("");
  $("#dl-emp-active").innerHTML = withId.filter(H.isActive).map((r) => opt(r[C.ID], empLabel(r))).join("");
  $("#dl-dept").innerHTML = [...st.depts.values()].map((d) => opt(d[0], `${d[1]} · ${d[2]}`)).join("");
  $("#dl-job").innerHTML = [...st.jobs.values()].map((j) => opt(j[0], j[1] + (j[2] ? " ⚑" : ""))).join("");
}

function empCard(id, title) {
  const r = H.emp(st, s(id));
  if (!s(id)) return "";
  if (!r) return `<div class="peek bad">⚠ ID tidak ditemukan di MASTER</div>`;
  const pill = H.isActive(r) ? `<span class="pill on">Active</span>` : `<span class="pill off">Inactive</span>`;
  return `<div class="card" style="margin:8px 0 0">${title ? `<div class="muted" style="font-size:11px;margin-bottom:6px">${title}</div>` : ""}
    <dl class="kv">
      <dt>Nama</dt><dd>${esc(r[C.NM]) || '<span class="muted">(kosong)</span>'} ${pill}</dd>
      <dt>Department</dt><dd>${esc(r[C.DEPTN])} <span class="muted">· ${esc(r[C.DIV])}</span></dd>
      <dt>Job Title</dt><dd>${esc(r[C.JOBN])} <span class="muted">(${esc(r[C.JOB])})</span></dd>
      <dt>Atasan</dt><dd>${esc(r[C.DR])} ${esc(r[C.DRN])} <span class="muted">${esc(r[C.DRJN])}</span></dd>
      <dt>Join Date</dt><dd>${esc(H.fmtDate(r[C.JOIN]))}</dd>
    </dl></div>`;
}

function deptPeek(code) {
  if (!s(code)) return "";
  const d = st.depts.get(s(code));
  return d ? `<b>${esc(d[1])}</b> · ${esc(d[2])}` : `<span class="bad">⚠ Kode tidak ada di REF_DEPT</span>`;
}
function jobPeek(code) {
  if (!s(code)) return "";
  const j = st.jobs.get(s(code));
  if (!j) return `<span class="bad">⚠ Kode tidak ada di REF_JOB</span>`;
  return `<b>${esc(j[1])}</b>` + (j[2] ? `<br><span class="flag">⚑ ${esc(j[2])}</span>` : "");
}
function setPeek(id, html) {
  const el = $(id);
  el.innerHTML = html;
  el.classList.toggle("bad", html.includes('class="bad"'));
}

function field(id, label, opts = {}) {
  const type = opts.type || "text";
  const list = opts.list ? ` list="${opts.list}"` : "";
  const req = opts.optional ? ` <span class="opt">(opsional)</span>` : " *";
  const input = type === "textarea"
    ? `<textarea id="${id}" placeholder="${esc(opts.ph || "")}"></textarea>`
    : `<input id="${id}" type="${type}"${list} autocomplete="off" placeholder="${esc(opts.ph || "")}" />`;
  return `<label for="${id}">${label}${req}</label>${input}${opts.peek ? `<div class="peek" id="${id}-peek"></div>` : ""}`;
}
const footerFields = (dateId, dateLabel) => `
  <div class="row2"><div>${field(dateId, dateLabel, { type: "date" })}</div><div>${field("f-by", "Submitted By", { ph: "Nama Anda" })}</div></div>
  ${field("f-rem", "Remarks", { type: "textarea", optional: true, ph: "Alasan / nomor SK / catatan" })}
  <button class="btn primary" id="submit">Submit</button><div id="msg"></div>`;

/* ======================= VIEWS ======================= */
const VIEWS = {
  home() {
    const act = st.master.filter((r) => s(r[C.ID]) && H.isActive(r)).length;
    const ina = st.master.filter((r) => s(r[C.ID]) && !H.isActive(r)).length;
    const recent = data.log.slice(-8).reverse();
    return `
      <h2>Ringkasan</h2><p class="sub">Data langsung dari sheet MASTER${layer === DataDemo ? " (demo — data fiktif)" : ""}.</p>
      <div class="kpis">
        <div class="kpi"><div class="n">${act}</div><div class="l">Active</div></div>
        <div class="kpi"><div class="n">${ina}</div><div class="l">Inactive</div></div>
        <div class="kpi"><div class="n">${act + ina}</div><div class="l">Total</div></div>
      </div>
      <div class="card"><div class="actions">
        <button class="btn ghost" data-go="onboarding">+ Onboarding</button>
        <button class="btn ghost" data-go="changeOrg">Change Org</button>
        <button class="btn ghost" data-go="changeReport">Direct Report</button>
        <button class="btn ghost" data-go="offboarding">Offboarding</button>
      </div></div>
      <div class="card"><b>Perubahan terakhir</b>
        ${recent.length ? `<ul class="list" style="margin-top:6px">${recent.map((r) => `
          <li><b>${esc(r[1])}</b> · ${esc(r[2])} ${esc(r[3])}<br>${esc(r[4])}: ${esc(r[5]) || "—"} → ${esc(r[6])}
          <div class="meta">${esc(H.fmtDate(Math.floor(r[0])))} · eff ${esc(H.fmtDate(r[7]))} · oleh ${esc(r[8])}</div></li>`).join("")}</ul>`
          : `<p class="muted" style="margin:6px 0 0">Belum ada perubahan tercatat.</p>`}
      </div>
      <button class="btn ghost" id="reload">↻ Refresh data</button>`;
  },

  onboarding() {
    return `<h2>Onboarding</h2><p class="sub">Karyawan baru masuk ke MASTER sebagai Active.</p>
      <div class="card">
        ${field("f-id", "Employee ID", { peek: true })}
        ${field("f-name", "Employee Name")}
        ${field("f-dept", "Dept Code", { list: "dl-dept", peek: true, ph: "Ketik kode / nama dept" })}
        ${field("f-job", "Job Title Code", { list: "dl-job", peek: true, ph: "Ketik kode / nama job" })}
        ${field("f-dr", "Direct Report ID", { list: "dl-emp-active", peek: true, ph: "Cari atasan" })}
        ${footerFields("f-date", "Join Date")}
      </div>`;
  },

  changeOrg() {
    return `<h2>Change Organization</h2><p class="sub">Isi New Department dan/atau New Job Title.</p>
      <div class="card">
        ${field("f-id", "Employee ID", { list: "dl-emp-active", peek: true, ph: "Cari karyawan" })}
        ${field("f-dept", "New Dept Code", { list: "dl-dept", peek: true, optional: true })}
        ${field("f-job", "New Job Title Code", { list: "dl-job", peek: true, optional: true })}
        ${footerFields("f-date", "Effective Date")}
      </div>`;
  },

  changeReport() {
    return `<h2>Change Direct Report</h2><p class="sub">Pindahkan karyawan ke atasan langsung yang baru.</p>
      <div class="card">
        ${field("f-id", "Employee ID", { list: "dl-emp-active", peek: true, ph: "Cari karyawan" })}
        ${field("f-dr", "New Direct Report ID", { list: "dl-emp-active", peek: true, ph: "Cari atasan baru" })}
        ${footerFields("f-date", "Effective Date")}
      </div>`;
  },

  offboarding() {
    return `<h2>Offboarding</h2><p class="sub">Status menjadi Inactive. Bawahan aktif harus dipindah dulu.</p>
      <div class="card">
        ${field("f-id", "Employee ID", { list: "dl-emp-active", peek: true, ph: "Cari karyawan" })}
        <div id="subs"></div>
        ${footerFields("f-date", "Termination Date")}
      </div>`;
  },
};

/* ======================= LIVE PREVIEW ======================= */
function wireLive(view) {
  const v = (id) => ($(id) ? s($(id).value) : "");
  const update = () => {
    const id = v("#f-id");
    if (view === "onboarding") {
      setPeek("#f-id-peek", id && H.emp(st, id) ? `<span class="bad">⚠ Employee ID sudah ada di MASTER</span>` : "");
      setPeek("#f-dept-peek", deptPeek(v("#f-dept")));
      setPeek("#f-job-peek", jobPeek(v("#f-job")));
      setPeek("#f-dr-peek", v("#f-dr") ? empCard(v("#f-dr"), "Atasan") : "");
    }
    if (view === "changeOrg") {
      setPeek("#f-id-peek", empCard(id, "Saat ini"));
      setPeek("#f-dept-peek", deptPeek(v("#f-dept")));
      let jp = jobPeek(v("#f-job"));
      const n = H.emp(st, id) ? st.master.filter((r) => s(r[C.DR]) === id && s(r[C.ID]) !== id).length : 0;
      if (v("#f-job") && st.jobs.get(v("#f-job")) && n) jp += `<br><span class="flag">↳ Direct Report Job Title ${n} bawahan ikut diperbarui</span>`;
      setPeek("#f-job-peek", jp);
    }
    if (view === "changeReport") {
      setPeek("#f-id-peek", empCard(id, "Saat ini"));
      let p = v("#f-dr") ? empCard(v("#f-dr"), "Atasan baru") : "";
      if (id && v("#f-dr") && H.emp(st, v("#f-dr")) && H.makesCycle(st, id, v("#f-dr")))
        p += `<div class="peek bad">⚠ Atasan baru adalah bawahan dari karyawan ini (rantai berputar)</div>`;
      setPeek("#f-dr-peek", p);
    }
    if (view === "offboarding") {
      setPeek("#f-id-peek", empCard(id));
      const subs = H.emp(st, id) ? H.activeSubordinates(st, id) : [];
      $("#subs").innerHTML = subs.length ? `<div class="alert err"><b>⚠ ${subs.length} bawahan aktif</b> — pindahkan dulu:
        <ul class="list subs" style="margin-top:6px">${subs.map((r) => `<li><span>${esc(r[C.ID] || "(tanpa ID)")} ${esc(r[C.NM])}<br><span class="meta">${esc(r[C.JOBN])}</span></span>
        ${s(r[C.ID]) ? `<button class="btn ghost" data-move="${esc(r[C.ID])}">Pindahkan →</button>` : ""}</li>`).join("")}</ul></div>` : "";
    }
  };
  $("#app").querySelectorAll("input").forEach((el) => el.addEventListener("input", update));
  update();
  return update;
}

/* ======================= SUBMIT ======================= */
async function submit(view) {
  const v = (id) => ($(id) ? $(id).value : "");
  const btn = $("#submit"), msg = $("#msg");
  btn.disabled = true; btn.textContent = "Memproses…"; msg.innerHTML = "";
  try {
    setBy(s(v("#f-by")));
    await refresh(); // selalu pakai data terbaru sebelum validasi (kalau ada user lain yang mengubah)
    const date = H.isoToSerial(v("#f-date"));
    const base = { id: v("#f-id"), by: v("#f-by"), remarks: v("#f-rem") };
    const res =
      view === "onboarding" ? H.onboarding(st, { ...base, name: v("#f-name"), join: date, dept: v("#f-dept"), job: v("#f-job"), dr: v("#f-dr") }) :
      view === "changeOrg" ? H.changeOrg(st, { ...base, newDept: v("#f-dept"), newJob: v("#f-job"), eff: date }) :
      view === "changeReport" ? H.changeReport(st, { ...base, newDr: v("#f-dr"), eff: date }) :
      H.offboarding(st, { ...base, term: date });

    if (res.errors.length) {
      msg.innerHTML = `<div class="alert err"><b>Belum bisa disubmit:</b><ul>${res.errors.map((e) => `<li>${esc(e)}</li>`).join("")}</ul></div>`;
      return;
    }
    await layer.commit(res);
    await refresh();
    const id = s(v("#f-id"));
    render(view);
    $("#msg").innerHTML = `<div class="alert ok">✓ Tersimpan untuk <b>${esc(id)}</b> — ${res.logs.length} baris dicatat di CHANGE_LOG.</div>`;
  } catch (e) {
    console.error(e);
    msg.innerHTML = `<div class="alert err"><b>Error:</b> ${esc(e.message || e)}</div>`;
  } finally {
    if ($("#submit")) { $("#submit").disabled = false; $("#submit").textContent = "Submit"; }
  }
}

/* ======================= RENDER / ROUTING ======================= */
function render(view, prefill = {}) {
  currentView = view;
  document.querySelectorAll("#nav button").forEach((b) => b.classList.toggle("on", b.dataset.view === view));
  $("#app").innerHTML = VIEWS[view]();
  if ($("#f-by")) $("#f-by").value = getBy();
  if ($("#f-date") && view !== "onboarding") $("#f-date").value = todayIso();
  Object.entries(prefill).forEach(([k, val]) => { if ($(k)) $(k).value = val; });
  if (view !== "home") {
    const update = wireLive(view);
    $("#submit").addEventListener("click", () => submit(view));
    update();
  } else {
    $("#reload").addEventListener("click", async () => { $("#app").innerHTML = `<div class="loading">Memuat data…</div>`; await refresh(); render("home"); });
    $("#app").querySelectorAll("[data-go]").forEach((b) => b.addEventListener("click", () => render(b.dataset.go)));
  }
}

async function start(inExcel) {
  layer = inExcel ? DataExcel : DataDemo;
  const badge = $("#mode");
  badge.textContent = inExcel ? "Excel" : "Demo mode";
  badge.classList.toggle("demo", !inExcel);
  document.querySelectorAll("#nav button").forEach((b) => b.addEventListener("click", () => render(b.dataset.view)));
  // tombol "Pindahkan →" di daftar bawahan (Offboarding) → buka Change Direct Report
  $("#app").addEventListener("click", (e) => {
    const mv = e.target.closest("[data-move]");
    if (mv) render("changeReport", { "#f-id": mv.dataset.move });
  });
  try {
    await refresh();
    const q = new URLSearchParams(location.search).get("view");
    render(VIEWS[q] ? q : "home");
  } catch (e) {
    console.error(e);
    $("#app").innerHTML = `<div class="alert err"><b>Gagal membaca workbook.</b><br>${esc(e.message || e)}<br><br>
      Pastikan file yang terbuka adalah <b>SSOT Employee Master.xlsx</b> (berisi tabel tblMaster, tblDept, tblJob, tblLog).</div>`;
  }
}

if (typeof Office !== "undefined") {
  Office.onReady((info) => start(info.host === Office.HostType.Excel));
} else {
  start(false);
}
