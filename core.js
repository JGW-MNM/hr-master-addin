/*
 * core.js — logika bisnis SSOT Employee Master (tanpa Excel, tanpa UI).
 * Semua fungsi menerima data biasa (array) dan mengembalikan hasil:
 *   { errors: [], updates: [{index, row}], newRows: [], logs: [] }
 * Karena murni, file ini bisa dites di Node tanpa Excel.
 */
(function (root) {
  "use strict";

  // Kolom MASTER (0-based)
  const C = { ST: 0, ID: 1, NM: 2, JOIN: 3, TERM: 4, DEPT: 5, DEPTN: 6, DIV: 7,
    JOB: 8, JOBN: 9, DR: 10, DRN: 11, DRJ: 12, DRJN: 13 };
  const MASTER_FMT = ["@", "@", "@", "dd-mmm-yyyy", "dd-mmm-yyyy", "@", "@", "@", "@", "@", "@", "@", "@", "@"];
  const LOG_FMT = ["dd-mmm-yyyy hh:mm", "@", "@", "@", "@", "@", "@", "dd-mmm-yyyy", "@", "@"];

  const s = (v) => (v === null || v === undefined ? "" : String(v).trim());
  const isDate = (v) => typeof v === "number" && v > 1;

  // ---- tanggal: Excel serial <-> "yyyy-mm-dd" (input type=date) ----
  const EPOCH = Date.UTC(1899, 11, 30);
  function isoToSerial(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s(iso));
    return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000 : "";
  }
  function serialToIso(v) {
    if (!isDate(v)) return "";
    return new Date(EPOCH + Math.round(v * 86400000)).toISOString().slice(0, 10);
  }
  function fmtDate(v) {
    if (!isDate(v)) return s(v);
    const d = new Date(EPOCH + Math.round(v * 86400000));
    const mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getUTCMonth()];
    return ("0" + d.getUTCDate()).slice(-2) + "-" + mon + "-" + d.getUTCFullYear();
  }
  function nowSerial() {
    const d = new Date();
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / 86400000;
  }

  // ---- state ----
  function buildState(master, deptRows, jobRows) {
    const idx = new Map();
    master.forEach((r, i) => { const k = s(r[C.ID]); if (k) idx.set(k, i); });
    const depts = new Map(), jobs = new Map();
    deptRows.forEach((r) => { if (s(r[0])) depts.set(s(r[0]), r.map(s)); });
    jobRows.forEach((r) => { if (s(r[0])) jobs.set(s(r[0]), r.map(s)); });
    return { master: master.map((r) => r.slice()), idx, depts, jobs };
  }

  const emp = (st, id) => (st.idx.has(id) ? st.master[st.idx.get(id)] : undefined);
  const isActive = (row) => !!row && s(row[C.ST]) === "Active";

  function activeSubordinates(st, id) {
    const self = st.idx.get(id);
    return st.master.filter((r, i) => i !== self && s(r[C.DR]) === id && isActive(r));
  }

  /** true jika `empId` ditaruh di bawah `newDr` membuat rantai atasan berputar */
  function makesCycle(st, empId, newDr) {
    const seen = new Set();
    let cur = newDr;
    while (cur && !seen.has(cur)) {
      if (cur === empId) return true;
      seen.add(cur);
      const r = emp(st, cur);
      if (!r) return false;
      const next = s(r[C.DR]);
      if (next === cur) return false; // puncak (mis. CEO melapor ke diri sendiri)
      cur = next;
    }
    return false;
  }

  function logRow(action, id, name, field, oldV, newV, eff, by, remarks) {
    return [nowSerial(), action, id, name, field, oldV, newV, isDate(eff) ? eff : "", by, remarks];
  }

  function result() { return { errors: [], updates: [], newRows: [], logs: [] }; }

  // ---- 1. ONBOARDING ----
  function onboarding(st, f) {
    const r = result();
    const id = s(f.id), name = s(f.name), join = f.join, dept = s(f.dept), job = s(f.job),
      dr = s(f.dr), by = s(f.by), rem = s(f.remarks);
    if (!id) r.errors.push("Employee ID wajib diisi");
    else if (st.idx.has(id)) r.errors.push("Employee ID sudah ada di MASTER");
    if (!name) r.errors.push("Employee Name wajib diisi");
    if (!isDate(join)) r.errors.push("Join Date wajib diisi");
    const d = st.depts.get(dept), j = st.jobs.get(job);
    if (!dept) r.errors.push("Dept Code wajib diisi"); else if (!d) r.errors.push("Dept Code tidak ada di REF_DEPT");
    if (!job) r.errors.push("Job Title Code wajib diisi"); else if (!j) r.errors.push("Job Title Code tidak ada di REF_JOB");
    const drRow = emp(st, dr);
    if (!dr) r.errors.push("Direct Report ID wajib diisi");
    else if (dr === id) r.errors.push("Atasan tidak boleh diri sendiri");
    else if (!drRow) r.errors.push("Direct Report ID tidak ada di MASTER");
    else if (!isActive(drRow)) r.errors.push("Atasan berstatus Inactive");
    if (!by) r.errors.push("Submitted By wajib diisi");
    if (r.errors.length) return r;

    r.newRows.push(["Active", id, name, join, "", dept, d[1], d[2], job, j[1],
      dr, s(drRow[C.NM]), s(drRow[C.JOB]), s(drRow[C.JOBN])]);
    r.logs.push(logRow("ONBOARDING", id, name, "New hire", "",
      `${j[1]} (${job}) | ${d[1]} | atasan ${dr}`, join, by, rem));
    return r;
  }

  // ---- 2. CHANGE ORGANIZATION ----
  function changeOrg(st, f) {
    const r = result();
    const id = s(f.id), newDept = s(f.newDept), newJob = s(f.newJob), eff = f.eff, by = s(f.by), rem = s(f.remarks);
    const row = emp(st, id);
    if (!id) r.errors.push("Employee ID wajib diisi");
    else if (!row) r.errors.push("Employee ID tidak ada di MASTER");
    else if (!isActive(row)) r.errors.push("Karyawan berstatus Inactive");
    if (!newDept && !newJob) r.errors.push("Isi minimal salah satu: New Department atau New Job Title");
    const d = newDept ? st.depts.get(newDept) : null;
    const j = newJob ? st.jobs.get(newJob) : null;
    if (newDept && !d) r.errors.push("New Dept Code tidak ada di REF_DEPT");
    if (newJob && !j) r.errors.push("New Job Title Code tidak ada di REF_JOB");
    if (row && newDept && newDept === s(row[C.DEPT])) r.errors.push("New Department sama dengan yang sekarang");
    if (row && newJob && newJob === s(row[C.JOB])) r.errors.push("New Job Title sama dengan yang sekarang");
    if (!isDate(eff)) r.errors.push("Effective Date wajib diisi");
    if (!by) r.errors.push("Submitted By wajib diisi");
    if (r.errors.length) return r;

    const i = st.idx.get(id), nr = row.slice(), name = s(row[C.NM]);
    if (d) {
      r.logs.push(logRow("CHANGE ORG", id, name, "Department",
        `${s(row[C.DEPTN])} (${s(row[C.DEPT])}) / ${s(row[C.DIV])}`, `${d[1]} (${newDept}) / ${d[2]}`, eff, by, rem));
      nr[C.DEPT] = newDept; nr[C.DEPTN] = d[1]; nr[C.DIV] = d[2];
    }
    if (j) {
      r.logs.push(logRow("CHANGE ORG", id, name, "Job Title",
        `${s(row[C.JOBN])} (${s(row[C.JOB])})`, `${j[1]} (${newJob})`, eff, by, rem));
      nr[C.JOB] = newJob; nr[C.JOBN] = j[1];
      let n = 0;
      st.master.forEach((sub, si) => {
        if (si !== i && s(sub[C.DR]) === id) {
          const ns = sub.slice(); ns[C.DRJ] = newJob; ns[C.DRJN] = j[1];
          r.updates.push({ index: si, row: ns }); n++;
        }
      });
      if (n) r.logs.push(logRow("CHANGE ORG", id, name, "Direct Report Job Title (bawahan)", "",
        `${n} bawahan diperbarui → ${j[1]}`, eff, by, "otomatis"));
    }
    r.updates.push({ index: i, row: nr });
    return r;
  }

  // ---- 3. CHANGE DIRECT REPORT ----
  function changeReport(st, f) {
    const r = result();
    const id = s(f.id), newDr = s(f.newDr), eff = f.eff, by = s(f.by), rem = s(f.remarks);
    const row = emp(st, id), drRow = emp(st, newDr);
    if (!id) r.errors.push("Employee ID wajib diisi");
    else if (!row) r.errors.push("Employee ID tidak ada di MASTER");
    else if (!isActive(row)) r.errors.push("Karyawan berstatus Inactive");
    if (!newDr) r.errors.push("New Direct Report wajib diisi");
    else if (newDr === id) r.errors.push("Atasan tidak boleh diri sendiri");
    else if (!drRow) r.errors.push("New Direct Report ID tidak ada di MASTER");
    else if (!isActive(drRow)) r.errors.push("Atasan baru berstatus Inactive");
    if (row && newDr && newDr === s(row[C.DR])) r.errors.push("Atasan baru sama dengan atasan sekarang");
    if (row && drRow && newDr !== id && makesCycle(st, id, newDr))
      r.errors.push("Rantai atasan jadi berputar (atasan baru adalah bawahan dari karyawan ini)");
    if (!isDate(eff)) r.errors.push("Effective Date wajib diisi");
    if (!by) r.errors.push("Submitted By wajib diisi");
    if (r.errors.length) return r;

    const nr = row.slice();
    const label = (a, b, c) => [a, b, c ? `(${c})` : ""].filter((x) => x).join(" ");
    r.logs.push(logRow("CHANGE REPORT", id, s(row[C.NM]), "Direct Report",
      label(s(row[C.DR]), s(row[C.DRN]), s(row[C.DRJN])), label(newDr, s(drRow[C.NM]), s(drRow[C.JOBN])), eff, by, rem));
    nr[C.DR] = newDr; nr[C.DRN] = s(drRow[C.NM]); nr[C.DRJ] = s(drRow[C.JOB]); nr[C.DRJN] = s(drRow[C.JOBN]);
    r.updates.push({ index: st.idx.get(id), row: nr });
    return r;
  }

  // ---- 4. OFFBOARDING ----
  function offboarding(st, f) {
    const r = result();
    const id = s(f.id), term = f.term, by = s(f.by), rem = s(f.remarks);
    const row = emp(st, id);
    if (!id) r.errors.push("Employee ID wajib diisi");
    else if (!row) r.errors.push("Employee ID tidak ada di MASTER");
    else if (!isActive(row)) r.errors.push("Karyawan sudah Inactive");
    if (!isDate(term)) r.errors.push("Termination Date wajib diisi");
    else if (row && isDate(row[C.JOIN]) && term < row[C.JOIN]) r.errors.push("Termination Date lebih awal dari Join Date");
    if (!by) r.errors.push("Submitted By wajib diisi");
    if (row) {
      const subs = activeSubordinates(st, id);
      if (subs.length) r.errors.push(`Masih punya ${subs.length} bawahan aktif — pindahkan dulu lewat Change Direct Report`);
    }
    if (r.errors.length) return r;

    const nr = row.slice();
    r.logs.push(logRow("OFFBOARDING", id, s(row[C.NM]), "Status", "Active", "Inactive", term, by, rem));
    r.logs.push(logRow("OFFBOARDING", id, s(row[C.NM]), "Termination Date", fmtDate(row[C.TERM]), fmtDate(term), term, by, rem));
    nr[C.ST] = "Inactive"; nr[C.TERM] = term;
    r.updates.push({ index: st.idx.get(id), row: nr });
    return r;
  }

  const api = { C, MASTER_FMT, LOG_FMT, s, isDate, isoToSerial, serialToIso, fmtDate, buildState, emp, isActive,
    activeSubordinates, makesCycle, onboarding, changeOrg, changeReport, offboarding };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.HRCore = api;
})(this);
