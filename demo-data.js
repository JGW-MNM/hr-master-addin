/*
 * demo-data.js — data FIKTIF untuk Demo Mode (saat halaman dibuka di browser biasa, di luar Excel).
 * Jangan taruh data karyawan asli di repo publik.
 */
(function (root) {
  const d = (iso) => HRCore.isoToSerial(iso);
  const DEPT = [
    ["DPT-100", "MANAGEMENT", "MANAGEMENT"],
    ["DPT-210", "HUMAN RESOURCES", "HR"],
    ["DPT-211", "HR OPERATIONS", "HR"],
    ["DPT-212", "TALENT ACQUISITION", "HR"],
    ["DPT-310", "FINANCE & ACCOUNTING", "FIN/ACC/TAX"],
    ["DPT-410", "PLANTING", "AGRI"],
    ["DPT-411", "UPKEEP & MANURING", "AGRI"],
  ];
  const JOB = [
    ["JT-001", "CEO", ""],
    ["JT-010", "Head of Human Resources", ""],
    ["JT-011", "HR Operations Lead", ""],
    ["JT-012", "Payroll Officer", ""],
    ["JT-013", "HRIS Specialist", ""],
    ["JT-014", "Talent Acquisition Lead", ""],
    ["JT-015", "Recruitment Officer", ""],
    ["JT-020", "Finance Manager", ""],
    ["JT-021", "Accounting Officer", ""],
    ["JT-030", "General Manager Agri", ""],
    ["JT-031", "Manager Planting", ""],
    ["JT-032", "Asisten Planting", ""],
    ["JT-033", "Manager Upkeep", ""],
    ["JT-034", "Asisten Upkeep", ""],
    ["JT-035", "Asisten Upkeep", "Contoh flag: nama sama dengan JT-034"],
  ];
  const dept = (c) => DEPT.find((x) => x[0] === c);
  const job = (c) => JOB.find((x) => x[0] === c);
  const people = [
    // id, name, join, dept, job, direct report id
    ["D0001", "Bambang Santoso", "2019-01-07", "DPT-100", "JT-001", "D0001"],
    ["D0002", "Ratna Wijaya", "2020-03-02", "DPT-210", "JT-010", "D0001"],
    ["D0003", "Hendra Gunawan", "2021-06-14", "DPT-211", "JT-011", "D0002"],
    ["D0004", "Sari Lestari", "2022-02-01", "DPT-211", "JT-012", "D0003"],
    ["D0005", "Yoga Pratama", "2023-08-21", "DPT-211", "JT-013", "D0003"],
    ["D0006", "Maya Kusuma", "2021-11-01", "DPT-212", "JT-014", "D0002"],
    ["D0007", "Fajar Nugroho", "2024-01-15", "DPT-212", "JT-015", "D0006"],
    ["D0008", "Dewi Anggraini", "2024-05-06", "DPT-212", "JT-015", "D0006"],
    ["D0009", "Agus Salim", "2020-09-01", "DPT-310", "JT-020", "D0001"],
    ["D0010", "Rina Marlina", "2023-03-13", "DPT-310", "JT-021", "D0009"],
    ["D0011", "Wahyu Hidayat", "2019-10-01", "DPT-410", "JT-030", "D0001"],
    ["D0012", "Teguh Saputra", "2021-04-05", "DPT-410", "JT-031", "D0011"],
    ["D0013", "Imam Fauzi", "2024-07-01", "DPT-410", "JT-032", "D0012"],
    ["D0014", "Rudi Hartono", "2025-02-03", "DPT-410", "JT-032", "D0012"],
    ["D0015", "Siti Rahma", "2022-10-10", "DPT-411", "JT-033", "D0011"],
    ["D0016", "Andi Setiawan", "2025-06-02", "DPT-411", "JT-034", "D0015"],
    ["D0017", "Putri Ayu", "2025-09-15", "DPT-411", "JT-035", "D0015"],
  ];
  const byId = Object.fromEntries(people.map((p) => [p[0], p]));
  const MASTER = people.map(([id, name, join, dc, jc, dr]) => {
    const dp = dept(dc), jb = job(jc), boss = byId[dr], bj = job(boss[4]);
    return ["Active", id, name, d(join), "", dc, dp[1], dp[2], jc, jb[1], dr, boss[1], bj[0], bj[1]];
  });
  root.HR_DEMO = { master: MASTER, dept: DEPT, job: JOB, log: [] };
})(this);
