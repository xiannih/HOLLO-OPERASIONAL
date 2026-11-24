/**
 * ============================================================================
 * HOLLO COFFEE & ROASTERY'S — MASTER BACKEND ENGINE V24 (DATA-SAFE)
 * Strict Data Trimmer, Sheet Normalizer, Dual-Way Sync & Native Charts
 * System Architect : Boy (RAJA_UTAMA)
 * Entity           : PT. HOLLO COFFEE
 * ----------------------------------------------------------------------------
 * V24: Perbaikan data-safe & keamanan. repair TIDAK PERNAH menghapus kolom,
 * baris, atau sheet; formula hanya dipasang pada baris ber-ID; read_all &
 * doPost wajib token login; kolom pin tidak lagi dikirim ke frontend;
 * jalur update menulis ulang formula; sinkron stok log berlaku untuk
 * insert, update, dan delete.
 * ============================================================================
 */

// V24: Peta sheet -> tab untuk otorisasi tulis berdasarkan "akses_tabs".
const SHEET_TO_TAB_ = {
  "DATA_PERUSAHAAN": "perusahaan",
  "USERS": "perusahaan",
  "JADWAL": "kalender",
  "LOG_KALENDER": "kalender",
  "MENU": "menu",
  "STOK": "stok",
  "CHECKLIST": "checklist",
  "SERAH_TERIMA": "serah-terima",
  "DAILY_REPORT": "daily-report"
};

// V24: Buat token sesi (UUID) dan simpan username di Script Cache, TTL 6 jam.
function issueToken_(username) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("auth_" + token, String(username), 6 * 60 * 60);
  return token;
}

// V24: Validasi token. Mengembalikan objek user (dari sheet USERS) atau null.
function requireAuth_(token) {
  if (!token) return null;
  const username = CacheService.getScriptCache().get("auth_" + String(token).trim());
  if (!username) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("USERS");
  if (!sh) return null;
  const users = getSheetRows(sh);
  const user = users.find(x => String(x.username).toLowerCase() === String(username).toLowerCase());
  return user || null;
}

// V24: Cek hak tulis user untuk sheet tertentu berdasarkan "akses_tabs".
function canWriteSheet_(user, sheetName) {
  if (!user) return false;
  if (Number(user.level) === 6 || String(user.role) === "RAJA_UTAMA") return true;
  const tabId = SHEET_TO_TAB_[sheetName];
  if (!tabId) return false;
  const s = String(user["akses_tabs"] || "");
  if (!s.trim()) return false;
  return s.split(',').map(x => x.trim()).indexOf(tabId + ":edit") !== -1;
}

function repairAndSyncDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // V24: cleanAndFormatSheet tidak lagi menghapus apa pun; kolom ekstra
  // hanya dicatat di sini dan dilaporkan di pesan hasil repair.
  const extraNotes = [];

  // 1. DATA PERUSAHAAN (5 KOLOM)
  const shCorp = getOrCreateSheet(ss, "DATA_PERUSAHAAN");
  const extraCorp = cleanAndFormatSheet(shCorp, [
    "KATEGORI", "PARAMETER_KUNCI", "NILAI", "LEVEL_MIN", "KETERANGAN"
  ]);
  if (extraCorp) extraNotes.push(extraCorp);
  if (shCorp.getLastRow() <= 1) {
    const corpSeed = [
      ["PROFIL", "NAMA_USAHA", "HOLLO COFFEE & ROASTERY'S", 1, "Brand Utama"],
      ["PROFIL", "BADAN_HUKUM", "PT. HOLLO COFFEE", 4, "Badan Usaha Resmi"],
      ["TARGET", "TARGET_BULANAN", 150000000, 4, "Target Omset Bulanan"],
      ["TARGET", "TARGET_HARIAN", 5000000, 3, "Target Rata-rata Harian"],
      ["JAM_OPS", "PREP_SHIFT_1", "08:30", 1, "Mulai Prep S1"],
      ["JAM_OPS", "SERAH_TERIMA_1_2", "16:30", 1, "Serah Terima S1 ke S2"],
      ["JAM_OPS", "CLOSE_ORDER_KASIR", "00:00", 1, "Close Order Malam"]
    ];
    shCorp.getRange(2, 1, corpSeed.length, 5).setValues(corpSeed);
    shCorp.getRange(4, 3, 2, 1).setNumberFormat("Rp #,##0");
  }

  // 2. USERS (11 KOLOM TEPAT)
  const shUsers = getOrCreateSheet(ss, "USERS");
  const userHeaders = ["user_id", "username", "nama_lengkap", "role", "level", "divisi", "tipe_kerja", "pin", "status", "akses_perusahaan", "akses_tabs"];
  const extraUsers = cleanAndFormatSheet(shUsers, userHeaders);
  if (extraUsers) extraNotes.push(extraUsers);
  if (shUsers.getLastRow() <= 1) {
    const userSeed = [
      ["USR-001", "boy", "Boy", "RAJA_UTAMA", 6, "MANAJEMEN", "Tetap", "0000", "Aktif", "YA", "dashboard:edit,kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit,perusahaan:edit,info-sistem:edit"],
      ["USR-002", "deni", "Deni", "OWNER", 5, "MANAJEMEN", "Tetap", "8899", "Aktif", "YA", "dashboard:view,kalender:view,checklist:view,stok:view,serah-terima:view,daily-report:view,menu:view"],
      ["USR-003", "munif", "Munif", "MANAGER", 4, "MANAJEMEN", "Tetap", "5566", "Aktif", "YA", "dashboard:edit,kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit,perusahaan:edit,info-sistem:edit"],
      ["USR-004", "sofia", "Sofia Salsa", "PURCHASING", 3, "MANAJEMEN", "Tetap", "3344", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit"],
      ["USR-005", "ateng", "Ateng", "HRD", 3, "MANAJEMEN", "Tetap", "2233", "Aktif", "YA", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit"],
      ["USR-006", "pace", "Pace", "HEAD", 2, "BARISTA", "Tetap", "1122", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit"],
      ["USR-007", "el", "El", "BARISTA", 1, "BARISTA", "Tetap", "1001", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit"],
      ["USR-008", "christy", "Christy", "BARISTA", 1, "BARISTA", "Tetap", "1002", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit"],
      ["USR-009", "fatir", "Fatir", "BARISTA", 1, "BARISTA", "Tetap", "1003", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit"],
      ["USR-010", "eko", "Eko", "HEAD", 2, "KITCHEN", "Tetap", "4455", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit,serah-terima:edit,daily-report:edit,menu:edit"],
      ["USR-011", "afdal", "Afdal", "KITCHEN", 1, "KITCHEN", "Tetap", "2001", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit"],
      ["USR-012", "rizky", "Rizky", "KITCHEN", 1, "KITCHEN", "Tetap", "2002", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit"],
      ["USR-013", "inka", "Inka", "SERVER", 1, "SERVER", "Tetap", "3001", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit"],
      ["USR-014", "zidan", "Zidan", "SERVER", 1, "SERVER", "Tetap", "3002", "Aktif", "TIDAK", "kalender:edit,checklist:edit,stok:edit"]
    ];
    shUsers.getRange(2, 1, userSeed.length, userHeaders.length).setValues(userSeed);
  }
  // V24.3: PIN boy TIDAK lagi dipaksa '0000'. Dulu baris ini mereset PIN
  // RAJA_UTAMA setiap repair dijalankan — rotasi PIN akan hilang diam-diam.
  // Sekarang hanya format kolomnya yang dipastikan teks.
  shUsers.getRange(2, 8).setNumberFormat("@");

  // 3. MENU (11 KOLOM)
  const shMenu = getOrCreateSheet(ss, "MENU");
  const menuHeaders = ["menu_id", "jenis", "divisi", "kategori", "nama_produk", "hpp", "harga_jual", "margin_rp", "margin_pct", "food_cost_pct", "status_data"];
  const extraMenu = cleanAndFormatSheet(shMenu, menuHeaders);
  if (extraMenu) extraNotes.push(extraMenu);
  if (shMenu.getLastRow() <= 1) {
    const rawMenu = [
      ["MNU-001", "Minuman", "BARISTA", "Black", "Americano Hot", 4694, 17000],
      ["MNU-002", "Minuman", "BARISTA", "Black", "Americano Ice", 5914, 18000],
      ["MNU-003", "Minuman", "BARISTA", "White", "Cafe Latte Ice", 10320, 25000],
      ["MNU-004", "Minuman", "BARISTA", "Signature", "Es Kopi Susu Setempat", 9863, 24000],
      ["MNU-005", "Makanan", "KITCHEN", "Snack", "French Fries", 7101, 18000],
      ["MNU-006", "Makanan", "KITCHEN", "Snack", "Tahu Konslet Hollo", 6560, 20000],
      ["MNU-007", "Makanan", "KITCHEN", "Main Course", "Ultimate Crazy Rice Hollo", 13635, 28000],
      ["MNU-008", "Makanan", "KITCHEN", "Pasta", "Hollo Pasta Bolognese", 13309, 35000]
    ];
    for (let i = 0; i < rawMenu.length; i++) {
      let r = i + 2;
      shMenu.getRange(r, 1, 1, 7).setValues([rawMenu[i]]);
      shMenu.getRange(r, 8).setFormula(`=IF(AND(ISNUMBER(G${r}),ISNUMBER(F${r})), G${r}-F${r}, 0)`);
      shMenu.getRange(r, 9).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), H${r}/G${r}, 0)`);
      shMenu.getRange(r, 10).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), F${r}/G${r}, 0)`);
      shMenu.getRange(r, 11).setValue("OK");
    }
  } else {
    // V24: Formula hanya dipasang pada baris yang menu_id-nya tidak kosong.
    for (let r = 2; r <= shMenu.getLastRow(); r++) {
      if (String(shMenu.getRange(r, 1).getValue()).trim() === "") continue;
      shMenu.getRange(r, 8).setFormula(`=IF(AND(ISNUMBER(G${r}),ISNUMBER(F${r})), G${r}-F${r}, 0)`);
      shMenu.getRange(r, 9).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), H${r}/G${r}, 0)`);
      shMenu.getRange(r, 10).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), F${r}/G${r}, 0)`);
    }
  }
  shMenu.getRange(2, 6, Math.max(shMenu.getLastRow() - 1, 1), 3).setNumberFormat("Rp #,##0");
  shMenu.getRange(2, 9, Math.max(shMenu.getLastRow() - 1, 1), 2).setNumberFormat("0.0%");

  // 4. STOK (PERSIS 21 KOLOM STANDAR WEB APP)
  const shStok = getOrCreateSheet(ss, "STOK");
  const stokHeaders = [
    "item_id", "divisi", "kategori", "nama_barang", "satuan_makro", "isi_konversi", "satuan_mikro",
    "min_stok", "stok_gudang", "tersedia_pagi", "refill_s1", "refill_s2",
    "opname_s1", "opname_s2", "waste_mikro",
    "stok_mikro", "pemakaian_hari_ini", "total_makro", "status", "tgl_update", "jam_update"
  ];
  const extraStok = cleanAndFormatSheet(shStok, stokHeaders);
  if (extraStok) extraNotes.push(extraStok);
  if (shStok.getLastRow() <= 1) {
    const rawStok = [
      ["BAR-001", "BARISTA", "Biji Kopi", "BBB-Beans Dark Roast", "pack", 1000, "gram", 1000, 2, 500, 500, 0, 550, 250, 50, "", "", "", "", "2026-11-09", "08:30"],
      ["BAR-002", "BARISTA", "Susu & Krimer", "Fresh Milk Greenfield", "Kotak", 950, "Mililiter", 950, 2, 400, 550, 0, 500, 300, 0, "", "", "", "", "2026-11-09", "08:30"],
      ["BAR-003", "BARISTA", "Biji Kopi", "BBB-Reguler", "pack", 1000, "gram", 1000, 2, 300, 0, 0, 250, 200, 0, "", "", "", "", "2026-11-09", "08:30"],
      ["KIT-001", "KITCHEN", "Protein & Frozen", "Chicken Shumai Frozen", "pack", 50, "pcs", 50, 2, 25, 25, 0, 20, 15, 0, "", "", "", "", "2026-11-09", "08:30"],
      ["KIT-002", "KITCHEN", "Karbo & Roti", "Kentang Shoestring", "pack", 2500, "gram", 7500, 5, 2500, 2500, 0, 2200, 1500, 200, "", "", "", "", "2026-11-09", "08:30"],
      ["SRV-001", "SERVER", "Tisu & Kertas", "Tisu Meja Dining", "dus", 30, "pack", 10, 1, 9, 0, 0, 4, 5, 0, "", "", "", "", "2026-11-09", "08:30"]
    ];
    shStok.getRange(2, 1, rawStok.length, 21).setValues(rawStok);
  }
  // V24: Formula hanya dipasang pada baris yang item_id-nya tidak kosong.
  for (let r = 2; r <= shStok.getLastRow(); r++) {
    if (String(shStok.getRange(r, 1).getValue()).trim() === "") continue;
    shStok.getRange(r, 16).setFormula(`=IF(ISNUMBER(N${r}),N${r},IF(ISNUMBER(M${r}),M${r},MAX(0,J${r}+K${r}+L${r}-O${r})))`);
    shStok.getRange(r, 17).setFormula(`=MAX(0,(J${r}+K${r}+L${r})-P${r})`);
    shStok.getRange(r, 18).setFormula(`=(I${r}*F${r})+P${r}`);
    shStok.getRange(r, 19).setFormula(`=IF(R${r}<=H${r},"RESTOCK","AMAN")`);
  }
  shStok.getRange(2, 6, Math.max(shStok.getLastRow() - 1, 1), 14).setNumberFormat("#,##0");

  // 5. CHECKLIST (10 KOLOM)
  const shChk = getOrCreateSheet(ss, "CHECKLIST");
  const extraChk = cleanAndFormatSheet(shChk, ["task_id", "tanggal", "shift", "station", "fase", "jam_target", "tugas", "status", "petugas", "jam_selesai"]);
  if (extraChk) extraNotes.push(extraChk);
  if (shChk.getLastRow() <= 1) {
    const chkSeed = [
      ["CHK-001", "2026-11-09", "1", "BARISTA & KASIR", "Prep", "08:30", "Nyalakan mesin espresso & grinder, cek air RO", "Belum", "pace", ""],
      ["CHK-002", "2026-11-09", "1", "BARISTA & KASIR", "Prep", "08:35", "Dial-in espresso, kalibrasi rasa & timbangan", "Belum", "el", ""],
      ["CHK-003", "2026-11-09", "1", "BARISTA & KASIR", "Prep", "08:30", "Hitung modal kas kecil & kertas thermal EDC", "Belum", "el", ""],
      ["CHK-004", "2026-11-09", "1", "KITCHEN", "Prep", "08:30", "Cek safety kompor gas & defrost ayam", "Belum", "eko", ""],
      ["CHK-005", "2026-11-09", "1", "SERVER", "Prep", "08:30", "Sanitasi meja kursi, AC & playlist kafe", "Belum", "inka", ""]
    ];
    shChk.getRange(2, 1, chkSeed.length, 10).setValues(chkSeed);
  }

  // 6. JADWAL (10 KOLOM)
  const shJdw = getOrCreateSheet(ss, "JADWAL");
  const extraJdw = cleanAndFormatSheet(shJdw, ["id", "tanggal", "shift", "username", "nama", "divisi", "status_hadir", "jam_masuk", "jam_pulang", "catatan"]);
  if (extraJdw) extraNotes.push(extraJdw);
  if (shJdw.getLastRow() <= 1) {
    const jdwSeed = [
      ["JDW-001", "2026-11-09", "1", "pace", "Pace", "BARISTA", "Hadir", "08:30", "16:30", "Opening Barista"],
      ["JDW-002", "2026-11-09", "1", "el", "El", "BARISTA", "Hadir", "08:30", "16:30", "Barista & Kasir S1"],
      ["JDW-003", "2026-11-09", "2", "fatir", "Fatir", "BARISTA", "Hadir", "16:30", "00:00", "Closing Shift 2"]
    ];
    shJdw.getRange(2, 1, jdwSeed.length, 10).setValues(jdwSeed);
  }

  // 7. LOG_KALENDER (13 KOLOM)
  // V24: Kolom "shift" DITAMBAHKAN DI AKHIR (kolom ke-13, setelah
  // "tindak_lanjut") agar kolom lama tidak bergeser sama sekali.
  const shLog = getOrCreateSheet(ss, "LOG_KALENDER");
  const extraLog = cleanAndFormatSheet(shLog, ["log_id", "tanggal", "jam", "divisi", "kategori", "item_terkait", "item_id", "jumlah_angka", "satuan", "detail", "pic", "tindak_lanjut", "shift"]);
  if (extraLog) extraNotes.push(extraLog);

  // 8. SERAH_TERIMA (19 KOLOM)
  const shSt = getOrCreateSheet(ss, "SERAH_TERIMA");
  const stHeaders = [
    "id", "tanggal", "jam", "tahap", "shift", "petugas_serah", "petugas_terima",
    "penjualan_tunai", "non_tunai_qris", "online_food", "total_omset_shift",
    "kas_modal_stay", "pengeluaran_kas_kecil", "keterangan_pengeluaran", "uang_setor_bersih",
    "prep_barista", "prep_kitchen", "prep_server", "catatan_evaluasi"
  ];
  const extraSt = cleanAndFormatSheet(shSt, stHeaders);
  if (extraSt) extraNotes.push(extraSt);
  if (shSt.getLastRow() > 1) {
    for (let r = 2; r <= shSt.getLastRow(); r++) {
      // V24: Baris ber-ID dipasangi formula; baris TANPA ID yang terlanjur
      // berisi formula hitungan cukup dikosongkan sel hitungannya saja
      // (barisnya TIDAK dihapus) agar tidak terbaca sebagai data hantu.
      if (String(shSt.getRange(r, 1).getValue()).trim() === "") {
        shSt.getRange(r, 11).setValue("");
        shSt.getRange(r, 15).setValue("");
        continue;
      }
      shSt.getRange(r, 11).setFormula(`=H${r}+I${r}+J${r}`);
      shSt.getRange(r, 15).setFormula(`=H${r}-M${r}`);
    }
    shSt.getRange(2, 8, Math.max(shSt.getLastRow() - 1, 1), 8).setNumberFormat("Rp #,##0");
  }

  // 9. DAILY_REPORT (26 KOLOM — V24.6: +tim_hadir, tim_libur, status_stok, isu_harian)
  const shDr = getOrCreateSheet(ss, "DAILY_REPORT");
  const drHeaders = [
    "report_id", "tanggal", "rev_shift1", "rev_shift2", "total_omset",
    "qris_edc", "online_food", "target_harian", "capaian_pct",
    "total_kas_keluar", "rincian_pengeluaran",
    "kalibrasi_beans", "waste_bahan", "barang_masuk", "insiden_alat",
    "belanja_barista", "belanja_kitchen", "belanja_server",
    "crew_meals", "komplimen_barter", "evaluasi_harian", "pelapor",
    "tim_hadir", "tim_libur", "status_stok", "isu_harian"
  ];
  const extraDr = cleanAndFormatSheet(shDr, drHeaders);
  if (extraDr) extraNotes.push(extraDr);
  if (shDr.getLastRow() > 1) {
    for (let r = 2; r <= shDr.getLastRow(); r++) {
      // V24: Sama seperti SERAH_TERIMA — formula hanya untuk baris ber-ID;
      // sel hitungan pada baris tanpa ID dikosongkan, baris tidak dihapus.
      if (String(shDr.getRange(r, 1).getValue()).trim() === "") {
        shDr.getRange(r, 5).setValue("");
        shDr.getRange(r, 9).setValue("");
        continue;
      }
      shDr.getRange(r, 5).setFormula(`=C${r}+D${r}`);
      shDr.getRange(r, 9).setFormula(`=IF(AND(ISNUMBER(H${r}),H${r}>0), E${r}/H${r}, 0)`);
    }
    shDr.getRange(2, 3, Math.max(shDr.getLastRow() - 1, 1), 6).setNumberFormat("Rp #,##0");
    shDr.getRange(2, 9, Math.max(shDr.getLastRow() - 1, 1), 1).setNumberFormat("0.0%");
    shDr.getRange(2, 10, Math.max(shDr.getLastRow() - 1, 1), 1).setNumberFormat("Rp #,##0");
  }

  // 10. DASHBOARD UTAMA (REBUILD LENGKAP DENGAN 2 GRAFIK NATIVE GOOGLE SHEETS)
  const shDash = getOrCreateSheet(ss, "DASHBOARD");
  shDash.clear();
  shDash.setHiddenGridlines(false);
  const oldCharts = shDash.getCharts();
  for (let i = 0; i < oldCharts.length; i++) shDash.removeChart(oldCharts[i]);

  shDash.getRange("A1:K1").merge().setValue("HOLLO COFFEE & ROASTERY'S — DASHBOARD OPERASIONAL")
    .setBackground("#09090b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(13).setVerticalAlignment("middle");
  shDash.setRowHeight(1, 45);

  shDash.getRange("A3").setValue("TANGGAL:").setFontWeight("bold").setFontColor("#71717a");
  shDash.getRange("B3").setValue("=TEXT(TODAY(),\"dddd, dd mmm yyyy\")").setFontWeight("bold");
  shDash.getRange("D3").setValue("JAM LIVE:").setFontWeight("bold").setFontColor("#71717a");
  shDash.getRange("E3").setValue("=TEXT(NOW(),\"HH:mm:ss\")").setFontWeight("bold");
  shDash.getRange("G3").setValue("FASE SHIFT:").setFontWeight("bold").setFontColor("#71717a");
  // V24: Label fase diperbaiki — sebelum 08:30 "TUTUP", 08:30–16:30 "SHIFT 1", sesudahnya "SHIFT 2".
  shDash.getRange("H3").setValue("=IF(NOW()-TODAY()<TIME(8,30,0),\"TUTUP\",IF(NOW()-TODAY()<TIME(16,30,0),\"SHIFT 1\",\"SHIFT 2\"))").setFontWeight("bold");

  const kpiHeaders = [["CAPAIAN TARGET MTD", "PENJUALAN HARI INI (OMSET)", "STOK PERLU RESTOCK", "CHECKLIST SELESAI"]];
  const kpiValues = [[
    // V24: MTD hanya menjumlah DAILY_REPORT pada bulan berjalan (via TEXT bulan), bukan seluruh data.
    '=SUMPRODUCT((TEXT(DAILY_REPORT!B2:B1000,"yyyy-mm")=TEXT(TODAY(),"yyyy-mm"))*IFERROR(DAILY_REPORT!E2:E1000,0))',
    // V24: Omset hari ini memakai SATU kriteria teks tanggal (SUMPRODUCT) agar tidak dobel hitung.
    '=SUMPRODUCT((TEXT(SERAH_TERIMA!B2:B1000,"yyyy-mm-dd")=TEXT(TODAY(),"yyyy-mm-dd"))*IFERROR(SERAH_TERIMA!K2:K1000,0))',
    '=COUNTIF(STOK!S:S, "RESTOCK")',
    '=IFERROR(TEXT(COUNTIF(CHECKLIST!H:H,"Selesai")/COUNTA(CHECKLIST!A2:A),"0%"), "0%")'
  ]];
  const kpiSubs = [["Target: Rp 150.000.000", "Tunai + QRIS/EDC + Online", "Batas minimum", "Seluruh Stasiun Kerja"]];

  shDash.getRange("A5:D5").setValues(kpiHeaders).setBackground("#18181b").setFontColor("#a1a1aa").setFontSize(8).setFontWeight("bold").setHorizontalAlignment("center");
  shDash.getRange("A6:D6").setValues(kpiValues).setBackground("#f4f4f5").setFontColor("#09090b").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  shDash.getRange("A7:D7").setValues(kpiSubs).setBackground("#f4f4f5").setFontColor("#71717a").setFontSize(7).setHorizontalAlignment("center");
  shDash.getRange("A5:D7").setBorder(true, true, true, true, true, true, "#e4e4e7", SpreadsheetApp.BorderStyle.SOLID);
  shDash.getRange("A6:B6").setNumberFormat("Rp #,##0");

  shDash.getRange("A9:D9").merge().setValue("1. BREAKDOWN PENJUALAN").setBackground("#18181b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(9);
  const salesRows = [
    ["SALURAN", "NOMINAL", "STATUS", "KETERANGAN"],
    ["Tunai", '=IFERROR(SUM(SERAH_TERIMA!H2:H), 0)', "Laci", "Kas Fisik"],
    ["QRIS / EDC", '=IFERROR(SUM(SERAH_TERIMA!I2:I), 0)', "Bank", "Rekening"],
    ["Online Food", '=IFERROR(SUM(SERAH_TERIMA!J2:J), 0)', "Portal", "Grab/Gojek/Shopee"]
  ];
  shDash.getRange("A10:D13").setValues(salesRows).setFontSize(8);
  shDash.getRange("B11:B13").setNumberFormat("Rp #,##0");
  shDash.getRange("A10:D13").setBorder(true, true, true, true, true, true, "#e4e4e7", SpreadsheetApp.BorderStyle.SOLID);
  shDash.getRange("A10:D10").setBackground("#f4f4f5").setFontWeight("bold");

  shDash.getRange("F9:I9").merge().setValue("2. KESEHATAN STOK PER DIVISI").setBackground("#18181b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(9);
  const stokRows = [
    ["DIVISI", "AMAN", "RESTOCK", "TOTAL"],
    ["BARISTA", '=COUNTIFS(STOK!B2:B,"BARISTA",STOK!S2:S,"AMAN")', '=COUNTIFS(STOK!B2:B,"BARISTA",STOK!S2:S,"RESTOCK")', '=COUNTIF(STOK!B2:B,"BARISTA")'],
    ["KITCHEN", '=COUNTIFS(STOK!B2:B,"KITCHEN",STOK!S2:S,"AMAN")', '=COUNTIFS(STOK!B2:B,"KITCHEN",STOK!S2:S,"RESTOCK")', '=COUNTIF(STOK!B2:B,"KITCHEN")'],
    ["SERVER",  '=COUNTIFS(STOK!B2:B,"SERVER",STOK!S2:S,"AMAN")',  '=COUNTIFS(STOK!B2:B,"SERVER",STOK!S2:S,"RESTOCK")',  '=COUNTIF(STOK!B2:B,"SERVER")']
  ];
  shDash.getRange("F10:I13").setValues(stokRows).setFontSize(8);
  shDash.getRange("F10:I13").setBorder(true, true, true, true, true, true, "#e4e4e7", SpreadsheetApp.BorderStyle.SOLID);
  shDash.getRange("F10:I10").setBackground("#f4f4f5").setFontWeight("bold");

  try {
    const chartSales = shDash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(shDash.getRange("A10:B13"))
      .setPosition(15, 1, 0, 0)
      .setOption('title', 'Saluran Penjualan')
      .setOption('colors', ['#18181b'])
      .setOption('legend', { position: 'none' })
      .build();
    shDash.insertChart(chartSales);

    const chartStok = shDash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(shDash.getRange("F10:H13"))
      .setPosition(15, 6, 0, 0)
      .setOption('title', 'Kesehatan Stok per Divisi (Aman vs Restock)')
      .setOption('colors', ['#10b981', '#ef4444'])
      .setOption('legend', { position: 'top' })
      .build();
    shDash.insertChart(chartStok);
  } catch(e) {}

  for (let c = 1; c <= 10; c++) shDash.autoResizeColumn(c);

  SpreadsheetApp.flush();
  // V24: Kolom ekstra tidak dihapus; cukup dilaporkan di pesan hasil.
  let msg = "DATABASE V24 BERHASIL DISINKRONKAN (DATA-SAFE): Tidak ada kolom/baris/sheet yang dihapus, formula hanya dipasang pada baris ber-ID, dan grafik Dashboard aktif!";
  if (extraNotes.length) msg += " Catatan kolom ekstra yang DIBIARKAN apa adanya: " + extraNotes.join(" | ");
  return msg;
}

// V24: PEMBERSIH AMAN — hanya menulis header, freeze, dan memastikan jumlah
// kolom minimal. Kolom ekstra di luar target TIDAK dihapus; fungsi ini
// mengembalikan catatan teks bila ada kolom ekstra (untuk dilaporkan).
function cleanAndFormatSheet(sheet, targetHeaders) {
  sheet.setHiddenGridlines(false);
  sheet.setFrozenRows(1);

  const maxCols = sheet.getMaxColumns();
  let extraNote = "";
  if (maxCols < targetHeaders.length) {
    sheet.insertColumnsAfter(maxCols, targetHeaders.length - maxCols);
  } else if (maxCols > targetHeaders.length) {
    extraNote = sheet.getName() + " punya " + (maxCols - targetHeaders.length) + " kolom ekstra di luar target (" + maxCols + " > " + targetHeaders.length + "), dibiarkan utuh";
  }

  // Tulis Header Bersih
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders])
    .setBackground("#0d0d11")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setFontSize(9)
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  for (let c = 1; c <= targetHeaders.length; c++) {
    sheet.autoResizeColumn(c);
    if (sheet.getColumnWidth(c) < 110) sheet.setColumnWidth(c, 110);
  }
  return extraNote;
}

function getOrCreateSheet(ss, name) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

// V24: Tulis ulang formula kolom hitungan untuk SATU baris. Dipakai jalur
// insert dan jalur update, dengan rumus yang sama persis seperti V23,
// supaya edit dari web tidak lagi menimpa formula dengan angka mentah.
function applyRowFormulas_(sheetName, sh, r) {
  if (sheetName === "STOK") {
    sh.getRange(r, 16).setFormula(`=IF(ISNUMBER(N${r}),N${r},IF(ISNUMBER(M${r}),M${r},MAX(0,J${r}+K${r}+L${r}-O${r})))`);
    sh.getRange(r, 17).setFormula(`=MAX(0,(J${r}+K${r}+L${r})-P${r})`);
    sh.getRange(r, 18).setFormula(`=(I${r}*F${r})+P${r}`);
    sh.getRange(r, 19).setFormula(`=IF(R${r}<=H${r},"RESTOCK","AMAN")`);
  }
  if (sheetName === "MENU") {
    const headers = sh.getDataRange().getValues()[0];
    sh.getRange(r, headers.indexOf("margin_rp") + 1).setFormula(`=IF(AND(ISNUMBER(G${r}),ISNUMBER(F${r})), G${r}-F${r}, 0)`);
    sh.getRange(r, headers.indexOf("margin_pct") + 1).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), H${r}/G${r}, 0)`);
    sh.getRange(r, headers.indexOf("food_cost_pct") + 1).setFormula(`=IF(AND(ISNUMBER(G${r}),G${r}>0), F${r}/G${r}, 0)`);
  }
  if (sheetName === "SERAH_TERIMA") {
    sh.getRange(r, 11).setFormula(`=H${r}+I${r}+J${r}`);
    sh.getRange(r, 15).setFormula(`=H${r}-M${r}`);
  }
  if (sheetName === "DAILY_REPORT") {
    sh.getRange(r, 5).setFormula(`=C${r}+D${r}`);
    sh.getRange(r, 9).setFormula(`=IF(AND(ISNUMBER(H${r}),H${r}>0), E${r}/H${r}, 0)`);
  }
}

// V24: Terapkan (sign=+1) atau balikkan (sign=-1) efek satu log kalender
// ke stok. Dipakai bersama oleh insert, update, dan delete LOG_KALENDER.
function applyLogStockEffect_(ss, logObj, sign) {
  if (!logObj) return;
  const itemId = logObj.item_id;
  const kat = String(logObj.kategori || "");
  const qty = Number(logObj.jumlah_angka) || 0;
  if (!itemId || qty <= 0) return;

  const shStok = ss.getSheetByName("STOK");
  if (!shStok) return; // V24: Cek null sheet STOK (sebelumnya bisa crash).
  const stokData = shStok.getDataRange().getValues();
  const sHeaders = stokData[0];
  const sItemIdx = sHeaders.indexOf("item_id");
  const sWasteIdx = sHeaders.indexOf("waste_mikro");
  const sGudangIdx = sHeaders.indexOf("stok_gudang");
  const sRefillIdx = sHeaders.indexOf(String(logObj.shift || "1") === "2" ? "refill_s2" : "refill_s1"); // V24.2: refill masuk kolom shift-nya
  const sKonvIdx = sHeaders.indexOf("isi_konversi");
  if (sItemIdx === -1) return;

  for (let sr = 1; sr < stokData.length; sr++) {
    if (String(stokData[sr][sItemIdx]) !== String(itemId)) continue;
    const konv = Number(stokData[sr][sKonvIdx]) || 1;
    if (kat.includes("Waste") || kat.includes("Spoil") || kat.includes("Kalibrasi")) {
      let cur = Number(stokData[sr][sWasteIdx]) || 0;
      shStok.getRange(sr + 1, sWasteIdx + 1).setValue(Math.max(0, cur + (sign * qty)));
    } else if (kat.includes("Barang Masuk")) {
      let cur = Number(stokData[sr][sGudangIdx]) || 0;
      shStok.getRange(sr + 1, sGudangIdx + 1).setValue(Math.max(0, cur + (sign * qty)));
    } else if (kat.includes("Bongkar") || kat.includes("Refill")) {
      let curG = Number(stokData[sr][sGudangIdx]) || 0;
      let curR = Number(stokData[sr][sRefillIdx]) || 0;
      shStok.getRange(sr + 1, sGudangIdx + 1).setValue(Math.max(0, curG - (sign * qty)));
      shStok.getRange(sr + 1, sRefillIdx + 1).setValue(Math.max(0, curR + (sign * (qty * konv))));
    }
    break;
  }
}

function doGet(e) {
  const p = e.parameter || {};
  const action = p.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (action === "ping") {
      return jsonResp({
        success: true,
        message: "Koneksi Google Sheets Aktif!",
        sheet_name: ss.getName(),
        sheet_url: ss.getUrl(),
        sheet_id: ss.getId()
      });
    }

    if (action === "repair_db") {
      // V24: repair_db wajib token milik user level 6 / RAJA_UTAMA.
      const authUser = requireAuth_(p.token);
      if (!authUser) return jsonResp({ success: false, message: "Sesi tidak valid, silakan login ulang" });
      if (Number(authUser.level) !== 6 && String(authUser.role) !== "RAJA_UTAMA") {
        return jsonResp({ success: false, message: "Repair database hanya untuk RAJA_UTAMA" });
      }
      const msg = repairAndSyncDatabase();
      return jsonResp({ success: true, message: msg });
    }

    if (action === "verify_pin") {
      const u = String(p.username || "").toLowerCase().trim();
      const pin = String(p.pin || "").trim();
      const users = getSheetRows(ss.getSheetByName("USERS"));
      const user = users.find(x => x.username.toLowerCase() === u && String(x.pin) === pin);
      if (user) {
        // V24: Login sukses menerbitkan token sesi untuk read_all & doPost.
        const token = issueToken_(user.username);
        return jsonResp({
          success: true,
          token: token,
          user: {
            user_id: user.user_id,
            username: user.username,
            nama_lengkap: user.nama_lengkap,
            role: user.role,
            level: Number(user.level),
            divisi: user.divisi,
            akses_perusahaan: user.akses_perusahaan,
            "akses_tabs": user["akses_tabs"] || ""
          }
        });
      }
      return jsonResp({ success: false, message: "Username atau PIN salah!" });
    }

    if (action === "read_all") {
      // V24: read_all wajib token valid; tanpa token data tidak dibagikan.
      const authUser = requireAuth_(p.token);
      if (!authUser) return jsonResp({ success: false, message: "Sesi tidak valid, silakan login ulang" });
      const sheets = ['DATA_PERUSAHAAN', 'USERS', 'JADWAL', 'LOG_KALENDER', 'MENU', 'STOK', 'CHECKLIST', 'SERAH_TERIMA', 'DAILY_REPORT'];
      let bundle = {};
      sheets.forEach(name => {
        const sh = ss.getSheetByName(name);
        let rows = sh ? getSheetRows(sh) : [];
        // V24: Kolom pin TIDAK PERNAH dikirim ke frontend.
        if (name === 'USERS') rows.forEach(r => { delete r.pin; });
        bundle[name] = rows;
      });
      return jsonResp({
        success: true,
        meta: { sheet_name: ss.getName(), sheet_url: ss.getUrl(), sheet_id: ss.getId() },
        data: bundle
      });
    }

    return jsonResp({ success: false, message: "Aksi GET tidak dikenal" });
  } catch (err) {
    return jsonResp({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // V24: Lock diambil di dalam try; releaseLock hanya bila berhasil diambil.
  let lockAcquired = false;

  try {
    lock.waitLock(15000);
    lockAcquired = true;

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // V24: Semua aksi POST wajib token valid.
    const authUser = requireAuth_(payload.auth_token);
    if (!authUser) return jsonResp({ success: false, message: "Sesi tidak valid, silakan login ulang" });

    const sh = ss.getSheetByName(payload.sheet);
    if (!sh) throw new Error("Sheet tidak ditemukan: " + payload.sheet);

    // V24: Otorisasi tulis per tab berdasarkan "akses_tabs" user.
    if (!canWriteSheet_(authUser, payload.sheet)) {
      return jsonResp({ success: false, message: "Anda tidak punya hak edit untuk data ini" });
    }

    let data = sh.getDataRange().getValues();
    let headers = data[0];

    if (action === "update") {
      const keyIdx = headers.indexOf(payload.key);
      if (keyIdx === -1) throw new Error("Key column tidak ditemukan");

      for (let r = 1; r < data.length; r++) {
        if (String(data[r][keyIdx]).trim() === String(payload.val).trim()) {
          // V24: Untuk LOG_KALENDER, simpan log lama dulu agar efek stoknya
          // bisa dibalikkan sebelum efek log baru diterapkan.
          let oldLog = null;
          if (payload.sheet === "LOG_KALENDER") {
            oldLog = {};
            headers.forEach((h, i) => { oldLog[h] = data[r][i]; });
          }
          for (let col in payload.updates) {
            let cIdx = headers.indexOf(col);
            if (cIdx !== -1) sh.getRange(r + 1, cIdx + 1).setValue(payload.updates[col]);
          }
          // V24: Tulis ulang formula kolom hitungan baris ini, menimpa
          // angka mentah kiriman frontend (STOK/MENU/SERAH_TERIMA/DAILY_REPORT).
          applyRowFormulas_(payload.sheet, sh, r + 1);
          SpreadsheetApp.flush();
          if (payload.sheet === "LOG_KALENDER") {
            const newLog = {};
            headers.forEach(h => { newLog[h] = oldLog[h]; });
            for (let col2 in payload.updates) newLog[col2] = payload.updates[col2];
            applyLogStockEffect_(ss, oldLog, -1);
            applyLogStockEffect_(ss, newLog, +1);
          }
          return jsonResp({ success: true, message: "Update berhasil" });
        }
      }
      throw new Error("Item tidak ditemukan untuk update");
    }

    if (action === "insert") {
      // V24.7: Tolak insert bila ID sudah ada — cegah baris duplikat dari klik ganda / kirim ulang.
      const KEYMAP_ = { USERS: "user_id", MENU: "menu_id", STOK: "item_id", JADWAL: "id", CHECKLIST: "task_id", LOG_KALENDER: "log_id", SERAH_TERIMA: "id", DAILY_REPORT: "report_id" };
      const keyCol_ = KEYMAP_[payload.sheet];
      if (keyCol_ && payload.data && payload.data[keyCol_] !== undefined && String(payload.data[keyCol_]).trim() !== "") {
        const kIdx_ = headers.indexOf(keyCol_);
        if (kIdx_ >= 0) {
          for (let r = 1; r < data.length; r++) {
            if (String(data[r][kIdx_]).trim() === String(payload.data[keyCol_]).trim()) {
              return jsonResp({ success: false, message: "ID sudah dipakai: " + payload.data[keyCol_] + " (baris duplikat ditolak)" });
            }
          }
        }
      }
      let row = [];
      const newRowNum = sh.getLastRow() + 1;

      headers.forEach(h => {
        if (payload.sheet === "STOK") {
          if (h === "stok_mikro") { row.push(`=IF(ISNUMBER(N${newRowNum}),N${newRowNum},IF(ISNUMBER(M${newRowNum}),M${newRowNum},MAX(0,J${newRowNum}+K${newRowNum}+L${newRowNum}-O${newRowNum})))`); return; }
          if (h === "pemakaian_hari_ini") { row.push(`=MAX(0,(J${newRowNum}+K${newRowNum}+L${newRowNum})-P${newRowNum})`); return; }
          if (h === "total_makro") { row.push(`=(I${newRowNum}*F${newRowNum})+P${newRowNum}`); return; }
          if (h === "status") { row.push(`=IF(R${newRowNum}<=H${newRowNum},"RESTOCK","AMAN")`); return; }
        }
        if (payload.sheet === "MENU") {
          if (h === "margin_rp") { row.push(`=IF(AND(ISNUMBER(G${newRowNum}),ISNUMBER(F${newRowNum})), G${newRowNum}-F${newRowNum}, 0)`); return; }
          if (h === "margin_pct") { row.push(`=IF(AND(ISNUMBER(G${newRowNum}),G${newRowNum}>0), H${newRowNum}/G${newRowNum}, 0)`); return; }
          if (h === "food_cost_pct") { row.push(`=IF(AND(ISNUMBER(G${newRowNum}),G${newRowNum}>0), F${newRowNum}/G${newRowNum}, 0)`); return; }
        }
        if (payload.sheet === "SERAH_TERIMA") {
          if (h === "total_omset_shift") { row.push(`=H${newRowNum}+I${newRowNum}+J${newRowNum}`); return; }
          if (h === "uang_setor_bersih") { row.push(`=H${newRowNum}-M${newRowNum}`); return; }
        }
        if (payload.sheet === "DAILY_REPORT") {
          if (h === "total_omset") { row.push(`=C${newRowNum}+D${newRowNum}`); return; }
          if (h === "capaian_pct") { row.push(`=IF(AND(ISNUMBER(H${newRowNum}),H${newRowNum}>0), E${newRowNum}/H${newRowNum}, 0)`); return; }
        }
        row.push(payload.data[h] !== undefined ? payload.data[h] : "");
      });

      sh.appendRow(row);
      SpreadsheetApp.flush();
      // V24: Efek stok log baru diterapkan lewat helper bersama (insert).
      if (payload.sheet === "LOG_KALENDER") applyLogStockEffect_(ss, payload.data, +1);
      return jsonResp({ success: true, message: "Data berhasil ditambahkan" });
    }

    if (action === "delete") {
      const keyIdx = headers.indexOf(payload.key);
      // V24: Validasi keyIdx pada delete (sebelumnya tidak dicek).
      if (keyIdx === -1) throw new Error("Key column tidak ditemukan: " + payload.key);
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][keyIdx]).trim() === String(payload.val).trim()) {
          // V24: Simpan log yang akan dihapus agar efek stoknya dibalikkan.
          let deletedLog = null;
          if (payload.sheet === "LOG_KALENDER") {
            deletedLog = {};
            headers.forEach((h, i) => { deletedLog[h] = data[r][i]; });
          }
          sh.deleteRow(r + 1);
          SpreadsheetApp.flush();
          if (payload.sheet === "LOG_KALENDER") applyLogStockEffect_(ss, deletedLog, -1);
          return jsonResp({ success: true, message: "Data berhasil dihapus" });
        }
      }
      throw new Error("Baris tidak ditemukan");
    }

    return jsonResp({ success: false, message: "Aksi tidak dikenali" });
  } catch (err) {
    return jsonResp({ success: false, error: err.toString() });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// FUNGSI PEMBERSIH BARIS KOSONG: Hanya mengembalikan baris yang benar-benar memiliki ID atau Nama
function getSheetRows(sh) {
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getDisplayValues();
  const headers = data[0].map(h => String(h).trim());

  return data.slice(1).filter(r => {
    // Abaikan jika baris kosong total
    return r.some(cell => String(cell).trim() !== "");
  }).map(r => {
    let o = {};
    headers.forEach((h, i) => {
      if (h) o[h] = (r[i] !== undefined) ? String(r[i]).trim() : "";
    });
    return o;
  });
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
