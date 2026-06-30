/**
 * ============================================================
 * TJAP BOENGA - STOCK OPNAME SYSTEM
 * Code.gs — BASE ORIGINAL + FITUR KALIBRASI/JATAH/WASTE
 * ============================================================
 *
 * RINGKASAN PERUBAHAN VERSI INI:
 * Base tetap 100% logic original Anda (validasi strict, 4-tier
 * Keterangan, dst — TIDAK diubah). Ditambahkan murni:
 *
 * 1. CONFIG_ADJ — konstanta terpisah untuk tabel KALIBRASI/JATAH/
 *    WASTE di "SO daily input" (No=E, Nama Barang=G, Type=J,
 *    Qty=N, Keterangan=O; header row 23, data row 24-34, 11 baris).
 * 2. CONFIG_ADJ_LOG — konstanta untuk Table2 di "SO daily log"
 *    (Waktu Submit=P, Tanggal SO=Q, No=R, Nama Barang=S, Type=T,
 *    Qty=U, Keterangan=V; data mulai row 2).
 * 3. readAdjustmentRows() — baca baris adjustment yang TERISI saja
 *    (skip baris kosong), validasi ringan (Nama Barang wajib kalau
 *    Type/Qty diisi, dan sebaliknya — supaya tidak ada baris "separuh
 *    isi" yang nyangkut).
 * 4. submitSO() — DIPERLUAS (bukan diganti total): setelah submit
 *    tabel utama sukses, lanjut insert baris adjustment ke Table2,
 *    lalu clear tabel adjustment di input. Semua dalam SATU alur
 *    submit & SATU dialog konfirmasi (adjustment ikut ditampilkan
 *    di summary sebelum user klik YES).
 * 5. clearAdjustmentTable() — reset 11 baris adjustment (No tetap
 *    1-11, kolom lain di-clear) setelah submit berhasil.
 *
 * TIDAK DIUBAH dari original: validateRows, buildWarningSummary,
 * generateWeeklyReport, onEdit, resetSistem (logic intinya sama
 * persis seperti yang Anda kirim). resetSistem ditambah 1 baris
 * untuk ikut clear tabel adjustment, supaya konsisten.
 *
 * CATATAN UNTUK PENGEMBANGAN WEEKLY/MONTHLY REPORT (sesuai poin 7
 * permintaan Anda): nilai adjustment SUDAH tersimpan rapi di Table2
 * dengan Tanggal SO eksplisit per baris, jadi siap dibaca oleh
 * function report manapun nanti (cukup getRange P:V di SO daily log,
 * filter by Tanggal SO sesuai range minggu/bulan, group by Type).
 * Fungsi pembacaan itu BELUM ditulis di sini karena belum diminta —
 * lihat catatan di akhir dokumentasi soal ini.
 * ============================================================
 */

const CONFIG = {
  INPUT_SHEET: "SO daily input",
  LOG_SHEET: "SO daily log",
  START_ROW: 4,

  COL_NO: 5,
  COL_KODE: 6,
  COL_NAMA: 7,
  COL_SATUAN: 8,
  COL_STOK_AWAL: 9,
  COL_MASUK: 10,
  COL_STOK_AKHIR: 11,
  COL_SYSTEM_BASE: 12,
  COL_TERPAKAI: 13,
  COL_SELISIH: 14,
  COL_KET: 15
};

/**
 * BARU: konstanta tabel KALIBRASI/JATAH/WASTE di "SO daily input".
 * Posisi dikonfirmasi langsung dari sheet (bukan tebakan):
 * No=E(5), Nama Barang=G(7), Type=J(10), Qty=N(14), Keterangan=O(15).
 * Header di row 23, data row 24 s.d. 34 (11 baris, fixed).
 */
const CONFIG_ADJ = {
  TITLE_ROW: 22,
  HEADER_ROW: 23,
  START_ROW: 24,
  MAX_ROWS: 11, // row 24-34

  COL_NO: 5,    // E
  COL_NAMA: 7,  // G
  COL_TYPE: 10, // J
  COL_QTY: 14,  // N
  COL_KET: 15   // O
};

/**
 * BARU: konstanta Table2 (smart table) di "SO daily log" untuk
 * menyimpan histori adjustment. Header row 1, data mulai row 2.
 * Kolom dikonfirmasi: Waktu Submit=P(16), Tanggal SO=Q(17), No=R(18),
 * Nama Barang=S(19), Type=T(20), Qty=U(21), Keterangan=V(22).
 */
const CONFIG_ADJ_LOG = {
  HEADER_ROW: 1,
  START_ROW: 2,

  COL_WAKTU_SUBMIT: 16, // P
  COL_TANGGAL_SO: 17,   // Q
  COL_NO: 18,           // R
  COL_NAMA: 19,         // S
  COL_TYPE: 20,         // T
  COL_QTY: 21,          // U
  COL_KET: 22           // V
};

const CONFIG_BM = {
  SHEET_NAME: "SO barang masuk",
  START_ROW: 5,
  COL_NAMA: 2, // B
  COL_SATUAN: 3, // C
  COL_DAY_START: 4, // D (Day 1)
  COL_DAY_END: 34 // AH (Day 31)
};

const CONFIG_BM_LOG = {
  SHEET_NAME: "SO log barang masuk",
  START_ROW: 2,
  COL_WAKTU_SUBMIT: 1, // A
  COL_TANGGAL_MASUK: 2, // B
  COL_NAMA: 3, // C
  COL_SATUAN: 4, // D
  COL_QTY: 5 // E
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tjap Boenga")
    .addItem("✅ Submit SO Hari Ini", "submitSO")
    .addItem("🔄 Reset Sistem", "resetSistem")
    .addItem("📊 Generate Weekly Report", "generateWeeklyReportMenu")
    .addItem("📈 Generate Monthly Report", "generateMonthlyReportMenu")
    .addSeparator()
    .addItem("📦 Closing Barang Masuk Bulanan", "submitBarangMasukBulanan")
    .addToUi();
}

function getJumlahBarisItem(sheet) {
  const data = sheet.getRange(CONFIG.START_ROW, CONFIG.COL_KODE, 500, 2).getValues();

  let count = 0;
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) break;
    count++;
  }
  return count;
}

function isNumeric(value) {
  if (value === "" || value === null) return false;
  return !isNaN(value) && isFinite(value);
}

function normalizeDate(dateObj) {
  return Utilities.formatDate(
    new Date(dateObj),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function isDuplicateSubmission(logSheet, tanggalRaw, outlet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return false;

  const data = logSheet.getRange(2, 2, lastRow - 1, 2).getValues();
  const targetDate = normalizeDate(tanggalRaw);

  for (let row of data) {
    if (!row[0]) continue;
    if (normalizeDate(row[0]) === targetDate && row[1] === outlet) {
      return true;
    }
  }

  return false;
}

function getLastSubmitInfo(logSheet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return null;

  const row = logSheet
    .getRange(lastRow, 1, 1, 4)
    .getValues()[0];

  return {
    waktu: row[0],
    tanggal: row[1],
    outlet: row[2],
    pic: row[3]
  };
}

function validateRows(sheet, jumlahBaris) {
  const errors = [];
  let firstErrorRow = null;

  for (let i = 0; i < jumlahBaris; i++) {
    const row = CONFIG.START_ROW + i;

    const kode = sheet.getRange(row, CONFIG.COL_KODE).getValue();
    const masuk = sheet.getRange(row, CONFIG.COL_MASUK).getValue();
    const stokAwal = Number(
      sheet.getRange(row, CONFIG.COL_STOK_AWAL).getValue()
    ) || 0;

    const stokAkhir = sheet.getRange(row, CONFIG.COL_STOK_AKHIR).getValue();
    const systemBase = sheet.getRange(row, CONFIG.COL_SYSTEM_BASE).getValue();

    if (!kode) continue;

    const rowErrors = [];

    if (masuk !== "" && !isNumeric(masuk)) {
      rowErrors.push("Barang Masuk harus angka");
    }

    if (Number(masuk) < 0) {
      rowErrors.push("Barang Masuk tidak boleh negatif");
    }

    if (stokAkhir === "") {
      rowErrors.push("Stok Akhir kosong");
    } else if (!isNumeric(stokAkhir)) {
      rowErrors.push("Stok Akhir harus angka");
    } else if (Number(stokAkhir) < 0) {
      rowErrors.push("Stok Akhir tidak boleh negatif");
    }

    if (systemBase === "") {
      rowErrors.push("System Base kosong");
    } else if (!isNumeric(systemBase)) {
      rowErrors.push("System Base harus angka");
    } else if (Number(systemBase) < 0) {
      rowErrors.push("System Base tidak boleh negatif");
    }

    const maxAvailable = stokAwal + (Number(masuk) || 0);

    if (
      stokAkhir !== "" &&
      isNumeric(stokAkhir) &&
      Number(stokAkhir) > maxAvailable
    ) {
      rowErrors.push(
        `Stok Akhir melebihi stok tersedia (${maxAvailable})`
      );
    }

    if (rowErrors.length > 0) {
      if (firstErrorRow === null) {
        firstErrorRow = row;
      }

      errors.push(
        `• Row ${row} (${kode}) → ${rowErrors.join(", ")}`
      );
    }
  }

  return {
    valid: errors.length === 0,
    errors,
    firstErrorRow
  };
}

// ============================================================
// BARU: VALIDASI & PEMBACAAN TABEL KALIBRASI/JATAH/WASTE
// ============================================================

/**
 * Membaca seluruh baris di tabel adjustment (Kalibrasi/Jatah/Waste),
 * HANYA mengembalikan baris yang benar-benar terisi (skip kosong).
 *
 * Definisi "terisi": Nama Barang ATAU Type ATAU Qty ada isinya.
 * Baris yang benar-benar kosong semua (No saja yang terisi, karena
 * No memang sudah pre-filled 1-11 dari awal) otomatis di-skip.
 *
 * Validasi ringan disertakan: kalau salah satu dari Nama Barang/
 * Type/Qty diisi tapi yang lain kosong, baris itu dianggap "partial"
 * dan dimasukkan ke errors supaya user diminta melengkapi sebelum
 * submit (mencegah data adjustment yang tidak lengkap masuk log).
 */
function readAdjustmentRows(sheet) {
  const rows = [];
  const errors = [];

  for (let i = 0; i < CONFIG_ADJ.MAX_ROWS; i++) {
    const row = CONFIG_ADJ.START_ROW + i;

    const no = sheet.getRange(row, CONFIG_ADJ.COL_NO).getValue();
    const nama = sheet.getRange(row, CONFIG_ADJ.COL_NAMA).getValue();
    const type = sheet.getRange(row, CONFIG_ADJ.COL_TYPE).getValue();
    const qty = sheet.getRange(row, CONFIG_ADJ.COL_QTY).getValue();
    const ket = sheet.getRange(row, CONFIG_ADJ.COL_KET).getValue();

    const namaFilled = nama !== "" && nama !== null;
    const typeFilled = type !== "" && type !== null;
    const qtyFilled = qty !== "" && qty !== null;

    // Baris benar-benar kosong (tidak ada satupun field utama terisi) -> skip total.
    if (!namaFilled && !typeFilled && !qtyFilled) {
      continue;
    }

    // Baris "partial" -> wajib lengkap Nama Barang + Type + Qty.
    const rowErrors = [];
    if (!namaFilled) rowErrors.push("Nama Barang kosong");
    if (!typeFilled) rowErrors.push("Type belum dipilih");
    if (!qtyFilled) {
      rowErrors.push("Qty kosong");
    } else if (!isNumeric(qty)) {
      rowErrors.push("Qty harus angka");
    } else if (Number(qty) < 0) {
      rowErrors.push("Qty tidak boleh negatif");
    }

    if (rowErrors.length > 0) {
      errors.push(`• Adjustment row ${row} (No ${no}) → ${rowErrors.join(", ")}`);
      continue; // tidak dimasukkan ke rows kalau tidak valid
    }

    rows.push({
      row,
      no,
      nama,
      type,
      qty: Number(qty),
      ket: ket || ""
    });
  }

  return {
    valid: errors.length === 0,
    errors,
    rows
  };
}

/**
 * Membangun ringkasan singkat tabel adjustment untuk ditampilkan
 * di dialog konfirmasi submit (digabung dengan summary tabel utama).
 */
function buildAdjustmentSummary(adjustmentRows) {
  if (adjustmentRows.length === 0) {
    return "Kalibrasi/Jatah/Waste: tidak ada entry hari ini.";
  }

  const countByType = {};
  adjustmentRows.forEach(r => {
    countByType[r.type] = (countByType[r.type] || 0) + 1;
  });

  const typeSummary = Object.keys(countByType)
    .map(t => `${t}: ${countByType[t]} item`)
    .join(", ");

  return (
    `Kalibrasi/Jatah/Waste: ${adjustmentRows.length} entry (${typeSummary})\n` +
    adjustmentRows
      .map(r => `  • [${r.type}] ${r.nama} — ${r.qty}${r.ket ? " (" + r.ket + ")" : ""}`)
      .join("\n")
  );
}

/**
 * Insert baris adjustment yang sudah divalidasi ke Table2 pada
 * "SO daily log". Ditulis sebagai append langsung di bawah baris
 * terakhir Table2 yang terisi, supaya Smart Table Google Sheets
 * otomatis melebar mencakup baris baru (tidak meninggalkan gap).
 *
 * waktuSubmit & tanggalSO SAMA untuk seluruh baris dalam 1x submit
 * (timestamp submit & tanggal SO dari setting B2), sesuai requirement.
 */
function insertAdjustmentToLog(logSheet, adjustmentRows, waktuSubmit, tanggalSO) {
  if (adjustmentRows.length === 0) return;

  // Cari baris kosong pertama di Table2 berdasarkan kolom Waktu Submit (P).
  // Tidak pakai logSheet.getLastRow() global karena itu bisa terpengaruh
  // oleh Table1 (A-O) yang barisnya jauh lebih banyak/berbeda panjang.
  const colP = CONFIG_ADJ_LOG.COL_WAKTU_SUBMIT;
  const existingData = logSheet
    .getRange(CONFIG_ADJ_LOG.START_ROW, colP, Math.max(logSheet.getLastRow(), CONFIG_ADJ_LOG.START_ROW), 1)
    .getValues();

  let nextRow = CONFIG_ADJ_LOG.START_ROW;
  for (let i = 0; i < existingData.length; i++) {
    if (existingData[i][0] === "" || existingData[i][0] === null) {
      nextRow = CONFIG_ADJ_LOG.START_ROW + i;
      break;
    }
    nextRow = CONFIG_ADJ_LOG.START_ROW + i + 1;
  }

  const values = adjustmentRows.map(r => [
    waktuSubmit,        // P - Waktu Submit
    tanggalSO,           // Q - Tanggal SO
    r.no,                 // R - No
    r.nama,               // S - Nama Barang
    r.type,               // T - Type
    r.qty,                // U - Qty
    r.ket                 // V - Keterangan
  ]);

  logSheet
    .getRange(nextRow, CONFIG_ADJ_LOG.COL_WAKTU_SUBMIT, values.length, 7)
    .setValues(values);
}

/**
 * Reset tabel KALIBRASI/JATAH/WASTE di "SO daily input" setelah
 * submit berhasil. Kolom No (E) TIDAK di-clear (tetap 1-11, sudah
 * pre-filled secara statis), hanya Nama Barang/Type/Qty/Keterangan
 * yang dibersihkan.
 */
function clearAdjustmentTable(sheet) {
  sheet
    .getRange(CONFIG_ADJ.START_ROW, CONFIG_ADJ.COL_NAMA, CONFIG_ADJ.MAX_ROWS, 1)
    .clearContent();

  sheet
    .getRange(CONFIG_ADJ.START_ROW, CONFIG_ADJ.COL_TYPE, CONFIG_ADJ.MAX_ROWS, 1)
    .clearContent();

  sheet
    .getRange(CONFIG_ADJ.START_ROW, CONFIG_ADJ.COL_QTY, CONFIG_ADJ.MAX_ROWS, 1)
    .clearContent();

  sheet
    .getRange(CONFIG_ADJ.START_ROW, CONFIG_ADJ.COL_KET, CONFIG_ADJ.MAX_ROWS, 1)
    .clearContent();
}

// ============================================================
// SUMMARY / WARNING (tabel utama — TIDAK DIUBAH dari original)
// ============================================================

function buildSubmitSummary(sheet, jumlahBaris, tanggalDisplay, outlet, pic) {
  let totalMasuk = 0;
  let totalTerpakai = 0;
  let totalSelisih = 0;

  for (let i = 0; i < jumlahBaris; i++) {
    const row = CONFIG.START_ROW + i;

    totalMasuk += Number(
      sheet.getRange(row, CONFIG.COL_MASUK).getValue()
    ) || 0;

    totalTerpakai += Number(
      sheet.getRange(row, CONFIG.COL_TERPAKAI).getValue()
    ) || 0;

    totalSelisih += Number(
      sheet.getRange(row, CONFIG.COL_SELISIH).getValue()
    ) || 0;
  }

  return (
    "Konfirmasi Submit SO\n\n" +
    "Tanggal      : " + tanggalDisplay + "\n" +
    "Outlet       : " + outlet + "\n" +
    "PIC          : " + pic + "\n" +
    "Total Item   : " + jumlahBaris + "\n" +
    "Total Masuk  : " + totalMasuk + "\n" +
    "Total Usage  : " + totalTerpakai + "\n" +
    "Total Selisih: " + totalSelisih + "\n\n" +
    "Pastikan data sudah benar."
  );
}

function buildWarningSummary(sheet, jumlahBaris) {
  let akur = 0;
  let smallDiff = 0;
  let bigDiff = 0;

  let allSystemBaseZero = true;
  const anomalyItems = [];

  for (let i = 0; i < jumlahBaris; i++) {
    const row = CONFIG.START_ROW + i;

    const kode = sheet.getRange(row, CONFIG.COL_KODE).getValue();
    if (!kode) continue;

    const systemBase =
      Number(sheet.getRange(row, CONFIG.COL_SYSTEM_BASE).getValue()) || 0;

    const selisih =
      Number(sheet.getRange(row, CONFIG.COL_SELISIH).getValue()) || 0;

    const ket = String(
      sheet.getRange(row, CONFIG.COL_KET).getValue()
    );

    if (systemBase !== 0) {
      allSystemBaseZero = false;
    }

    if (ket.includes("🟢")) akur++;
    else if (ket.includes("🟡")) smallDiff++;
    else if (ket.includes("🟠") || ket.includes("🔴")) bigDiff++;

    if (systemBase > 0) {
      const diffPercent = Math.abs(selisih) / systemBase;

      if (diffPercent > 0.3) {
        anomalyItems.push(
          `${kode} (${Math.round(diffPercent * 100)}%)`
        );
      }
    }
  }

  const total = akur + smallDiff + bigDiff;
  const accuracy =
    total > 0 ? ((akur / total) * 100).toFixed(1) : 0;

  const warnings = [];

  if (allSystemBaseZero) {
    warnings.push(
      "⚠ Semua System Base bernilai 0 (data POS belum diinput?)"
    );
  }

  if (anomalyItems.length > 0) {
    warnings.push(
      `⚠ ${anomalyItems.length} item punya selisih besar (>30%)`
    );
  }

  return {
    warnings,
    summary:
      "Ringkasan SO\n\n" +
      `🟢 Akur          : ${akur}\n` +
      `🟡 Selisih kecil : ${smallDiff}\n` +
      `🟠/🔴 Selisih besar: ${bigDiff}\n` +
      `Accuracy Rate   : ${accuracy}%`
  };
}

// ============================================================
// SUBMIT FLOW — DIPERLUAS untuk mencakup tabel adjustment
// ============================================================

function submitSO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName(CONFIG.INPUT_SHEET);
  const log = ss.getSheetByName(CONFIG.LOG_SHEET);
  const ui = SpreadsheetApp.getUi();

  const tanggalRaw = input.getRange("B2").getValue();
  const outlet = input.getRange("B3").getValue();
  const pic = input.getRange("B4").getValue();

  const tanggalDisplay = Utilities.formatDate(
    tanggalRaw,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  if (!tanggalRaw || !outlet || !pic) {
    ui.alert("Tanggal / Outlet / PIC belum lengkap");
    return;
  }

  if (isDuplicateSubmission(log, tanggalRaw, outlet)) {
    const last = getLastSubmitInfo(log);

    let message =
      `SO untuk outlet ${outlet}\n` +
      `tanggal ${tanggalDisplay}\n` +
      `SUDAH DIKUNCI.\n\n`;

    if (last) {
      const jam = Utilities.formatDate(
        new Date(last.waktu),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      );

      message +=
        `Last Submit:\n` +
        `${jam}\n` +
        `PIC: ${last.pic}`;
    }

    ui.alert(message);
    return;
  }

  const jumlahBaris = getJumlahBarisItem(input);

  const validationResult = validateRows(input, jumlahBaris);

  if (!validationResult.valid) {
    ui.alert(
      "Validation Error",
      "Ditemukan " + validationResult.errors.length + " masalah:\n\n" +
      validationResult.errors.join("\n"),
      ui.ButtonSet.OK
    );

    if (validationResult.firstErrorRow) {
      input.setActiveRange(
        input.getRange(
          validationResult.firstErrorRow,
          CONFIG.COL_MASUK
        )
      );
    }

    return;
  }

  // BARU: validasi tabel adjustment (Kalibrasi/Jatah/Waste).
  // Kalau ada baris "partial" (terisi sebagian), submit dihentikan
  // dulu supaya user melengkapi — sama seperti perlakuan tabel utama.
  const adjustmentResult = readAdjustmentRows(input);

  if (!adjustmentResult.valid) {
    ui.alert(
      "Validation Error — Kalibrasi/Jatah/Waste",
      "Ditemukan " + adjustmentResult.errors.length + " masalah:\n\n" +
      adjustmentResult.errors.join("\n"),
      ui.ButtonSet.OK
    );
    return;
  }

  const summary = buildSubmitSummary(
    input,
    jumlahBaris,
    tanggalDisplay,
    outlet,
    pic
  );

  const warningData = buildWarningSummary(
    input,
    jumlahBaris
  );

  // BARU: gabungkan ringkasan adjustment ke summary konfirmasi,
  // supaya user lihat dulu sebelum klik YES.
  const adjustmentSummary = buildAdjustmentSummary(adjustmentResult.rows);

  let finalSummary = summary + "\n\n" + warningData.summary + "\n\n" + adjustmentSummary;

  if (warningData.warnings.length > 0) {
    finalSummary +=
      "\n\nWarnings:\n" +
      warningData.warnings.join("\n");
  }

  const confirm = ui.alert(
    "Submit SO",
    finalSummary,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // ---- Submit tabel utama (logic ASLI, tidak diubah) ----
  const data = input.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_KODE,
    jumlahBaris,
    10
  ).getValues();

  // ---- BARU: Sync Barang Masuk to Matrix ----
  try {
    syncBarangMasukToMatrix(tanggalRaw, data);
  } catch(e) {
    ui.alert("Warning Sync", "Gagal sinkronisasi ke SO Barang Masuk: " + e.message, ui.ButtonSet.OK);
  }

  const logs = [];
  const newStocks = [];

  for (let row of data) {
    const kode = row[0];
    if (!kode) continue;

    logs.push([
      new Date(),
      tanggalRaw,
      outlet,
      pic,
      row[0],
      row[1],
      row[2],
      row[3],
      Number(row[4]) || 0,
      row[5],
      row[6],
      row[7],
      row[8],
      row[9]
    ]);

    newStocks.push([Number(row[5]) || 0]);
  }

  const insertRow = Math.max(log.getLastRow() + 1, 2);

  log.getRange(insertRow, 1, logs.length, logs[0].length)
    .setValues(logs);

  SpreadsheetApp.flush();

  input.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_STOK_AWAL,
    newStocks.length,
    1
  ).setValues(newStocks);

  SpreadsheetApp.flush();

  input.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_MASUK,
    jumlahBaris,
    3
  ).clearContent();

  input.getRange("B7").setValue("Last Submit");
  input.getRange("B8").setValue(
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    )
  );
  input.getRange("B9").setValue(pic);

  // ---- BARU: Submit tabel adjustment (Kalibrasi/Jatah/Waste) ----
  // Pakai timestamp & tanggal SO YANG SAMA dengan submit utama,
  // supaya satu kali klik Submit = satu waktu submit untuk semua data.
  const waktuSubmit = new Date();

  insertAdjustmentToLog(log, adjustmentResult.rows, waktuSubmit, tanggalRaw);

  clearAdjustmentTable(input);

  ui.alert("✔ Submit berhasil");
}

function resetSistem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = ss.getSheetByName(CONFIG.INPUT_SHEET);
  const log = ss.getSheetByName(CONFIG.LOG_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (ui.alert("Reset?", "Hapus log dan reset stok?", ui.ButtonSet.YES_NO) !== ui.Button.YES)
    return;

  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn())
      .clearContent();
  }

  const jumlahBaris = getJumlahBarisItem(input);
  const zeros = Array(jumlahBaris).fill([0]);

  input.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_STOK_AWAL,
    jumlahBaris,
    1
  ).setValues(zeros);

  input.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_MASUK,
    jumlahBaris,
    3
  ).clearContent();

  // BARU: ikut reset tabel adjustment supaya konsisten dengan
  // tabel utama saat reset sistem dilakukan.
  clearAdjustmentTable(input);

  ui.alert("✔ Reset berhasil");
}

function generateWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("SO daily log");
  const reportSheet = ss.getSheetByName("SO weekly report");

  if (!logSheet || !reportSheet) {
    SpreadsheetApp.getUi().alert("Sheet report/log tidak ditemukan");
    return;
  }

  const weekValue = reportSheet.getRange("B6").getDisplayValue();
  const monthValue = reportSheet.getRange("B7").getDisplayValue();
  
  const inputSheet = ss.getSheetByName(CONFIG.INPUT_SHEET);
  const outletFilter = inputSheet ? inputSheet.getRange("B3").getValue() : "";
  reportSheet.getRange("B8").setValue(outletFilter);

  // Parse week range
  const weekRange = parseWeekRange(weekValue);

  if (!weekRange) {
    SpreadsheetApp.getUi().alert("Format week selector invalid");
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) {
    SpreadsheetApp.getUi().alert("Belum ada data log");
    return;
  }

  const filtered = [];
  const itemMap = {};

  let totalUsage = 0;
  let totalMasuk = 0;
  let totalSelisih = 0;
  let akurCount = 0;

  const uniqueDates = new Set();

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];

    const tanggal = row[1];
    const outlet = row[2];

    if (!tanggal) continue;

    const dateObj = new Date(tanggal);
    if (
      dateObj < weekRange.start ||
      dateObj > weekRange.end
    ) {
      continue;
    }

    if (outlet !== outletFilter) {
      continue;
    }

    filtered.push(row);
    uniqueDates.add(normalizeDate(tanggal));

    const kode = row[4];
    const nama = row[5];
    const satuan = row[6];
    const stokAwal = Number(row[7]) || 0;
    const masuk = Number(row[8]) || 0;
    const stokAkhir = Number(row[9]) || 0;
    const terpakai = Number(row[11]) || 0;
    const selisih = Number(row[12]) || 0;

    totalUsage += terpakai;
    totalMasuk += masuk;
    totalSelisih += Math.abs(selisih);

    if (!itemMap[kode]) {
      itemMap[kode] = {
        week: weekValue,
        outlet,
        kode,
        nama,
        satuan,
        awal: stokAwal,
        masuk: 0,
        akhir: stokAkhir,
        usage: 0,
        selisih: 0,
        entries: 0
      };
    }

    itemMap[kode].masuk += masuk;
    itemMap[kode].akhir = stokAkhir;
    itemMap[kode].usage += terpakai;
    itemMap[kode].selisih += Math.abs(selisih);
    itemMap[kode].entries++;
  }

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("Info", "Tidak ada data untuk minggu ini.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Calculate Overall Accuracy by Volume
  let accuracyRate = 1;
  if (totalUsage > 0) {
    accuracyRate = Math.max(0, 1 - (totalSelisih / totalUsage));
  } else if (totalSelisih > 0) {
    accuracyRate = 0;
  }

  const items = Object.values(itemMap);

  const worstItem = [...items].sort(
    (a, b) => b.selisih - a.selisih
  )[0];

  // KPI
  reportSheet.getRange("E7").setValue(uniqueDates.size);
  reportSheet.getRange("G7").setValue(totalUsage);
  reportSheet.getRange("I7").setValue(totalMasuk);
  reportSheet.getRange("K7").setValue(totalSelisih);
  reportSheet.getRange("M7").setValue(accuracyRate);
  reportSheet.getRange("O7").setValue(worstItem?.nama || "-");

  reportSheet.getRange("E7:K7").setNumberFormat("#,##0");
  reportSheet.getRange("M7").setNumberFormat("0.0%");

  // Top 5 Usage
  const topUsage = [...items]
    .sort((a, b) => b.usage - a.usage)
    .slice(0, 5);

  reportSheet.getRange("A13:E17").clearContent();

  topUsage.forEach((item, idx) => {
    const row = 13 + idx;
    reportSheet.getRange(`A${row}`).setValue(idx + 1);
    reportSheet.getRange(`B${row}`).setValue(item.kode);
    reportSheet.getRange(`C${row}`).setValue(item.nama);
    reportSheet.getRange(`E${row}`).setValue(item.usage);
  });

  // Top 5 Selisih
  const topSelisih = [...items]
    .sort((a, b) => b.selisih - a.selisih)
    .slice(0, 5);

  reportSheet.getRange("G13:L17").clearContent();

  topSelisih.forEach((item, idx) => {
    const row = 13 + idx;
    reportSheet.getRange(`G${row}`).setValue(idx + 1);
    reportSheet.getRange(`I${row}`).setValue(item.kode);
    reportSheet.getRange(`J${row}`).setValue(item.nama);
    reportSheet.getRange(`L${row}`).setValue(item.selisih);
  });

  // Notes
  const notes = [
    `• Accuracy minggu ini ${(accuracyRate * 100).toFixed(1)}%`,
    `• ${topSelisih.length} item punya selisih`,
    `• Worst item: ${worstItem?.nama || "-"}`,
    `• Total usage ${totalUsage.toLocaleString()}`
  ].join("\n");

  reportSheet.getRange("O11").setValue(notes);

  // Summary table
  reportSheet.getRange("A22:Q500").clearContent();

  items.forEach((item, idx) => {
    const row = 22 + idx;
    let itemAccuracy = 1;
    if (item.usage > 0) {
      itemAccuracy = Math.max(0, 1 - (item.selisih / item.usage));
    } else if (item.selisih > 0) {
      itemAccuracy = 0;
    }

    reportSheet.getRange(`A${row}`).setValue(item.week);
    reportSheet.getRange(`B${row}`).setValue(item.outlet);
    reportSheet.getRange(`C${row}`).setValue(item.kode);
    reportSheet.getRange(`E${row}`).setValue(item.nama);
    reportSheet.getRange(`G${row}`).setValue(item.satuan);
    reportSheet.getRange(`H${row}`).setValue(item.awal);
    reportSheet.getRange(`J${row}`).setValue(item.masuk);
    reportSheet.getRange(`L${row}`).setValue(item.akhir);
    reportSheet.getRange(`N${row}`).setValue(item.usage);
    reportSheet.getRange(`P${row}`).setValue(item.selisih);
    reportSheet.getRange(`Q${row}`).setValue(itemAccuracy);
  });

  reportSheet.getRange("Q22:Q500").setNumberFormat("0.0%");

  // Update Generated At & Generated By
  const picName = inputSheet ? inputSheet.getRange("B4").getValue() : "System";
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  reportSheet.getRange("Q2").setValue(nowStr);
  reportSheet.getRange("Q3").setValue(picName);

  SpreadsheetApp.getUi().alert("Sukses", "Weekly Report berhasil di-generate!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateWeeklyReportMenu() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Generate Weekly Report",
    "Pilih mode:\n1 = Current Week\n2 = Custom Week",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const mode = response.getResponseText().trim();

  if (mode === "1") {
    try {
      prepareCurrentWeek();
      generateWeeklyReport();
    } catch (e) { return; }
    return;
  }

  if (mode === "2") {
    try {
      prepareCustomWeek();
      generateWeeklyReport();
    } catch (e) {
      return;
    }
    return;
  }

  ui.alert("Input harus 1 atau 2");
}

function parseWeekRange(weekText) {
  // Example:
  // Week 2 (8 - 14 June 2026)

  const regex = /\((\d+)\s*-\s*(\d+)\s+([A-Za-z]+)\s+(\d{4})\)/;
  const match = weekText.match(regex);

  if (!match) return null;

  const startDay = Number(match[1]);
  const endDay = Number(match[2]);
  const monthName = match[3];
  const year = Number(match[4]);

  const monthMap = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11,
    Januari: 0,
    Februari: 1,
    Maret: 2,
    April: 3,
    Mei: 4,
    Juni: 5,
    Juli: 6,
    Agustus: 7,
    September: 8,
    Oktober: 9,
    November: 10,
    Desember: 11
  };

  const month = monthMap[monthName];

  if (month === undefined) return null;

  return {
    start: new Date(year, month, startDay),
    end: new Date(year, month, endDay, 23, 59, 59)
  };
}

function prepareCurrentWeek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("SO weekly report");

  const today = new Date();
  const day = today.getDate();
  const month = today.getMonth();
  const year = today.getFullYear();

  let weekNum;
  let startDay;
  let endDay;

  if (day <= 7) {
    weekNum = 1;
    startDay = 1;
    endDay = 7;
  } else if (day <= 14) {
    weekNum = 2;
    startDay = 8;
    endDay = 14;
  } else if (day <= 21) {
    weekNum = 3;
    startDay = 15;
    endDay = 21;
  } else if (day <= 28) {
    weekNum = 4;
    startDay = 22;
    endDay = 28;
  } else {
    weekNum = 5;
    startDay = 29;
    endDay = new Date(year, month + 1, 0).getDate();
  }

  const monthName = Utilities.formatDate(
    today,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  );

  const weekText =
    `Week ${weekNum} (${startDay} - ${endDay} ` +
    Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM yyyy") +
    `)`;

  reportSheet.getRange("B6").setValue(weekText);
  reportSheet.getRange("B7").setValue(monthName);
}

function prepareCustomWeek() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("SO weekly report");

  const monthPrompt = ui.prompt(
    "Custom Month",
    "Masukkan bulan (contoh: June 2026)",
    ui.ButtonSet.OK_CANCEL
  );

  if (monthPrompt.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Cancelled");
  }

  const monthText = monthPrompt.getResponseText().trim();

  const weekPrompt = ui.prompt(
    "Custom Week",
    "Pilih week (1-5)",
    ui.ButtonSet.OK_CANCEL
  );

  if (weekPrompt.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Cancelled");
  }

  const weekNum = Number(weekPrompt.getResponseText());

  if (weekNum < 1 || weekNum > 5) {
    throw new Error("Week harus 1-5");
  }

  const parts = monthText.split(" ");
  const monthName = parts[0];
  const year = Number(parts[1]);

  const monthMap = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11
  };

  const month = monthMap[monthName];

  let startDay, endDay;

  if (weekNum === 1) {
    startDay = 1; endDay = 7;
  } else if (weekNum === 2) {
    startDay = 8; endDay = 14;
  } else if (weekNum === 3) {
    startDay = 15; endDay = 21;
  } else if (weekNum === 4) {
    startDay = 22; endDay = 28;
  } else {
    startDay = 29;
    endDay = new Date(year, month + 1, 0).getDate();
  }

  const weekText =
    `Week ${weekNum} (${startDay} - ${endDay} ${monthText})`;

  reportSheet.getRange("B6").setValue(weekText);
  reportSheet.getRange("B7").setValue(monthText);
}

function onEdit(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.INPUT_SHEET) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < CONFIG.START_ROW) return;

  const watchedCols = [
    CONFIG.COL_MASUK,
    CONFIG.COL_STOK_AKHIR,
    CONFIG.COL_SYSTEM_BASE
  ];

  if (!watchedCols.includes(col)) return;

  // Self-healing: bersihkan semua background merah lama
  const totalRows = getJumlahBarisItem(sheet);

  sheet.getRange(
    CONFIG.START_ROW,
    CONFIG.COL_NO,
    totalRows,
    CONFIG.COL_KET - CONFIG.COL_NO + 1
  ).setBackground(null);

    for (let i = 0; i < totalRows; i++) {
    const currentRow = CONFIG.START_ROW + i;

    const kode = sheet.getRange(currentRow, CONFIG.COL_KODE).getValue();
    if (!kode) continue;

    const masukRaw = sheet.getRange(currentRow, CONFIG.COL_MASUK).getValue();
    const stokAkhirRaw = sheet.getRange(currentRow, CONFIG.COL_STOK_AKHIR).getValue();
    const systemBaseRaw = sheet.getRange(currentRow, CONFIG.COL_SYSTEM_BASE).getValue();
    const stokAwalRaw = sheet.getRange(currentRow, CONFIG.COL_STOK_AWAL).getValue();

    let invalid = false;

    const stokAwal = Number(stokAwalRaw) || 0;
    const masuk = Number(masukRaw) || 0;
    const stokAkhir = Number(stokAkhirRaw);
    const systemBase = Number(systemBaseRaw);

    // Skip row yang masih belum selesai diisi user
    if (masukRaw === "" || stokAkhirRaw === "" || systemBaseRaw === "") {
      continue;
    }

    if (!isNumeric(masukRaw) || masuk < 0) invalid = true;
    if (!isNumeric(stokAkhirRaw) || stokAkhir < 0) invalid = true;
    if (!isNumeric(systemBaseRaw) || systemBase < 0) invalid = true;

    const maxAvailable = stokAwal + masuk;
    if (!isNaN(stokAkhir) && stokAkhir > maxAvailable) {
      invalid = true;
    }

    if (invalid) {
      sheet.getRange(
        currentRow,
        CONFIG.COL_NO,
        1,
        CONFIG.COL_KET - CONFIG.COL_NO + 1
      ).setBackground("#ffd7d7");
    }
  }
}

// ============================================================
// BARU: INTEGRASI BARANG MASUK & CLOSING BULANAN
// ============================================================

/**
 * Sinkronisasi kolom Barang Masuk dari SO Daily Input ke sheet SO Barang Masuk (Matrix)
 * @param {Date} tanggalRaw - Objek Date dari input "Tanggal"
 * @param {Array} dataInput - Array 2D dari getValues() tabel utama SO Daily (kolom KODE s.d KET)
 */
function syncBarangMasukToMatrix(tanggalRaw, dataInput) {
  if (!tanggalRaw || !(tanggalRaw instanceof Date)) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bmSheet = ss.getSheetByName(CONFIG_BM.SHEET_NAME);
  
  if (!bmSheet) {
    throw new Error(`Sheet '${CONFIG_BM.SHEET_NAME}' tidak ditemukan.`);
  }

  const dayOfMonth = tanggalRaw.getDate(); // 1-31
  if (dayOfMonth < 1 || dayOfMonth > 31) return;

  const targetCol = CONFIG_BM.COL_DAY_START + dayOfMonth - 1; // Jika day 1, target col D(4)

  const lastRowBM = bmSheet.getLastRow();
  if (lastRowBM < CONFIG_BM.START_ROW) return;

  // Cache seluruh data matrix (mulai Start Row s.d Target Col) agar read/write efisien
  const matrixData = bmSheet.getRange(
    CONFIG_BM.START_ROW, 
    CONFIG_BM.COL_NAMA, 
    lastRowBM - CONFIG_BM.START_ROW + 1, 
    targetCol - CONFIG_BM.COL_NAMA + 1
  ).getValues();

  let hasChanges = false;

  // dataInput array map:
  // [0]=Kode, [1]=Nama, [2]=Satuan, [3]=Stok Awal, [4]=Masuk (Col 10), dst...
  for (let i = 0; i < dataInput.length; i++) {
    const namaInput = dataInput[i][1];
    const masukQty = Number(dataInput[i][4]) || 0;

    if (!namaInput || masukQty <= 0) continue; // Hanya sync yang terisi barang masuk

    // Cari item di matrix
    for (let j = 0; j < matrixData.length; j++) {
      const namaMatrix = matrixData[j][0]; // Kolom NAMA (index 0 relatif dari pencarian B s.d targetCol)
      if (namaMatrix === namaInput) {
        // Tulis value qty masuk ke cell day ybs
        const dayColIndex = targetCol - CONFIG_BM.COL_NAMA; // index array targetCol
        const currentMatrixVal = Number(matrixData[j][dayColIndex]) || 0;
        
        matrixData[j][dayColIndex] = currentMatrixVal + masukQty;
        hasChanges = true;
        break; // Lanjut ke item berikutnya di dataInput
      }
    }
  }

  // Jika ada perubahan, flush/tulis ulang HANYA ke kolom day terkait agar tidak over-write kolom lain
  if (hasChanges) {
    const colToUpdate = [];
    for (let j = 0; j < matrixData.length; j++) {
      const dayColIndex = targetCol - CONFIG_BM.COL_NAMA;
      colToUpdate.push([matrixData[j][dayColIndex]]);
    }
    
    bmSheet.getRange(
      CONFIG_BM.START_ROW, 
      targetCol, 
      colToUpdate.length, 
      1
    ).setValues(colToUpdate);
    
    SpreadsheetApp.flush();
  }
}

/**
 * Menu untuk Closing Bulanan SO Barang Masuk
 */
function submitBarangMasukBulanan() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bmSheet = ss.getSheetByName(CONFIG_BM.SHEET_NAME);
  const logSheet = ss.getSheetByName(CONFIG_BM_LOG.SHEET_NAME);

  if (!bmSheet || !logSheet) {
    ui.alert("Error", `Sheet '${CONFIG_BM.SHEET_NAME}' atau '${CONFIG_BM_LOG.SHEET_NAME}' tidak ditemukan!`, ui.ButtonSet.OK);
    return;
  }

  const promptBulan = ui.prompt(
    "Closing Barang Masuk Bulanan",
    "Masukkan Bulan & Tahun Closing\n(Contoh: Juni 2026)",
    ui.ButtonSet.OK_CANCEL
  );

  if (promptBulan.getSelectedButton() !== ui.Button.OK) return;
  
  const strBulanTahun = promptBulan.getResponseText().trim();
  if (!strBulanTahun) {
    ui.alert("Bulan & Tahun tidak boleh kosong!");
    return;
  }

  const confirm = ui.alert(
    "Konfirmasi Closing",
    `Anda akan melakukan closing barang masuk untuk periode: ${strBulanTahun}.\n\n` +
    `Proses ini akan:\n1. Menyimpan data matrix ke sheet Log.\n2. MENGHAPUS nilai 1-31 di matrix Barang Masuk agar kembali kosong.\n\nLanjutkan?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const lastRow = bmSheet.getLastRow();
  if (lastRow < CONFIG_BM.START_ROW) {
    ui.alert("Tidak ada data untuk di-closing.");
    return;
  }

  const matrixWidth = CONFIG_BM.COL_DAY_END - CONFIG_BM.COL_NAMA + 1;
  const dataBM = bmSheet.getRange(CONFIG_BM.START_ROW, CONFIG_BM.COL_NAMA, lastRow - CONFIG_BM.START_ROW + 1, matrixWidth).getValues();

  const waktuSubmit = new Date();
  const logsToInsert = [];

  // Parse dummy date untuk base bulan
  // Fallback: kita attach text strBulanTahun + " - Day X" jika parsing date rumit, 
  // atau user bebas. Kita asumsikan format "Day [X] - [Bulan Tahun]" agar tetap informatif.
  for (let i = 0; i < dataBM.length; i++) {
    const namaBarang = dataBM[i][0];
    const satuan = dataBM[i][1];
    
    if (!namaBarang) continue;

    for (let day = 1; day <= 31; day++) {
      const colIndex = (day - 1) + (CONFIG_BM.COL_DAY_START - CONFIG_BM.COL_NAMA);
      
      const qty = Number(dataBM[i][colIndex]) || 0; // cell kosong jadi 0
      
      // Sesuai persetujuan di implementasi: Filter out qty 0 untuk hemat space log
      if (qty > 0) {
        const tanggalStr = `${day} ${strBulanTahun}`;
        logsToInsert.push([
          waktuSubmit,      // A
          tanggalStr,       // B
          namaBarang,       // C
          satuan,           // D
          qty               // E
        ]);
      }
    }
  }

  if (logsToInsert.length > 0) {
    const insertRow = Math.max(logSheet.getLastRow() + 1, CONFIG_BM_LOG.START_ROW);
    logSheet.getRange(insertRow, 1, logsToInsert.length, 5).setValues(logsToInsert);
    SpreadsheetApp.flush();
  }

  // Clear data di matrix
  bmSheet.getRange(
    CONFIG_BM.START_ROW,
    CONFIG_BM.COL_DAY_START,
    lastRow - CONFIG_BM.START_ROW + 1,
    31
  ).clearContent();

  ui.alert("Sukses", `Closing Barang Masuk untuk periode ${strBulanTahun} berhasil.\nSebanyak ${logsToInsert.length} data record dimasukkan ke Log.`, ui.ButtonSet.OK);
}

// =========================================================================
// FITUR BARU: MONTHLY REPORT
// =========================================================================

function prepareCurrentMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("SO monthly report");

  const today = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];
  
  reportSheet.getRange("B6").setValue(today.getFullYear());
  reportSheet.getRange("B7").setValue(monthNames[today.getMonth()]);
}

function prepareCustomMonth() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("SO monthly report");

  const responseTahun = ui.prompt(
    "Custom Year",
    "Masukkan tahun (Contoh: 2026):",
    ui.ButtonSet.OK_CANCEL
  );

  if (responseTahun.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Dibatalkan");
  }

  const responseBulan = ui.prompt(
    "Custom Month",
    "Masukkan bulan (Contoh: June atau Juni):",
    ui.ButtonSet.OK_CANCEL
  );

  if (responseBulan.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Dibatalkan");
  }

  reportSheet.getRange("B6").setValue(responseTahun.getResponseText().trim());
  reportSheet.getRange("B7").setValue(responseBulan.getResponseText().trim());
}

function generateMonthlyReportMenu() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Generate Monthly Report",
    "Pilih mode:\n1 = Current Month\n2 = Custom Month",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const mode = response.getResponseText().trim();

  if (mode === "1") {
    try {
      prepareCurrentMonth();
      generateMonthlyReport();
    } catch (e) { return; }
    return;
  }

  if (mode === "2") {
    try {
      prepareCustomMonth();
      generateMonthlyReport();
    } catch (e) {
      return;
    }
    return;
  }

  ui.alert("Mode tidak dikenali. Pilih 1 atau 2.");
}

function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("SO monthly report");
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);

  const yearValue = reportSheet.getRange("B6").getValue();
  const monthText = reportSheet.getRange("B7").getValue(); 
  
  const inputSheet = ss.getSheetByName(CONFIG.INPUT_SHEET);
  const outletFilter = inputSheet ? inputSheet.getRange("B3").getValue() : "";
  reportSheet.getRange("B8").setValue(outletFilter);

  const monthName = String(monthText).trim();
  const year = Number(yearValue);

  if (!monthName || !year) {
    SpreadsheetApp.getUi().alert("Bulan atau Tahun tidak boleh kosong!");
    return;
  }
  
  const displayMonthText = monthName + " " + year;

  const monthMap = {
    January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
    July: 6, August: 7, September: 8, October: 9, November: 10, December: 11,
    Januari: 0, Februari: 1, Maret: 2, Mei: 4, Juni: 5, Juli: 6,
    Agustus: 7, Oktober: 9, Desember: 11
  };
  const targetMonth = monthMap[monthName];

  if (targetMonth === undefined) {
    SpreadsheetApp.getUi().alert("Nama bulan tidak dikenali: " + monthName);
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const filtered = [];
  const uniqueDates = new Set();
  
  let totalUsage = 0;
  let totalMasuk = 0;
  let totalSelisih = 0;

  const itemMap = {};

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const rawDate = row[1];
    const outlet = row[2];

    if (!rawDate) continue;
    
    let dateObj;
    if (typeof rawDate === "string") {
      const dateParts = rawDate.split("/");
      if (dateParts.length === 3) {
        dateObj = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
      } else {
        dateObj = new Date(rawDate);
      }
    } else {
      dateObj = new Date(rawDate);
    }

    if (!dateObj || isNaN(dateObj.getTime())) continue;

    if (dateObj.getMonth() !== targetMonth || dateObj.getFullYear() !== year) {
      continue;
    }

    if (outlet !== outletFilter) {
      continue;
    }

    filtered.push(row);
    uniqueDates.add(dateObj.getTime());

    const kode = row[4];
    const nama = row[5];
    const satuan = row[6];
    const stokAwal = Number(row[7]) || 0;
    const masuk = Number(row[8]) || 0;
    const stokAkhir = Number(row[9]) || 0;
    const terpakai = Number(row[11]) || 0;
    const selisih = Number(row[12]) || 0;

    totalUsage += terpakai;
    totalMasuk += masuk;
    totalSelisih += Math.abs(selisih);

    if (!itemMap[kode]) {
      itemMap[kode] = {
        month: displayMonthText,
        outlet: outletFilter,
        kode: kode,
        nama: nama,
        satuan: satuan,
        awal: stokAwal, 
        masuk: 0,
        akhir: stokAkhir, 
        usage: 0,
        selisih: 0,
        entries: 0
      };
    }

    itemMap[kode].masuk += masuk;
    itemMap[kode].akhir = stokAkhir;
    itemMap[kode].usage += terpakai;
    itemMap[kode].selisih += Math.abs(selisih);
    itemMap[kode].entries++;
  }

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("Info", "Tidak ada data untuk bulan ini.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let accuracyRate = 1;
  if (totalUsage > 0) {
    accuracyRate = Math.max(0, 1 - (totalSelisih / totalUsage));
  } else if (totalSelisih > 0) {
    accuracyRate = 0;
  }

  const items = Object.values(itemMap);
  const worstItem = [...items].sort((a, b) => b.selisih - a.selisih)[0];

  reportSheet.getRange("E7").setValue(uniqueDates.size);
  reportSheet.getRange("G7").setValue(totalUsage);
  reportSheet.getRange("I7").setValue(totalMasuk);
  reportSheet.getRange("K7").setValue(totalSelisih);
  reportSheet.getRange("M7").setValue(accuracyRate);
  reportSheet.getRange("O7").setValue(worstItem ? worstItem.nama : "-");
  reportSheet.getRange("O8").setValue(worstItem ? `Selisih: ${worstItem.selisih}` : "");

  const topUsage = [...items].sort((a, b) => b.usage - a.usage).slice(0, 5);
  reportSheet.getRange("A13:E17").clearContent();
  topUsage.forEach((item, idx) => {
    const row = 13 + idx;
    reportSheet.getRange(`A${row}`).setValue(idx + 1);
    reportSheet.getRange(`B${row}`).setValue(item.kode);
    reportSheet.getRange(`C${row}`).setValue(item.nama);
    reportSheet.getRange(`E${row}`).setValue(item.usage);
  });

  const topSelisih = [...items].sort((a, b) => b.selisih - a.selisih).slice(0, 5);
  reportSheet.getRange("G13:L17").clearContent();
  topSelisih.forEach((item, idx) => {
    const row = 13 + idx;
    reportSheet.getRange(`G${row}`).setValue(idx + 1);
    reportSheet.getRange(`I${row}`).setValue(item.kode);
    reportSheet.getRange(`J${row}`).setValue(item.nama);
    reportSheet.getRange(`L${row}`).setValue(item.selisih);
  });

  const notes = [
    `• Accuracy bulan ini ${(accuracyRate * 100).toFixed(1)}%`,
    `• ${topSelisih.length} item punya selisih`,
    `• Worst item: ${worstItem?.nama || "-"}`,
    `• Total usage ${totalUsage.toLocaleString()}`
  ].join("\n");
  reportSheet.getRange("O11").setValue(notes);

  reportSheet.getRange("A22:Q500").clearContent();
  items.forEach((item, idx) => {
    const row = 22 + idx;
    let itemAccuracy = 1;
    if (item.usage > 0) {
      itemAccuracy = Math.max(0, 1 - (item.selisih / item.usage));
    } else if (item.selisih > 0) {
      itemAccuracy = 0;
    }

    reportSheet.getRange(`A${row}`).setValue(item.month);
    reportSheet.getRange(`B${row}`).setValue(item.outlet);
    reportSheet.getRange(`C${row}`).setValue(item.kode);
    reportSheet.getRange(`E${row}`).setValue(item.nama);
    reportSheet.getRange(`G${row}`).setValue(item.satuan);
    reportSheet.getRange(`H${row}`).setValue(item.awal);
    reportSheet.getRange(`J${row}`).setValue(item.masuk);
    reportSheet.getRange(`L${row}`).setValue(item.akhir);
    reportSheet.getRange(`N${row}`).setValue(item.usage);
    reportSheet.getRange(`P${row}`).setValue(item.selisih);
    reportSheet.getRange(`Q${row}`).setValue(itemAccuracy);
  });

  reportSheet.getRange("Q22:Q500").setNumberFormat("0.0%");

  const picName = inputSheet ? inputSheet.getRange("B4").getValue() : "System";
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  reportSheet.getRange("Q2").setValue(nowStr);
  reportSheet.getRange("Q3").setValue(picName);

  SpreadsheetApp.getUi().alert("Sukses", "Monthly Report berhasil di-generate!", SpreadsheetApp.getUi().ButtonSet.OK);
}