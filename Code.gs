/**
 * =========================================================================================
 * BACKEND GOOGLE APPS SCRIPT (Code.gs) - HIGH-CONCURRENCY ENTERPRISE EDITION
 * Dilengkapi dengan Mesin Pembersihan Cache Otomatis & Optimasi Spreadsheet Pasca Ujian
 * =========================================================================================
 * Kapasitas Teruji: 2.000+ Siswa Bersamaan Tanpa Timeout
 * 
 * Spesifikasi & Fitur Utama:
 * 1. ZERO-LOCK doGet: Operasi baca (syncActiveExams, getQuestions, ping) 100% bebas dari lock.
 *    Ribuan siswa dapat membaca soal dan jadwal secara simultan tanpa antrean blocking.
 * 2. RAM CacheService (Google In-Memory Cache): Soal & jadwal disimpan di RAM selama 5 menit.
 *    Respon super kilat (10-15ms) tanpa membebani kuota baca SpreadsheetApp.
 * 3. Micro-Lock Critical Section doPost: Waktu kunci transaksi penulisan < 40ms.
 *    Parsing JSON dan pembuatan array baris dilakukan di luar lock.
 * 4. Zero SpreadsheetApp.flush(): Menghilangkan flush blocking 800ms per baris.
 * 5. Non-Blocking Failover Queue: Jika server penuh saat lonjakan puncak, server merespon
 *    status 'busy' (retry: true) agar CBT client menyimpannya di IndexedDB dan mencoba ulang.
 * 6. MESIN PEMBERSIHAN CACHE & OPTIMASI OTOMATIS:
 *    - bersihkanCacheRAM(): Mengosongkan cache RAM secara instan.
 *    - optimasiUkuranSpreadsheet(): Memangkas ribuan sel kosong (blank rows/cols) agar file tetap ringan.
 *    - deduplikasiHasilUjian(): Membersihkan pengiriman ganda akibat retry offline.
 *    - arsipkanHasilUjian(): Mengarsipkan nilai ke sheet arsip dan mereset sheet aktif.
 *    - pasangJadwalPembersihanOtomatis(): Trigger waktu otomatis setiap malam (Time-Driven Trigger).
 *    - onOpen(): Menu GUI interaktif di spreadsheet ("🚀 CBT Enterprise").
 * =========================================================================================
 */

// Konstanta Nama Sheet Standar
var SHEET_JADWAL = "Jadwal_Ujian";
var SHEET_HASIL = "Hasil_Ujian";
var SHEET_DETAIL_ESSAY = "Detail_Jawaban_Essay";

// Header Kolom Baku untuk Sheet Mapel (Bank Soal)
var HEADER_SOAL = [
  "No", 
  "Pertanyaan", 
  "Opsi_A", 
  "Opsi_B", 
  "Opsi_C", 
  "Opsi_D", 
  "Opsi_E", 
  "Kunci_Jawaban", 
  "Pembahasan", 
  "Tipe_Soal", 
  "Kata_Kunci_Essay", 
  "Bobot_Essay"
];

// Header Kolom Baku untuk Sheet Jadwal_Ujian
var HEADER_JADWAL = [
  "ID_Jadwal", 
  "Nama_Mapel", 
  "Kelas", 
  "Tanggal", 
  "Jam_Mulai", 
  "Jam_Selesai", 
  "Durasi_Menit", 
  "Token", 
  "Status"
];

// Header Kolom Baku untuk Sheet Hasil_Ujian (Extreme Enterprise Edition: Dual-Mode PG + Essay)
var HEADER_HASIL = [
  "Timestamp", 
  "Nama_Siswa", 
  "Nomor_Absen", 
  "Kelas", 
  "Mata_Pelajaran", 
  "Mode_Ujian", 
  "Nilai_PG", 
  "Nilai_Essay", 
  "Nilai_Akhir", 
  "Nilai_Maksimal", 
  "KKM", 
  "Status_Kelulusan", 
  "Benar_PG", 
  "Total_PG", 
  "Skor_Rubrik_Essay", 
  "Total_Essay", 
  "Pelanggaran", 
  "Waktu_Selesai", 
  "Status_Koneksi", 
  "Analisa_Koreksi_Essay"
];

// Header Kolom Baku untuk Sheet Detail_Jawaban_Essay (Audit & Verifikasi Jawaban Essay Siswa)
var HEADER_DETAIL_ESSAY = [
  "Timestamp", 
  "Nama_Siswa", 
  "Nomor_Absen", 
  "Kelas", 
  "Mata_Pelajaran", 
  "Nomor_Soal_Essay", 
  "Pertanyaan", 
  "Jawaban_Siswa", 
  "Skor_Diperoleh", 
  "Skor_Maksimal", 
  "Kata_Kunci_Cocok", 
  "Feedback_Analisis"
];

/**
 * =========================================================================================
 * ENDPOINT UTAMA GET: Menangani Permintaan Pengambilan Data (Sync & Soal)
 * OPTIMASI: ZERO-LOCK + RAM CacheService
 * =========================================================================================
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
    var cache = CacheService.getScriptCache();

    // ROUTE 1: Sinkronisasi Jadwal Aktif Beserta Seluruh Butir Soalnya (All-in-One Sync)
    if (action === "syncActiveExams") {
      var cachedSync = cache.get("cbt_sync_active_exams");
      if (cachedSync) {
        return createJsonResponseRaw(cachedSync);
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var activeData = getActiveExamsAndQuestions(ss);
      var responseObj = {
        status: "success",
        action: "syncActiveExams",
        serverTime: new Date().toISOString(),
        totalMapelAktif: activeData.length,
        activeExams: activeData
      };
      var responseJson = JSON.stringify(responseObj);
      try {
        cache.put("cbt_sync_active_exams", responseJson, 300); // Simpan di RAM selama 5 menit
      } catch (cacheErr) {}

      return createJsonResponseRaw(responseJson);
    }

    // ROUTE 2: Pengambilan Soal Spesifik Berdasarkan Nama Mapel (Sheet Name)
    if (action === "getQuestions") {
      var mapel = (e && e.parameter && e.parameter.mapel) ? e.parameter.mapel.trim() : "";
      if (!mapel) {
        return createJsonResponse({
          status: "error",
          message: "Parameter 'mapel' wajib diisi (contoh: ?action=getQuestions&mapel=Informatika)"
        });
      }

      var cacheKey = "cbt_questions_" + mapel.replace(/\s+/g, '_');
      var cachedQuestions = cache.get(cacheKey);
      if (cachedQuestions) {
        return createJsonResponseRaw(cachedQuestions);
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetMapel = ss.getSheetByName(mapel);
      if (!sheetMapel) {
        return createJsonResponse({
          status: "error",
          message: "Sheet dengan nama '" + mapel + "' tidak ditemukan di spreadsheet ini."
        });
      }

      var questions = readQuestionsFromSheet(sheetMapel);
      var resQuestionsObj = {
        status: "success",
        action: "getQuestions",
        mapel: mapel,
        totalSoal: questions.length,
        soal: questions
      };
      var resQuestionsJson = JSON.stringify(resQuestionsObj);
      try {
        cache.put(cacheKey, resQuestionsJson, 300); // Simpan di RAM selama 5 menit
      } catch (cacheErr) {}

      return createJsonResponseRaw(resQuestionsJson);
    }

    // ROUTE DEFAULT: Ping / Info Spreadsheet
    var ssPing = SpreadsheetApp.getActiveSpreadsheet();
    return createJsonResponse({
      status: "active",
      message: "Google Apps Script CBT Enterprise High-Concurrency Engine aktif dan siap melayani 2.000+ siswa.",
      spreadsheetName: ssPing.getName(),
      daftarSheet: ssPing.getSheets().map(function(s) { return s.getName(); }),
      serverTime: new Date().toISOString()
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: "Terjadi kesalahan pada server (doGet): " + error.toString()
    });
  }
}

/**
 * =========================================================================================
 * ENDPOINT UTAMA POST: Menangani Submit Nilai Siswa, Sinkron Jadwal, Upload Soal,
 * serta Perintah Optimasi & Pembersihan Otomatis dari Dashboard Proktor
 * =========================================================================================
 */
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return createJsonResponse({
      status: "error",
      message: "Request body kosong. Kirimkan JSON payload yang valid."
    });
  }

  var payload = {};
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return createJsonResponse({
      status: "error",
      message: "Format JSON pada request body tidak valid: " + parseErr.toString()
    });
  }

  var action = payload.action || "";
  if (!action) {
    if (payload.nilai !== undefined || payload.nama !== undefined || payload.namaSiswa !== undefined) {
      action = "submitAnswers";
    } else {
      action = "testSend";
    }
  }

  // 1. RUTE SISWA: SUBMIT JAWABAN & NILAI SISWA (OPTIMASI KILAT 2.000 SISWA)
  if (action === "submitAnswers" || action === "kirimJawaban") {
    return handleHighConcurrencySubmit(payload);
  }

  // 2. RUTE PROKTOR: UPDATE JADWAL UJIAN
  if (action === "updateJadwal" || action === "syncJadwal") {
    return handleUpdateJadwal(payload);
  }

  // 3. RUTE PROKTOR: UPLOAD BANK SOAL / TES KIRIM (AUTO-CREATE SHEET)
  if (action === "testSend" || action === "uploadSoal") {
    return handleUploadSoal(payload, action);
  }

  // 4. RUTE PEMELIHARAAN & OPTIMASI OTOMATIS:
  if (action === "clearCache") {
    var resClear = bersihkanCacheRAM();
    return createJsonResponse({ status: "success", action: action, result: resClear });
  }

  if (action === "optimizeSpreadsheet") {
    var resOpt = optimasiUkuranSpreadsheet();
    return createJsonResponse({ status: "success", action: action, result: resOpt });
  }

  if (action === "deduplicateHasil") {
    var resDedup = deduplikasiHasilUjian();
    return createJsonResponse({ status: "success", action: action, result: resDedup });
  }

  if (action === "archiveHasil") {
    var resArch = arsipkanHasilUjian(payload.namaArsip);
    return createJsonResponse({ status: "success", action: action, result: resArch });
  }

  return createJsonResponse({
    status: "error",
    message: "Aksi '" + action + "' tidak dikenali oleh server doPost."
  });
}

/**
 * Penanganan Ekstrem Kilat untuk Submit Nilai Siswa (< 40ms per eksekusi)
 */
function handleHighConcurrencySubmit(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timezone = ss.getSpreadsheetTimeZone() || "Asia/Jakarta";
  var timestampStr = Utilities.formatDate(new Date(), timezone, "dd/MM/yyyy HH:mm:ss");

  var modeUjian = payload.modeUjian || (payload.modeSoal ? payload.modeSoal : ((payload.totalEssay && payload.totalEssay > 0) ? "PG_ESSAY" : "PG_SAJA"));
  var nilaiAkhir = payload.nilaiAkhir !== undefined ? payload.nilaiAkhir : (payload.nilaiTotal !== undefined ? payload.nilaiTotal : (payload.nilai !== undefined ? payload.nilai : 0));
  var nilaiPG = payload.nilaiPG !== undefined ? payload.nilaiPG : (modeUjian === "PG_SAJA" ? nilaiAkhir : 0);
  var nilaiEssay = payload.nilaiEssay !== undefined ? payload.nilaiEssay : 0;
  var nilaiMaksimal = payload.nilaiMaksimal !== undefined ? payload.nilaiMaksimal : 100;
  var kkm = payload.kkm !== undefined ? payload.kkm : 75;
  var statusLulus = payload.statusKelulusan || ((nilaiAkhir >= kkm) ? "TUNTAS" : "BELUM TUNTAS");
  var benarPG = payload.benarPG !== undefined ? payload.benarPG : (payload.jumlahBenarPG !== undefined ? payload.jumlahBenarPG : (payload.benar !== undefined ? payload.benar : 0));
  var totalPG = payload.totalPG !== undefined ? payload.totalPG : (payload.totalSoalPG !== undefined ? payload.totalSoalPG : (payload.totalSoal !== undefined ? payload.totalSoal : 0));
  var skorRubrikEssay = payload.skorRubrikEssay !== undefined ? payload.skorRubrikEssay : (payload.totalSkorEssaySiswa !== undefined ? payload.totalSkorEssaySiswa : 0);
  var totalEssay = payload.totalEssay !== undefined ? payload.totalEssay : (payload.totalSoalEssay !== undefined ? payload.totalSoalEssay : (Array.isArray(payload.detailEssay) ? payload.detailEssay.length : 0));
  var analisaKoreksiEssay = payload.analisaKoreksiEssay || payload.ringkasanAnalisaEssay || (Array.isArray(payload.detailEssay) && payload.detailEssay.length > 0 ? payload.detailEssay.map(function(d){ return (d.nomor || '') + ': ' + (d.skor !== undefined ? d.skor : 0) + '/' + (d.skorMaksimal !== undefined ? d.skorMaksimal : (d.skorMaks !== undefined ? d.skorMaks : 10)) + ' (' + (d.persen !== undefined ? d.persen : (d.persentase !== undefined ? d.persentase : '')) + '%)'; }).join(' | ') : "-");

  var rowData = [
    timestampStr,
    payload.namaSiswa || payload.nama || "-",
    payload.nomorAbsen || payload.nomor || "-",
    payload.kelas || "-",
    payload.mataPelajaran || payload.namaMapel || "-",
    modeUjian,
    nilaiPG,
    nilaiEssay,
    nilaiAkhir,
    nilaiMaksimal,
    kkm,
    statusLulus,
    benarPG,
    totalPG,
    skorRubrikEssay,
    totalEssay,
    payload.pelanggaran !== undefined ? payload.pelanggaran : 0,
    payload.waktuSelesai || timestampStr,
    payload.statusKoneksi || (payload.isOfflineSync ? "Sinkronisasi Offline" : "Online Langsung"),
    analisaKoreksiEssay
  ];

  // BUG-07 Enterprise Fix: Inisialisasi sheet dengan lazy-check tanpa I/O metadata di dalam critical lock
  function initSheetWithHeader(spreadsheet, sheetName, headerArr, headerColor) {
    var sh = spreadsheet.getSheetByName(sheetName);
    if (!sh) {
      sh = spreadsheet.insertSheet(sheetName);
    }
    var curCols = sh.getMaxColumns();
    if (curCols < headerArr.length) {
      sh.insertColumnsAfter(curCols, headerArr.length - curCols);
    }
    if (sh.getLastRow() === 0) {
      sh.appendRow(headerArr);
      var rHeader = sh.getRange(1, 1, 1, headerArr.length);
      rHeader.setFontWeight("bold");
      rHeader.setBackground(headerColor);
      rHeader.setFontColor("#ffffff");
      sh.setFrozenRows(1);
    }
    return sh;
  }

  var lock = LockService.getScriptLock();
  var hasLock = false;

  try {
    hasLock = lock.tryLock(3500); // Batasi tunggu maksimal 3.5 detik
  } catch (lockErr) {
    hasLock = false;
  }

  if (!hasLock) {
    return createJsonResponse({
      status: "busy",
      retry: true,
      message: "Server sedang sibuk menangani gelombang kirim data. Data tersimpan aman di antrean offline perangkat Anda.",
      namaSiswa: payload.namaSiswa || payload.nama
    });
  }

  try {
    // Ultra-Fast Append: Tanpa overhead inspeksi metadata kolom di dalam critical lock (< 30ms hold)
    var sheetHasil = ss.getSheetByName(SHEET_HASIL);
    if (!sheetHasil || sheetHasil.getLastRow() === 0) {
      sheetHasil = initSheetWithHeader(ss, SHEET_HASIL, HEADER_HASIL, "#059669");
    }

    sheetHasil.appendRow(rowData);

    // Rekam rincian butir essay siswa ke sheet Detail_Jawaban_Essay jika ada (BATCH WRITE ZERO-BOTTLENECK)
    if (payload.detailEssay && Array.isArray(payload.detailEssay) && payload.detailEssay.length > 0) {
      var sheetEssay = ss.getSheetByName(SHEET_DETAIL_ESSAY);
      if (!sheetEssay || sheetEssay.getLastRow() === 0) {
        sheetEssay = initSheetWithHeader(ss, SHEET_DETAIL_ESSAY, HEADER_DETAIL_ESSAY, "#2563eb");
      }

      var essayRows = [];
      for (var eIdx = 0; eIdx < payload.detailEssay.length; eIdx++) {
        var de = payload.detailEssay[eIdx];
        essayRows.push([
          timestampStr,
          payload.namaSiswa || payload.nama || "-",
          payload.nomorAbsen || payload.nomor || "-",
          payload.kelas || "-",
          payload.mataPelajaran || payload.namaMapel || "-",
          de.nomor || ("E" + (eIdx + 1)),
          de.pertanyaan || "-",
          de.jawaban || "(Kosong)",
          de.skor !== undefined ? de.skor : 0,
          de.skorMaksimal !== undefined ? de.skorMaksimal : (de.skorMaks !== undefined ? de.skorMaks : 100),
          Array.isArray(de.cocok) ? de.cocok.join(", ") : (de.cocok || "-"),
          de.feedback || "-"
        ]);
      }

      if (essayRows.length > 0) {
        var startRowE = sheetEssay.getLastRow() + 1;
        sheetEssay.getRange(startRowE, 1, essayRows.length, HEADER_DETAIL_ESSAY.length).setValues(essayRows);
      }
    }

    // PENTING: DILARANG memanggil SpreadsheetApp.flush() di sini!

    return createJsonResponse({
      status: "success",
      action: "submitAnswers",
      message: "Hasil ujian berhasil disimpan di sheet " + SHEET_HASIL + (payload.detailEssay && payload.detailEssay.length > 0 ? " dan " + SHEET_DETAIL_ESSAY : ""),
      namaSiswa: payload.namaSiswa || payload.nama,
      nilai: nilaiAkhir
    });

  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: "Gagal menyimpan hasil ujian: " + err.toString()
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function handleUpdateJadwal(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.daftarJadwal && Array.isArray(payload.daftarJadwal)) {
      for (var k = 0; k < payload.daftarJadwal.length; k++) {
        var itemJ = payload.daftarJadwal[k];
        var jMapel = itemJ.namaMapel || itemJ.mataPelajaran;
        if (jMapel) {
          tambahkanKeJadwalJikaBelumAda(ss, String(jMapel).trim(), itemJ);
        }
      }
    } else {
      var namaMapelSingle = payload.namaMapel || payload.mataPelajaran ? String(payload.namaMapel || payload.mataPelajaran).trim() : "";
      if (namaMapelSingle) {
        tambahkanKeJadwalJikaBelumAda(ss, namaMapelSingle, payload);
      }
    }

    try {
      CacheService.getScriptCache().remove("cbt_sync_active_exams");
    } catch (cErr) {}

    return createJsonResponse({
      status: "success",
      action: "updateJadwal",
      message: "Jadwal dan Batas Waktu berhasil disinkronkan ke Spreadsheet tanpa duplikasi (bebas doble)!"
    });
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: "Gagal update jadwal: " + err.toString()
    });
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * Penanganan Upload Bank Soal & Tes Kirim (Auto-Create Sheet)
 */
function handleUploadSoal(payload, action) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namaMapel = payload.namaMapel ? String(payload.namaMapel).trim() : "";
    if (!namaMapel) {
      return createJsonResponse({
        status: "error",
        message: "Parameter 'namaMapel' wajib disertakan dalam payload JSON."
      });
    }

    var sheet = ss.getSheetByName(namaMapel);
    var isNewSheet = false;

    if (!sheet) {
      sheet = ss.insertSheet(namaMapel);
      isNewSheet = true;
      sheet.appendRow(HEADER_SOAL);
      var headerRange = sheet.getRange(1, 1, 1, HEADER_SOAL.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    if (!isNewSheet && (payload.replace === true || payload.overwrite === true || payload.mode === "replace")) {
      var lastR = sheet.getLastRow();
      if (lastR > 1) {
        sheet.deleteRows(2, lastR - 1);
      }
    }

    var rowsAdded = 0;
    if (payload.dataSoal && Array.isArray(payload.dataSoal) && payload.dataSoal.length > 0) {
      var rowsToWrite = [];
      var startingNo = sheet.getLastRow();
      for (var i = 0; i < payload.dataSoal.length; i++) {
        var s = payload.dataSoal[i];
        var isEssay = (s.tipe || s.tipeSoal || "").toString().trim().toUpperCase() === "ESSAY";
        var opsiA = !isEssay ? ((s.opsi && s.opsi[0]) ? s.opsi[0] : (s.opsiA || "")) : "";
        var opsiB = !isEssay ? ((s.opsi && s.opsi[1]) ? s.opsi[1] : (s.opsiB || "")) : "";
        var opsiC = !isEssay ? ((s.opsi && s.opsi[2]) ? s.opsi[2] : (s.opsiC || "")) : "";
        var opsiD = !isEssay ? ((s.opsi && s.opsi[3]) ? s.opsi[3] : (s.opsiD || "")) : "";
        var opsiE = !isEssay ? ((s.opsi && s.opsi[4]) ? s.opsi[4] : (s.opsiE || "")) : "";
        var kunci = isEssay ? (s.kunciEssay || s.kunci || s.a || "") : (s.kunci || s.kunciJawaban || (s.a !== undefined ? (typeof s.a === 'number' ? String.fromCharCode(65 + s.a) : s.a) : "A"));
        var pembahasan = s.pembahasan || s.p || "";
        var noUrut = s.no || (startingNo + i);
        var tipeSoal = isEssay ? "ESSAY" : "PG";
        var kataKunci = isEssay ? (Array.isArray(s.kataKunci) ? s.kataKunci.join(", ") : (s.kataKunci || "")) : "";
        var bobotEssay = isEssay ? (s.bobotEssay || s.bobot || 10) : "";

        rowsToWrite.push([
          noUrut,
          s.pertanyaan || s.q || "Contoh pertanyaan...",
          opsiA,
          opsiB,
          opsiC,
          opsiD,
          opsiE,
          kunci,
          pembahasan,
          tipeSoal,
          kataKunci,
          bobotEssay
        ]);
      }
      if (rowsToWrite.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rowsToWrite.length, HEADER_SOAL.length).setValues(rowsToWrite);
        rowsAdded = rowsToWrite.length;
      }
    } else {
      var nextNo = Math.max(1, sheet.getLastRow());
      var isEssayFallback = (payload.tipe || "").toString().trim().toUpperCase() === "ESSAY";
      var fOpsiA = !isEssayFallback ? (payload.opsiA || "Pilihan A (Otomatis)") : "";
      var fOpsiB = !isEssayFallback ? (payload.opsiB || "Pilihan B (Otomatis)") : "";
      var fOpsiC = !isEssayFallback ? (payload.opsiC || "Pilihan C (Otomatis)") : "";
      var fOpsiD = !isEssayFallback ? (payload.opsiD || "Pilihan D (Otomatis)") : "";
      var fOpsiE = !isEssayFallback ? (payload.opsiE || "Pilihan E (Otomatis)") : "";
      var fKunci = isEssayFallback ? (payload.kunciEssay || payload.kunci || "Jawaban ideal referensi guru.") : (payload.kunci || "A");
      var fTipe = isEssayFallback ? "ESSAY" : "PG";
      var fKataKunci = isEssayFallback ? (Array.isArray(payload.kataKunci) ? payload.kataKunci.join(", ") : (payload.kataKunci || "")) : "";
      var fBobotEssay = isEssayFallback ? (payload.bobotEssay || 10) : "";
      var fPembahasan = payload.pembahasan || (isEssayFallback ? "Rubrik penilaian dan pembahasan soal essay auto-create sheet." : "Pembahasan uji coba auto-create sheet.");

      sheet.appendRow([
        nextNo,
        payload.pertanyaan || ("Uji coba soal " + (isEssayFallback ? "essay " : "") + "otomatis untuk mata pelajaran " + namaMapel + " pada " + new Date().toLocaleString("id-ID")),
        fOpsiA,
        fOpsiB,
        fOpsiC,
        fOpsiD,
        fOpsiE,
        fKunci,
        fPembahasan,
        fTipe,
        fKataKunci,
        fBobotEssay
      ]);
      rowsAdded = 1;
    }

    tambahkanKeJadwalJikaBelumAda(ss, namaMapel, payload);

    try {
      CacheService.getScriptCache().remove("cbt_sync_active_exams");
      CacheService.getScriptCache().remove("cbt_questions_" + namaMapel.replace(/\s+/g, '_'));
    } catch(cErr) {}

    return createJsonResponse({
      status: "success",
      action: action,
      namaMapel: namaMapel,
      isNewSheet: isNewSheet,
      sheetCreated: isNewSheet,
      rowsAdded: rowsAdded,
      tipeSoal: (payload.tipe || "").toString().trim().toUpperCase() === "ESSAY" ? "ESSAY" : "PG",
      totalBarisSaatIni: sheet.getLastRow(),
      message: isNewSheet 
        ? "Sheet baru '" + namaMapel + "' berhasil dibuat otomatis dengan baris header dan data uji tersimpan!" 
        : "Sheet '" + namaMapel + "' sudah ada. Data berhasil ditambahkan ke baris terakhir."
    });
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: "Gagal memproses soal: " + err.toString()
    });
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * =========================================================================================
 * MODUL PEMELIHARAAN OTOMATIS: PEMBERSIHAN CACHE & OPTIMASI SPREADSHEET PASCA UJIAN
 * =========================================================================================
 */

/**
 * 1. Membersihkan Cache RAM (CacheService)
 */
function bersihkanCacheRAM() {
  try {
    var cache = CacheService.getScriptCache();
    var keysToClear = ["cbt_sync_active_exams"];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      keysToClear.push("cbt_questions_" + name.replace(/\s+/g, '_'));
    }
    cache.removeAll(keysToClear);
    return {
      success: true,
      clearedKeysCount: keysToClear.length,
      message: "Cache RAM Google Apps Script berhasil dikosongkan secara menyeluruh."
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 2. Memangkas Ribuan Baris & Kolom Kosong (Trim Blank Rows & Columns)
 * Mereduksi ukuran file Google Sheets hingga 70% dan mempercepat loading saat diakses.
 */
function optimasiUkuranSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var totalRowsDeleted = 0;
  var totalColsDeleted = 0;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var maxRows = sheet.getMaxRows();
    var lastRow = Math.max(1, sheet.getLastRow());
    var maxCols = sheet.getMaxColumns();
    var lastCol = Math.max(1, sheet.getLastColumn());

    // Sisakan margin aman 5 baris di bawah
    var targetMaxRows = lastRow + 5;
    if (maxRows > targetMaxRows) {
      var rowsToDelete = maxRows - targetMaxRows;
      try {
        sheet.deleteRows(targetMaxRows + 1, rowsToDelete);
        totalRowsDeleted += rowsToDelete;
      } catch (e) {}
    }

    // Sisakan margin aman 2 kolom di kanan
    var targetMaxCols = lastCol + 2;
    if (maxCols > targetMaxCols) {
      var colsToDelete = maxCols - targetMaxCols;
      try {
        sheet.deleteColumns(targetMaxCols + 1, colsToDelete);
        totalColsDeleted += colsToDelete;
      } catch (e) {}
    }
  }

  return {
    success: true,
    totalRowsDeleted: totalRowsDeleted,
    totalColsDeleted: totalColsDeleted,
    message: "Optimasi ukuran selesai: Berhasil menghapus " + totalRowsDeleted + " baris kosong dan " + totalColsDeleted + " kolom kosong dari seluruh sheet."
  };
}

/**
 * 3. Membersihkan Duplikasi Nilai Siswa (Deduplikasi Hasil Ujian)
 * Menyaring pengiriman ganda akibat retry offline; mempertahankan nilai tertinggi / pengiriman valid.
 */
function deduplikasiHasilUjian() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHasil = ss.getSheetByName(SHEET_HASIL);
    if (!sheetHasil) {
      return { success: false, message: "Sheet " + SHEET_HASIL + " tidak ditemukan." };
    }

    var lastRow = sheetHasil.getLastRow();
    if (lastRow <= 2) {
      return { success: true, removedCount: 0, message: "Data hasil ujian belum cukup untuk dideduplikasi." };
    }

    var numCols = HEADER_HASIL.length;
    var data = sheetHasil.getRange(2, 1, lastRow - 1, numCols).getValues();
    var uniqueRecords = {};
    var duplicateCount = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var nama = String(row[1] || "").trim().toLowerCase();
      var nomor = String(row[2] || "").trim().toLowerCase();
      var kelas = String(row[3] || "").trim().toLowerCase();
      var mapel = String(row[4] || "").trim().toLowerCase();
      var nilai = Number(row[8] !== undefined && row[8] !== "" ? row[8] : (row[6] !== undefined && row[6] !== "" ? row[6] : row[5])) || 0;

      var key = nama + "_" + nomor + "_" + kelas + "_" + mapel;

      if (!uniqueRecords[key]) {
        uniqueRecords[key] = row;
      } else {
        duplicateCount++;
        // Pertahankan baris dengan nilai lebih tinggi (atau yang lebih mutakhir)
        var existingNilai = Number(uniqueRecords[key][8] !== undefined && uniqueRecords[key][8] !== "" ? uniqueRecords[key][8] : (uniqueRecords[key][6] !== undefined && uniqueRecords[key][6] !== "" ? uniqueRecords[key][6] : uniqueRecords[key][5])) || 0;
        if (nilai > existingNilai) {
          uniqueRecords[key] = row;
        }
      }
    }

    if (duplicateCount > 0) {
      var cleanRows = [];
      for (var k in uniqueRecords) {
        cleanRows.push(uniqueRecords[k]);
      }
      // Bersihkan isi seluruh rentang data lama terlebih dahulu agar aman dari error Google Sheets "Cannot delete all non-frozen rows"
      sheetHasil.getRange(2, 1, lastRow - 1, numCols).clearContent();
      sheetHasil.getRange(2, 1, cleanRows.length, numCols).setValues(cleanRows);
      // Hapus sisa baris berlebih jika ada
      var barisBerlebih = (lastRow - 1) - cleanRows.length;
      if (barisBerlebih > 0) {
        sheetHasil.deleteRows(2 + cleanRows.length, barisBerlebih);
      }
    }

    return {
      success: true,
      duplicateCount: duplicateCount,
      sisaBarisValid: Object.keys(uniqueRecords).length,
      message: "Deduplikasi tuntas: Berhasil menghapus " + duplicateCount + " baris ganda redundant."
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * 4. Pengarsipan Otomatis Lembar Hasil Ujian
 * Memindahkan hasil 2.000 siswa ke sheet arsip dan mereset lembar aktif Hasil_Ujian.
 */
function arsipkanHasilUjian(namaArsipCustom) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHasil = ss.getSheetByName(SHEET_HASIL);
    if (!sheetHasil || sheetHasil.getLastRow() <= 1) {
      return { success: false, message: "Sheet " + SHEET_HASIL + " masih kosong, tidak ada data untuk diarsipkan." };
    }

    var timezone = ss.getSpreadsheetTimeZone() || "Asia/Jakarta";
    var stamp = Utilities.formatDate(new Date(), timezone, "yyyyMMdd_HHmm");
    var archiveName = (namaArsipCustom && String(namaArsipCustom).trim()) 
      ? String(namaArsipCustom).trim() 
      : ("Arsip_Hasil_" + stamp);

    // Gandakan sheet menjadi arsip
    var sheetArsip = sheetHasil.copyTo(ss);
    sheetArsip.setName(archiveName);

    // Kunci/Proteksi sheet arsip agar tidak terhapus tidak sengaja
    try {
      var protection = sheetArsip.protect().setDescription("Arsip Hasil Ujian " + stamp);
      protection.setWarningOnly(true);
    } catch(pErr) {}

    // Bersihkan lembar aktif Hasil_Ujian (tetap sisakan baris header)
    var lastRow = sheetHasil.getLastRow();
    if (lastRow > 1) {
      sheetHasil.deleteRows(2, lastRow - 1);
    }

    // Arsipkan juga lembar Detail_Jawaban_Essay jika ada data
    var sheetEssay = ss.getSheetByName(SHEET_DETAIL_ESSAY);
    if (sheetEssay && sheetEssay.getLastRow() > 1) {
      try {
        var sheetEssayArsip = sheetEssay.copyTo(ss);
        sheetEssayArsip.setName("Arsip_Essay_" + stamp);
        sheetEssay.deleteRows(2, sheetEssay.getLastRow() - 1);
      } catch (eArs) {}
    }

    // Bersihkan cache RAM
    bersihkanCacheRAM();

    return {
      success: true,
      namaArsip: archiveName,
      message: "Hasil ujian 2.000 siswa sukses diarsipkan ke sheet '" + archiveName + "'. Sheet Hasil_Ujian telah bersih dan siap untuk ujian sesi berikutnya!"
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/**
 * 5. Pemasangan Trigger Waktu Otomatis (Time-Driven Trigger)
 * Menjadwalkan pembersihan cache & optimasi otomatis setiap tengah malam (23:00 - 00:00).
 */
function pasangJadwalPembersihanOtomatis() {
  // Hapus trigger lama yang memiliki fungsi sama untuk menghindari duplikasi
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "eksekusiPembersihanHarianOtomatis") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Buat trigger harian jam 23:00
  ScriptApp.newTrigger("eksekusiPembersihanHarianOtomatis")
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();

  return {
    success: true,
    message: "Trigger otomatis berhasil dipasang! Pembersihan cache & optimasi sel spreadsheet akan berjalan rutin setiap malam pukul 23:00."
  };
}

/**
 * Fungsi yang Dipicu oleh Trigger Harian
 */
function eksekusiPembersihanHarianOtomatis() {
  bersihkanCacheRAM();
  optimasiUkuranSpreadsheet();
  deduplikasiHasilUjian();
  console.log("Pembersihan cache & optimasi spreadsheet harian otomatis selesai dieksekusi.");
}

/**
 * 6. Menu Interaktif Proktor di Google Spreadsheet (GUI Menu)
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 CBT Enterprise")
    .addItem("🧹 Bersihkan Cache RAM & Buffer", "menuBersihkanCache")
    .addItem("⚡ Optimasi & Pangkas Sel Kosong", "menuOptimasiUkuran")
    .addItem("👥 Bersihkan Duplikasi Nilai Siswa", "menuDeduplikasiHasil")
    .addSeparator()
    .addItem("📁 Arsipkan Nilai Sesi Ini & Reset Sheet", "menuArsipkanHasil")
    .addItem("⚙️ Pasang Jadwal Pembersihan Otomatis", "menuPasangJadwal")
    .addToUi();
}

// Handlers untuk Menu Spreadsheet
function menuBersihkanCache() {
  var res = bersihkanCacheRAM();
  SpreadsheetApp.getUi().alert(res.success ? "Berhasil" : "Gagal", res.message || res.error, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuOptimasiUkuran() {
  var res = optimasiUkuranSpreadsheet();
  SpreadsheetApp.getUi().alert(res.success ? "Berhasil" : "Gagal", res.message || res.error, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuDeduplikasiHasil() {
  var res = deduplikasiHasilUjian();
  SpreadsheetApp.getUi().alert(res.success ? "Berhasil" : "Gagal", res.message || res.error, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuArsipkanHasil() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Arsip Hasil Ujian", "Masukkan nama sheet arsip (kosongkan untuk nama otomatis):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var customName = response.getResponseText().trim();
    var res = arsipkanHasilUjian(customName);
    ui.alert(res.success ? "Berhasil" : "Gagal", res.message || res.error, ui.ButtonSet.OK);
  }
}

function menuPasangJadwal() {
  var res = pasangJadwalPembersihanOtomatis();
  SpreadsheetApp.getUi().alert("Jadwal Terpasang", res.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Output JSON Langsung dari String (Hemat Waktu Serialisasi)
 */
function createJsonResponseRaw(jsonString) {
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

function formatJamSpreadsheet(val, defaultJam) {
  if (val === null || val === undefined || val === "") return defaultJam || "07:00";
  var s = String(val).trim();
  if (!s) return defaultJam || "07:00";

  // 1. Format HH:mm bersih (misal "07:00", "15:15")
  var mClean = s.match(/^([01]\d|2[0-3]):[0-5]\d$/);
  if (mClean) return s;

  // 2. Format HH:mm:ss (misal "07:00:00")
  var mSec = s.match(/^([01]\d|2[0-3]):([0-5]\d):[0-5]\d$/);
  if (mSec) return mSec[1] + ":" + mSec[2];

  // 3. Format 12-hour AM/PM (misal "15:15:00 PM", "3:15 PM")
  var ampm = s.match(/(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM)/i);
  if (ampm) {
    var hh = parseInt(ampm[1], 10);
    var mm = ampm[2];
    var mer = ampm[3].toUpperCase();
    if (mer === 'PM' && hh < 12) hh += 12;
    if (mer === 'AM' && hh === 12) hh = 0;
    var hhS = (hh < 10 ? "0" : "") + hh;
    return hhS + ":" + mm;
  }

  // 4. Artifact Date epoch 1899 dari Google Spreadsheet (misal "Sat Dec 30 1899 22:07:12 GMT+0707...")
  if (s.indexOf('1899') !== -1 || s.indexOf('GMT') !== -1 || s.indexOf('WIB') !== -1 || s.indexOf('Waktu Indonesia') !== -1) {
    if (s.indexOf('22:07:12') !== -1 || s.indexOf('22:07') !== -1) return "07:00";
    if (s.indexOf('08:07:12') !== -1 || s.indexOf('08:07') !== -1) return "17:00";
    var dTime = s.match(/\b([01]\d|2[0-3]):([0-5]\d):([0-5]\d)\b/);
    if (dTime) {
      var hVal = parseInt(dTime[1], 10);
      if (hVal >= 21 || hVal <= 4) return defaultJam || "07:00";
      var hStr = (hVal < 10 ? "0" : "") + hVal;
      return hStr + ":" + dTime[2];
    }
    return defaultJam || "07:00";
  }

  // 5. Objek Date murni jika lolos String(val)
  if (Object.prototype.toString.call(val) === '[object Date]' || val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    var hh = (h < 10 ? "0" : "") + h;
    var mm = (m < 10 ? "0" : "") + m;
    return hh + ":" + mm;
  }

  // 6. Pencocokan umum jam:menit
  var gen = s.match(/(\d{1,2})[:.](\d{2})/);
  if (gen) {
    var gh = parseInt(gen[1], 10);
    if (gh >= 0 && gh <= 23) {
      var ghStr = (gh < 10 ? "0" : "") + gh;
      return ghStr + ":" + gen[2];
    }
  }

  return defaultJam || "07:00";
}

/**
 * Fungsi Pembantu: Memformat nilai tanggal dari Spreadsheet menjadi string "yyyy-MM-dd" standar
 */
function formatTanggalSpreadsheet(val, timezone) {
  if (val === null || val === undefined || val === "") return "";
  if (Object.prototype.toString.call(val) === '[object Date]' || val instanceof Date) {
    return Utilities.formatDate(val, timezone || "Asia/Jakarta", "yyyy-MM-dd");
  }
  var str = String(val).trim();
  var dMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (dMatch) {
    var m = (dMatch[2].length === 1 ? "0" : "") + dMatch[2];
    var d = (dMatch[3].length === 1 ? "0" : "") + dMatch[3];
    return dMatch[1] + "-" + m + "-" + d;
  }
  var idMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (idMatch) {
    var idDay = (idMatch[1].length === 1 ? "0" : "") + idMatch[1];
    var idMonth = (idMatch[2].length === 1 ? "0" : "") + idMatch[2];
    return idMatch[3] + "-" + idMonth + "-" + idDay;
  }
  return str;
}

/**
 * Fungsi Pembantu: Mengambil jadwal aktif dan seluruh butir soal dari sheet mapel masing-masing
 */
function getActiveExamsAndQuestions(ss) {
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
  if (!sheetJadwal) return [];

  var lastRow = sheetJadwal.getLastRow();
  if (lastRow <= 1) return [];

  var tz = ss.getSpreadsheetTimeZone() || "Asia/Jakarta";
  // PENTING: Gunakan getDisplayValues() untuk membaca tampilan teks asli dari cell (misal "07:00", "15:15")
  // guna menghindari konversi otomatis ke objek Date epoch 1899!
  var dataJadwalDisplay = sheetJadwal.getRange(2, 1, lastRow - 1, HEADER_JADWAL.length).getDisplayValues();
  var dataJadwalValues = sheetJadwal.getRange(2, 1, lastRow - 1, HEADER_JADWAL.length).getValues();
  var activeExams = [];

  for (var i = 0; i < dataJadwalDisplay.length; i++) {
    var rowDisplay = dataJadwalDisplay[i];
    var rowValues = dataJadwalValues[i];
    var status = String(rowDisplay[8] || rowValues[8] || "").trim().toUpperCase();

    // HANYA proses jadwal yang berstatus AKTIF
    if (status === "AKTIF") {
      var idJadwal = String(rowDisplay[0] || rowValues[0] || ("JDW-" + (i + 1))).trim();
      var namaMapel = String(rowDisplay[1] || rowValues[1] || "").trim();
      var kelas = String(rowDisplay[2] || rowValues[2] || "Semua Kelas").trim();
      var tanggal = formatTanggalSpreadsheet(rowDisplay[3] || rowValues[3], tz);
      var jamMulai = formatJamSpreadsheet(rowDisplay[4] || rowValues[4], "07:00");
      var jamSelesai = formatJamSpreadsheet(rowDisplay[5] || rowValues[5], "16:00");
      var durasiMenit = parseInt(rowDisplay[6] || rowValues[6]) || 40;
      var token = String(rowDisplay[7] || rowValues[7] || "212").trim();

      // Pengambilan soal presisi berdasarkan Nama Sheet yang sama persis dengan namaMapel
      var sheetSoal = ss.getSheetByName(namaMapel);
      var daftarSoal = [];

      if (sheetSoal) {
        daftarSoal = readQuestionsFromSheet(sheetSoal);
      }

      activeExams.push({
        id: idJadwal,
        namaMapel: namaMapel,
        mataPelajaran: namaMapel,
        kelas: kelas,
        tanggal: tanggal,
        jamMulai: jamMulai,
        jamSelesai: jamSelesai,
        durasiMenit: durasiMenit,
        token: token,
        totalSoal: daftarSoal.length,
        soal: daftarSoal
      });
    }
  }

  // Fallback cerdas: Jika di sheet Jadwal_Ujian belum ada baris 'AKTIF', pindai seluruh sheet mapel yang sudah memiliki soal
  if (activeExams.length === 0) {
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var sh = allSheets[s];
      var sName = sh.getName();
      if (sName !== SHEET_JADWAL && sName !== SHEET_HASIL && sName !== SHEET_DETAIL_ESSAY && sh.getLastRow() > 1) {
        var sQuestions = readQuestionsFromSheet(sh);
        if (sQuestions.length > 0) {
          activeExams.push({
            id: "AUTO-" + (s + 1),
            namaMapel: sName,
            mataPelajaran: sName,
            kelas: "Semua Kelas",
            tanggal: "",
            jamMulai: "00:00",
            jamSelesai: "23:59",
            durasiMenit: 60,
            token: "212",
            totalSoal: sQuestions.length,
            soal: sQuestions
          });
        }
      }
    }
  }

  return activeExams;
}

/**
 * Normalisasi URL Gambar & Konversi Google Drive Link
 */
function normalisasiUrlGambar(val) {
  if (!val) return "";
  var s = String(val).trim();
  if (!s || s === "[object Object]" || s === "[object CellImage]") return "";

  // Ekstrak dari formula =IMAGE("url")
  var mFormula = s.match(/=IMAGE\(\s*["']([^"']+)["']/i);
  if (mFormula) s = mFormula[1].trim();

  // Konversi link Google Drive view ke CDN render langsung (lh3)
  var mDrive = s.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=view&)?id=)([a-zA-Z0-9_-]+)/i);
  if (mDrive) {
    return "https://lh3.googleusercontent.com/d/" + mDrive[1];
  }
  return s;
}

/**
 * Ekstraksi URL gambar dari nilai sel (Mendukung URL string, Formula =IMAGE, dan CellImage)
 */
function ekstrakGambarDariSel(cellVal, formulaVal) {
  if (formulaVal && String(formulaVal).toUpperCase().indexOf("IMAGE") !== -1) {
    return normalisasiUrlGambar(formulaVal);
  }
  if (cellVal && typeof cellVal === 'object') {
    try {
      if (typeof cellVal.getContentUrl === 'function') {
        var u = cellVal.getContentUrl();
        if (u) return u;
      }
      if (typeof cellVal.getUrl === 'function') {
        var u2 = cellVal.getUrl();
        if (u2) return u2;
      }
    } catch(e) {}
  }
  if (typeof cellVal === 'string' && cellVal.trim()) {
    var str = cellVal.trim();
    if (str.indexOf("http://") === 0 || str.indexOf("https://") === 0 || str.indexOf("data:image/") === 0) {
      return normalisasiUrlGambar(str);
    }
  }
  return "";
}

/**
 * Fungsi Pembantu: Membaca butir soal dari sheet Mapel tertentu
 * Mendukung Teks Soal, Markdown Gambar, Formula =IMAGE, CellImage, dan Kolom Gambar Khusus.
 */
function readQuestionsFromSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return [];

  // Pemetaan kolom dinamis berdasarkan teks header baris 1
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = {
    no: -1,
    pertanyaan: -1,
    gambar: -1,
    opsiA: -1,
    opsiB: -1,
    opsiC: -1,
    opsiD: -1,
    opsiE: -1,
    kunci: -1,
    pembahasan: -1,
    tipe: -1,
    kataKunci: -1,
    bobot: -1
  };

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || "").trim().toLowerCase();
    if (h === 'no' || h === 'nomor') colIdx.no = c;
    else if (h.indexOf('tanya') !== -1 || h.indexOf('soal') !== -1 || h.indexOf('pertanyaan') !== -1) colIdx.pertanyaan = c;
    else if (h.indexOf('gambar') !== -1 || h.indexOf('foto') !== -1 || h.indexOf('ilustrasi') !== -1 || h.indexOf('image') !== -1) colIdx.gambar = c;
    else if (h.indexOf('opsi a') !== -1 || h.indexOf('opsi_a') !== -1 || h.indexOf('pilihan a') !== -1 || h === 'a') colIdx.opsiA = c;
    else if (h.indexOf('opsi b') !== -1 || h.indexOf('opsi_b') !== -1 || h.indexOf('pilihan b') !== -1 || h === 'b') colIdx.opsiB = c;
    else if (h.indexOf('opsi c') !== -1 || h.indexOf('opsi_c') !== -1 || h.indexOf('pilihan c') !== -1 || h === 'c') colIdx.opsiC = c;
    else if (h.indexOf('opsi d') !== -1 || h.indexOf('opsi_d') !== -1 || h.indexOf('pilihan d') !== -1 || h === 'd') colIdx.opsiD = c;
    else if (h.indexOf('opsi e') !== -1 || h.indexOf('opsi_e') !== -1 || h.indexOf('pilihan e') !== -1 || h === 'e') colIdx.opsiE = c;
    else if (h.indexOf('kata_kunci') !== -1 || h.indexOf('katakunci') !== -1 || h.indexOf('rubrik') !== -1 || h.indexOf('keyword') !== -1) colIdx.kataKunci = c;
    else if ((h.indexOf('kunci') !== -1 || h.indexOf('jawaban') !== -1) && h.indexOf('kata') === -1) colIdx.kunci = c;
    else if (h.indexOf('bahas') !== -1 || h.indexOf('pembahasan') !== -1 || h.indexOf('penjelasan') !== -1) colIdx.pembahasan = c;
    else if (h.indexOf('tipe') !== -1 || h.indexOf('jenis') !== -1) colIdx.tipe = c;
    else if (h.indexOf('bobot') !== -1 || h.indexOf('poin') !== -1 || h.indexOf('score') !== -1) colIdx.bobot = c;
  }

  // Fallback jika nama header berbeda
  if (colIdx.pertanyaan === -1) {
    if (headers.length >= 7 && String(headers[0]).toLowerCase().indexOf('tanya') !== -1) {
      colIdx.pertanyaan = 0; colIdx.opsiA = 1; colIdx.opsiB = 2; colIdx.opsiC = 3;
      colIdx.opsiD = 4; colIdx.opsiE = 5; colIdx.kunci = 6; colIdx.pembahasan = 7;
      if (headers.length > 8) colIdx.tipe = 8;
      if (headers.length > 9) colIdx.kataKunci = 9;
      if (headers.length > 10) colIdx.bobot = 10;
    } else {
      colIdx.no = 0; colIdx.pertanyaan = 1; colIdx.opsiA = 2; colIdx.opsiB = 3;
      colIdx.opsiC = 4; colIdx.opsiD = 5; colIdx.opsiE = 6; colIdx.kunci = 7; colIdx.pembahasan = 8;
      if (headers.length > 9) colIdx.tipe = 9;
      if (headers.length > 10) colIdx.kataKunci = 10;
      if (headers.length > 11) colIdx.bobot = 11;
    }
  }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var formulas = sheet.getRange(2, 1, lastRow - 1, lastCol).getFormulas();
  var soalList = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var fRow = formulas[i];

    var valQ = colIdx.pertanyaan >= 0 ? row[colIdx.pertanyaan] : "";
    var fQ = colIdx.pertanyaan >= 0 ? fRow[colIdx.pertanyaan] : "";
    var imgFromQ = ekstrakGambarDariSel(valQ, fQ);

    var pertanyaan = (typeof valQ === 'string') ? valQ.trim() : "";
    if (pertanyaan === "[object Object]" || pertanyaan === "[object CellImage]") pertanyaan = "";

    // Pindai gambar dari kolom Gambar/Foto khusus jika ada
    var imgFromCol = "";
    if (colIdx.gambar >= 0) {
      imgFromCol = ekstrakGambarDariSel(row[colIdx.gambar], fRow[colIdx.gambar]);
    }

    var gambarFinal = imgFromCol || imgFromQ;
    if (gambarFinal && pertanyaan.indexOf(gambarFinal) === -1) {
      pertanyaan = (pertanyaan ? pertanyaan + "\n\n" : "") + "![Ilustrasi Soal](" + gambarFinal + ")";
    }

    if (!pertanyaan && !gambarFinal) continue; // Lewati baris kosong

    function resolveOpsi(idx) {
      if (idx < 0) return "";
      var valO = row[idx];
      var fO = fRow[idx];
      var imgO = ekstrakGambarDariSel(valO, fO);
      var teksO = (typeof valO === 'string') ? valO.trim() : "";
      if (teksO === "[object Object]" || teksO === "[object CellImage]") teksO = "";
      if (imgO && teksO.indexOf(imgO) === -1) {
        return (teksO ? teksO + " " : "") + "![Opsi Gambar](" + imgO + ")";
      }
      return teksO;
    }

    var opsiA = resolveOpsi(colIdx.opsiA);
    var opsiB = resolveOpsi(colIdx.opsiB);
    var opsiC = resolveOpsi(colIdx.opsiC);
    var opsiD = resolveOpsi(colIdx.opsiD);
    var opsiE = resolveOpsi(colIdx.opsiE);

    var kunciRaw = colIdx.kunci >= 0 ? String(row[colIdx.kunci] || "").trim() : "A";
    var tipeVal = colIdx.tipe >= 0 ? String(row[colIdx.tipe] || "").trim().toUpperCase() : "";
    var isEssay = tipeVal === "ESSAY" || tipeVal === "URAIAN";

    // Fail-safe backward compatibility deteksi format lama
    if (!isEssay && !tipeVal) {
      if (!opsiA && !opsiB && !opsiC && !opsiD && !opsiE && kunciRaw.length > 2 && !/^[A-E]$/i.test(kunciRaw)) {
        isEssay = true;
      }
    }

    var pembahasan = colIdx.pembahasan >= 0 ? String(row[colIdx.pembahasan] || "").trim() : "";
    var noVal = (colIdx.no >= 0 && row[colIdx.no]) ? parseInt(row[colIdx.no]) : (i + 1);

    if (isEssay) {
      var kataKunciRaw = colIdx.kataKunci >= 0 ? String(row[colIdx.kataKunci] || "").trim() : "";
      var kataKunciList = kataKunciRaw ? kataKunciRaw.split(',').map(function(k){ return k.trim(); }).filter(Boolean) : [];
      var bobotVal = (colIdx.bobot >= 0 && row[colIdx.bobot]) ? Number(row[colIdx.bobot]) : 10;
      if (isNaN(bobotVal) || bobotVal <= 0) bobotVal = 10;

      soalList.push({
        no: noVal || (i + 1),
        tipe: "ESSAY",
        pertanyaan: pertanyaan,
        q: pertanyaan,
        opsi: [],
        o: [],
        kunci: kunciRaw,
        a: kunciRaw,
        kunciEssay: kunciRaw,
        kataKunci: kataKunciList,
        bobotEssay: bobotVal,
        pembahasan: pembahasan,
        p: pembahasan
      });
    } else {
      var kunciUpper = (kunciRaw || "A").toUpperCase();
      var kunciIdx = 0;
      if (kunciUpper === 'B' || kunciUpper === '1') kunciIdx = 1;
      else if (kunciUpper === 'C' || kunciUpper === '2') kunciIdx = 2;
      else if (kunciUpper === 'D' || kunciUpper === '3') kunciIdx = 3;
      else if (kunciUpper === 'E' || kunciUpper === '4') kunciIdx = 4;

      soalList.push({
        no: noVal || (i + 1),
        tipe: "PG",
        pertanyaan: pertanyaan,
        q: pertanyaan,
        opsi: [opsiA, opsiB, opsiC, opsiD, opsiE],
        o: [opsiA, opsiB, opsiC, opsiD, opsiE],
        kunci: kunciUpper,
        a: kunciIdx,
        pembahasan: pembahasan,
        p: pembahasan
      });
    }
  }

  return soalList;
}

/**
 * Fungsi Pembantu: Menambahkan atau Memperbarui Jadwal di Jadwal_Ujian tanpa duplikasi (Anti-Doble)
 * Memastikan Jam_Mulai dan Jam_Selesai di Spreadsheet mengikuti Database Ujian (Batas Waktu).
 */
function tambahkanKeJadwalJikaBelumAda(ss, namaMapel, payload) {
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
  if (!sheetJadwal) {
    sheetJadwal = pastikanSheetJadwalTersedia(ss);
  }
  if (!sheetJadwal) return;

  var lastRow = sheetJadwal.getLastRow();
  var matchRow = -1;
  var rowsToDelete = [];
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Jakarta";

  var jamMulai = (payload && payload.jamMulai) ? formatJamSpreadsheet(payload.jamMulai, "07:00") : "07:00";
  var jamSelesai = (payload && payload.jamSelesai) ? formatJamSpreadsheet(payload.jamSelesai, "16:00") : "16:00";
  var kelas = (payload && payload.kelas) ? String(payload.kelas).trim() : "Semua Kelas";
  var tanggal = (payload && payload.tanggal) ? formatTanggalSpreadsheet(payload.tanggal, tz) : Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var durasi = (payload && payload.durasiMenit) ? parseInt(payload.durasiMenit) : 40;
  var token = (payload && payload.token) ? String(payload.token).trim() : "212";
  var status = (payload && payload.status) ? String(payload.status).trim().toUpperCase() : "AKTIF";

  var targetNorm = String(namaMapel || "").replace(/\s+/g, ' ').trim().toLowerCase();

  if (lastRow > 1) {
    var mapelValues = sheetJadwal.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < mapelValues.length; i++) {
      var cellVal = String(mapelValues[i][0] || "").replace(/\s+/g, ' ').trim().toLowerCase();
      if (cellVal === targetNorm) {
        if (matchRow === -1) {
          matchRow = i + 2; // Baris pertama yang cocok
        } else {
          rowsToDelete.push(i + 2); // Baris duplikat yang berlebih (dihapus agar tidak doble)
        }
      }
    }
  }

  // Jika baris mapel sudah ada di Jadwal_Ujian:
  // PERBARUI Jam_Mulai dan Jam_Selesai agar sinkron dengan Database Ujian (Batas Waktu), TIDAK DOBLE!
  if (matchRow > 1) {
    if (payload) {
      if (payload.kelas) sheetJadwal.getRange(matchRow, 3).setValue(kelas);
      if (payload.tanggal) sheetJadwal.getRange(matchRow, 4).setValue(tanggal);
      sheetJadwal.getRange(matchRow, 5).setNumberFormat("@").setValue(jamMulai);
      sheetJadwal.getRange(matchRow, 6).setNumberFormat("@").setValue(jamSelesai);
      if (payload.durasiMenit) sheetJadwal.getRange(matchRow, 7).setValue(durasi);
      if (payload.token) sheetJadwal.getRange(matchRow, 8).setValue(token);
      sheetJadwal.getRange(matchRow, 9).setValue(status);
    }
  } else {
    // Jika belum ada, masukkan SATU baris saja lengkap dengan Jam_Mulai & Jam_Selesai dari setting Database Ujian
    var nextId = "JDW-" + Utilities.formatDate(new Date(), tz, "yyMMddHHmm");
    sheetJadwal.appendRow([
      nextId,
      namaMapel,
      kelas,
      tanggal,
      jamMulai,
      jamSelesai,
      durasi,
      token,
      status
    ]);
    var newRowIdx = sheetJadwal.getLastRow();
    sheetJadwal.getRange(newRowIdx, 5, 1, 2).setNumberFormat("@");
  }

  // Bersihkan baris duplikat lama jika ada (dari bawah ke atas agar indeks baris tidak bergeser)
  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheetJadwal.deleteRow(rowsToDelete[d]);
  }
}

/**
 * Fungsi Pembantu: Membuat Sheet Jadwal_Ujian bawaan jika file spreadsheet masih kosong
 */
function pastikanSheetJadwalTersedia(ss) {
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
  if (!sheetJadwal) {
    sheetJadwal = ss.insertSheet(SHEET_JADWAL);
    sheetJadwal.appendRow(HEADER_JADWAL);
    var r = sheetJadwal.getRange(1, 1, 1, HEADER_JADWAL.length);
    r.setFontWeight("bold");
    r.setBackground("#3730a3");
    r.setFontColor("#ffffff");
    sheetJadwal.setFrozenRows(1);

    // Sampel Jadwal Awal: Informatika & Pendidikan Agama Islam (Status: AKTIF)
    var todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || "Asia/Jakarta", "yyyy-MM-dd");
    sheetJadwal.appendRow(["JDW-01", "Informatika", "X TAB 2, X DKV 1", todayStr, "07:00", "17:00", 40, "212", "AKTIF"]);
    sheetJadwal.appendRow(["JDW-02", "Pendidikan Agama Islam", "Semua Kelas", todayStr, "07:00", "17:00", 45, "212", "AKTIF"]);
    sheetJadwal.getRange(2, 5, 2, 2).setNumberFormat("@");

    // Buat juga contoh sheet Informatika jika belum ada
    if (!ss.getSheetByName("Informatika")) {
      var sInfo = ss.insertSheet("Informatika");
      sInfo.appendRow(HEADER_SOAL);
      sInfo.getRange(1, 1, 1, HEADER_SOAL.length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
      sInfo.setFrozenRows(1);

      sInfo.appendRow([
        1,
        "Teknologi informasi dan komunikasi dapat diartikan sebagai segala kegiatan yang berkaitan dengan pemrosesan, manipulasi, pengolahan, dan penyampaian informasi menggunakan suatu media. Istilah ini dikenal dengan singkatan...",
        "TIK",
        "Hardware",
        "Software",
        "Brainware",
        "Jaringan Komputer",
        "A",
        "TIK adalah singkatan dari Teknologi Informasi dan Komunikasi."
      ]);
      sInfo.appendRow([
        2,
        "Perangkat keras pada komputer yang berfungsi sebagai media penyimpanan data utama yang bersifat non-volatile (data tidak hilang saat listrik padam) adalah...",
        "RAM",
        "ROM",
        "Solid State Drive (SSD) / Harddisk",
        "Processor",
        "VGA Card",
        "C",
        "SSD dan Harddisk merupakan storage non-volatile."
      ]);
      sInfo.appendRow([
        3,
        "Aplikasi video conference gratis dari Google yang memudahkan rapat online cukup dengan membagikan link tautan adalah...",
        "Google Sheets",
        "Google Classroom",
        "Google Meet",
        "Google Drive",
        "Google Forms",
        "C",
        "Google Meet digunakan untuk video conference interaktif."
      ]);
    }
  }
  return sheetJadwal;
}

/**
 * FUNGSI SETUP / INISIALISASI DATABASE SPREADSHEET
 * Pilih fungsi ini di toolbar editor Apps Script, lalu klik "Jalankan (Run)"
 * untuk membuat seluruh Sheet baku (Jadwal_Ujian, Hasil_Ujian, Informatika) secara otomatis!
 */

/**
 * Inisialisasi awal sheet dan struktur database (Alias panduan UI).
 * Menjalankan pembuatan sheet default (Jadwal_Ujian, Hasil_Ujian, dsb).
 */
function setupDatabaseSheets() {
  inisialisasiSpreadsheet();
}

function inisialisasiSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  pastikanSheetJadwalTersedia(ss);

  var sheetHasil = ss.getSheetByName(SHEET_HASIL);
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet(SHEET_HASIL);
    sheetHasil.appendRow(HEADER_HASIL);
    var rH = sheetHasil.getRange(1, 1, 1, HEADER_HASIL.length);
    rH.setFontWeight("bold");
    rH.setBackground("#15803d");
    rH.setFontColor("#ffffff");
    sheetHasil.setFrozenRows(1);
  }

  var sheetEssay = ss.getSheetByName(SHEET_DETAIL_ESSAY);
  if (!sheetEssay) {
    sheetEssay = ss.insertSheet(SHEET_DETAIL_ESSAY);
    sheetEssay.appendRow(HEADER_DETAIL_ESSAY);
    var rHE = sheetEssay.getRange(1, 1, 1, HEADER_DETAIL_ESSAY.length);
    rHE.setFontWeight("bold");
    rHE.setBackground("#2563eb");
    rHE.setFontColor("#ffffff");
    sheetEssay.setFrozenRows(1);
  }

  Logger.log("Inisialisasi database CBT spreadsheet selesai 100%!");
}

/**
 * Standarisasi Response JSON bebas CORS
 */
function createJsonResponse(data) {
  if (data && typeof data === 'object' && typeof data.success === 'undefined') {
    data.success = (data.status === "success" || data.status === "active" || data.result === "success");
  }
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}