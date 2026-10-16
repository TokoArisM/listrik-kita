/**
 * =========================================================================
 * LISTRIK KITA - POS BACKEND V2 (Google Apps Script)
 * =========================================================================
 * Fitur Tambahan:
 * 1. Manajemen Produk (Tambah, Edit, Hapus)
 * 2. Ambil Riwayat Transaksi & Ringkasan Omzet
 * 3. Dukungan Multi Metode Pembayaran, Diskon, dan Nama Pelanggan
 * 4. Proteksi Keamanan Token / PIN Toko
 * =========================================================================
 */

// =========================================================================
// KEAMANAN TOKO: Token / PIN Rahasia
// Ubah PIN ini sesuka Anda. Hanya aplikasi kasir yang memiliki PIN ini
// yang dapat melihat data produk, omzet, dan melakukan transaksi.
// =========================================================================
const SECURITY_TOKEN = "LK-LISTRIK-KITA-2026";

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Produk
  let sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet("Produk");
    sheetProduk.appendRow(["ID", "Nama", "Kategori", "Harga", "Stok", "Satuan", "HargaModal", "Foto"]);
    sheetProduk.appendRow(["P001", "Lampu LED Philips 10W Putih", "Lampu", 35000, 50, "pcs", 27000, ""]);
    sheetProduk.appendRow(["P002", "Stop Kontak Broco 4L + Kabel 3M", "Kabel & Stopkontak", 65000, 20, "pcs", 50000, ""]);
    sheetProduk.appendRow(["P003", "Kabel NYM Eterna 2x1.5mm (Meter)", "Kabel & Stopkontak", 12000, 100, "meter", 9500, ""]);
    sheetProduk.appendRow(["P004", "MCB Schneider 10A 1 Phase", "MCB & Box", 55000, 15, "pcs", 42000, ""]);
    sheetProduk.appendRow(["P005", "Saklar Ganda Panasonic Seri WESJ", "Saklar", 38000, 30, "pcs", 29000, ""]);
    sheetProduk.appendRow(["P006", "Isolasi Listrik Unibell Hitam", "Aksesoris", 7000, 80, "roll", 4500, ""]);
    sheetProduk.appendRow(["P007", "Fitting Lampu Plafon Broco Standard", "Lampu", 15000, 45, "pcs", 10000, ""]);
    sheetProduk.appendRow(["P008", "Tespen Listrik Obeng Bolak Balik", "Alat & Tool", 18000, 25, "pcs", 12000, ""]);
  } else {
    // Tambah header Foto otomatis jika belum ada di kolom 8
    const lastCol = sheetProduk.getLastColumn();
    if (lastCol < 8) {
      sheetProduk.getRange(1, 8).setValue("Foto");
    }
  }

  // 2. Setup Sheet Transaksi
  let sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) {
    sheetTransaksi = ss.insertSheet("Transaksi");
    sheetTransaksi.appendRow([
      "ID Transaksi", 
      "Waktu", 
      "Pelanggan", 
      "Detail Item", 
      "Subtotal",
      "Diskon",
      "Total Akhir", 
      "Bayar", 
      "Kembalian", 
      "Metode Pembayaran",
      "RawItemsJSON"
    ]);
  }
}

/**
 * doGet() - Menangani:
 * - ?action=products (default: daftar produk & stok)
 * - ?action=transactions (daftar riwayat transaksi)
 */
function doGet(e) {
  try {
    const token = (e && e.parameter && e.parameter.token) || "";
    if (SECURITY_TOKEN && token !== SECURITY_TOKEN) {
      return createJsonResponse({
        status: "error",
        message: "Akses ditolak: PIN / Token keamanan toko tidak valid."
      });
    }

    const action = (e && e.parameter && e.parameter.action) || "products";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "transactions") {
      return getTransactionsData(ss);
    }

    // Default: Ambil data produk
    const sheetProduk = ss.getSheetByName("Produk");
    if (!sheetProduk) {
      return createJsonResponse({
        status: "error",
        message: "Sheet 'Produk' tidak ditemukan. Jalankan setupDatabase() terlebih dahulu."
      });
    }

    const data = sheetProduk.getDataRange().getValues();
    const products = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      products.push({
        id: String(row[0]).trim(),
        nama: String(row[1] || "").trim(),
        kategori: String(row[2] || "Lainnya").trim(),
        harga: Number(row[3]) || 0,
        stok: Number(row[4]) || 0,
        satuan: String(row[5] || "pcs").trim(),
        hargaModal: Number(row[6]) || 0,
        foto: String(row[7] || "").trim()
      });
    }

    return createJsonResponse({
      status: "success",
      count: products.length,
      data: products
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * Mengambil maksimal 100 transaksi terakhir & ringkasan omzet (Optimized O(1) Read)
 */
function getTransactionsData(ss) {
  const sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) {
    return createJsonResponse({ status: "success", data: [], summary: { totalOmzet: 0, count: 0 } });
  }

  const lastRow = sheetTransaksi.getLastRow();
  if (lastRow <= 1) {
    return createJsonResponse({ status: "success", data: [], summary: { totalOmzet: 0, count: 0 } });
  }

  // Cukup ambil maksimal 100 baris terakhir, tidak membaca seluruh sheet
  const numRows = Math.min(lastRow - 1, 100);
  const startRow = lastRow - numRows + 1;
  const values = sheetTransaksi.getRange(startRow, 1, numRows, 11).getValues();

  const transactions = [];
  let totalOmzet = 0;

  // Baca dari baris paling baru ke baris lama
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (!row[0]) continue;

    const totalAkhir = Number(row[6]) || 0;
    totalOmzet += totalAkhir;

    let itemsParsed = [];
    if (row[10]) {
      try {
        itemsParsed = JSON.parse(row[10]);
      } catch (e) {}
    }

    transactions.push({
      transactionId: String(row[0]),
      waktu: String(row[1]),
      pelanggan: String(row[2] || "Umum"),
      detailItem: String(row[3]),
      subtotal: Number(row[4]) || 0,
      diskon: Number(row[5]) || 0,
      total: totalAkhir,
      bayar: Number(row[7]) || 0,
      kembalian: Number(row[8]) || 0,
      metode: String(row[9] || "Tunai"),
      items: itemsParsed
    });
  }

  return createJsonResponse({
    status: "success",
    count: transactions.length,
    data: transactions,
    summary: {
      totalOmzet: totalOmzet,
      count: transactions.length
    }
  });
}

/**
 * doPost() - Menangani:
 * - action: 'checkout' (default) -> Catat transaksi & potong stok
 * - action: 'save_product' -> Tambah / Update produk
 * - action: 'delete_product' -> Hapus produk
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(15000);
  if (!hasLock) {
    return createJsonResponse({ status: "error", message: "Server sedang sibuk, ulangi dalam beberapa detik." });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "Payload request kosong." });
    }

    const postData = JSON.parse(e.postData.contents);
    const token = (postData && postData.token) || (e.parameter && e.parameter.token) || "";
    if (SECURITY_TOKEN && token !== SECURITY_TOKEN) {
      return createJsonResponse({
        status: "error",
        message: "Akses ditolak: PIN / Token keamanan toko tidak valid."
      });
    }

    const action = postData.action || "checkout";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "save_product") {
      return handleSaveProduct(ss, postData);
    } else if (action === "restock_product") {
      return handleRestockProduct(ss, postData);
    } else if (action === "delete_product") {
      return handleDeleteProduct(ss, postData);
    } else if (action === "restore_product") {
      return handleRestoreProduct(ss, postData);
    } else if (action === "pay_debt") {
      return handlePayDebt(ss, postData);
    } else {
      return handleCheckout(ss, postData);
    }
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle Checkout Transaksi (Optimized Batch Update)
 */
function handleCheckout(ss, data) {
  const {
    transactionId,
    items,
    subtotal,
    diskon,
    total,
    bayar,
    kembalian,
    metodePembayaran,
    namaPelanggan
  } = data;

  if (!items || !items.length) {
    return createJsonResponse({ status: "error", message: "Keranjang belanja kosong." });
  }

  const sheetTransaksi = ss.getSheetByName("Transaksi");
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetTransaksi || !sheetProduk) {
    throw new Error("Sheet Transaksi atau Produk belum disiapkan.");
  }

  const now = new Date();
  const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
  const invoiceId = transactionId || ("TRX-" + Utilities.formatDate(now, "Asia/Jakarta", "yyyyMMdd-HHmmss"));
  const summaryItems = items.map(it => `${it.nama} (${it.qty}x @${it.harga}) = Rp${it.qty * it.harga}`).join("\n");

  // 1. Simpan ke sheet Transaksi
  sheetTransaksi.appendRow([
    invoiceId,
    timestampStr,
    namaPelanggan || "Umum",
    summaryItems,
    Number(subtotal) || Number(total) || 0,
    Number(diskon) || 0,
    Number(total) || 0,
    Number(bayar) || 0,
    Number(kembalian) || 0,
    metodePembayaran || "Tunai",
    JSON.stringify(items)
  ]);

  // 2. Potong stok otomatis di sheet Produk dengan 1 kali Batch Write (sangat cepat)
  const lastProdRow = sheetProduk.getLastRow();
  if (lastProdRow > 1) {
    const numProdRows = lastProdRow - 1;
    // Ambil kolom ID (kolom 1) dan Stok (kolom 5)
    const idList = sheetProduk.getRange(2, 1, numProdRows, 1).getValues();
    const stockRange = sheetProduk.getRange(2, 5, numProdRows, 1);
    const stockValues = stockRange.getValues();

    const idMap = {};
    for (let r = 0; r < idList.length; r++) {
      const pid = String(idList[r][0]).trim();
      if (pid) idMap[pid] = r;
    }

    let hasStockChange = false;
    for (const item of items) {
      const rowIdx = idMap[String(item.id).trim()];
      if (rowIdx !== undefined) {
        const curStok = Number(stockValues[rowIdx][0]) || 0;
        stockValues[rowIdx][0] = Math.max(0, curStok - (Number(item.qty) || 0));
        hasStockChange = true;
      }
    }

    // Tulis kembali seluruh stok sekaligus hanya dengan 1 panggilan API
    if (hasStockChange) {
      stockRange.setValues(stockValues);
    }
  }

  return createJsonResponse({
    status: "success",
    message: "Transaksi berhasil dicatat dan stok dipotong.",
    transactionId: invoiceId
  });
}

/**
 * Simpan base64 foto ke folder Google Drive toko dan kembalikan URL langsungnya
 */
function saveImageToDrive(base64Data, filename) {
  try {
    const folderName = "FOTO_PRODUK_LISTRIK_KITA";
    let folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    let contentType = "image/webp";
    let stripped = base64Data;
    if (base64Data.indexOf(",") > -1) {
      const parts = base64Data.split(",");
      const match = parts[0].match(/:(.*?);/);
      if (match) contentType = match[1];
      stripped = parts[1];
    }
    const decoded = Utilities.base64Decode(stripped);
    const blob = Utilities.newBlob(decoded, contentType, filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Gunakan URL Google User Content yang sangat cepat dimuat
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (err) {
    Logger.log("Gagal simpan ke Drive: " + err.toString());
    return "";
  }
}

/**
 * Handle Tambah / Update Produk (Optimized Batch Update & Logging)
 */
function handleSaveProduct(ss, payload) {
  const { id, nama, kategori, harga, stok, satuan, hargaModal, foto, fotoBase64, isEdit } = payload;
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) throw new Error("Sheet Produk tidak ditemukan.");

  // Proses upload foto ke Drive jika ada payload fotoBase64 baru
  let finalFotoUrl = foto || "";
  if (fotoBase64 && String(fotoBase64).startsWith("data:image")) {
    const targetFilename = `PROD_${(id || "NEW").replace(/[^a-zA-Z0-9]/g, "_")}_${Date.now()}.webp`;
    const uploadedUrl = saveImageToDrive(fotoBase64, targetFilename);
    if (uploadedUrl) {
      finalFotoUrl = uploadedUrl;
    }
  }

  const values = sheetProduk.getDataRange().getValues();
  let targetRow = -1;
  let maxPidNum = 0;
  let oldStock = 0;

  for (let r = 1; r < values.length; r++) {
    const existingId = String(values[r][0]).trim();
    if (existingId === String(id).trim()) {
      targetRow = r + 1;
      oldStock = Number(values[r][4]) || 0;
      break;
    }
    const match = existingId.match(/^P(\d+)$/i);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxPidNum) maxPidNum = n;
    }
  }

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

  if (targetRow > 0) {
    // Update data produk sekaligus dalam 1 batch call (kolom 2 s/d 8)
    sheetProduk.getRange(targetRow, 2, 1, 7).setValues([[
      nama,
      kategori || "Lainnya",
      Number(harga) || 0,
      Number(stok) || 0,
      satuan || "pcs",
      Number(hargaModal) || 0,
      finalFotoUrl
    ]]);

    // Log jika ada perubahan stok saat edit
    if (oldStock !== Number(stok)) {
      logStockChange(ss, {
        waktu: nowStr,
        id: id,
        nama: nama,
        stokLama: oldStock,
        perubahan: Number(stok) - oldStock,
        stokBaru: Number(stok),
        keterangan: "Edit info produk"
      });
    }

    return createJsonResponse({ 
      status: "success", 
      message: `Produk ${id} berhasil diperbarui.`,
      fotoUrl: finalFotoUrl
    });
  } else {
    // Tambah baris baru dengan ID anti-bentrok
    const newId = id || ("P" + String(maxPidNum + 1).padStart(3, "0"));
    sheetProduk.appendRow([
      newId,
      nama,
      kategori || "Lainnya",
      Number(harga) || 0,
      Number(stok) || 0,
      satuan || "pcs",
      Number(hargaModal) || 0,
      finalFotoUrl
    ]);

    logStockChange(ss, {
      waktu: nowStr,
      id: newId,
      nama: nama,
      stokLama: 0,
      perubahan: Number(stok) || 0,
      stokBaru: Number(stok) || 0,
      keterangan: "Produk Baru Ditambahkan"
    });

    return createJsonResponse({ 
      status: "success", 
      message: `Produk ${newId} berhasil ditambahkan.`,
      fotoUrl: finalFotoUrl
    });
  }
}

/**
 * Handle Restock / Tambah Stok Cepat (Kulakan)
 */
function handleRestockProduct(ss, payload) {
  const { id, addQty, newHargaModal, keterangan } = payload;
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) throw new Error("Sheet Produk tidak ditemukan.");

  const values = sheetProduk.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(id).trim()) {
      const oldStock = Number(values[r][4]) || 0;
      const addition = Number(addQty) || 0;
      const newStock = Math.max(0, oldStock + addition);
      const nama = String(values[r][1]);

      sheetProduk.getRange(r + 1, 5).setValue(newStock);

      if (Number(newHargaModal) > 0) {
        sheetProduk.getRange(r + 1, 7).setValue(Number(newHargaModal));
      }

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
      logStockChange(ss, {
        waktu: nowStr,
        id: id,
        nama: nama,
        stokLama: oldStock,
        perubahan: addition,
        stokBaru: newStock,
        keterangan: keterangan || "Restock / Kulakan"
      });

      return createJsonResponse({ 
        status: "success", 
        message: `Stok ${nama} bertambah +${addition} (Total: ${newStock}).`,
        newStock: newStock
      });
    }
  }

  return createJsonResponse({ status: "error", message: `Produk ${id} tidak ditemukan.` });
}

/**
 * Handle Hapus Produk dengan Soft Delete (Pindah ke Sheet 'Produk_Sampah')
 */
function handleDeleteProduct(ss, payload) {
  const { id, reason } = payload;
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) throw new Error("Sheet Produk tidak ditemukan.");

  let sheetSampah = ss.getSheetByName("Produk_Sampah");
  if (!sheetSampah) {
    sheetSampah = ss.insertSheet("Produk_Sampah");
    sheetSampah.appendRow(["ID", "Nama Produk", "Kategori", "Harga Jual", "Stok Terakhir", "Satuan", "Harga Modal", "Foto", "Waktu Dihapus", "Alasan"]);
    sheetSampah.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#fee2e2");
  }

  const values = sheetProduk.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(id).trim()) {
      const rowData = values[r];
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

      // Simpan cadangan ke Produk_Sampah sebelum dihapus dari katalog utama
      sheetSampah.appendRow([
        rowData[0], // ID
        rowData[1], // Nama
        rowData[2], // Kategori
        rowData[3], // Harga
        rowData[4], // Stok
        rowData[5], // Satuan
        rowData[6] || 0, // Modal
        rowData[7] || "", // Foto
        nowStr,
        reason || "Dihapus oleh pengguna"
      ]);

      // Catat log
      logStockChange(ss, {
        waktu: nowStr,
        id: id,
        nama: String(rowData[1]),
        stokLama: Number(rowData[4]) || 0,
        perubahan: -(Number(rowData[4]) || 0),
        stokBaru: 0,
        keterangan: "Dihapus ke Tong Sampah"
      });

      // Hapus dari sheet Produk utama
      sheetProduk.deleteRow(r + 1);
      return createJsonResponse({ 
        status: "success", 
        message: `Produk ${id} berhasil dipindahkan ke Tong Sampah (Aman & dapat dipulihkan).` 
      });
    }
  }

  return createJsonResponse({ status: "error", message: `Produk dengan ID ${id} tidak ditemukan.` });
}

/**
 * Handle Pulihkan Produk dari Tong Sampah (Restore)
 */
function handleRestoreProduct(ss, payload) {
  const { id } = payload;
  const sheetSampah = ss.getSheetByName("Produk_Sampah");
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetSampah || !sheetProduk) throw new Error("Sheet Produk atau Produk_Sampah tidak ditemukan.");

  const values = sheetSampah.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(id).trim()) {
      const row = values[r];
      // Kembalikan ke sheet Produk utama (termasuk kolom foto di index 7)
      sheetProduk.appendRow([row[0], row[1], row[2], row[3], row[4], row[5], row[6] || 0, row[7] || ""]);
      sheetSampah.deleteRow(r + 1);

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
      logStockChange(ss, {
        waktu: nowStr,
        id: id,
        nama: String(row[1]),
        stokLama: 0,
        perubahan: Number(row[4]) || 0,
        stokBaru: Number(row[4]) || 0,
        keterangan: "Dipulihkan dari Tong Sampah"
      });

      return createJsonResponse({ status: "success", message: `Produk ${id} (${row[1]}) berhasil dipulihkan.` });
    }
  }

  return createJsonResponse({ status: "error", message: `Produk ${id} tidak ada di Tong Sampah.` });
}

/**
 * Mencatat perubahan stok ke sheet 'Log_Stok' untuk pelacakan audit
 */
function logStockChange(ss, entry) {
  try {
    let sheetLog = ss.getSheetByName("Log_Stok");
    if (!sheetLog) {
      sheetLog = ss.insertSheet("Log_Stok");
      sheetLog.appendRow(["Waktu", "ID Produk", "Nama Produk", "Stok Lama", "Perubahan", "Stok Baru", "Keterangan"]);
      sheetLog.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e2e8f0");
    }
    sheetLog.appendRow([
      entry.waktu,
      entry.id,
      entry.nama,
      entry.stokLama,
      entry.perubahan >= 0 ? `+${entry.perubahan}` : entry.perubahan,
      entry.stokBaru,
      entry.keterangan
    ]);
  } catch (e) {
    console.warn("Gagal menulis Log_Stok:", e);
  }
}

/**
 * Handle Pelunasan Kasbon / Hutang
 */
function handlePayDebt(ss, payload) {
  const { transactionId, bayarNominal, metodePelunasan } = payload;
  const sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) throw new Error("Sheet Transaksi tidak ditemukan.");

  const values = sheetTransaksi.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(transactionId).trim()) {
      const totalTagihan = Number(values[r][6]) || 0;
      const bayarLama = Number(values[r][7]) || 0;
      const bayarBaru = bayarNominal ? (bayarLama + Number(bayarNominal)) : totalTagihan;
      const kembalian = Math.max(0, bayarBaru - totalTagihan);

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm");
      const metodeStr = bayarBaru >= totalTagihan 
        ? `Lunas (${metodePelunasan || "Tunai"} - ${nowStr})` 
        : `Cicil (${bayarBaru}/${totalTagihan} - ${nowStr})`;

      // Update kolom Bayar (8), Kembalian (9), Metode (10)
      sheetTransaksi.getRange(r + 1, 8).setValue(bayarBaru);
      sheetTransaksi.getRange(r + 1, 9).setValue(kembalian);
      sheetTransaksi.getRange(r + 1, 10).setValue(metodeStr);

      return createJsonResponse({ 
        status: "success", 
        message: `Kasbon invoice ${transactionId} berhasil diperbarui (Status: ${metodeStr}).`,
        total: totalTagihan,
        bayar: bayarBaru,
        metode: metodeStr
      });
    }
  }

  return createJsonResponse({ status: "error", message: `Transaksi ${transactionId} tidak ditemukan.` });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
