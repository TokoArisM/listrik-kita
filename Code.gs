/**
 * =========================================================================
 * LISTRIK KITA - POS BACKEND V2 (Google Apps Script)
 * =========================================================================
 * Fitur Tambahan:
 * 1. Manajemen Produk (Tambah, Edit, Hapus)
 * 2. Ambil Riwayat Transaksi & Ringkasan Omzet
 * 3. Dukungan Multi Metode Pembayaran, Diskon, dan Nama Pelanggan
 * =========================================================================
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Produk
  let sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet("Produk");
    sheetProduk.appendRow(["ID", "Nama", "Kategori", "Harga", "Stok", "Satuan"]);
    sheetProduk.appendRow(["P001", "Lampu LED Philips 10W Putih", "Lampu", 35000, 50, "pcs"]);
    sheetProduk.appendRow(["P002", "Stop Kontak Broco 4L + Kabel 3M", "Kabel & Stopkontak", 65000, 20, "pcs"]);
    sheetProduk.appendRow(["P003", "Kabel NYM Eterna 2x1.5mm (Meter)", "Kabel & Stopkontak", 12000, 100, "meter"]);
    sheetProduk.appendRow(["P004", "MCB Schneider 10A 1 Phase", "MCB & Box", 55000, 15, "pcs"]);
    sheetProduk.appendRow(["P005", "Saklar Ganda Panasonic Seri WESJ", "Saklar", 38000, 30, "pcs"]);
    sheetProduk.appendRow(["P006", "Isolasi Listrik Unibell Hitam", "Aksesoris", 7000, 80, "roll"]);
    sheetProduk.appendRow(["P007", "Fitting Lampu Plafon Broco Standard", "Lampu", 15000, 45, "pcs"]);
    sheetProduk.appendRow(["P008", "Tespen Listrik Obeng Bolak Balik", "Alat & Tool", 18000, 25, "pcs"]);
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
        satuan: String(row[5] || "pcs").trim()
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
 * Mengambil 100 transaksi terakhir & ringkasan omzet
 */
function getTransactionsData(ss) {
  const sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) {
    return createJsonResponse({ status: "success", data: [], summary: { totalOmzet: 0, count: 0 } });
  }

  const values = sheetTransaksi.getDataRange().getValues();
  if (values.length <= 1) {
    return createJsonResponse({ status: "success", data: [], summary: { totalOmzet: 0, count: 0 } });
  }

  const transactions = [];
  let totalOmzet = 0;

  // Baca dari baris paling baru (bawah ke atas), ambil maksimal 100 data terakhir
  const startRow = Math.max(1, values.length - 100);
  for (let i = values.length - 1; i >= startRow; i--) {
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
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action || "checkout";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "save_product") {
      return handleSaveProduct(ss, postData);
    } else if (action === "delete_product") {
      return handleDeleteProduct(ss, postData);
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
 * Handle Checkout Transaksi
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

  // 2. Potong stok otomatis di sheet Produk
  const produkValues = sheetProduk.getDataRange().getValues();
  const idMap = {};
  for (let r = 1; r < produkValues.length; r++) {
    const pid = String(produkValues[r][0]).trim();
    if (pid) idMap[pid] = r + 1;
  }

  for (const item of items) {
    const rowIdx = idMap[String(item.id).trim()];
    if (rowIdx) {
      const cellStok = sheetProduk.getRange(rowIdx, 5);
      const currentStok = Number(cellStok.getValue()) || 0;
      cellStok.setValue(Math.max(0, currentStok - Number(item.qty)));
    }
  }

  return createJsonResponse({
    status: "success",
    message: "Transaksi berhasil dicatat dan stok dipotong.",
    transactionId: invoiceId
  });
}

/**
 * Handle Tambah / Update Produk
 */
function handleSaveProduct(ss, payload) {
  const { id, nama, kategori, harga, stok, satuan, isEdit } = payload;
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) throw new Error("Sheet Produk tidak ditemukan.");

  const values = sheetProduk.getDataRange().getValues();
  let targetRow = -1;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(id).trim()) {
      targetRow = r + 1;
      break;
    }
  }

  if (targetRow > 0) {
    // Update data produk yang ada
    sheetProduk.getRange(targetRow, 2).setValue(nama);
    sheetProduk.getRange(targetRow, 3).setValue(kategori || "Lainnya");
    sheetProduk.getRange(targetRow, 4).setValue(Number(harga) || 0);
    sheetProduk.getRange(targetRow, 5).setValue(Number(stok) || 0);
    sheetProduk.getRange(targetRow, 6).setValue(satuan || "pcs");
    return createJsonResponse({ status: "success", message: `Produk ${id} berhasil diperbarui.` });
  } else {
    // Tambah baris baru
    const newId = id || ("P" + String(values.length).padStart(3, "0"));
    sheetProduk.appendRow([
      newId,
      nama,
      kategori || "Lainnya",
      Number(harga) || 0,
      Number(stok) || 0,
      satuan || "pcs"
    ]);
    return createJsonResponse({ status: "success", message: `Produk ${newId} berhasil ditambahkan.` });
  }
}

/**
 * Handle Hapus Produk
 */
function handleDeleteProduct(ss, payload) {
  const { id } = payload;
  const sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) throw new Error("Sheet Produk tidak ditemukan.");

  const values = sheetProduk.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(id).trim()) {
      sheetProduk.deleteRow(r + 1);
      return createJsonResponse({ status: "success", message: `Produk ${id} berhasil dihapus.` });
    }
  }

  return createJsonResponse({ status: "error", message: `Produk dengan ID ${id} tidak ditemukan.` });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
