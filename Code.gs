/**
 * =========================================================================
 * LISTRIK KITA - POS BACKEND (Google Apps Script)
 * =========================================================================
 * Petunjuk Penyiapan Google Sheets:
 * 1. Buat Spreadsheet baru di Google Drive bernama "DB POS LISTRIK KITA".
 * 2. Buat sheet bernama "Produk" dengan header kolom baris 1:
 *    [ID, Nama, Kategori, Harga, Stok, Satuan]
 * 3. Buat sheet bernama "Transaksi" dengan header kolom baris 1:
 *    [ID Transaksi, Waktu, Detail Item, Total Belanja, Bayar, Kembalian, Metode]
 * 4. Buka menu Extensions -> Apps Script, tempelkan skrip ini.
 * 5. Klik "Deploy" -> "New deployment" -> pilih tipe "Web app".
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 6. Salin Web App URL dan tempelkan ke variabel CONFIG.API_URL di frontend.
 * =========================================================================
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Produk jika belum ada
  let sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet("Produk");
    sheetProduk.appendRow(["ID", "Nama", "Kategori", "Harga", "Stok", "Satuan"]);
    
    // Sample Data Toko Alat Listrik
    sheetProduk.appendRow(["P001", "Lampu LED Philips 10W Putih", "Lampu", 35000, 50, "pcs"]);
    sheetProduk.appendRow(["P002", "Stop Kontak Broco 4 Lubang + Kabel 3M", "Kabel & Stopkontak", 65000, 20, "pcs"]);
    sheetProduk.appendRow(["P003", "Kabel NYM Eterna 2x1.5mm (per meter)", "Kabel & Stopkontak", 12000, 100, "meter"]);
    sheetProduk.appendRow(["P004", "MCB Schneider 10A 1 Phase", "MCB & Box", 55000, 15, "pcs"]);
    sheetProduk.appendRow(["P005", "Saklar Ganda Panasonic Seri WESJ", "Saklar", 38000, 30, "pcs"]);
    sheetProduk.appendRow(["P006", "Isolasi Listrik Unibell Hitam", "Aksesoris", 7000, 80, "roll"]);
    sheetProduk.appendRow(["P007", "Fitting Lampu Plafon Broco Standard", "Lampu", 15000, 45, "pcs"]);
    sheetProduk.appendRow(["P008", "Tespen Listrik Obeng Bolak Balik", "Alat & Tool", 18000, 25, "pcs"]);
  }

  // 2. Setup Sheet Transaksi jika belum ada
  let sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) {
    sheetTransaksi = ss.insertSheet("Transaksi");
    sheetTransaksi.appendRow(["ID Transaksi", "Waktu", "Detail Item", "Total Belanja", "Bayar", "Kembalian", "Metode"]);
  }
}

/**
 * doGet() - Mengambil data produk & stok terkini
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetProduk = ss.getSheetByName("Produk");

    if (!sheetProduk) {
      return createJsonResponse({
        status: "error",
        message: "Sheet 'Produk' tidak ditemukan. Jalankan setupDatabase() terlebih dahulu."
      });
    }

    const data = sheetProduk.getDataRange().getValues();
    if (data.length <= 1) {
      return createJsonResponse({ status: "success", data: [] });
    }

    const headers = data[0];
    const products = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Lewati baris kosong
      products.push({
        id: String(row[0]),
        nama: String(row[1] || ""),
        kategori: String(row[2] || "Lainnya"),
        harga: Number(row[3]) || 0,
        stok: Number(row[4]) || 0,
        satuan: String(row[5] || "pcs")
      });
    }

    return createJsonResponse({
      status: "success",
      count: products.length,
      data: products
    });
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: err.toString()
    });
  }
}

/**
 * doPost() - Menyimpan transaksi baru & mengurangi stok produk
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Kunci eksekusi selama maksimal 10 detik agar pemotongan stok & nomor invoice tidak bentrok
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {
    return createJsonResponse({
      status: "error",
      message: "Server sedang sibuk, silakan ulangi beberapa saat lagi."
    });
  }

  try {
    const postData = JSON.parse(e.postData.contents);
    const {
      transactionId,
      items,
      total,
      bayar,
      kembalian,
      metodePembayaran
    } = postData;

    if (!items || !items.length) {
      return createJsonResponse({
        status: "error",
        message: "Keranjang belanja kosong."
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTransaksi = ss.getSheetByName("Transaksi");
    const sheetProduk = ss.getSheetByName("Produk");

    if (!sheetTransaksi || !sheetProduk) {
      throw new Error("Sheet Transaksi atau Produk belum disiapkan.");
    }

    const now = new Date();
    const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    const invoiceId = transactionId || ("TRX-" + Utilities.formatDate(now, "Asia/Jakarta", "yyyyMMdd-HHmmss"));

    // Rangkum item menjadi string pembacaan cepat untuk sheet
    const summaryItems = items.map(it => `${it.nama} (${it.qty}x @${it.harga}) = Rp${it.qty * it.harga}`).join("\n");

    // 1. Catat ke sheet Transaksi
    sheetTransaksi.appendRow([
      invoiceId,
      timestampStr,
      summaryItems,
      Number(total) || 0,
      Number(bayar) || 0,
      Number(kembalian) || 0,
      metodePembayaran || "Tunai"
    ]);

    // 2. Potong stok di sheet Produk secara otomatis
    const produkRange = sheetProduk.getDataRange();
    const produkValues = produkRange.getValues();
    const idMap = {};

    // Buat index baris produk (baris sheet = index array + 1)
    for (let r = 1; r < produkValues.length; r++) {
      const pId = String(produkValues[r][0]);
      if (pId) {
        idMap[pId] = r + 1; // 1-based row number
      }
    }

    for (const item of items) {
      const rowIdx = idMap[String(item.id)];
      if (rowIdx) {
        // Kolom E adalah Stok (Kolom ke-5)
        const cellStok = sheetProduk.getRange(rowIdx, 5);
        const currentStok = Number(cellStok.getValue()) || 0;
        const newStok = Math.max(0, currentStok - Number(item.qty));
        cellStok.setValue(newStok);
      }
    }

    return createJsonResponse({
      status: "success",
      message: "Transaksi berhasil disimpan dan stok terpotong.",
      transactionId: invoiceId
    });

  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: err.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper untuk response JSON dengan header CORS
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
