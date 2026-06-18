/**
 * SI-REKAP (Sistem Informasi Rekap Kas & Cetak Voucher)
 * Backend: Google Apps Script
 * Database: Google Spreadsheet (Menyatu dalam file yang sama)
 */

function doGet(e) {
  const action = e.parameter.action;
  let jsonResponse;
  
  try {
    setupDatabase(); // Memastikan tabel database siap jika belum ada 
    
    // Routing API berdasarkan parameter ?action=
    if (action === 'getCompanySettings') {
      jsonResponse = getCompanySettings(); [cite: 80]
    } else if (action === 'getAllData') {
      jsonResponse = getAllData(); [cite: 90]
    } else if (action === 'getTransactionDetails') {
      const id = e.parameter.id_transaksi;
      jsonResponse = getTransactionDetails(id); [cite: 94]
    } else {
      jsonResponse = JSON.stringify({status: 'error', message: 'Action GET tidak ditemukan'});
    }
  } catch (err) {
    jsonResponse = JSON.stringify({status: 'error', message: err.message});
  }
  
  // Mengembalikan output JSON dengan Header CORS agar bisa ditembak oleh Vercel
  return ContentService.createTextOutput(jsonResponse)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = e.parameter.action;
  let jsonResponse;
  
  try {
    // Mengambil data payload JSON yang dikirim dari Vercel
    const payload = JSON.parse(e.postData.contents); [cite: 31]
    
    // Routing API untuk pemrosesan data (POST)
    if (action === 'saveData' || action === 'submitTransaction') {
      jsonResponse = saveData(payload); [cite: 28, 31]
    } else if (action === 'submitPemasukan') {
      // Jika Anda memiliki fungsi simpan pemasukan tersendiri, panggil di sini
      jsonResponse = typeof submitPemasukan === 'function' ? submitPemasukan(payload) : JSON.stringify({status: 'error', message: 'Fungsi tidak ditemukan'});
    } else {
      jsonResponse = JSON.stringify({status: 'error', message: 'Action POST tidak ditemukan'});
    }
  } catch (err) {
    jsonResponse = JSON.stringify({status: 'error', message: err.message});
  }
  
  return ContentService.createTextOutput(jsonResponse)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupDatabase() {
  const ss = getSpreadsheet();
  const tables = {
    'db_transactions': ['id_transaksi', 'tanggal', 'company_pc', 'dept_utama', 'paid_to', 'total_amount', 'terbilang', 'status', 'created_by', 'created_at'],
    'db_transaction_details': ['id_detail', 'id_transaksi', 'account_no', 'dept_item', 'description', 'amount'],
    'db_audit_log': ['id_log', 'timestamp', 'user', 'action', 'target_id', 'description']
  };

  for (const sheetName in tables) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(tables[sheetName]);
      sheet.getRange(1, 1, 1, tables[sheetName].length).setFontWeight("bold").setBackground("#f8f9fa");
      sheet.setFrozenRows(1);
    }
  }

  // 1. SETUP DATABASE KONFIGURASI PERUSAHAAN (SETTING)
  let sheetSetting = ss.getSheetByName('SETTING');
  if (!sheetSetting) {
    sheetSetting = ss.insertSheet('SETTING');
    sheetSetting.appendRow(['Key', 'Value']);
    sheetSetting.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f8f9fa");
    sheetSetting.setFrozenRows(1);
    
    // Nilai Default
    const defaultSettings = [
      ['company_name', 'PT BERLINA Tbk'],
      ['company_address', 'Jl. Contoh Alamat No. 123'],
      ['company_phone', '021-12345678'],
      ['company_email', 'info@berlina.co.id'],
      ['company_website', 'www.berlina.co.id'],
      ['company_city', 'Jakarta'],
      ['company_npwp', '01.234.567.8-901.000'],
      ['company_logo', '1tC6F9joGlCAqY-_jnTF-mFD3YSiAl6Hy'] // Default File ID
    ];
    
    sheetSetting.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
    sheetSetting.setColumnWidth(1, 200);
    sheetSetting.setColumnWidth(2, 400);
  }
}

// 2. FUNGSI MENARIK IDENTITAS & MENGUBAH LOGO JADI BASE64
function getCompanySettings() {
  try {
    const ss = getSpreadsheet();
    const sheetSetting = ss.getSheetByName('SETTING');
    if (!sheetSetting) return JSON.stringify({status: 'error', message: 'Sheet SETTING belum dibuat.'});
    
    const data = sheetSetting.getDataRange().getValues();
    let settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }
    
    // Konversi File ID Google Drive -> Gambar Base64 Anti-CORS
    if (settings['company_logo']) {
      try {
        let fileId = settings['company_logo'].toString().trim();
        // Fitur keamanan: Jika admin memasukkan URL panjang, sistem otomatis ekstrak File ID-nya
        let match = fileId.match(/[-\w]{25,}/); 
        if (match) fileId = match[0];
        
        let blob = DriveApp.getFileById(fileId).getBlob();
        let base64 = Utilities.base64Encode(blob.getBytes());
        let mimeType = blob.getContentType();
        
        settings['logo_base64'] = "data:" + mimeType + ";base64," + base64;
      } catch (e) {
        settings['logo_base64'] = ""; // Kosongkan jika file di-lock/salah ID
      }
    } else {
      settings['logo_base64'] = "";
    }
    
    return JSON.stringify({status: 'success', data: settings});
  } catch(error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function getAllData() {
  try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const dataTx = sheetTx.getDataRange().getValues();
    
    if (dataTx.length <= 1) return JSON.stringify({status: 'success', data: []});
    
    const headers = dataTx[0];
    const rows = dataTx.slice(1);
    
    const result = rows.map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
    
    return JSON.stringify({status: 'success', data: result.reverse()});
  } catch (error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function getTransactionDetails(id_transaksi) {
  try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const dataTx = sheetTx.getDataRange().getValues();
    const headTx = dataTx[0];
    let txData = null;
    for(let i=1; i<dataTx.length; i++){
      if(dataTx[i][0] == id_transaksi){
        txData = {};
        headTx.forEach((h, idx) => txData[h] = dataTx[i][idx]);
        break;
      }
    }

    const sheetDt = ss.getSheetByName('db_transaction_details');
    const dataDt = sheetDt.getDataRange().getValues();
    const headDt = dataDt[0];
    let details = [];
    
    for(let i=1; i<dataDt.length; i++){
      if(dataDt[i][1] == id_transaksi){
        let d = {};
        headDt.forEach((h, idx) => d[h] = dataDt[i][idx]);
        details.push(d);
      }
    }
    
    return JSON.stringify({status: 'success', data: { header: txData, details: details }});
  } catch (error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function saveData(payload) {
  try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const sheetDt = ss.getSheetByName('db_transaction_details');
    const user = Session.getActiveUser().getEmail() || 'Admin Berlina';
    const timestamp = new Date();
    
    let isUpdate = payload.id_transaksi ? true : false;
    let id_transaksi = isUpdate ? payload.id_transaksi : "TRX-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss") + Math.floor(Math.random() * 100);
    let terbilangText = terbilang(payload.total_amount) + " Rupiah";

    if (isUpdate) {
      const dataTx = sheetTx.getDataRange().getValues();
      let rowIndex = -1;
      for(let i=1; i<dataTx.length; i++){
        if(dataTx[i][0] == id_transaksi){
          rowIndex = i + 1;
          break;
        }
      }
      if(rowIndex > -1){
        sheetTx.getRange(rowIndex, 2, 1, 6).setValues([[
          payload.tanggal, payload.company_pc, payload.dept_utama, payload.paid_to, payload.total_amount, terbilangText
        ]]);
        let dtData = sheetDt.getDataRange().getValues();
        for(let i = dtData.length - 1; i >= 1; i--){
          if(dtData[i][1] == id_transaksi){
            sheetDt.deleteRow(i + 1);
          }
        }
      }
      logAudit(user, 'UPDATE', id_transaksi, 'Memperbarui transaksi');
    } else {
      sheetTx.appendRow([
        id_transaksi, payload.tanggal, payload.company_pc, payload.dept_utama, payload.paid_to, payload.total_amount, terbilangText, 'Pending', user, timestamp
      ]);
      logAudit(user, 'CREATE', id_transaksi, 'Membuat transaksi baru');
    }
    
    let dtRows = payload.details.map(d => {
      let id_detail = "DT-" + Utilities.getUuid();
      return [id_detail, id_transaksi, d.account_no, d.dept_item, d.description, d.amount];
    });
    
    if(dtRows.length > 0) {
      sheetDt.getRange(sheetDt.getLastRow() + 1, 1, dtRows.length, dtRows[0].length).setValues(dtRows);
    }
    
    return JSON.stringify({status: 'success', message: 'Data berhasil disimpan!', id_transaksi: id_transaksi});
  } catch (error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function deleteData(id_transaksi) {
  try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const dataTx = sheetTx.getDataRange().getValues();
    
    for(let i=1; i<dataTx.length; i++){
      if(dataTx[i][0] == id_transaksi){
        sheetTx.deleteRow(i + 1);
        break;
      }
    }
    const user = Session.getActiveUser().getEmail() || 'Admin Berlina';
    logAudit(user, 'DELETE', id_transaksi, 'Menghapus transaksi');
    return JSON.stringify({status: 'success', message: 'Data berhasil dihapus!'});
  } catch (error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function getDashboardSummary() {
  try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const dataTx = sheetTx.getDataRange().getValues();
    
    let totalData = 0, totalAmount = 0, countToday = 0, pendingVoucher = 0;
    let today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    let recent = [];

    if (dataTx.length > 1) {
      totalData = dataTx.length - 1;
      for (let i = 1; i < dataTx.length; i++) {
        let tgl = dataTx[i][1];
        if (tgl) {
            let tglFormatted = '';
            try { tglFormatted = Utilities.formatDate(new Date(tgl), "GMT+7", "yyyy-MM-dd"); }catch(e){}
            if (tglFormatted === today) countToday++;
        }
        totalAmount += (parseFloat(dataTx[i][5]) || 0);
        if (dataTx[i][7] === 'Pending') pendingVoucher++;
      }
      let rows = dataTx.slice(1).reverse();
      let limit = rows.length > 5 ? 5 : rows.length;
      for(let i=0; i<limit; i++){
        recent.push({ id: rows[i][0], tanggal: rows[i][1], paid_to: rows[i][4], amount: rows[i][5] });
      }
    }
    return JSON.stringify({ status: 'success', data: { totalData, totalAmount, countToday, pendingVoucher, recent } });
  } catch (error) {
    return JSON.stringify({status: 'error', message: error.message});
  }
}

function logAudit(user, action, target, desc) {
  try {
    const ss = getSpreadsheet();
    const sheetLog = ss.getSheetByName('db_audit_log');
    sheetLog.appendRow([Utilities.getUuid(), new Date(), user, action, target, desc]);
  } catch(e) {}
}

function updatePrintStatus(id_transaksi) {
   try {
    const ss = getSpreadsheet();
    const sheetTx = ss.getSheetByName('db_transactions');
    const dataTx = sheetTx.getDataRange().getValues();
    for(let i=1; i<dataTx.length; i++){
      if(dataTx[i][0] == id_transaksi){
        sheetTx.getRange(i + 1, 8).setValue('Printed');
        break;
      }
    }
    return JSON.stringify({status: 'success'});
   } catch(e) {
     return JSON.stringify({status: 'error'});
   }
}

function terbilang(angka) {
  let bilangan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];
  let temp = "";
  angka = Math.abs(parseInt(angka, 10));
  if (angka < 12) { temp = " " + bilangan[angka]; }
  else if (angka < 20) { temp = terbilang(angka - 10) + " Belas"; }
  else if (angka < 100) { temp = terbilang(Math.floor(angka / 10)) + " Puluh" + terbilang(angka % 10); }
  else if (angka < 200) { temp = " Seratus" + terbilang(angka - 100); }
  else if (angka < 1000) { temp = terbilang(Math.floor(angka / 100)) + " Ratus" + terbilang(angka % 100); }
  else if (angka < 2000) { temp = " Seribu" + terbilang(angka - 1000); }
  else if (angka < 1000000) { temp = terbilang(Math.floor(angka / 1000)) + " Ribu" + terbilang(angka % 1000); }
  else if (angka < 1000000000) { temp = terbilang(Math.floor(angka / 1000000)) + " Juta" + terbilang(angka % 1000000); }
  else if (angka < 1000000000000) { temp = terbilang(Math.floor(angka / 1000000000)) + " Milyar" + terbilang(angka % 1000000000); }
  else if (angka < 1000000000000000) { temp = terbilang(Math.floor(angka / 1000000000000)) + " Trilyun" + terbilang(angka % 1000000000000); }
  return temp.trim();
}