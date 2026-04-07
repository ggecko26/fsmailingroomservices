//GERBANG UTAMA UI
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Mailing Services Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

//SISTEM VERIFIKASI LOGIN
function verifyUserLogin(inputEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var googleEmail = Session.getActiveUser().getEmail(); 
  var email = inputEmail.trim().toLowerCase();
  
  if (email.indexOf('@superindo.co.id') === -1) {
    throw new Error("Wajib menggunakan email berdomain @superindo.co.id");
  }

  var sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) throw new Error("Sheet 'Users' tidak ditemukan. Hubungi IT.");

  var data = sheetUsers.getDataRange().getValues();
  var userFound = null;

  // Mencari user di data sheet (start dari baris 2/index 1)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      userFound = {
        role: (data[i][1] || 'Staff').toString().charAt(0).toUpperCase() + (data[i][1] || 'Staff').toString().slice(1).toLowerCase()
      };
      break;
    }
  }

  if (!userFound) throw new Error("Email " + email + " belum terdaftar di sistem.");

  // Jam Operasional (08:00 - 15:00)
  var currentHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);
  if (userFound.role !== 'Owner' && userFound.role !== 'Admin') {
    if (currentHour < 8 || currentHour >= 15) {
      throw new Error("Akses ditutup. Portal hanya aktif pukul 08:00 - 15:00 WIB.");
    }
  }

  var logo = "";
  var settingSheet = ss.getSheetByName('Settings');
  if (settingSheet) { logo = settingSheet.getRange('B1').getValue(); }

  return { 
    superindoEmail: email, 
    googleEmail: googleEmail, 
    role: userFound.role, 
    logo: logo 
  };
}

//FUNGSI DATABASE & MASTER
function getMasterDivisi() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MasterDivisi');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] !== "").map(r => ({
    region: r[0], divisi: r[1], costCenter: r[2], alamat: r[3]
  }));
}

function saveData(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Database');
  if (!sheet) throw new Error("Sheet 'Database' tidak ditemukan.");

  // 1. Ambil semua header di baris pertama
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 2. Fungsi pembantu untuk mencari index berdasarkan nama kolom
  // Kita tambah +1 karena index array mulai dari 0, sedangkan kolom Spreadsheet mulai dari 1
  function getCol(name) {
    var index = headers.indexOf(name);
    if (index === -1) throw new Error("Kolom '" + name + "' tidak ditemukan!");
    return index;
  }

  var timestamp = new Date();
  var googleEmailBackend = Session.getActiveUser().getEmail() || "Guest/Anonymous";

  // 3. Mapping data ke baris sesuai urutan header yang ada di Sheet
  var rowsToInsert = payload.recipients.map(function(rec) {
    // Kita buat array kosong sepanjang jumlah kolom yang ada
    var row = new Array(headers.length);
    
    // Isi data ke index yang tepat berdasarkan nama kolomnya
    row[getCol("Timestamp")] = timestamp;
    row[getCol("Email Superindo")] = payload.superindoEmail;
    row[getCol("Gmail")] = googleEmailBackend;
    row[getCol("Role")] = payload.userRole;
    row[getCol("Employee Number")] = payload.empNumber;
    row[getCol("Nama Pengirim")] = payload.senderName;
    row[getCol("Email Pengirim")] = payload.senderEmail;
    row[getCol("Nomor HP Pengirim")] = payload.senderPhone;
    row[getCol("Region Pengirim")] = payload.senderRegion;
    row[getCol("Divisi Pengirim")] = payload.senderDivisi;
    row[getCol("Cost Center Pengirim")] = payload.senderCostCenter;
    row[getCol("Alamat Pengirim")] = payload.senderAlamat;
    row[getCol("Pembebanan")] = payload.pembebanan;
    row[getCol("Jumlah Tujuan Region")] = payload.totalRecipients;

    row[getCol("Tujuan Region")] = rec.region;
    row[getCol("Tujuan Divisi")] = rec.divisi;
    row[getCol("Nama Penerima")] = rec.name;
    row[getCol("Jenis Barang")] = rec.jenisPaket;
    row[getCol("Jumlah Barang")] = rec.qty;
    row[getCol("Asuransi")] = rec.asuransi;
    row[getCol("Nilai Barang")] = rec.hargaBarang;
    row[getCol("Packing Kayu")] = rec.packingKayu;
    row[getCol("Packing Bubble")] = rec.packingBubble;
    row[getCol("Layanan")] = rec.layanan;

    return row;
  });

  // 4. Insert data (urutan sudah otomatis mengikuti template header)
  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, headers.length).setValues(rowsToInsert);
  
  return "Data berhasil disimpan!";
  
  // //Menggunakan Array untuk Bulk Insert
  // var rowsToInsert = payload.recipients.map(function(rec) {
  //   return [
  //     timestamp, 
  //     payload.superindoEmail, 
  //     googleEmailBackend,
  //     payload.userRole,
  //     payload.empNumber, 
  //     payload.senderName, 
  //     payload.senderEmail, 
  //     payload.senderPhone, 
  //     payload.senderRegion, 
  //     payload.senderDivisi, 
  //     payload.senderCostCenter, 
  //     payload.senderAlamat,
  //     payload.pembebanan, 
  //     payload.totalRecipients,
  //     rec.region, 
  //     rec.divisi, 
  //     rec.name, 
  //     rec.jenisPaket, 
  //     rec.qty, 
  //     rec.asuransi, 
  //     rec.hargaBarang, 
  //     rec.packingKayu, 
  //     rec.packingBubble, 
  //     rec.layanan
  //   ];
  // });

  // sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
  
  // return "Data berhasil disimpan!";
}

function saveLogo(base64Data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Settings');
  sheet.getRange('A1:B1').setValues([['Company Logo', base64Data]]);
  return "Logo diperbarui!";
}

//LOGIKA DASHBOARD & FILTERING
function getDashboardData(startDate, endDate, currentUser) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Database");
    if (!sheet) return { list: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { list: [] };

    // 1. Ambil Header & Data (Tetap sama)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rangeSize = Math.min(lastRow - 1, 2000); 
    const startRow = lastRow - rangeSize + 1; 
    const dataRows = sheet.getRange(startRow, 1, rangeSize, sheet.getLastColumn()).getValues();

    const groups = {};

    const getCol = (name) => {
      const idx = headers.indexOf(name);
      if (idx === -1) throw new Error(`Kolom '${name}' tidak ditemukan.`);
      return idx;
    };

    const colIdx = {
      timestamp: getCol("Timestamp"),
      emailSuperIndo: getCol("Email Superindo"),
      gmail: getCol("Gmail"),
      employeeNumber: getCol("Employee Number"),
      senderName: getCol("Nama Pengirim"),
      senderEmail: getCol("Email Pengirim"),
      senderPhone: getCol("Nomor HP Pengirim"),
      senderRegion: getCol("Region Pengirim"),
      senderDivisi: getCol("Divisi Pengirim"),
      senderCostCenter: getCol("Cost Center Pengirim"),
      senderAlamat: getCol("Alamat Pengirim"),
      pembebanan: getCol("Pembebanan"),
      totalRecipients: getCol("Jumlah Tujuan Region"),
    };

    dataRows.forEach(row => {
      let ts = row[colIdx.timestamp];
      if (!ts) return;
      if (!(ts instanceof Date)) ts = new Date(ts);
      if (isNaN(ts.getTime())) return;

      // --- TAMBAHAN LOGIKA FILTER ROLE ---
      const emailDiRow = row[colIdx.emailSuperIndo] || "";
      
      // Jika role bukan "Owner" dan bukan "Admin", cek apakah email di baris ini sama dengan email user
      if (currentUser.role !== "Owner" && currentUser.role !== "Admin") {
        if (emailDiRow.toLowerCase() !== currentUser.superindoEmail.toLowerCase()) {
          return; // Skip baris ini jika bukan milik si Owner
        }
      }
      // ------------------------------------

      const dateStr = Utilities.formatDate(ts, "Asia/Jakarta", "yyyy-MM-dd");
      if (startDate && dateStr < startDate) return;
      if (endDate && dateStr > endDate) return;

      const timeKey = Utilities.formatDate(ts, "Asia/Jakarta", "dd-MM-yyyy HH:mm:ss");
      const userKey = row[colIdx.gmail] || "-";
      const senderName = row[colIdx.senderName] || "-";
      const emailSuperIndo = emailDiRow || "-";
      const employeeNumber = row[colIdx.employeeNumber] || "-";
      const senderEmail = row[colIdx.senderEmail] || "-";
      const senderPhone = row[colIdx.senderPhone] || "-";

      const groupKey = timeKey + "_" + userKey + "_" + employeeNumber + "_" + emailSuperIndo + "_" + senderName + "_" + senderEmail + "_" + senderPhone;

      if (!groups[groupKey]) {
        groups[groupKey] = {
          waktu: timeKey,
          gmail: userKey,
          emailSuperIndo: emailSuperIndo,
          employeeNumber: employeeNumber,
          senderName: senderName,
          senderEmail: senderEmail,
          senderPhone: senderPhone,
          senderRegion: row[colIdx.senderRegion] || "-",
          senderDivisi: row[colIdx.senderDivisi] || "-",
          senderCostCenter: row[colIdx.senderCostCenter] || "-",
          senderAlamat: row[colIdx.senderAlamat] || "-",
          pembebanan: row[colIdx.pembebanan] || "-",
          totalRecipients: row[colIdx.totalRecipients] || 0,
          rawTime: ts.getTime()
        };
      }
    });

    const result = Object.values(groups).sort((a, b) => b.rawTime - a.rawTime);
    return { list: result };

  } catch (err) {
    console.error("Dashboard Error: " + err.message);
    return { list: [], error: err.message };
  }
}

function getExportData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  if (!sheet) return [["Data Kosong"]];
  return sheet.getDataRange().getDisplayValues();
}

function getDetailTransaksi(groupKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Database");
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    // 1. Ambil Header untuk Mapping
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 2. Ambil data terbaru (sesuaikan limitnya dengan dashboard, misal 2000 baris)
    const rangeSize = Math.min(lastRow - 1, 2000);
    const startRow = lastRow - rangeSize + 1;
    const dataRows = sheet.getRange(startRow, 1, rangeSize, sheet.getLastColumn()).getValues();

    const getCol = (name) => {
      const idx = headers.indexOf(name);
      if (idx === -1) throw new Error(`Kolom '${name}' tidak ditemukan.`);
      return idx;
    };

    // 3. Definisikan Index Kolom
    const colIdx = {
      timestamp: getCol("Timestamp"),
      gmail: getCol("Gmail"),
      emailSuperIndo: getCol("Email Superindo"),
      employeeNumber: getCol("Employee Number"),
      senderName: getCol("Nama Pengirim"),
      senderEmail: getCol("Email Pengirim"),
      senderPhone: getCol("Nomor HP Pengirim"),
      senderRegion: getCol("Region Pengirim"),
      senderDivisi: getCol("Divisi Pengirim"),
      senderCC: getCol("Cost Center Pengirim"),
      senderAlamat: getCol("Alamat Pengirim"),
      pembebanan: getCol("Pembebanan"),
      // Recipient Index
      recRegion: getCol("Tujuan Region"),
      recDivisi: getCol("Tujuan Divisi"),
      recName: getCol("Nama Penerima"),
      recJenis: getCol("Jenis Barang"),
      recQty: getCol("Jumlah Barang"),
      recAsuransi: getCol("Asuransi"),
      recNilai: getCol("Nilai Barang"),
      recKayu: getCol("Packing Kayu"),
      recBubble: getCol("Packing Bubble"),
      recLayanan: getCol("Layanan")
    };

    // 4. Filter baris menggunakan skema GroupKey yang sama dengan Dashboard
    const matches = dataRows.filter(row => {
      const ts = row[colIdx.timestamp];
      if (!ts) return false;
      
      const timeStr = Utilities.formatDate(
        ts instanceof Date ? ts : new Date(ts), 
        "Asia/Jakarta", 
        "dd-MM-yyyy HH:mm:ss"
      );
      
      const userKey = row[colIdx.gmail] || "-";
      const empNum = row[colIdx.employeeNumber] || "-";
      const superIndo = row[colIdx.emailSuperIndo] || "-";
      const sName = row[colIdx.senderName] || "-";
      const sEmail = row[colIdx.senderEmail] || "-";
      const sPhone = row[colIdx.senderPhone] || "-";

      // Membangun key baris ini untuk dibandingkan dengan groupKey dari parameter
      const currentRowKey = timeStr + "_" + userKey + "_" + empNum + "_" + superIndo + "_" + sName + "_" + sEmail + "_" + sPhone;

      return currentRowKey === groupKey;
    });

    if (matches.length === 0) return null;

    // 5. Kembalikan Object Terstruktur
    const firstMatch = matches[0];
    return {
      empNumber: firstMatch[colIdx.employeeNumber],
      senderName: firstMatch[colIdx.senderName],
      senderEmail: firstMatch[colIdx.senderEmail],
      senderPhone: firstMatch[colIdx.senderPhone],
      senderRegion: firstMatch[colIdx.senderRegion],
      senderDivisi: firstMatch[colIdx.senderDivisi],
      senderCostCenter: firstMatch[colIdx.senderCC],
      senderAlamat: firstMatch[colIdx.senderAlamat],
      pembebanan: firstMatch[colIdx.pembebanan],
      recipients: matches.map(r => ({
        region: r[colIdx.recRegion],
        divisi: r[colIdx.recDivisi],
        name: r[colIdx.recName],
        jenisPaket: r[colIdx.recJenis],
        qty: r[colIdx.recQty],
        asuransi: r[colIdx.recAsuransi],
        hargaBarang: r[colIdx.recNilai],
        packingKayu: r[colIdx.recKayu],
        packingBubble: r[colIdx.recBubble],
        layanan: r[colIdx.recLayanan]
      }))
    };

  } catch (err) {
    console.error("Error Detail: " + err.message);
    return null;
  }
}

// function getDetailTransaksi(waktuStr, emailUser) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("Database");
//   if (!sheet) return null;

//   const data = sheet.getDataRange().getValues();
//   const rows = data.slice(1);

//   //Filter baris yang cocok
//   const matches = rows.filter(row => {
//     if (!row[0]) return false;
    
//     // Konversi waktu di row menjadi string yang sama formatnya dengan di tabel dashboard
//     const rowTime = Utilities.formatDate(new Date(row[0]), "Asia/Jakarta", "dd-MM-yyyy HH:mm:ss");
//     const rowEmail = row[2] ? row[2].toString().trim() : "";
    
//     //Bandingkan: Waktu harus sama DAN Email harus sama
//     return rowTime === waktuStr && rowEmail === emailUser;
//   });

//   if (matches.length === 0) return null;

//   const firstRow = matches[0];
//   return {
//     empNumber: firstRow[4],
//     senderName: firstRow[5],
//     senderEmail: firstRow[6],
//     senderPhone: firstRow[7],
//     senderRegion: firstRow[8],
//     senderDivisi: firstRow[9],
//     senderCostCenter: firstRow[10],
//     senderAlamat: firstRow[11],
//     pembebanan: firstRow[12],
//     recipients: matches.map(r => ({
//       region: r[14],
//       divisi: r[15],
//       name: r[16],
//       jenisPaket: r[17],
//       qty: r[18],
//       asuransi: r[19],
//       hargaBarang: r[20],
//       packingKayu: r[21],
//       packingBubble: r[22],
//       layanan: r[23]
//     }))
