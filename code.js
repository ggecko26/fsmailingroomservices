// //KONFIGURASI SUPABASE
// Buka Google Apps Script editor. -> saved Key supaya ga langsung terlihat di skrip
// Klik ikon Project Settings (ikon roda gigi ⚙️) di panel sebelah kiri.
// Scroll ke bawah sampai menemukan bagian Script Properties.
// Klik Edit Script Properties, lalu tambahkan dua baris baru:
// Property: SUPABASE_URL | Value: https://xxxxxxxxx.supabase.co
// Property: SUPABASE_KEY | Value: xxxxxx

// Mengambil service properti
const scriptProperties = PropertiesService.getScriptProperties();

// MASKING: Memanggil data dari Script Properties
const SUPABASE_URL = scriptProperties.getProperty('SUPABASE_URL');
const SUPABASE_KEY = scriptProperties.getProperty('SUPABASE_KEY');

// INCLUDE FILE UNTUK PISAH FILE
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// FUNCTION GENERATE UNIQUE ID
function generateUUID() {
  const id = Utilities.getUuid();
  return id;
} // Contoh output: "123e4567-e89b-12d3-a456-426614174000"

//HELPER BUAT KOMUNIKASI SAMA REST API SUPABASE
function callSupabase(tableName, method = "GET", payload = null, queryParams = "") {

  const endpoint = `${SUPABASE_URL}/rest/v1/${tableName}${queryParams}`;
  const options = {
    method: method,
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/json",
      "Prefer": method === "POST" ? "return=representation" : "" 
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode >= 400) {
    throw new Error(`Supabase Error (${responseCode}): ${response.getContentText()}`);
  }
  
  const responseText = response.getContentText();
  return responseText ? JSON.parse(responseText) : [];
}

//MODUL INTEGRASI EKSPEDISI
function requestPickup(payload) {
//PANGGIL SISTEM SIMULASI
  return simulasiEkspedisiOtomatis(payload);
}

function simulasiEkspedisiOtomatis(payload) {
//MEMBUAT NOMOR RESI UNIK BERDASARKAN WAKTU(TIMESTAMP)
  var randomCode = new Date().getTime().toString().slice(-6);
  // var randomCode = (new Date(payload.created_at) ?? new Date()).getTime().toString().slice(-6);
  var systemAWB = "SYS-" + randomCode + "-OK"; 
  
  return {
    status: "success",
    message: "Pickup berhasil dijadwalkan oleh sistem",
    awb: systemAWB,
    courier: "Internal System"
  };
}

//UI DEPAN & VERIFIKASI LOGIN
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('1_Index');
  
  //AMBIL LOGO DARI SUPABASE
  var logo = "";
  try {
    var settingsData = callSupabase("settings", "GET", null, `?key=eq.company_logo`);
    if (settingsData && settingsData.length > 0) {
      logo = settingsData[0].value;
    }
  } catch (err) {
    console.error("Gagal mengambil logo di doGet: " + err.message);
  }
  
//KIRIM KE TEMPLATE HTML
  template.appLogo = logo;
  
  return template
    .evaluate()
    .setTitle('Mailing Services Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// cek role saat input email(onblur)
function cekRoleEmail(email) {
  if (!email) return "Staff";

  try {
    const query = `?email=eq.${email.toLowerCase().trim()}&status=eq.Active&select=role`;
    const data = callSupabase("users", "GET", null, query);
    
    if (data && data.length > 0) {
      return data[0].role; // Mengembalikan 'Admin', 'Owner', atau 'Staff'
    }
    return "Staff"; 
  } catch (e) {
    console.error("Error dalam pengecekan role pengguna: " + e.toString());
    return "Staff";
  }
}

function getAllUsers() {
  try {
    var queryParams = "?select=*&order=updated_at.desc";
    var data = callSupabase("users", "GET", null, queryParams);
    if (!data) return [];

    return data.map(row => ({
      email: row.email,
      role: row.role,
      status: row.status || 'Active',
      updatedAt: row.updated_at // Ambil data waktu terbaru
    }));
  } catch (err) {
    console.error("Error getAllUsers: " + err.message);
    return [];
  }
}

// Gantikan dua fungsi lama dengan ini agar lebih efisien
function verifyUserLogin(inputEmail, inputPassword) {
  if (!inputEmail) throw new Error("Email wajib diisi.");
  
  var email = inputEmail.trim().toLowerCase();
  if (email.indexOf('@superindo.co.id') === -1) {
    throw new Error("Wajib menggunakan email @superindo.co.id");
  }

  // Ambil data user - Cukup sekali panggil database
  var usersData = callSupabase("users", "GET", null, `?email=eq.${encodeURIComponent(email)}`);
  
  if (!usersData || usersData.length === 0) {
    throw new Error("Email belum terdaftar di sistem.");
  }

  var userData = usersData[0];
  var userRole = userData.role; // 'Admin', 'Owner', atau 'Staff'

  // 1. Cek Status Akun
  if (userData.status !== 'Active') {
    throw new Error("Akun Anda dinonaktifkan. Hubungi administrator untuk meng-aktifkan kembali.");
  }

  // //JAM OPERASIONAL
  // var currentHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);
  // if (userRole !== 'Owner' && userRole !== 'Admin') {
  //   if (currentHour < 8 || currentHour >= 15) {
  //     throw new Error("Akses ditutup. Portal hanya aktif pukul 08:00 - 15:00 WIB.");
  //   }
  // }

  // 2. Logika Password (Hanya jika Admin/Owner)
  if (userRole === 'Admin' || userRole === 'Owner') {
    if (!inputPassword) throw new Error("Password wajib diisi untuk akses " + userRole);
    
    if (hashPassword(inputPassword) !== userData.password) {
      throw new Error("Password salah.");
    }
  }
  
  // 3. Ambil Logo (Optional)
  var logo = "";
  try {
    var settingsData = callSupabase("settings", "GET", null, `?key=eq.company_logo`);
    if (settingsData && settingsData.length > 0) logo = settingsData[0].value;
  } catch (e) {}

  var serverHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);

  return { 
    superindoEmail: email, 
    role: userRole, 
    logo: logo,
    serverHour: serverHour // Kirim jam server ke client
  };
}

/** * Fungsi pembantu untuk mengambil Role User saja (dipanggil saat login/cek privilege) */
function getUserRoleOnly(email) {
  try {
    var data = callSupabase("users", "GET", null, "?email=eq." + encodeURIComponent(email) + "&select=role");
    if (data && data.length > 0) {
      return data[0].role;
    }
    return "Staff";
  } catch (err) {
    return "Staff";
  }
}

// INPUT TRANSAKSI PAGE
function saveData(payload) {
  try {
    // 1. VALIDASI OPERATIONAL TIME (Contoh: 08:00 - 15:00 WIB)
    var serverHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);

    var isOperational = (serverHour >= 8 && serverHour < 15);

    if (!isOperational // tidak jam operasional
        && (payload.userRole != "Admin" && payload.userRole != "Owner") // dan bukan Admin maupun Owner
    ) {
      throw new Error("Transaksi gagal. Sistem hanya beroperasi pada pukul 08:00 - 15:00 WIB. Silahkan coba kembali di jam operasional.");
    }

    // 2. SETUP DATA DASAR
    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ssXXX");
    var googleEmailBackend = Session.getActiveUser().getEmail() || "Guest/Anonymous";
    var transactionId = generateUUID();

    // 3. PANGGIL SISTEM EKSPEDISI (Bungkus di try-catch jika ingin tetap lanjut meski ekspedisi gagal)
    var expeditionResult;
    try {
      expeditionResult = requestPickup(payload);
    } catch (e) {
      console.error("Gagal panggil ekspedisi: " + e.message);
      throw new Error("Gagal menghubungkan ke sistem ekspedisi. Silakan coba lagi.");
    }

    // 4. MAPPING DATA PAYLOAD
    var rowsToInsert = payload.recipients.map(function(rec) {
      return {
        transaction_id: transactionId,
        created_at: timestamp,
        email_superindo: payload.superindoEmail,
        gmail: googleEmailBackend,
        role: payload.userRole,
        employee_number: payload.empNumber,
        sender_name: payload.senderName,
        sender_email: payload.senderEmail,
        sender_phone: payload.senderPhone,
        sender_region: payload.senderRegion,
        sender_divisi: payload.senderDivisi,
        sender_cost_center: payload.senderCostCenter,
        sender_alamat: payload.senderAlamat,
        pembebanan: payload.pembebanan,
        total_recipients: payload.totalRecipients,
        tujuan_region: rec.region,
        tujuan_divisi: rec.divisi,
        alamat_tujuan_divisi: rec.alamatTujuan, 
        cost_center_tujuan_divisi: rec.ccTujuan,
        nama_penerima: rec.name,
        nomor_telepon_penerima: rec.phone,
        jenis_barang: rec.jenisPaket,
        detail_barang: rec.detailPaket,
        jumlah_barang: rec.qty,
        asuransi: rec.asuransi,
        nilai_barang: rec.hargaBarang,
        packing_kayu: rec.packingKayu,
        packing_bubble: rec.packingBubble,
        layanan: rec.layanan
      };
    });

    // 5. BULK INSERT KE SUPABASE
    var dbResponse = callSupabase("database_transaksi", "POST", rowsToInsert);
    
    // Asumsi callSupabase melempar error jika response code bukan 2xx
    console.log('Insert Success:', dbResponse);

    return "Data berhasil disimpan!";

  } catch (err) {
    // LOGGING ERROR KE CONSOLE GOOGLE CLOUD
    console.error("Error di saveData: " + err.message);
    
    // MELEMPAR PESAN ERROR KE FRONT-END (agar muncul di alert/toast)
    throw new Error(err.message); 
  }
}

// COMPONENTS NAVIGATION - CHANGE / UPDATE LOGO
function saveLogo(base64Data) {
  var payload = { key: "company_logo", value: base64Data };
  callSupabase("settings", "POST", payload, "?on_conflict=key");
  return "Logo diperbarui!";
}

// VIEW DASHBOARD PAGE
function getDashboardData(startDate, endDate, currentUser) { 
  try {
    let filters = ["order=created_at.desc", "limit=2000"];

    // 2. Filter Privasi (OR) untuk Staff
    if (currentUser && currentUser.role !== "Owner" && currentUser.role !== "Admin") {
      // Pastikan nama kolom di database sesuai (sender_email atau email_superindo)
      filters.push(`or=(email_superindo.eq.${currentUser.superindoEmail},sender_email.eq.${currentUser.superindoEmail})`);
    }

    // Perbaikan Filter Tanggal dengan menyertakan Offset WIB (+07:00)
    // Tujuannya agar jam 00:00:00 di Jakarta dikonversi dengan benar oleh DB
    if (startDate) {
      filters.push(`created_at=gte.${startDate}T00:00:00%2B07:00`);
    }
    if (endDate) {
      filters.push(`created_at=lte.${endDate}T23:59:59%2B07:00`);
    }

    const queryParams = "?" + filters.join("&");
    var dataRows = callSupabase("database_transaksi", "GET", null, queryParams);
    
    if (!dataRows || dataRows.length === 0) return { list: [] };

    // --- PROSES MAPPING DATA (TANPA GROUPING) ---
    const resultList = dataRows.map(row => {
      let dateString = row.created_at;
      
      // Menggunakan parse manual agar aman dari berbagai format string ISO
      let ts = new Date(dateString);

      return {
        transactionDetailId: row.transaction_detail_id || "no-detail-id",
        transactionId: row.transaction_id || "no-id",
        // Menambahkan info detail tambahan jika ada (misal: nama produk/penerima spesifik di baris tersebut)
        timeDisplay: Utilities.formatDate(ts, "Asia/Jakarta", "dd-MM-yyyy HH:mm:ss"),
        rawTime: ts.getTime(),
        emailSuperIndo: row.email_superindo || "-",
        userGmail: row.gmail || "-",
        role: row.role || "-",
        employeeNumber: row.employee_number || "-",
        senderName: row.sender_name || "-",
        senderEmail: row.sender_email || "-",
        senderPhone: row.sender_phone || "-",
        senderRegion: row.sender_region || "-",
        senderDivisi: row.sender_divisi || "-",
        senderCostCenter: row.sender_cost_center || "-",
        senderAlamat: row.sender_alamat || "-",
        pembebanan: row.pembebanan || "-",
        totalRecipients: row.total_recipients || "-",
        tujuanRegion: row.tujuan_region || "-", 
        tujuanDivisi: row.tujuan_divisi || "-", 
        alamatTujuanDivisi: row.alamat_tujuan_divisi || "-", 
        costCenterTujuanDivisi: row.cost_center_tujuan_divisi || "-", 
        recipientName: row.nama_penerima || "-", 
        recipientPhone: row.nomor_telepon_penerima || "-", 
        jenisBarang: row.jenis_barang || "-",
        detailBarang: row.detail_barang || "-",
        jumlahBarang: row.jumlah_barang || "-",
        asuransi: row.asuransi || "-",
        nilaiBarang: row.nilai_barang || "-",
        packingKayu: row.packing_kayu || "-",
        packingBubble: row.packing_bubble || "-",
        layanan: row.layanan || "-",
        status: row.status || "-"
      };
    });

    // Data sudah urut dari query (order=created_at.desc), 
    // tapi kita pastikan lagi dengan sort di sisi client jika diperlukan
    resultList.sort((a, b) => b.rawTime - a.rawTime);

    return { list: resultList };

  } catch (err) {
    console.error("Dashboard Error: " + err.message);
    return { list: [], error: err.message };
  }
}

function getDetailTransaksi(transaction_id) {
  try {
    // 1. Filter langsung di level Database menggunakan query param 'eq'
    // Format Supabase/PostgREST: column=eq.value
    var queryParams = "?transaction_id=eq." + transaction_id;
    
    var dataRows = callSupabase("database_transaksi", "GET", null, queryParams);
    
    // Jika data tidak ditemukan
    if (!dataRows || dataRows.length === 0) return null;

    // Karena sudah difilter di DB, dataRows hanya berisi transaksi yang dicari
    const firstMatch = dataRows[0];
    
    return {
      transactionId: firstMatch.transaction_id,
      rawDate: firstMatch.created_at,
      empNumber: firstMatch.employee_number,
      senderName: firstMatch.sender_name,
      senderEmail: firstMatch.sender_email,
      senderPhone: firstMatch.sender_phone,
      senderRegion: firstMatch.sender_region,
      senderDivisi: firstMatch.sender_divisi,
      senderCostCenter: firstMatch.sender_cost_center,
      senderAlamat: firstMatch.sender_alamat,
      pembebanan: firstMatch.pembebanan,
      // Mapping semua recipient jika satu transaction_id punya banyak baris
      recipients: dataRows.map(r => ({
        region: r.tujuan_region, 
        divisi: r.tujuan_divisi, 
        name: r.nama_penerima, 
        phone: r.nomor_telepon_penerima, 
        jenisPaket: r.jenis_barang, 
        detailPaket: r.detail_barang, 
        qty: r.jumlah_barang, 
        asuransi: r.asuransi, 
        hargaBarang: r.nilai_barang, 
        packingKayu: r.packing_kayu, 
        packingBubble: r.packing_bubble, 
        layanan: r.layanan
      }))
    };
  } catch (err) {
    console.error("Error Detail: " + err.message);
    return null;
  }
}

/** * Fungsi yang dipanggil oleh tombol 'Export Filtered' dari Dashboard  */
function getExportDataByRange(startDate, endDate, user) {
  // Kita teruskan objek user ke core logic
  return processExportLogic(startDate, endDate, user);
}

/** * Fungsi yang dipanggil oleh tombol 'Export Semua Data' (Hanya Admin/Owner) */
function getExportData() {
  // Karena ini tombol "Semua Data", kita tidak kirim filter tanggal
  // Namun kita tetap butuh informasi user untuk verifikasi keamanan di logic bawah
  // (Asumsi: Tombol ini hanya muncul di UI Admin/Owner)
  return processExportLogic(null, null, null); 
}

/** * CORE LOGIC: Satu fungsi untuk semua jenis export Excel * Mengambil data alamat langsung dari snapshot database transaksi*/
function processExportLogic(startDate, endDate, user) {
  try {
    let filters = ["order=created_at.desc"];
    
    // --- 1. FILTER TANGGAL ---
    if (startDate) filters.push(`created_at=gte.${startDate}T00:00:00`);
    if (endDate)   filters.push(`created_at=lte.${endDate}T23:59:59`);
    
    // --- 2. FILTER PRIVASI ---
    if (user && user.role !== "Owner" && user.role !== "Admin") {
      filters.push(`or=(email_superindo.eq.${user.superindoEmail},sender_email.eq.${user.superindoEmail})`);
    }
    
    const queryParams = "?" + filters.join("&");
    var dataRows = callSupabase("database_transaksi", "GET", null, queryParams);
    
    // masterDivisi dihapus karena kita pakai snapshot dari database_transaksi
    
    if (!dataRows || dataRows.length === 0) return [["Data Tidak Ditemukan"]];

    // DEFINISI KOLOM
    var columnsToExport = {
      "email_superindo": "User Email",
      "gmail": "User Gmail",
      "role": "User Role",

      // DATA PENGIRIM
      "employee_number": "NIK Pengirim",
      "sender_name": "Nama Pengirim",
      "sender_phone": "No HP Pengirim",
      "sender_email": "Email Pengirim",
      "sender_region": "Region Asal",
      "sender_divisi": "Divisi Asal",
      "sender_cost_center": "Cost Center Pengirim",
      "sender_alamat": "Alamat Pengirim",

      // DATA PENERIMA / TUJUAN
      "tujuan_region": "Region Tujuan",
      "tujuan_divisi": "Divisi/Toko Tujuan",
      "cost_center_tujuan_divisi": "Cost Center Divisi/Toko Tujuan",
      "alamat_tujuan_divisi": "Alamat Divisi/Toko Tujuan", 
      "nama_penerima": "Nama Penerima",
      "nomor_telepon_penerima": "No HP Penerima",

      // DATA DETAIL TRANSAKSI / PAKET
      "pembebanan": "Pembebanan Biaya",
      "jenis_barang": "Jenis Paket",
      "detail_barang": "Isi/Detail Paket",
      "jumlah_barang": "Qty (Koli)",
      "asuransi": "Gunakan Asuransi",
      "nilai_barang": "Harga Barang (Rp)",
      "packing_kayu": "Packing Kayu",
      "packing_bubble": "Bubble Wrap",
      "layanan": "Layanan",

      // DATA IDENTIFIKATOR TRANSAKSI DI DB
      "transaction_id": "ID Transaksi",
      "transaction_detail_id": "ID Detail Transaksi",
    };

    var selectedKeys = Object.keys(columnsToExport);
    
    // Header row
    var headerRow = ["Tanggal Input", "Jam Input"];
    selectedKeys.forEach(key => {
      headerRow.push(columnsToExport[key]);
    });
    
    var exportArray = [headerRow];

    dataRows.forEach(row => {
      var ts = new Date(row.created_at);
      var dateVal = Utilities.formatDate(ts, "Asia/Jakarta", "yyyy-MM-dd");
      var timeVal = Utilities.formatDate(ts, "Asia/Jakarta", "HH:mm:ss");

      var rowData = [dateVal, timeVal];
      
      selectedKeys.forEach(key => {
        var value = row[key];

        // LOGIKA ALAMAT: Gunakan snapshot database
        if (key === "alamat_tujuan_divisi") {
          if (row.tujuan_region === "OTHER") {
            value = row.tujuan_divisi || "-";
          } else {
            // Langsung ambil dari kolom database transaksi tanpa lookup master
            value = row.alamat_tujuan_divisi || "-";
          }
        }
        
        // Formatting Angka
        if (key === "nilai_barang" || key === "jumlah_barang") {
          rowData.push(value !== null && value !== undefined ? Number(value) : 0);
        } else {
          rowData.push((value !== null && value !== undefined && value !== "") ? value : "-");
        }
      });
      
      exportArray.push(rowData);
    });

    return exportArray;
  } catch(e) {
    throw new Error("Export Error: " + e.message);
  }
}

/** * Menghapus transaksi berdasarkan transactionDetailId */
function deleteTransaction(userRole, transactionDetailId) {
  try {
    var serverHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);
    var isOperational = (serverHour >= 8 && serverHour < 15);

    if (!isOperational // tidak jam operasional
        && (userRole != "Admin" && userRole != "Owner") // dan bukan Admin maupun Owner
    ) {
      throw new Error("Transaksi gagal. Sistem hanya beroperasi pada pukul 08:00 - 15:00 WIB. Silahkan coba kembali di jam operasional.");
    }
    
    // Sesuaikan nama tabel 'database_transaksi' dengan nama tabel Anda di Supabase
    var queryParams = "?transaction_detail_id=eq." + encodeURIComponent(transactionDetailId);
    
    // Pastikan fungsi callSupabase Anda sudah mendukung metode "DELETE"
    var response = callSupabase("database_transaksi", "DELETE", null, queryParams);
    
    return "OK";
  } catch (err) {
    throw new Error("Gagal menghapus transaksi di database: " + err.message);
  }
}

/** * Mengambil 1 baris detail transaksi spesifik untuk diedit */
function getSingleTransactionDetail(transactionDetailId) {
  try {
    var queryParams = "?transaction_detail_id=eq." + encodeURIComponent(transactionDetailId);
    var dataRows = callSupabase("database_transaksi", "GET", null, queryParams);
    
    if (!dataRows || dataRows.length === 0) return null;
    return dataRows[0]; // Hanya return baris pertama yang cocok
  } catch (err) {
    console.error("Error Get Single Detail: " + err.message);
    throw new Error(err.message);
  }
}

/** * Melakukan update / PATCH pada transaksi detail spesifik  */
function updateTransactionDetail(userRole, transactionDetailId, updatePayload) {
  try {
    // Validasi jam operasional (Sama dengan aturan Delete)
    var serverHour = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "HH"), 10);
    var isOperational = (serverHour >= 8 && serverHour < 15);

    if (!isOperational && (userRole !== "Admin" && userRole !== "Owner")) {
      throw new Error("Sistem hanya beroperasi pada pukul 08:00 - 15:00 WIB. Perubahan gagal.");
    }
    
    var queryParams = "?transaction_detail_id=eq." + encodeURIComponent(transactionDetailId);
    
    // Asumsi fungsi callSupabase Anda mensupport parameter METHOD "PATCH"
    var response = callSupabase("database_transaksi", "PATCH", updatePayload, queryParams);
    
    return "OK";
  } catch (err) {
    console.error("Error Update Detail: " + err.message);
    throw new Error("Gagal mengupdate transaksi: " + err.message);
  }
}

// MANAGE USERS PAGE
/** * Menambah user baru ke Supabase */
function addNewUser(email) {
  try {
    // 1. Cek dulu apakah email sudah terdaftar
    // Asumsi kamu punya fungsi untuk fetch user berdasarkan email
    var existingUser = callSupabase("users", "GET", null, "?email=eq." + encodeURIComponent(email));
    
    if (existingUser && existingUser.length > 0) {
      // Jika ditemukan, lempar error agar ditangkap oleh withFailureHandler di client
      throw new Error("Email '" + email + "' sudah terdaftar dalam sistem.");
    }

    // 2. Jika belum ada, baru lakukan proses insert
    var payload = {
      email: email,
      role: "Staff",
      status: "Active",
      // created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };
    
    return callSupabase("users", "POST", payload);
  } catch (err) {
    // Pesan ini akan muncul di SweetAlert client
    throw new Error(err.message);
  }
}

/** * Update Role User berdasarkan Email */
function updateUserRole(email, newRole) {
  try {
    var queryParams = "?email=eq." + email;
    var payload = { 
      role: newRole,
      updated_at: new Date().toISOString() // BARIS INI WAJIB ADA
    };
    
    // PATCH (Update) data di Supabase
    return callSupabase("users", "PATCH", payload, queryParams);
  } catch (err) {
    throw new Error("Gagal update role: " + err.message);
  }
}

/** * Hapus user (Delete row) */
function deleteUser(email) {
  try {
    var queryParams = "?email=eq." + email;
    
    // DELETE data di Supabase
    return callSupabase("users", "DELETE", null, queryParams);
  } catch (err) {
    throw new Error("Gagal menghapus user: " + err.message);
  }
}

/** * Update Status User di Supabase */
function updateUserStatus(email, newStatus) {
  try {
    var queryParams = "?email=eq." + encodeURIComponent(email);
    var payload = { 
      status: newStatus,
      updated_at: new Date().toISOString() // BARIS INI WAJIB ADA
    };
    
    // Pastikan nama tabelnya benar (misal: "users")
    return callSupabase("users", "PATCH", payload, queryParams); 
  } catch (err) {
    throw new Error("Gagal update status: " + err.message);
  }
}

// Fungsi Helper untuk mengenkripsi password
function hashPassword(password) {
  const salt = "KODE_UNIK_SAYANG_LION_SUPER_INDO"; // Ganti dengan kata acak rahasia kamu
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  
  // Konversi byte array ke Hex string
  let hash = "";
  for (let i = 0; i < signature.length; i++) {
    let byte = signature[i];
    if (byte < 0) byte += 256;
    let byteStr = byte.toString(16);
    if (byteStr.length == 1) byteStr = "0" + byteStr;
    hash += byteStr;
  }
  return hash;
}

// Update fungsi updateUserPassword yang lama
function updateUserPassword(email, newPassword) {
  try {
    const hashedPw = hashPassword(newPassword); // HASH DULU SEBELUM KIRIM
    
    var queryParams = "?email=eq." + encodeURIComponent(email);
    var payload = { 
      password: hashedPw, 
      updated_at: new Date().toISOString()
    };
    
    return callSupabase("users", "PATCH", payload, queryParams);
  } catch (err) {
    throw new Error("Gagal update password: " + err.message);
  }
}

// MASTER DIVISI PAGE
/** * Mengambil semua data master divisi dari Supabase  */
function getMasterDivisi() {
  try {
    var queryParams = "?select=*&order=region.asc,divisi.asc";
    var data = callSupabase("master_divisi", "GET", null, queryParams);
    
    if (!data) return [];

    // Mapping lengkap sesuai struktur tabel database
    return data.map(row => ({
      id: row.id,
      costCenter: row.cost_center,
      region: row.region,
      divisi: row.divisi,
      workLocationSunfish: row.work_location_sunfish,
      workLocationSap: row.work_location_sap,
      kodeArea: row.kode_area,
      alamat: row.alamat,
      email: row.email,
      email2: row.email2,
      updatedAt: row.updated_at,
      organizationUnitSunfish: row.organization_unit_sunfish,
      workLocationCodeSunfish: row.work_location_code_sunfish
    }));
  } catch (err) {
    console.error("Error getMasterDivisi: " + err.message);
    return [];
  }
}

/** * Menyimpan data master divisi (Tambah Baru atau Update) */
function saveMasterDivisi(payload) {
  try {
    var dbPayload = {
      cost_center: payload.costCenter,
      region: payload.region,
      divisi: payload.divisi,
      kode_area: payload.kodeArea,
      alamat: payload.alamat,
      email: payload.email,
      email2: payload.email2,
      work_location_sap: payload.workLocationSap,
      work_location_sunfish: payload.workLocationSunfish,
      organization_unit_sunfish: payload.organizationUnitSunfish,
      work_location_code_sunfish: payload.workLocationCodeSunfish,
      updated_at: new Date().toISOString()
    };

    if (payload.id) {
      var queryParams = "?id=eq." + payload.id;
      return callSupabase("master_divisi", "PATCH", dbPayload, queryParams);
    } else {
      // Cek duplikasi
      var existing = callSupabase("master_divisi", "GET", null, 
        "?region=eq." + encodeURIComponent(payload.region) + "&divisi=eq." + encodeURIComponent(payload.divisi));
      
      if (existing && existing.length > 0) {
        throw new Error("Divisi '" + payload.divisi + "' di region '" + payload.region + "' sudah ada.");
      }
      return callSupabase("master_divisi", "POST", dbPayload);
    }
  } catch (err) {
    throw new Error(err.message);
  }
}

/** * Menghapus data master divisi berdasarkan ID */
function deleteMasterDivisi(id) {
  try {
    if (!id) throw new Error("ID tidak valid.");
    
    var queryParams = "?id=eq." + id;
    return callSupabase("master_divisi", "DELETE", null, queryParams);
  } catch (err) {
    throw new Error("Gagal menghapus data master: " + err.message);
  }
}

