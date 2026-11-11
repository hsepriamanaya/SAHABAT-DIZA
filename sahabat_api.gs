/* SAHABAT API - tahap 1: login + Hazard Report.
LANGKAH PASANG
1. Apps Script > Project Settings > Script Properties: tambah SAHABAT_SECRET = teks acak panjang (rahasia).
2. Tambahkan file ini ke project Hazard Tracker. Di fungsi doPost lama, ganti "return handleApiPost(e);" menjadi "return handleApiPostSahabat_(e);"
   (fungsi handleApiPost lama tidak perlu diubah sama sekali).
3. Deploy > Manage deployments > edit deployment yang ada > New version (URL tidak berubah).
4. Pendaftar muncul di tab "AKUN SAHABAT" berstatus "Menunggu Konfirmasi"; ubah ke "Disetujui" untuk mengaktifkan.
Kolom Q, R, AG, AH (ARRAYFORMULA) sengaja tidak ditulis. */
const SS_KARYAWAN = "1VK7JKDaGOu38ngzFHA9q2bJ2sz7jsew_cYzNa-Rv7s0";
const AREA = ["OFFICE","PIT","DISPOSAL","ROM","JALAN HAULING","TPS LB3","WORKSHOP / PIT SHOP","WAREHOUSE","SETLING POND","MESS","KANTIN","TANGKI TIMBUN","CONTAINER YARD (CY)","LAINNYA"];
const PERUSAHAAN = ["PT. DIZAMATRA POWERINDO","PT. PRIAMANAYA ENERGI","PT. PRIAMANAYA ENERGI (PLTU)","PT. ASTA MAHARANITA","PT. ANTAREJA MAHADA MAKMUR","PT. CIPTA KRIDATAMA","PT. KALIMANTAN PRIMA PERSADA","PT. SCADA","PT. SUCOFINDO","PT. G4S","PT. GUN"];
const KTA = ["Akomodasi tidak memadai","APD tidak standar","Bahaya Ledakan/Kebakaran","Drainase/saluran tidak memadai","Dumping point tidak standar","Fasilitas tidak memadai","Pencamaran B3","House keeping yang buruk","Jalan Tambang tidak sesuai standar","Kondisi lingkungan yang berbahaya","Loading point tidak standar","Ruang gerak terbatas","Sarana pengaman tidak memadai","Sistem peringatan tidak memadai","Terpajan kebisingan melebihi standar","Terpajan suhu tinggi atau rendah","Unit/alat/peralatan rusak"];
const TTA = ["Bercanda / Tidak serius","Dibawah pengaruh alkohol/narkoba","Melakukan perbaikan dengan kondisi mesin masih hidup","melanggar Prosedur/Standar/Instruksi Kerja","Melanggar Rambu-rambu","Menggunakan unit/alat/peralatan rusak","Menggunakan Unit/alat/peralatan tidak sesuai dengan fungsinya","Mengoperasikan unit melebihi batas kecepatan","Mengoperasikan unit tanpa izin","Merusak alat pengaman","Tidak melakukan pengamanan","Tidak memberikan peringatan","Tidak menggunakan APD"];

function nik_(s) { return String(s || "").replace(/[^0-9]/g, ""); }
function t_(s) { s = String(s == null ? "" : s).trim(); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function sig_(s) {
  var k = PropertiesService.getScriptProperties().getProperty("SAHABAT_SECRET");
  if (!k) throw new Error("SAHABAT_SECRET belum diatur di Script Properties.");
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, k));
}
function karyawan_() { // Data Karyawan: header baris 4, data mulai baris 5. A NO, B NIK, C NAME, D STATUS, E GRUP, F LVL, G DEPARTMENT, H JOB TITLE
  var sh = SpreadsheetApp.openById(SS_KARYAWAN).getSheetByName("Data Karyawan");
  return sh.getRange(5, 1, sh.getLastRow() - 4, 8).getValues().filter(function (r) { return nik_(r[1]); });
}
function akun_() {
  var ss = SpreadsheetApp.openById(SS_ID), sh = ss.getSheetByName("AKUN SAHABAT");
  if (!sh) { sh = ss.insertSheet("AKUN SAHABAT"); sh.appendRow(["NIK", "PIN (HMAC)", "STATUS", "WAKTU DAFTAR"]); }
  return sh;
}
function cekToken_(t) {
  var x = String(t || "").split(".");
  if (x.length !== 3 || Number(x[1]) < Date.now() || sig_(x[0] + "|" + x[1]) !== x[2]) return null;
  return karyawan_().filter(function (r) { return nik_(r[1]) === x[0]; })[0] || null;
}
function masters_() {
  var d = {}, kar = karyawan_(); kar.forEach(function (r) { var v = String(r[6]).trim(); if (v) d[v] = 1; });
  return { area: AREA, perusahaan: PERUSAHAAN, kta: KTA, tta: TTA, dept: Object.keys(d), emp: kar.map(function (r) { return [String(r[2]).trim(), String(r[7]).trim()]; }) };
}

function sahabatRoute_(a, p) {
  try {
    if (a === "sahabatDaftar") return sahabatDaftar_(p);
    if (a === "sahabatLogin") return sahabatLogin_(p);
    if (a === "sahabatMasters") return cekToken_(p.token) ? { success: true, data: masters_() } : { success: false, auth: false };
    if (a === "sahabatHazard") return sahabatHazard_(p);
    if (a === "sahabatForm") return sahabatForm_(p);
  } catch (e) { return { success: false, message: e.message }; }
  return null;
}

function sahabatDaftar_(p) {
  var nik = nik_(p.nik), pin = String(p.pin || "");
  if (!/^\d{4,6}$/.test(pin)) return { success: false, message: "PIN harus 4-6 angka." };
  var k = karyawan_().filter(function (r) { return nik_(r[1]) === nik; })[0];
  if (!nik || !k) return { success: false, message: "NIK tidak ditemukan di Data Karyawan." };
  var sh = akun_();
  if (sh.getDataRange().getValues().some(function (r, i) { return i && nik_(r[0]) === nik; })) return { success: false, message: "NIK sudah terdaftar." };
  sh.appendRow([String(k[1]), sig_("pin|" + nik + "|" + pin), "Menunggu Konfirmasi", new Date()]);
  return { success: true };
}

function sahabatLogin_(p) {
  var nik = nik_(p.nik), c = CacheService.getScriptCache(), ck = "lg" + nik, n = Number(c.get(ck) || 0);
  if (n >= 5) return { success: false, message: "Terlalu banyak percobaan. Coba lagi 15 menit." };
  var a = akun_().getDataRange().getValues().filter(function (r, i) { return i && nik_(r[0]) === nik; })[0];
  if (!nik || !a || a[1] !== sig_("pin|" + nik + "|" + p.pin)) { c.put(ck, n + 1, 900); return { success: false, message: "NIK atau PIN salah." }; }
  if (a[2] !== "Disetujui") return { success: false, message: "Akun menunggu persetujuan admin HSE." };
  var k = karyawan_().filter(function (r) { return nik_(r[1]) === nik; })[0], exp = Date.now() + 30 * 864e5;
  return { success: true, token: nik + "." + exp + "." + sig_(nik + "|" + exp), profil: { nik: String(k[1]), nama: k[2], jabatan: k[7], dept: k[6] } };
}

function sahabatHazard_(p) {
  var k = cekToken_(p.token);
  if (!k) return { success: false, auth: false, message: "Sesi berakhir, silakan masuk lagi." };
  var f = p.f || {};
  if (!p.id || !f.area || !f.bahaya || !f.kategori) return { success: false, message: "Data laporan belum lengkap." };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);
    var ids = sh.getRange(1, 28, sh.getMaxRows(), 1).getValues(), last = 1; // kolom AB (ID); getLastRow tidak dipakai karena ada ARRAYFORMULA
    for (var i = 0; i < ids.length; i++) {
      var v = String(ids[i][0]).trim();
      if (v) { if (v === String(p.id)) return { success: true, id: p.id }; last = i + 1; }
    }
    var folder = DriveApp.getFolderById(FOLDER_ID), urls = [];
    (p.fotos || []).slice(0, 3).forEach(function (b, n) {
      var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(String(b).split(",")[1]), "image/jpeg", "Hazard_" + p.id + "_" + (n + 1) + ".jpg"));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push("https://lh3.googleusercontent.com/d/" + file.getId());
    });
    var r = last + 1, put = function (c, vals) { sh.getRange(r, c, 1, vals.length).setValues([vals]); };
    sh.getRange(r, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss"); sh.getRange(r, 2).setNumberFormat("@"); sh.getRange(r, 3).setNumberFormat("mmm d, yyyy");
    put(1, [new Date(), String(f.jam), new Date(f.tanggal + "T00:00:00"), t_(k[2]), t_(k[1]), t_(k[7]), "", t_(k[6]),
      t_(f.area), t_(f.lokasi), t_(f.kategori), (f.kta || []).map(t_).join("\n"), (f.tta || []).map(t_).join("\n"), t_(f.bahaya), Number(f.prob), Number(f.sev)]); // A..P
    put(19, [t_(f.tindakan), t_(f.deptPic), "Open"]);          // S, T, U
    put(23, [urls.join(", "), t_(f.perusahaan)]);              // W, X
    sh.getRange(r, 28).setValue(Number(p.id));                 // AB
    SpreadsheetApp.flush();
    return { success: true, id: p.id };
  } finally { lock.releaseLock(); }
}

// ---------- TAHAP 2: Inspeksi, Observasi, Safety Talk ----------
// Rumus AG2 di DATA BASE INSPEKSI memecah kolom AD berdasarkan koma, jadi koma di teks isian diganti ";" dan "|" diganti "/".
function c_(s) { return String(s == null ? "" : s).replace(/[,|]/g, function (m) { return m === "," ? ";" : "/"; }).trim(); }
function simpanFoto_(arr, nama) {
  var folder = DriveApp.getFolderById(FOLDER_ID), urls = [];
  (arr || []).slice(0, 3).forEach(function (b, n) {
    var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(String(b).split(",")[1]), "image/jpeg", nama + "_" + (n + 1) + ".jpg"));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push("https://lh3.googleusercontent.com/d/" + file.getId());
  });
  return urls.join(", ");
}
function sahabatForm_(p) {
  var k = cekToken_(p.token);
  if (!k) return { success: false, auth: false, message: "Sesi berakhir, silakan masuk lagi." };
  var C = { inspeksi: ["DATA BASE INSPEKSI", 32], observasi: ["OBSERVASI", 17], safetytalk: ["SAFETY TALK", 15] }[p.form], f = p.f || {};
  if (!C || !p.id || !f.tanggal) return { success: false, message: "Data laporan belum lengkap." };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(C[0]);
    var ids = sh.getRange(1, C[1], sh.getMaxRows(), 1).getValues(), last = 1; // kolom Submission ID; ada formula spill, jadi getLastRow tidak dipakai
    for (var i = 0; i < ids.length; i++) {
      var v = String(ids[i][0]).trim();
      if (v) { if (v === String(p.id)) return { success: true, id: p.id }; last = i + 1; }
    }
    var r = last + 1, D = new Date(f.tanggal + "T00:00:00"), me = [new Date(), t_(k[2]), t_(k[1]), t_(k[7]), t_(k[6])]; // A..E
    var put = function (c, vals) { sh.getRange(r, c, 1, vals.length).setValues([vals]); };
    sh.getRange(r, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    var foto = simpanFoto_(p.fotos, p.form + "_" + p.id);
    if (p.form === "inspeksi") {
      var ad = (f.temuan || []).slice(0, 9).map(function (x) {
        return "Deskripsi Temuan: " + c_(x.d) + ", Tindakan yang dilakukan: " + c_(x.t) + ", Person In Charge: " + c_(x.p) + ", Departemen PIC: " + c_(x.dp) + ", Status Temuan: " + c_(x.s);
      }).join(", ");
      sh.getRange(r, 7).setNumberFormat("mm-dd-yyyy");
      put(1, me.concat([t_(f.judul), D, t_(f.area), t_(f.lokasi)]));        // A..I
      put(26, [simpanFoto_(p.fotos2, "checklist_" + p.id)]);               // Z bukti checklist
      put(30, [ad, foto, Number(p.id)]);                                   // AD temuan, AE dokumentasi, AF ID
    } else if (p.form === "observasi") {
      sh.getRange(r, 6).setNumberFormat("m/d/yyyy");
      put(1, me.concat([D, t_(f.area), t_(f.lokasi), t_(f.nama), t_(f.jabatan), t_(f.deskripsi), t_(f.tindakan), t_(f.feedback)])); // A..M
      put(14, [foto]); put(17, [Number(p.id)]);                            // N dokumentasi, Q ID
    } else {
      sh.getRange(r, 6).setNumberFormat("mm-dd-yyyy"); sh.getRange(r, 7).setNumberFormat("@");
      put(1, me.concat([D, String(f.jam), t_(f.area), t_(f.lokasi), t_(f.pemateri), t_(f.materi)])); // A..K
      put(14, [foto, Number(p.id)]);                                       // N dokumentasi, O ID
    }
    SpreadsheetApp.flush();
    return { success: true, id: p.id };
  } finally { lock.releaseLock(); }
}

// Pintu masuk baru: coba rute SAHABAT dulu, kalau bukan action SAHABAT teruskan ke handleApiPost lama.
function handleApiPostSahabat_(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var r = sahabatRoute_(payload.action, payload);
    if (r) return outputJSON(r);
  } catch (err) {}
  return handleApiPost(e);
}
