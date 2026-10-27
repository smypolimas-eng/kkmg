// test-docx-wscript.js — Uji menggunakan Windows Script Host (cscript)
// Jalankan: cscript test-docx-wscript.js
// Atau klik dua kali (wscript)

// Muat docx.js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName) + "\\docx.js", 1);
var code = f.ReadAll();
f.Close();
eval(code);

// Data ujian
var d = {
  nama: "Bengkel Kepimpinan Pelajar",
  anjuran: "UNIT HAL EHWAL PELAJAR",
  kluster: "Kepimpinan dan Sahsiah",
  peringkat: "Institusi",
  tarikh: "20 November 2026",
  masa: "8.00 pagi \u2013 1.00 petang",
  tempat: "Dewan Kolej Komuniti Mas Gading",
  sasaran: "Wakil pelajar aktif",
  peserta: "30",
  rakan: "",
  latar: "Program ini memberi pendedahan kemahiran kepimpinan kepada wakil pelajar.",
  objektif: "Meningkatkan kemahiran komunikasi.\nMembina semangat kerja berpasukan.",
  impak: "Pelajar dapat mengendalikan aktiviti dengan lebih berkesan.",
  kesimpulan: "Program ini diharap membentuk pemimpin yang bertanggungjawab.",
  sumber: "Peruntukan aktiviti pelajar",
  pegawai: "Pegawai HEP",
  telefon: "082-000000",
  emel: "hep@kkmg.edu.my",
  alamat: "Kolej Komuniti Mas Gading\n94000 Bau, Sarawak",
  perbatuan: [["Penyediaan kertas kerja", "1-5 November 2026"], ["Pelaksanaan", "20 November 2026"]],
  tentatif: [["8.00 pagi", "Pendaftaran"], ["9.00 pagi", "Bengkel"]],
  ajk: [["Pengerusi", "Pegawai A", "Menyelaras program"]],
  bajet: [["Bahan bengkel", "30", "5.00", "", "Bekalan"], ["Jamuan", "30", "10.00", "", "Makanan"]],
  senarai: [["Peserta A", "KK001", "Diploma IT"]]
};

try {
  var bytes = KKMG.build(d);
  
  // Check it's a Uint8Array with correct ZIP signature
  if (!(bytes instanceof Uint8Array)) throw new Error("Output bukan Uint8Array, jenis: " + typeof bytes);
  if (bytes.length < 1000) throw new Error("Saiz terlalu kecil: " + bytes.length);
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) throw new Error("Tandatangan ZIP salah");

  // Write to file using ADODB.Stream
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Type = 1; // binary
  stream.Open();
  // Convert Uint8Array to VBArray-compatible array
  var arr = new Array(bytes.length);
  for (var i = 0; i < bytes.length; i++) arr[i] = bytes[i];
  // Write bytes
  var outPath = fso.GetParentFolderName(WScript.ScriptFullName) + "\\test-output.docx";
  
  // Use a simpler approach: write using FileSystemObject via binary
  stream.Close();
  
  // Verify content by converting to string
  var textBytes = "";
  var checkStr = "";
  for (var j = 0; j < Math.min(bytes.length, 10000); j++) {
    checkStr += String.fromCharCode(bytes[j]);
  }
  
  var checks = ["[Content_Types].xml", "word/document.xml", "word/styles.xml", "_rels/.rels"];
  var allOk = true;
  
  // Check full buffer as latin1 string
  var fullStr = "";
  for (var k = 0; k < bytes.length; k++) fullStr += String.fromCharCode(bytes[k]);
  
  for (var c = 0; c < checks.length; c++) {
    if (fullStr.indexOf(checks[c]) >= 0) {
      WScript.Echo("  OK: " + checks[c]);
    } else {
      WScript.Echo("  GAGAL: " + checks[c] + " tidak dijumpai");
      allOk = false;
    }
  }
  
  if (fullStr.indexOf("w:document") >= 0 && fullStr.indexOf("KERTAS KERJA") >= 0) {
    WScript.Echo("  OK: XML Word dan kandungan hadir");
  } else {
    WScript.Echo("  GAGAL: XML Word atau kandungan tidak dijumpai");
    allOk = false;
  }
  
  if (fullStr.indexOf("Bengkel Kepimpinan Pelajar") >= 0) {
    WScript.Echo("  OK: Nama program hadir dalam dokumen");
  }
  
  if (allOk) {
    WScript.Echo("\nUJIAN LULUS - Penjana DOCX berfungsi (" + bytes.length + " bytes)");
  } else {
    WScript.Echo("\nBEBERAPA UJIAN GAGAL");
  }
} catch(e) {
  WScript.Echo("RALAT: " + e.message);
}
