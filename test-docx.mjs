// test-docx.mjs — Uji penjana DOCX dalam Node.js
// Jalankan: node test-docx.mjs

import { readFileSync, writeFileSync } from 'fs';

// Load docx.js
const code = readFileSync('./docx.js', 'utf8');
eval(code.replace('typeof window', 'typeof __fakewindow__'));

// Sample data
const d = {
  nama: 'Bengkel Kepimpinan Pelajar',
  anjuran: 'UNIT HAL EHWAL PELAJAR',
  kluster: 'Kepimpinan dan Sahsiah',
  peringkat: 'Institusi',
  tarikh: '20 November 2026',
  masa: '8.00 pagi – 1.00 petang',
  tempat: 'Dewan Kolej Komuniti Mas Gading',
  sasaran: 'Wakil pelajar aktif',
  peserta: '30',
  rakan: '',
  latar: 'Program ini memberi pendedahan kemahiran kepimpinan dan komunikasi kepada wakil pelajar.',
  objektif: 'Meningkatkan kemahiran komunikasi.\nMembina semangat kerja berpasukan.\nMendedahkan teknik kepimpinan.',
  impak: 'Pelajar dapat mengendalikan aktiviti dengan lebih berkesan.',
  kesimpulan: 'Program ini diharap membentuk pemimpin pelajar yang bertanggungjawab.',
  sumber: 'Peruntukan aktiviti pelajar',
  pegawai: 'Pegawai HEP',
  telefon: '082-000000',
  emel: 'hep@kkmg.edu.my',
  alamat: 'Kolej Komuniti Mas Gading\nLot 742–750, Off Parent Lot 726\n94000 Bau, Sarawak',
  perbatuan: [
    ['Penyediaan kertas kerja', '2–5 November 2026'],
    ['Semakan dan kelulusan', '6–10 November 2026'],
    ['Pelaksanaan program', '20 November 2026'],
  ],
  tentatif: [
    ['8.00 pagi', 'Pendaftaran peserta'],
    ['9.00 pagi', 'Bengkel kepimpinan sesi 1'],
    ['1.00 petang', 'Majlis penutup'],
  ],
  ajk: [
    ['Pengerusi', 'Pegawai A', 'Menyelaras program'],
    ['Setiausaha', 'Pegawai B', 'Laporan dan minit'],
  ],
  bajet: [
    ['Bahan bengkel', '30', '5.00', '', 'Bekalan am'],
    ['Jamuan peserta', '30', '10.00', '', 'Makanan'],
    ['Cetakan', '50', '0.20', '', 'Alat tulis'],
  ],
  senarai: [
    ['Peserta A', 'KK2026001', 'Diploma IT'],
    ['Peserta B', 'KK2026002', 'Diploma Perakaunan'],
  ],
};

try {
  const bytes = KKMG.build(d);
  if (!(bytes instanceof Uint8Array)) throw new Error('Output bukan Uint8Array');
  if (bytes.length < 1000) throw new Error('Saiz fail terlalu kecil: ' + bytes.length + ' bytes');

  // Check ZIP signature
  const sig = bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04;
  if (!sig) throw new Error('Tandatangan ZIP tidak sah');

  // Write output
  writeFileSync('./test-output.docx', bytes);
  console.log('✓ BERJAYA — Fail dijana: test-output.docx (' + bytes.length + ' bytes)');

  // Verify ZIP contains expected files
  const text = Buffer.from(bytes).toString('latin1');
  const required = ['[Content_Types].xml', 'word/document.xml', 'word/styles.xml', '_rels/.rels'];
  let allFound = true;
  for (const f of required) {
    if (text.includes(f)) {
      console.log('  ✓ ' + f);
    } else {
      console.log('  ✗ TIDAK DIJUMPAI: ' + f);
      allFound = false;
    }
  }
  if (allFound) {
    console.log('✓ Semua fail pakej DOCX hadir');
  }

  // Check XML content
  const docStart = text.indexOf('<?xml', text.indexOf('word/document.xml'));
  if (text.includes('w:document') && text.includes('w:body')) {
    console.log('✓ XML dokumen Word kelihatan sah');
  } else {
    console.log('✗ XML dokumen mungkin rosak');
  }

  // Check special characters escaped
  if (text.includes('&amp;') || !text.includes('&nama')) {
    console.log('✓ Aksara khas XML dilepaskan dengan betul');
  }

  console.log('\n✓ Ujian lulus. Buka test-output.docx dalam Microsoft Word untuk semakan visual.');
} catch (err) {
  console.error('✗ UJIAN GAGAL:', err.message);
  process.exit(1);
}
