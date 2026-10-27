/* Penjana DOCX tanpa pergantungan rangkaian. ZIP disimpan tanpa pemampatan.
   Versi 3 – Kolej Komuniti Mas Gading, Unit HEP
   Modul: Kertas Kerja, Memo, Surat Lantikan AJK, Surat Lantikan Penceramah,
          Sijil Bernombor Siri, Laporan Program */
(function (root) {
  'use strict';

  const enc = new TextEncoder();

  /* ── XML escape ─────────────────────────────────────────────── */
  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({
      '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'
    }[c]));
  }

  /* ── CRC32 ──────────────────────────────────────────────────── */
  function crc32(bytes) {
    let c = 0xffffffff;
    for (const b of bytes) { c ^= b; for (let j=0;j<8;j++) c=(c>>>1)^((c&1)?0xedb88320:0); }
    return (c ^ 0xffffffff) >>> 0;
  }

  /* ── ZIP builder (store, no compression) ───────────────────── */
  function zip(files) {
    const chunks=[], central=[];
    let offset=0;
    const dv=n=>new DataView(new ArrayBuffer(n)), u8=v=>new Uint8Array(v.buffer);
    for (const [name, text] of Object.entries(files)) {
      const n=enc.encode(name), d=enc.encode(text), c=crc32(d);
      const lh=dv(30);
      lh.setUint32(0,0x04034b50,true); lh.setUint16(4,20,true);
      lh.setUint32(14,c,true); lh.setUint32(18,d.length,true);
      lh.setUint32(22,d.length,true); lh.setUint16(26,n.length,true);
      chunks.push(u8(lh),n,d);
      const cd=dv(46);
      cd.setUint32(0,0x02014b50,true); cd.setUint16(4,20,true); cd.setUint16(6,20,true);
      cd.setUint32(16,c,true); cd.setUint32(20,d.length,true);
      cd.setUint32(24,d.length,true); cd.setUint16(28,n.length,true);
      cd.setUint32(42,offset,true);
      central.push(u8(cd),n);
      offset += 30+n.length+d.length;
    }
    const cdLen=central.reduce((s,b)=>s+b.length,0), eocd=dv(22);
    eocd.setUint32(0,0x06054b50,true);
    const count=Object.keys(files).length;
    eocd.setUint16(8,count,true); eocd.setUint16(10,count,true);
    eocd.setUint32(12,cdLen,true); eocd.setUint32(16,offset,true);
    const all=[...chunks,...central,u8(eocd)];
    const out=new Uint8Array(offset+cdLen+22); let pos=0;
    for (const b of all) { out.set(b,pos); pos+=b.length; }
    return out;
  }

  /* ── WordprocessingML helpers ───────────────────────────────── */
  const NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const xmlDecl='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  function para(text='', opts={}) {
    const {bold=false,center=false,size=22,spaceBefore=0,spaceAfter=120,italic=false,underline=false,indent=0}=opts;
    const pPr=[
      `<w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}"/>`,
      center?'<w:jc w:val="center"/>':'',
      indent?`<w:ind w:left="${indent}"/>`:'',
    ].join('');
    const rPr=[
      bold?'<w:b/>':'',
      italic?'<w:i/>':'',
      underline?'<w:u w:val="single"/>':'',
      size!==22?`<w:sz w:val="${size}"/>`:'',
    ].join('');
    const textRuns=String(text).split('\n').map((line,i)=>
      (i>0?'<w:br/>':'')+'<w:t xml:space="preserve">'+esc(line)+'</w:t>'
    ).join('');
    return `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr>${rPr}</w:rPr>${textRuns}</w:r></w:p>`;
  }

  function pageBreak() { return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'; }
  function hLine() { return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="888888"/></w:pBdr><w:spacing w:after="60"/></w:pPr></w:p>'; }
  function spacer(n=80) { return `<w:p><w:pPr><w:spacing w:after="${n}"/></w:pPr></w:p>`; }

  function table(headers, rows, opts={}) {
    const {colWidths}=opts;
    const cols=headers.length;
    const border=['top','left','bottom','right','insideH','insideV']
      .map(x=>`<w:${x} w:val="single" w:sz="4" w:color="555555"/>`).join('');
    function cell(v, isHeader, wi) {
      const w=colWidths?colWidths[wi]:Math.floor(9100/cols);
      return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${para(v,{bold:isHeader,size:20})}</w:tc>`;
    }
    function row(cells,isHeader=false) {
      const trPr=isHeader?'<w:trPr><w:tblHeader/></w:trPr>':'';
      return `<w:tr>${trPr}${cells.map((v,i)=>cell(v,isHeader,i)).join('')}</w:tr>`;
    }
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${border}</w:tblBorders></w:tblPr>`
      +row(headers,true)
      +rows.map(r=>row(Array.from({length:cols},(_,i)=>r[i]??''))).join('')
      +'</w:tbl>'+spacer(60);
  }

  function sectionHead(text) { return para(text,{bold:true,size:24,spaceBefore:200,spaceAfter:80}); }

  /* ── Styles XML ─────────────────────────────────────────────── */
  function stylesXml() {
    return xmlDecl+`<w:styles ${NS}>
<w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
</w:styles>`;
  }

  /* ── Assemble DOCX package ──────────────────────────────────── */
  function pack(bodyXml) {
    const doc=xmlDecl+`<w:document ${NS}><w:body>${bodyXml}`
      +`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>`
      +`<w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1701"/>`
      +`</w:sectPr></w:body></w:document>`;
    return zip({
      '[Content_Types].xml':xmlDecl+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        +'<Default Extension="xml" ContentType="application/xml"/>'
        +'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        +'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        +'</Types>',
      '_rels/.rels':xmlDecl+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        +'</Relationships>',
      'word/_rels/document.xml.rels':xmlDecl+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        +'</Relationships>',
      'word/styles.xml':stylesXml(),
      'word/document.xml':doc,
    });
  }

  /* ══════════════════════════════════════════════════════════════
     1. KERTAS KERJA
  ══════════════════════════════════════════════════════════════ */
  function buildKK(d) {
    let body='';
    // Cover
    body+=para('',{spaceAfter:400});
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:28,spaceAfter:60});
    body+=para('Unit Hal Ehwal Pelajar',{center:true,size:24,spaceAfter:400});
    body+=hLine();
    body+=para('KERTAS KERJA',{bold:true,center:true,size:32,spaceBefore:200,spaceAfter:120});
    body+=para(d.nama.toUpperCase(),{bold:true,center:true,size:26,spaceAfter:300});
    body+=hLine();
    body+=para('',{spaceAfter:200});
    body+=para('Tarikh Program',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.tarikh,{center:true,size:22,spaceAfter:160});
    body+=para('Tempat',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.tempat,{center:true,size:22,spaceAfter:160});
    body+=para('Penganjur',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.anjuran.toUpperCase(),{bold:true,center:true,size:22,spaceAfter:40});
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:22,spaceAfter:300});
    body+=hLine();
    body+=para('',{spaceAfter:200});
    body+=para('DISEDIAKAN OLEH:',{bold:true,center:true,size:22,spaceAfter:80});
    body+=para('Nama: __________________________________',{center:true,size:22,spaceAfter:40});
    body+=para('Jawatan: ________________________________',{center:true,size:22,spaceAfter:40});
    body+=para('Tarikh: _________________________________',{center:true,size:22});
    body+=pageBreak();
    // Content
    body+=para('PERMOHONAN PERUNTUKAN PROGRAM PEMBANGUNAN PELAJAR',{bold:true,center:true,size:24,spaceAfter:200});
    body+=table(['Perkara','Maklumat'],[
      ['Nama Program',d.nama],['Kluster Program',d.kluster],['Peringkat Program',d.peringkat],
      ['Tarikh / Masa',d.tarikh+(d.masa?' / '+d.masa:'')],['Tempat',d.tempat],
      ['Kumpulan Sasaran',d.sasaran],['Bilangan Peserta',d.peserta],
      ['Agensi / Rakan Kolaborasi',d.rakan||'Tiada'],
    ]);
    body+=sectionHead('1.0  LATAR BELAKANG / PENGENALAN');
    body+=para(d.latar,{spaceAfter:160});
    body+=sectionHead('2.0  OBJEKTIF PROGRAM');
    (d.objektif||'').split('\n').map(s=>s.trim()).filter(Boolean).forEach((o,i)=>{body+=para((i+1)+'. '+o,{spaceAfter:60});});
    body+=sectionHead('3.0  CARTA JAWATANKUASA');
    body+=para('Sila rujuk Lampiran 1.',{spaceAfter:160});
    body+=sectionHead('4.0  CARTA PERBATUAN');
    body+=table(['Aktiviti','Tarikh / Tempoh'],d.perbatuan||[]);
    body+=sectionHead('5.0  ANGGARAN PERUNTUKAN PROGRAM');
    body+=para('Sila rujuk Lampiran 2.',{spaceAfter:160});
    body+=sectionHead('6.0  PERINCIAN / TENTATIF PROGRAM');
    body+=table(['Masa','Aktiviti'],d.tentatif||[]);
    body+=sectionHead('7.0  IMPAK PROGRAM');
    body+=para(d.impak,{spaceAfter:160});
    body+=sectionHead('8.0  KESIMPULAN');
    body+=para(d.kesimpulan,{spaceAfter:160});
    // Signatures
    body+=pageBreak();
    body+=para('PENGESAHAN INSTITUSI',{bold:true,center:true,size:24,spaceAfter:200});
    for (const lbl of ['Disediakan oleh','Disemak oleh']) {
      body+=para(lbl+':',{bold:true,spaceAfter:60});
      body+=para('\n\n\n..........................................................',{spaceAfter:40});
      body+=para('Nama: __________________________________',{spaceAfter:20});
      body+=para('Jawatan: ________________________________',{spaceAfter:20});
      body+=para('Tarikh: __________________________________',{spaceAfter:160});
    }
    body+=sectionHead('SEMAKAN PERUNTUKAN KEWANGAN (JIKA BERKAITAN)');
    body+=para('Pengesahan baki peruntukan:  MENCUKUPI  /  TIDAK MENCUKUPI\n\nDi bawah: ________________________________\nBaki peruntukan: RM _______________________\n\nTandatangan: _____________________________\nNama / Jawatan: __________________________\nTarikh: __________________________________',{spaceAfter:160});
    body+=sectionHead('KELULUSAN');
    body+=para('DILULUSKAN  /  TIDAK DILULUSKAN\n\nUlasan: __________________________________\n_________________________________________\n\nTandatangan: _____________________________\nNama: __________________________________\nPengarah / Timbalan Pengarah\nTarikh: __________________________________',{spaceAfter:160});
    // Contact
    body+=pageBreak();
    body+=sectionHead('PEGAWAI UNTUK DIHUBUNGI');
    body+=para('Nama: '+(d.pegawai||'')+'\nNo. Telefon: '+(d.telefon||'')+'\nE-mel: '+(d.emel||''),{spaceAfter:120});
    body+=sectionHead('ALAMAT INSTITUSI');
    body+=para(d.alamat||'',{spaceAfter:120});
    // Lampiran 1
    body+=pageBreak();
    body+=para('Lampiran 1',{bold:true,size:24,spaceAfter:40});
    body+=para('SENARAI JAWATANKUASA PROGRAM',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.nama,{center:true,spaceAfter:120});
    body+=table(['Jawatan','Nama','Tugas / Tanggungjawab'],d.ajk||[]);
    // Lampiran 2
    body+=pageBreak();
    body+=para('Lampiran 2',{bold:true,size:24,spaceAfter:40});
    body+=para('ANGGARAN PERBELANJAAN MENGURUS',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.nama,{center:true,spaceAfter:120});
    let total=0;
    const budgetRows=(d.bajet||[]).map(r=>{
      const amt=Math.round((Number(r[1])||0)*(Number(r[2])||0)*100)/100;
      total=Math.round((total+amt)*100)/100;
      return [r[0],String(r[1]),(Number(r[2])||0).toFixed(2),amt.toFixed(2),r[3]||''];
    });
    body+=table(['Perbelanjaan','Kuantiti','Harga / Unit (RM)','Jumlah (RM)','Objek Kewangan'],budgetRows);
    body+=para('JUMLAH KESELURUHAN: RM '+total.toFixed(2),{bold:true,size:22,spaceAfter:80});
    body+=para('Sumber peruntukan: '+(d.sumber||''),{spaceAfter:80});
    // Lampiran 3
    body+=pageBreak();
    body+=para('Lampiran 3',{bold:true,size:24,spaceAfter:40});
    body+=para('SENARAI PESERTA PROGRAM',{bold:true,center:true,size:22,spaceAfter:40});
    body+=para(d.nama,{center:true,spaceAfter:120});
    body+=table(['Bil.','Nama','No. Pendaftaran','Program Pengajian'],
      (d.senarai||[]).map((r,i)=>[String(i+1),r[0]||'',r[1]||'',r[2]||'']));
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     2. MEMO
  ══════════════════════════════════════════════════════════════ */
  function buildMemo(d) {
    // d: {rujukan, tarikh, kepada, drpd, perkara, isi (multiline), nama_penandatangan, jawatan_penandatangan}
    let body='';
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:26,spaceAfter:60});
    body+=para('Unit Hal Ehwal Pelajar',{center:true,size:22,spaceAfter:200});
    body+=hLine();
    body+=para('MEMO',{bold:true,center:true,size:28,spaceBefore:160,spaceAfter:200});
    body+=hLine();
    body+=spacer(120);
    // Header table
    body+=table(['',''],
      [
        ['Rujukan','KKMG.HEP.'+(d.rujukan||'')],
        ['Tarikh',d.tarikh||''],
        ['Kepada',d.kepada||''],
        ['Daripada',d.drpd||''],
        ['Perkara',d.perkara||''],
      ],
      {colWidths:[2500,7000]}
    );
    body+=spacer(120);
    body+=hLine();
    body+=spacer(120);
    // Body paragraphs
    (d.isi||'').split('\n').filter(s=>s.trim()).forEach(line=>{
      body+=para(line,{spaceAfter:160,indent:720});
    });
    body+=spacer(400);
    // Signature
    body+=para('Sekian, terima kasih.',{spaceAfter:300});
    body+=para('"BERKHIDMAT UNTUK NEGARA"',{bold:true,center:true,spaceAfter:400});
    body+=para('Yang menjalankan tugas,',{spaceAfter:200});
    body+=spacer(300);
    body+=para('( '+(d.nama_penandatangan||'_________________________________')+' )',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penandatangan||'',{spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{spaceAfter:40});
    body+=para('Kolej Komuniti Mas Gading',{spaceAfter:40});
    body+=para('Tarikh: '+d.tarikh,{spaceAfter:40});
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     3. SURAT LANTIKAN AJK
  ══════════════════════════════════════════════════════════════ */
  function buildSuratAJK(d) {
    // d: {tarikh, no_ruj, nama_penerima, jawatan_penerima, alamat_penerima,
    //     nama_program, tarikh_program, tempat_program, jawatan_dilantik, skop, tempoh,
    //     nama_penandatangan, jawatan_penandatangan}
    let body='';
    // Letterhead
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:26,spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{center:true,size:22,spaceAfter:40});
    body+=para('Lot 742–750, Off Parent Lot 726, 94000 Bau, Sarawak',{center:true,size:20,spaceAfter:40});
    body+=hLine();
    body+=spacer(120);
    body+=para('Ruj. Kami: KKMG.HEP.'+(d.no_ruj||''),{spaceAfter:60});
    body+=para('Tarikh: '+(d.tarikh||''),{spaceAfter:200});
    // Recipient
    body+=para(d.nama_penerima||'',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penerima||'',{spaceAfter:40});
    (d.alamat_penerima||'').split('\n').forEach(line=>{ body+=para(line,{spaceAfter:20}); });
    body+=spacer(160);
    // Subject
    body+=para('Tuan / Puan,',{spaceAfter:120});
    body+=para('LANTIKAN JAWATANKUASA PROGRAM: '+((d.nama_program||'').toUpperCase()),{bold:true,underline:true,size:22,spaceAfter:160});
    // Body
    body+=para('Dengan segala hormatnya, pihak Unit Hal Ehwal Pelajar, Kolej Komuniti Mas Gading ingin memaklumkan bahawa tuan/puan telah dilantik sebagai Jawatankuasa Program bagi program yang dinyatakan di bawah:',{spaceAfter:160,indent:720});
    body+=table(['Perkara','Maklumat'],[
      ['Nama Program',d.nama_program||''],
      ['Tarikh',d.tarikh_program||''],
      ['Tempat',d.tempat_program||''],
      ['Jawatan Dilantik',d.jawatan_dilantik||''],
      ['Tempoh Lantikan',d.tempoh||''],
    ],{colWidths:[3000,6500]});
    body+=para('2. Skop tugas dan tanggungjawab tuan/puan adalah seperti berikut:',{spaceAfter:100,indent:720});
    (d.skop||'').split('\n').filter(s=>s.trim()).forEach((line,i)=>{
      body+=para(`${i+1}. ${line}`,{spaceAfter:60,indent:1440});
    });
    body+=spacer(100);
    body+=para('3. Kerjasama dan komitmen tuan/puan amat dihargai. Sebarang pertanyaan sila hubungi pejabat HEP.',{spaceAfter:300,indent:720});
    body+=para('Sekian, terima kasih.',{spaceAfter:300});
    body+=para('"BERKHIDMAT UNTUK NEGARA"',{bold:true,center:true,spaceAfter:400});
    body+=para('Yang menjalankan tugas,',{spaceAfter:300});
    body+=para('( '+(d.nama_penandatangan||'_________________________________')+' )',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penandatangan||'',{spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{spaceAfter:40});
    body+=para('Kolej Komuniti Mas Gading',{spaceAfter:40});
    body+=para('Tarikh: '+d.tarikh,{spaceAfter:40});
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     4. SURAT LANTIKAN PENCERAMAH
  ══════════════════════════════════════════════════════════════ */
  function buildSuratPenceramah(d) {
    // d: {tarikh, no_ruj, nama_penerima, jawatan_penerima, org_penerima, alamat_penerima,
    //     nama_program, tarikh_program, masa_program, tempat_program, tajuk_ceramah,
    //     bayaran, nama_penandatangan, jawatan_penandatangan}
    let body='';
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:26,spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{center:true,size:22,spaceAfter:40});
    body+=para('Lot 742–750, Off Parent Lot 726, 94000 Bau, Sarawak',{center:true,size:20,spaceAfter:40});
    body+=hLine();
    body+=spacer(120);
    body+=para('Ruj. Kami: KKMG.HEP.'+(d.no_ruj||''),{spaceAfter:60});
    body+=para('Tarikh: '+(d.tarikh||''),{spaceAfter:200});
    body+=para(d.nama_penerima||'',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penerima||'',{spaceAfter:40});
    if (d.org_penerima) body+=para(d.org_penerima,{spaceAfter:40});
    (d.alamat_penerima||'').split('\n').forEach(line=>{ body+=para(line,{spaceAfter:20}); });
    body+=spacer(160);
    body+=para('Tuan / Puan,',{spaceAfter:120});
    body+=para('JEMPUTAN MENJADI PENCERAMAH / FASILITATOR: '+((d.nama_program||'').toUpperCase()),{bold:true,underline:true,size:22,spaceAfter:160});
    body+=para('Dengan hormatnya, pihak kami dengan sukacitanya menjemput tuan/puan untuk menjadi Penceramah / Fasilitator bagi program seperti berikut:',{spaceAfter:160,indent:720});
    body+=table(['Perkara','Maklumat'],[
      ['Nama Program',d.nama_program||''],
      ['Tarikh',d.tarikh_program||''],
      ['Masa',d.masa_program||''],
      ['Tempat',d.tempat_program||''],
      ['Tajuk Ceramah / Sesi',d.tajuk_ceramah||''],
      ['Honorarium',d.bayaran?'RM '+d.bayaran:'Tiada'],
    ],{colWidths:[3000,6500]});
    body+=para('2. Kehadiran tuan/puan amat diperlukan. Sila sahkan penyertaan tuan/puan kepada pegawai yang dinyatakan di bawah.',{spaceAfter:300,indent:720});
    body+=para('Sekian, terima kasih.',{spaceAfter:300});
    body+=para('"BERKHIDMAT UNTUK NEGARA"',{bold:true,center:true,spaceAfter:400});
    body+=para('Yang menjalankan tugas,',{spaceAfter:300});
    body+=para('( '+(d.nama_penandatangan||'_________________________________')+' )',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penandatangan||'',{spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{spaceAfter:40});
    body+=para('Kolej Komuniti Mas Gading',{spaceAfter:40});
    body+=para('Tarikh: '+d.tarikh,{spaceAfter:40});
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     5. SIJIL BERNOMBOR SIRI
  ══════════════════════════════════════════════════════════════ */
  function buildSijil(d) {
    // d: {nama_program, tarikh_program, tempat_program, penandatangan, jawatan_penandatangan,
    //     tema (klasik|moden|ringkas), penerima:[{nama,no_siri,jawatan_penyertaan}]}
    const penerima = d.penerima || [];
    let body='';
    penerima.forEach((p, idx) => {
      if (idx>0) body+=pageBreak();
      // Border frame using table trick
      body+=spacer(200);
      body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:28,spaceAfter:60});
      body+=para('Unit Hal Ehwal Pelajar',{center:true,size:22,spaceAfter:200});
      body+=hLine();
      body+=spacer(200);
      body+=para('SIJIL PENYERTAAN',{bold:true,center:true,size:36,spaceAfter:80});
      body+=para('Nombor Siri: '+(p.no_siri||''),{center:true,size:20,spaceAfter:300});
      body+=para('Ini adalah untuk mengesahkan bahawa',{center:true,size:22,spaceAfter:100});
      body+=para(p.nama||'',{bold:true,center:true,size:32,spaceAfter:100});
      if (p.jawatan_penyertaan) body+=para(p.jawatan_penyertaan,{center:true,size:22,spaceAfter:100});
      body+=para('telah mengambil bahagian dalam program',{center:true,size:22,spaceAfter:100});
      body+=para(d.nama_program||'',{bold:true,center:true,size:26,spaceAfter:100});
      body+=para('pada '+d.tarikh_program+' di '+d.tempat_program,{center:true,size:22,spaceAfter:400});
      body+=hLine();
      body+=spacer(200);
      body+=para('\n\n\n.......................................',{center:true,spaceAfter:40});
      body+=para('( '+(d.penandatangan||'_______________________')+' )',{bold:true,center:true,spaceAfter:40});
      body+=para(d.jawatan_penandatangan||'',{center:true,spaceAfter:40});
      body+=para('Kolej Komuniti Mas Gading',{center:true,spaceAfter:40});
      body+=para('Tarikh: '+d.tarikh_program,{center:true,spaceAfter:40});
    });
    if (!penerima.length) body+=para('(Tiada penerima sijil)',{center:true});
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     6. LAPORAN PROGRAM
  ══════════════════════════════════════════════════════════════ */
  function buildLaporan(d) {
    // d: {nama_program, tarikh, tempat, penganjur, kluster, peringkat,
    //     bilangan_daftar, bilangan_hadir, peratus_hadir,
    //     latar, objektif (multiline), pelaksanaan (multiline),
    //     pencapaian (multiline), cabaran (multiline), cadangan (multiline),
    //     perbelanjaan_sebenar:[{perkara,amaun}], jumlah_perbelanjaan,
    //     gambar_keterangan (multiline), kesimpulan,
    //     nama_penyedia, jawatan_penyedia, tarikh_laporan}
    let body='';
    // Cover
    body+=para('',{spaceAfter:400});
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:28,spaceAfter:60});
    body+=para('Unit Hal Ehwal Pelajar',{center:true,size:24,spaceAfter:400});
    body+=hLine();
    body+=para('LAPORAN PROGRAM',{bold:true,center:true,size:32,spaceBefore:200,spaceAfter:120});
    body+=para((d.nama_program||'').toUpperCase(),{bold:true,center:true,size:26,spaceAfter:300});
    body+=hLine();
    body+=spacer(300);
    body+=para('Disediakan oleh:',{bold:true,center:true,spaceAfter:60});
    body+=para(d.nama_penyedia||'',{center:true,spaceAfter:40});
    body+=para(d.jawatan_penyedia||'',{center:true,spaceAfter:40});
    body+=para('Tarikh: '+(d.tarikh_laporan||''),{center:true,spaceAfter:40});
    body+=pageBreak();
    // Program info
    body+=sectionHead('1.0  MAKLUMAT PROGRAM');
    body+=table(['Perkara','Maklumat'],[
      ['Nama Program',d.nama_program||''],
      ['Kluster',d.kluster||''],
      ['Peringkat',d.peringkat||''],
      ['Tarikh Pelaksanaan',d.tarikh||''],
      ['Tempat',d.tempat||''],
      ['Penganjur',d.penganjur||''],
      ['Bilangan Peserta Didaftar',d.bilangan_daftar||''],
      ['Bilangan Peserta Hadir',d.bilangan_hadir||''],
      ['Peratusan Kehadiran',d.peratus_hadir?d.peratus_hadir+'%':''],
    ],{colWidths:[3000,6500]});
    body+=sectionHead('2.0  LATAR BELAKANG / PENGENALAN');
    body+=para(d.latar||'',{spaceAfter:160});
    body+=sectionHead('3.0  OBJEKTIF PROGRAM');
    (d.objektif||'').split('\n').filter(s=>s.trim()).forEach((o,i)=>{ body+=para((i+1)+'. '+o,{spaceAfter:60}); });
    body+=sectionHead('4.0  PELAKSANAAN PROGRAM');
    (d.pelaksanaan||'').split('\n').filter(s=>s.trim()).forEach((p,i)=>{ body+=para((i+1)+'. '+p,{spaceAfter:60}); });
    body+=sectionHead('5.0  PENCAPAIAN / HASIL PROGRAM');
    (d.pencapaian||'').split('\n').filter(s=>s.trim()).forEach((p,i)=>{ body+=para((i+1)+'. '+p,{spaceAfter:60}); });
    body+=sectionHead('6.0  CABARAN DAN MASALAH');
    (d.cabaran||'').split('\n').filter(s=>s.trim()).forEach((c,i)=>{ body+=para((i+1)+'. '+c,{spaceAfter:60}); });
    body+=sectionHead('7.0  CADANGAN PENAMBAHBAIKAN');
    (d.cadangan||'').split('\n').filter(s=>s.trim()).forEach((c,i)=>{ body+=para((i+1)+'. '+c,{spaceAfter:60}); });
    // Expenditure
    body+=sectionHead('8.0  PERBELANJAAN SEBENAR');
    const belanja=d.perbelanjaan_sebenar||[];
    body+=table(['Perkara','Amaun (RM)'],belanja.map(r=>[r[0]||'',(Number(r[1])||0).toFixed(2)]));
    body+=para('JUMLAH PERBELANJAAN SEBENAR: RM '+(Number(d.jumlah_perbelanjaan)||0).toFixed(2),{bold:true,spaceAfter:80});
    body+=sectionHead('9.0  DOKUMENTASI / GAMBAR');
    body+=para('(Sila lampirkan gambar program)',{spaceAfter:80,italic:true});
    (d.gambar_keterangan||'').split('\n').filter(s=>s.trim()).forEach((g,i)=>{ body+=para('Gambar '+(i+1)+': '+g,{spaceAfter:60}); });
    body+=sectionHead('10.0  KESIMPULAN');
    body+=para(d.kesimpulan||'',{spaceAfter:300});
    // Signature
    body+=hLine();
    body+=spacer(160);
    body+=para('Disediakan oleh:',{bold:true,spaceAfter:200});
    body+=para('( '+(d.nama_penyedia||'_________________________________')+' )',{bold:true,spaceAfter:40});
    body+=para(d.jawatan_penyedia||'',{spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar',{spaceAfter:40});
    body+=para('Kolej Komuniti Mas Gading',{spaceAfter:40});
    body+=para('Tarikh: '+(d.tarikh_laporan||''),{spaceAfter:40});
    body+=spacer(300);
    body+=para('Disemak oleh:',{bold:true,spaceAfter:200});
    body+=para('( ___________________________________ )',{bold:true,spaceAfter:40});
    body+=para('_________________________________',{spaceAfter:40});
    body+=para('Tarikh: __________________________',{spaceAfter:40});
    return pack(body);
  }

  /* ══════════════════════════════════════════════════════════════
     7. PENILAIAN KURSUS (borang dalam pelayar, tiada DOCX)
        buildPenilaian digunakan untuk jana ringkasan hasil sahaja
  ══════════════════════════════════════════════════════════════ */
  function buildRingkasanPenilaian(d) {
    // d: {nama_program, tarikh, soalan:[{soalan, purata, min, maks}], ulasan:[string], respons}
    let body='';
    body+=para('KOLEJ KOMUNITI MAS GADING',{bold:true,center:true,size:26,spaceAfter:40});
    body+=para('Unit Hal Ehwal Pelajar — Ringkasan Penilaian Kursus',{center:true,size:22,spaceAfter:200});
    body+=hLine();
    body+=spacer(120);
    body+=table(['Perkara','Maklumat'],[
      ['Nama Program',d.nama_program||''],
      ['Tarikh',d.tarikh||''],
      ['Jumlah Respons',String(d.respons||0)],
    ],{colWidths:[3000,6500]});
    body+=sectionHead('KEPUTUSAN PENILAIAN (Skala 1–5)');
    body+=table(
      ['Dimensi / Soalan','Purata','Min','Maks'],
      (d.soalan||[]).map(s=>[s.soalan||'',String(s.purata||''),String(s.min||''),String(s.maks||'')])
    );
    if ((d.ulasan||[]).length) {
      body+=sectionHead('ULASAN PESERTA');
      d.ulasan.forEach((u,i)=>{ body+=para(`"${u}"`,{italic:true,spaceAfter:80,indent:720}); });
    }
    body+=spacer(200);
    body+=para('Laporan ini dijana secara automatik oleh Sistem HEP KKMG.',{italic:true,center:true,size:20});
    return pack(body);
  }

  /* ── Public API ─────────────────────────────────────────────── */
  root.KKMG = {
    build: buildKK,
    buildMemo,
    buildSuratAJK,
    buildSuratPenceramah,
    buildSijil,
    buildLaporan,
    buildRingkasanPenilaian,
  };

})(typeof window === 'undefined' ? globalThis : window);
