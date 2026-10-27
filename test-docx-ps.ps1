$docxJs = Get-Content -Path "docx.js" -Raw -Encoding UTF8
$html   = Get-Content -Path "index.html" -Raw -Encoding UTF8
$pass = 0; $fail = 0

function Check($label, $content, $needle) {
    if ($content.Contains($needle)) {
        Write-Host "  LULUS  $label" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  GAGAL  $label  -- '$needle' tidak ditemui" -ForegroundColor Red
        $script:fail++
    }
}

Write-Host "`n=== docx.js ===" -ForegroundColor Cyan
Check "Tandatangan ZIP local"     $docxJs "0x04034b50"
Check "Tandatangan ZIP central"   $docxJs "0x02014b50"
Check "Tandatangan ZIP EOCD"      $docxJs "0x06054b50"
Check "Fungsi CRC32"              $docxJs "crc32"
Check "Fungsi build"              $docxJs "function build"
Check "Fungsi zip"                $docxJs "function zip"
Check "Fungsi para"               $docxJs "function para"
Check "Fungsi table"              $docxJs "function table"
Check "Export KKMG"               $docxJs "root.KKMG"
Check "Content_Types.xml"         $docxJs "[Content_Types].xml"
Check "word/document.xml"         $docxJs "word/document.xml"
Check "word/styles.xml"           $docxJs "word/styles.xml"
Check "_rels/.rels"               $docxJs "_rels/.rels"
Check "XML namespace Word"        $docxJs "wordprocessingml/2006/main"
Check "w:document"                $docxJs "w:document"
Check "w:body"                    $docxJs "w:body"
Check "w:tbl (jadual)"            $docxJs "w:tbl"
Check "Muka depan KERTAS KERJA"   $docxJs "KERTAS KERJA"
Check "Nama kolej"                $docxJs "KOLEJ KOMUNITI MAS GADING"
Check "Lampiran 1 AJK"            $docxJs "Lampiran 1"
Check "Lampiran 2 Bajet"          $docxJs "Lampiran 2"
Check "Lampiran 3 Peserta"        $docxJs "Lampiran 3"
Check "Pengesahan institusi"      $docxJs "PENGESAHAN INSTITUSI"
Check "Ruangan kelulusan"         $docxJs "KELULUSAN"
Check "Semakan kewangan"          $docxJs "SEMAKAN PERUNTUKAN KEWANGAN"
Check "Tandatangan kosong"        $docxJs "Disediakan oleh"
Check "Disemak kosong"            $docxJs "Disemak oleh"
Check "Carta perbatuan"           $docxJs "CARTA PERBATUAN"
Check "Tentatif program"          $docxJs "PERINCIAN / TENTATIF"
Check "Page break"                $docxJs "pageBreak"
Check "A4 page size"              $docxJs "11906"
Check "Section heading"           $docxJs "sectionHead"

Write-Host "`n=== index.html ===" -ForegroundColor Cyan
Check "Muat KKMG.build"          $html "KKMG.build"
Check "DynTable widget"          $html "DynTable"
Check "Butang jana"              $html "btn-generate"
Check "Simpan draf"              $html "btn-save"
Check "Buka draf"                $html "btn-load"
Check "Import fail"              $html "file-import"
Check "Perbatuan wrap"           $html "perbatuan-wrap"
Check "Tentatif wrap"            $html "tentatif-wrap"
Check "AJK wrap"                 $html "ajk-wrap"
Check "Bajet wrap"               $html "bajet-wrap"
Check "Senarai wrap"             $html "senarai-wrap"
Check "Jumlah bajet"             $html "budget-total"
Check "Kira bajet"               $html "calcBudget"
Check "Script docx.js"           $html 'src="docx.js"'
Check "Nav Kertas Kerja"         $html "Kertas Kerja"
Check "Nav Memo"                 $html "Memo"
Check "Nav Analisis aktif"       $html "nav-analisis"
Check "Nav Sijil"                $html "Penjanaan Sijil"
Check "Nav Laporan"              $html "Laporan Program"
Check "Nav Penilaian"            $html "Penilaian Kursus"
Check "Nav Analisis"             $html "Analisis Keberkesanan"
Check "Simpan draf label"        $html "Simpan draf"
Check "Buka draf label"          $html "Buka draf"
Check "Borang jana"              $html "btn-generate"
Check "Nota tandatangan"         $html "Tandatangan"
Check "Validate required"        $html "required"
Check "Format tarikh"            $html "formatDate"
Check "Borang seksi 01"          $html "01"
Check "Borang seksi 05"          $html "05"
Check "Select kluster"           $html "Pilih kluster"
Check "Kluster Kepimpinan"       $html "Kepimpinan"
Check "Kluster Sukan"            $html "Sukan"
Check "Kluster Sukarelawan"      $html "Sukarelawan"
Check "Kluster OKU"              $html "OKU"
Check "Select peringkat"         $html "Pilih peringkat"
Check "Peringkat Kolej"          $html "Kolej"
Check "Peringkat Kebangsaan"     $html "Kebangsaan"
Check "Peringkat Antarabangsa"   $html "Antarabangsa"
Check "Dashboard view"           $html "view-dash"
Check "Kertas Kerja view"        $html "view-kk"
Check "Stat total"               $html "chart-dash-jenis"
Check "Nav Dashboard"            $html "nav-dash"

Write-Host ""
Write-Host "KEPUTUSAN: $pass lulus, $fail gagal" -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Yellow" })
if ($fail -eq 0) {
    Write-Host "SEMUA UJIAN LULUS - Buka index.html dalam pelayar untuk ujian visual." -ForegroundColor Green
} else {
    Write-Host "ADA UJIAN GAGAL - Semak output di atas." -ForegroundColor Red
}
