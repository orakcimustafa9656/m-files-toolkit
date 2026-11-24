	Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# PROGRAM DURUMLARI
# ============================================================

$script:excelPath = $null
$script:programdanCik = $false

# ============================================================
# EXCEL DOSYASI SEÇ
# ============================================================

function Select-ExcelFile {

    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    $dialog.Filter = "Excel Dosyası (*.xlsx)|*.xlsx"
    $dialog.Title = "Makbuz için Excel dosyasını seçin"
    $dialog.Multiselect = $false
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

# ============================================================
# TUTARI YAZIYA ÇEVİR
# ============================================================

function Convert-NumberToText {

    param(
        [double]$number
    )

    $birler = @(
        "",
        "bir",
        "iki",
        "üç",
        "dört",
        "beş",
        "altı",
        "yedi",
        "sekiz",
        "dokuz"
    )

    $onlar = @(
        "",
        "on",
        "yirmi",
        "otuz",
        "kırk",
        "elli",
        "altmış",
        "yetmiş",
        "seksen",
        "doksan"
    )

    $binler = @(
        "",
        "bin",
        "milyon",
        "milyar"
    )

    if ($number -eq 0) {
        return "Sıfır lira"
    }

    # Türkiye formatı:
    # nokta = binlik
    # virgül = ondalık

    $numStr = "{0:N2}" -f [math]::Round($number, 2)

    $numStr = $numStr.Replace(".", "")

    $parts = $numStr -split ","

    $tam = [int]$parts[0]

    $kuruş = if ($parts.Count -gt 1) {
        [int]$parts[1]
    }
    else {
        0
    }

    $text = ""

    $i = 0

    while ($tam -gt 0) {

        $n = $tam % 1000

        if ($n -ne 0) {

            $yuz = [math]::Floor($n / 100)
            $on = [math]::Floor(($n % 100) / 10)
            $bir = $n % 10

            $kisim = ""

            # Yüzler
            if ($yuz -gt 0) {

                if ($yuz -eq 1) {
                    $kisim += "yüz "
                }
                else {
                    $kisim += $birler[$yuz] + " yüz "
                }
            }

            # Onlar
            if ($on -gt 0) {
                $kisim += $onlar[$on] + " "
            }

            # Birler
            if ($bir -gt 0) {
                $kisim += $birler[$bir] + " "
            }

            # Bin
            if ($i -eq 1 -and $n -eq 1) {
                $kisim = "bin "
            }
            else {
                $kisim += $binler[$i] + " "
            }

            $text = $kisim + $text
        }

        $tam = [math]::Floor($tam / 1000)
        $i++
    }

    $text = $text.Trim()

    # Kuruş
    if ($kuruş -gt 0) {

        $kuruşYazi = ""

        $on = [math]::Floor($kuruş / 10)
        $bir = $kuruş % 10

        if ($on -gt 0) {
            $kuruşYazi += $onlar[$on] + " "
        }

        if ($bir -gt 0) {
            $kuruşYazi += $birler[$bir] + " "
        }

        $kuruşYazi = $kuruşYazi.Trim()

        $text += " lira $kuruşYazi kuruş"
    }
    else {
        $text += " lira"
    }

    # İlk harfi büyüt
    if ($text.Length -gt 1) {

        $text =
            $text.Substring(0, 1).ToUpper() +
            $text.Substring(1)
    }

    return $text
}

# ============================================================
# MAKBUZLARI OLUŞTUR
# ============================================================

function Create-Receipts {

    if (
        [string]::IsNullOrWhiteSpace(
            $script:excelPath
        )
    ) {

        [System.Windows.Forms.MessageBox]::Show(
            "Önce [1] Excel Dosyası Seç bölümünden bir Excel dosyası seçmelisiniz.",
            "Uyarı",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return
    }

    $excel = $null
    $wb = $null
    $ws = $null
    $word = $null
    $doc = $null
    $sel = $null

    try {

        # ====================================================
        # EXCEL BAŞLAT
        # ====================================================

        try {

            $excel =
                New-Object -ComObject Excel.Application

            $excel.Visible = $false
            $excel.DisplayAlerts = $false
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Excel açılamadı. Kurulum veya COM hatası olabilir.",
                "Hata",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null

            return
        }

        # ====================================================
        # WORD BAŞLAT
        # ====================================================

        try {

            $word =
                New-Object -ComObject Word.Application

            $word.Visible = $true
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Word açılamadı. Kurulum veya COM hatası olabilir.",
                "Hata",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null

            return
        }

        # ====================================================
        # EXCEL DOSYASINI AÇ
        # ====================================================

        try {

            $wb =
                $excel.Workbooks.Open(
                    $script:excelPath,
                    $null,
                    $true
                )

            $ws =
                $wb.Worksheets.Item(1)
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Excel dosyası açılamadı.`n`nDosya kilitli veya bozuk olabilir.",
                "Hata",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null

            return
        }

        # ====================================================
        # MÜŞTERİ SAYISINI BUL
        # ====================================================

        $lastRow =
            $ws.UsedRange.Row +
            $ws.UsedRange.Rows.Count -
            1

        $customerCount = 0

        for ($r = 2; $r -le $lastRow; $r++) {

            $hesapAdi =
                [string]$ws.Cells.Item($r, 2).Text

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $hesapAdi
                )
            ) {

                $borc =
                    if (
                        $ws.Cells.Item($r, 6).Value() -ne $null
                    ) {
                        try {
                            [double]$ws.Cells.Item($r, 6).Value()
                        }
                        catch {
                            0
                        }
                    }
                    else {
                        0
                    }

                $alacak =
                    if (
                        $ws.Cells.Item($r, 7).Value() -ne $null
                    ) {
                        try {
                            [double]$ws.Cells.Item($r, 7).Value()
                        }
                        catch {
                            0
                        }
                    }
                    else {
                        0
                    }

                if (
                    ($borc -gt 0) -or
                    ($alacak -gt 0)
                ) {
                    $customerCount++
                }
            }
        }

        if ($customerCount -eq 0) {

            [System.Windows.Forms.MessageBox]::Show(
                "Excel dosyasında uygun müşteri/bakiye bilgisi bulunamadı.",
                "Bilgi",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            return
        }

        # ====================================================
        # ONAY
        # ====================================================

        $confirmMessage = @"
$customerCount adet makbuz oluşturulacak.

Word belgesi hazırlanacak.

Devam etmek istiyor musunuz?
"@

        $confirm =
            [System.Windows.Forms.MessageBox]::Show(
                $confirmMessage,
                "Makbuzları Oluştur",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

        if (
            $confirm -ne
            [System.Windows.Forms.DialogResult]::Yes
        ) {
            return
        }

        # ====================================================
        # WORD BELGESİ OLUŞTUR
        # ====================================================

        try {

            $doc =
                $word.Documents.Add()

            $sel =
                $word.Selection
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Word belgesi oluşturulamadı.",
                "Hata",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null

            return
        }

        # ====================================================
        # EXCEL SATIRLARINI OKU
        # ====================================================

        $row = 2

        while (
            $row -le $lastRow
        ) {

            try {

                $hesapAdi =
                    [string]$ws.Cells.Item($row, 2).Text

                if (
                    [string]::IsNullOrWhiteSpace(
                        $hesapAdi
                    )
                ) {
                    $row++
                    continue
                }

                $borcBakiye = 0

                if (
                    $ws.Cells.Item($row, 6).Value() -ne $null
                ) {

                    try {
                        $borcBakiye =
                            [double]$ws.Cells.Item($row, 6).Value()
                    }
                    catch {
                        $borcBakiye = 0
                    }
                }

                $alacakBakiye = 0

                if (
                    $ws.Cells.Item($row, 7).Value() -ne $null
                ) {

                    try {
                        $alacakBakiye =
                            [double]$ws.Cells.Item($row, 7).Value()
                    }
                    catch {
                        $alacakBakiye = 0
                    }
                }

            }
            catch {

                Write-Warning "Satır $row okunamadı, atlandı."

                $row++
                continue
            }

            # =================================================
            # TUTARI BELİRLE
            # =================================================

            if ($borcBakiye -gt 0) {

                $tutar = $borcBakiye
            }
            elseif ($alacakBakiye -gt 0) {

                $tutar = $alacakBakiye
            }
            else {

                $row++
                continue
            }

            # =================================================
            # MAKBUZ BAŞLANGICI
            # =================================================

            $sel.Font.Name = "Courier New"
            $sel.Font.Size = 11
            $sel.Font.Bold = 1
            $sel.ParagraphFormat.Alignment = 1

            $sel.TypeText(
                "TAHSİLAT/TEDİYE MAKBUZU"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "Sayın: $hesapAdi"
            )

            $sel.TypeParagraph()

            # =================================================
            # AYIRICI
            # =================================================

            $sel.ParagraphFormat.Alignment = 0
            $sel.Font.Bold = 0

            $sel.TypeText(
                "------------------------------------------------------------"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "Tarih   : ______"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "Seri    : ______"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "Sıra No : ______"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "------------------------------------------------------------"
            )

            $sel.TypeParagraph()

            # =================================================
            # DETAY BAŞLIĞI
            # =================================================

            $sel.TypeText(
                "No   Banka       Şube      Hesap/Iban     Çek No   Vade   Meblağ"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "------------------------------------------------------------"
            )

            $sel.TypeParagraph()

            # =================================================
            # 5 SATIRLIK DETAY ALANI
            # =================================================

            for ($i = 1; $i -le 5; $i++) {

                $rowData =
                    @(
                        $i,
                        '',
                        '',
                        '',
                        '',
                        '',
                        ''
                    )

                $line =
                    "{0,-4}{1,-12}{2,-10}{3,-15}{4,-8}{5,-7}{6,-10}" `
                    -f $rowData

                $sel.TypeText($line)

                $sel.TypeParagraph()
            }

            $sel.TypeText(
                "------------------------------------------------------------"
            )

            $sel.TypeParagraph()

            # =================================================
            # KAŞE / İMZA
            # =================================================

            $sel.TypeText(
                "Tahsil Eden & Tediye Eden İmza"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "------------------------------"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "|                              |"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "|                              |"
            )

            $sel.TypeParagraph()

            $sel.TypeText(
                "------------------------------"
            )

            $sel.TypeParagraph()

            # =================================================
            # ÖDEME TÜRÜ
            # =================================================

            $sel.TypeText(
                "Ödeme Türü : Nakit □   Kredi Kartı □   Çek □   Senet □"
            )

            $sel.TypeParagraph()

            # =================================================
            # TOPLAM
            # =================================================

            $sel.TypeText(
                "Toplam     : {0:N2} TL" -f $tutar
            )

            $sel.TypeParagraph()

            # =================================================
            # YAZI İLE
            # =================================================

            $yaziIle =
                Convert-NumberToText $tutar

            $sel.TypeText(
                "Yalnız $yaziIle ödenmiştir."
            )

            $sel.TypeParagraph()
            $sel.TypeParagraph()

            # =================================================
            # YENİ SAYFA
            # =================================================

            $sel.InsertBreak(7)

            $row++
        }

        # ====================================================
        # TAMAMLANDI
        # ====================================================

        [System.Windows.Forms.MessageBox]::Show(
            "$customerCount adet makbuz Word belgesinde oluşturuldu.",
            "İşlem Tamamlandı",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "İşlem sırasında bir hata oluştu.`n`n$($_.Exception.Message)",
            "Hata",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {

        # ====================================================
        # WORD BELGESİ KULLANILMAYA DEVAM EDECEĞİ İÇİN
        # BELGE VE WORD KAPATILMIYOR.
        # SADECE EXCEL KAPATILIYOR.
        # ====================================================

        if ($ws -ne $null) {

            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
            }
            catch {}

            $ws = $null
        }

        if ($wb -ne $null) {

            try {
                $wb.Close($false)
            }
            catch {}

            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            }
            catch {}

            $wb = $null
        }

        if ($excel -ne $null) {

            try {
                $excel.Quit()
            }
            catch {}

            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            catch {}

            $excel = $null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ============================================================
# ANA MENÜ
# ============================================================

function Show-Menu {

    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "🧾 MAKBUZ OLUŞTURUCU" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""

    # --------------------------------------------------------
    # MİZAN UYARISI
    # --------------------------------------------------------

    Write-Host "ÖNEMLİ:" `
        -ForegroundColor Yellow

    Write-Host "Luca'dan alınan mizan HAM VERİ olarak kullanılmalıdır." `
        -ForegroundColor Yellow

    Write-Host "İlk 5 Satır Üst Bilgi," `
        -ForegroundColor Yellow

    Write-Host "Grup hesapları ve Genel toplamlar" `
        -ForegroundColor Yellow

    Write-Host "SİLİNMELİDİR." `
        -ForegroundColor Yellow

    Write-Host ""

    # --------------------------------------------------------
    # EXCEL DURUMU
    # --------------------------------------------------------

    if (
        [string]::IsNullOrWhiteSpace(
            $script:excelPath
        )
    ) {

        Write-Host "Excel Dosyası : SEÇİLMEDİ" `
            -ForegroundColor Yellow
    }
    else {

        $fileName =
            Split-Path `
                $script:excelPath `
                -Leaf

        Write-Host "Excel Dosyası : $fileName" `
            -ForegroundColor Green
    }

    Write-Host ""

    # --------------------------------------------------------
    # MENÜ
    # --------------------------------------------------------

    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host "  1. 📂  Excel Dosyası Seç"
    Write-Host "  2. 🧾  Makbuzları Oluştur"
    Write-Host "  Q  🚪  Çıkış"
    Write-Host ("-" * 72) -ForegroundColor DarkGray

    Write-Host ""
}

# ============================================================
# MENÜ SONRASI BEKLE
# ============================================================

function Pause-Menu {

    Write-Host ""
    Write-Host "Devam etmek için bir tuşa basın..." `
        -ForegroundColor DarkGray

    [void][System.Console]::ReadKey($true)
}

# ============================================================
# ANA PROGRAM DÖNGÜSÜ
# ============================================================

do {

    Show-Menu

    $choice =
        Read-Host "Seçiminiz"

    switch ($choice) {

        # ====================================================
        # 1 - EXCEL DOSYASI SEÇ
        # ====================================================

        "1" {

            $selected =
                Select-ExcelFile

            if ($selected) {

                $script:excelPath =
                    $selected

                Write-Host ""
                Write-Host "Excel dosyası seçildi:" `
                    -ForegroundColor Green

                Write-Host $script:excelPath `
                    -ForegroundColor White

                Pause-Menu
            }
        }

        # ====================================================
        # 2 - MAKBUZLARI OLUŞTUR
        # ====================================================

        "2" {

            Create-Receipts
        }

        # ====================================================
        # Q - ÇIKIŞ
        # ====================================================

        "Q" {

            Clear-Host

            Write-Host ""
            Write-Host "Program kapatılıyor..." `
                -ForegroundColor Yellow

            Write-Host ""

            $script:programdanCik = $true
        }

        "q" {

            Clear-Host

            Write-Host ""
            Write-Host "Program kapatılıyor..." `
                -ForegroundColor Yellow

            Write-Host ""

            $script:programdanCik = $true
        }

        # ====================================================
        # GEÇERSİZ SEÇİM
        # ====================================================

        default {

            Write-Host ""
            Write-Host "Geçersiz seçim. Lütfen 1-2 veya Q seçin." `
                -ForegroundColor Red

            Pause-Menu
        }
    }

}
while (-not $script:programdanCik)

