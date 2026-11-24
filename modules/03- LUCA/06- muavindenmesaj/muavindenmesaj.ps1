# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

# muavindenmesaj_pdf_v3_4_1_HTML_PDF_TEMPLATE_SYNC
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# PROGRAM YOLLARI
# ============================================================

$programDir = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($programDir)) {
    $programDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

$settingsPath = Join-Path (Get-MFTModuleDataPath -ModuleId 'muavindenmesaj' -Category Config) 'muavindenmesaj_config.json'

function New-EmptySettings {
    return [ordered]@{
        MesajSablonu = ""
    }
}

function Normalize-LineEndings {
    param(
        [AllowEmptyString()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    $Text = $Text.Replace("`r`n", "`n")
    $Text = $Text.Replace("`r", "`n")
    $Text = $Text.Replace("`n", "`r`n")

    return $Text
}

function Load-Settings {

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        return New-EmptySettings
    }

    try {
        $json = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($json)) {
            return New-EmptySettings
        }

        $saved = $json | ConvertFrom-Json

        $mesajSablonu = ""

        if ($null -ne $saved.MesajSablonu) {
            $mesajSablonu = Normalize-LineEndings -Text ([string]$saved.MesajSablonu)
        }

        return [ordered]@{
            MesajSablonu = $mesajSablonu
        }
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "Ayar dosyası okunamadı.`n`n" +
            "Program boş ayarlarla devam edecek.`n`n" +
            "Dosya:`n$settingsPath`n`n" +
            "Hata:`n$($_.Exception.Message)",
            "Ayar Dosyası Hatası",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return New-EmptySettings
    }
}

$settings = Load-Settings

function Save-Settings {

    $tempPath = "$settingsPath.tmp"

    try {

        if (-not (Test-Path -LiteralPath $programDir)) {
            New-Item -ItemType Directory -Path $programDir -Force -ErrorAction Stop | Out-Null
        }

        $settings.MesajSablonu =
            Normalize-LineEndings -Text ([string]$settings.MesajSablonu)

        $json = $settings | ConvertTo-Json -Depth 5

        [System.IO.File]::WriteAllText(
            $tempPath,
            $json,
            (New-Object System.Text.UTF8Encoding($false))
        )

        if (Test-Path -LiteralPath $settingsPath) {
            Remove-Item -LiteralPath $settingsPath -Force -ErrorAction Stop
        }

        Move-Item -LiteralPath $tempPath -Destination $settingsPath -Force -ErrorAction Stop

        return $true
    }
    catch {

        if (Test-Path -LiteralPath $tempPath) {
            try {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
            catch {}
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Ayarlar kaydedilemedi.`n`n" +
            "Dosya:`n$settingsPath`n`n" +
            "Hata:`n$($_.Exception.Message)",
            "Ayar Kaydetme Hatası",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null

        return $false
    }
}

# ============================================================
# MUAVİN EXCEL DOSYASI SEÇ
# ============================================================

function Select-MuavinFile {

    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    $dialog.Title = "Muavin Kasa Excel Dosyasını Seçin"
    $dialog.Filter =
        "Excel Dosyaları (*.xlsx;*.xls)|*.xlsx;*.xls|Tüm Dosyalar (*.*)|*.*"
    $dialog.Multiselect = $false

    $desktopPath = [Environment]::GetFolderPath("Desktop")
    if (
        -not [string]::IsNullOrWhiteSpace($desktopPath) -and
        (Test-Path -LiteralPath $desktopPath)
    ) {
        $dialog.InitialDirectory = $desktopPath
    }

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

# ============================================================
# DEĞİŞKEN EKLEME
# ============================================================

function Insert-TextAtCursor {

    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$TextBox,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    try {

        $start = $TextBox.SelectionStart
        $length = $TextBox.SelectionLength

        if ($length -gt 0) {

            $TextBox.SelectedText = $Text

            $TextBox.SelectionStart = $start + $Text.Length
            $TextBox.SelectionLength = 0
        }
        else {

            $TextBox.Text = $TextBox.Text.Insert($start, $Text)

            $TextBox.SelectionStart = $start + $Text.Length
            $TextBox.SelectionLength = 0
        }

        $TextBox.Focus()
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "Değişken eklenirken hata oluştu.`n`n" +
            "$($_.Exception.Message)",
            "Hata",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# ============================================================
# MESAJ ŞABLONU
# ============================================================

function Edit-MessageTemplate {

    $form = New-Object System.Windows.Forms.Form

    $form.Text = "Mesaj Şablonunu Düzenle"
    $form.Size = New-Object System.Drawing.Size(900, 650)
    $form.MinimumSize = New-Object System.Drawing.Size(800, 580)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text =
        "Mesajınızı aşağıdaki alana yazın. Değişkenleri istediğiniz yere ekleyebilirsiniz."
    $infoLabel.Location = New-Object System.Drawing.Point(15, 15)
    $infoLabel.AutoSize = $true
    $infoLabel.Font =
        New-Object System.Drawing.Font(
            "Segoe UI",
            10,
            [System.Drawing.FontStyle]::Bold
        )
    $form.Controls.Add($infoLabel)

    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Text =
        "Değişkenler: {MUSTERI_ADI} = Müşteri adı    |    {TUTAR} = Güncel borç bakiyesi"
    $helpLabel.Location = New-Object System.Drawing.Point(15, 45)
    $helpLabel.AutoSize = $true
    $helpLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($helpLabel)

    $insertLabel = New-Object System.Windows.Forms.Label
    $insertLabel.Text = "Değişken Ekle:"
    $insertLabel.Location = New-Object System.Drawing.Point(15, 82)
    $insertLabel.AutoSize = $true
    $form.Controls.Add($insertLabel)

    $customerButton = New-Object System.Windows.Forms.Button
    $customerButton.Text = "Müşteri Adı"
    $customerButton.Size = New-Object System.Drawing.Size(130, 32)
    $customerButton.Location = New-Object System.Drawing.Point(115, 76)
    $form.Controls.Add($customerButton)

    $amountButton = New-Object System.Windows.Forms.Button
    $amountButton.Text = "Tutar"
    $amountButton.Size = New-Object System.Drawing.Size(100, 32)
    $amountButton.Location = New-Object System.Drawing.Point(255, 76)
    $form.Controls.Add($amountButton)

    $templateBox = New-Object System.Windows.Forms.TextBox
    $templateBox.Multiline = $true
    $templateBox.ScrollBars = "Both"
    $templateBox.AcceptsReturn = $true
    $templateBox.AcceptsTab = $true
    $templateBox.WordWrap = $true
    $templateBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $templateBox.Location = New-Object System.Drawing.Point(15, 120)
    $templateBox.Size = New-Object System.Drawing.Size(850, 420)
    $templateBox.Anchor =
        [System.Windows.Forms.AnchorStyles]::Top `
        -bor [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Left `
        -bor [System.Windows.Forms.AnchorStyles]::Right
    $templateBox.Text = $settings.MesajSablonu
    $form.Controls.Add($templateBox)

    $customerButton.Add_Click({
        Insert-TextAtCursor -TextBox $templateBox -Text "{MUSTERI_ADI}"
    })

    $amountButton.Add_Click({
        Insert-TextAtCursor -TextBox $templateBox -Text "{TUTAR}"
    })

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Kaydet"
    $saveButton.Size = New-Object System.Drawing.Size(110, 38)
    $saveButton.Anchor =
        [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Right
    $saveButton.Location = New-Object System.Drawing.Point(635, 555)

    $saveButton.Add_Click({

        if ([string]::IsNullOrWhiteSpace($templateBox.Text)) {

            [System.Windows.Forms.MessageBox]::Show(
                "Mesaj şablonu boş bırakılamaz.",
                "Uyarı",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $templateBox.Focus()
            return
        }

        $oldTemplate = $settings.MesajSablonu

        $newTemplate =
            Normalize-LineEndings -Text $templateBox.Text

        $settings.MesajSablonu = $newTemplate

        if (Save-Settings) {

            $templateBox.Text = $settings.MesajSablonu

            $form.DialogResult =
                [System.Windows.Forms.DialogResult]::OK

            $form.Close()
        }
        else {
            $settings.MesajSablonu = $oldTemplate
        }
    })

    $form.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "İptal"
    $cancelButton.Size = New-Object System.Drawing.Size(110, 38)
    $cancelButton.Anchor =
        [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Right
    $cancelButton.Location = New-Object System.Drawing.Point(755, 555)
    $cancelButton.DialogResult =
        [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $saveButton
    $form.CancelButton = $cancelButton

    [void]$form.ShowDialog()
}

# ============================================================
# FORMATLAMA
# ============================================================

function Format-TurkishAmount {

    param(
        [Parameter(Mandatory = $true)]
        [decimal]$Value
    )

    return $Value.ToString(
        "N2",
        [System.Globalization.CultureInfo]::GetCultureInfo("tr-TR")
    )
}

function Get-SafeFileName {

    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $invalidChars =
        [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($char in $invalidChars) {
        $Name = $Name.Replace([string]$char, "_")
    }

    $Name = $Name.Trim()

    while ($Name.EndsWith(".")) {
        $Name = $Name.Substring(0, $Name.Length - 1)
    }

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "Musteri"
    }

    return $Name
}

function Release-ComObject {

    param(
        [Parameter(Mandatory = $false)]
        $Object
    )

    if ($null -ne $Object) {

        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        }
        catch {}
    }
}

# ============================================================
# MUAVİN EXCEL'İNİ OKU
# ============================================================
#
# Excel yapısı:
#
# 120.A004 Adem Kayrak
# TARİH | TİP | FİŞ NO | AÇIKLAMA | BORÇ | ALACAK | BAKİYE | B/A
# ...
# Nakli Yekün Hariç
# Genel Toplam
# ...
# 120.A005 Ahmet Akyüz
#
# Program BAKİYE sütunundaki önbellek/formül değerine güvenmez.
# Güncel bakiye:
#
#     BORÇ toplamı - ALACAK toplamı
#
# şeklinde hesaplanır.
# Ayrıca satır satır çalışan bakiye yeniden oluşturulur.

function Get-MuavinCustomerData {

    param(
        [Parameter(Mandatory = $true)]
        [string]$ExcelPath
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    $customers =
        New-Object System.Collections.Generic.List[object]

    try {

        Write-Host ""
        Write-Host "Muavin Excel dosyası okunuyor..." -ForegroundColor Cyan

        $excel = New-Object -ComObject Excel.Application

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # Arka planda Excel'in hesaplamayı güncellemesi sağlanır.
        try {
            $excel.Calculate
        }
        catch {}

        $workbook =
            $excel.Workbooks.Open(
                $ExcelPath,
                $null,
                $true
            )

        $worksheet =
            $workbook.Worksheets.Item(1)

        $usedRange =
            $worksheet.UsedRange

        $usedStartRow = $usedRange.Row
        $usedStartCol = $usedRange.Column

        $usedValues = $usedRange.Value2

        # Tek hücrelik UsedRange durumuna karşı dizi/tek değer ayrımı.
        $isArray =
            $usedValues -is [System.Array]

        $rowCount =
            if ($isArray) {
                $usedValues.GetLength(0)
            }
            else {
                1
            }

        $colCount =
            if ($isArray) {
                $usedValues.GetLength(1)
            }
            else {
                1
            }

        function Get-CellValueLocal {
            param(
                [int]$ExcelRow,
                [int]$ExcelCol
            )

            $localRow = $ExcelRow - $usedStartRow + 1
            $localCol = $ExcelCol - $usedStartCol + 1

            if (
                $localRow -lt 1 -or
                $localRow -gt $rowCount -or
                $localCol -lt 1 -or
                $localCol -gt $colCount
            ) {
                return $null
            }

            if ($isArray) {
                return $usedValues[$localRow, $localCol]
            }

            return $usedValues
        }

        $current = $null

        for (
            $excelRow = $usedStartRow;
            $excelRow -lt ($usedStartRow + $rowCount);
            $excelRow++
        ) {

            $a =
                [string](Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 1)

            $b =
                [string](Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 2)

            $c =
                [string](Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 3)

            $d =
                [string](Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 4)

            $eRaw =
                Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 5

            $fRaw =
                Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 6

            $aTrim = $a.Trim()

            # ------------------------------------------------
            # YENİ CARİ HESAP BLOĞU
            # ------------------------------------------------

            if (
                $aTrim -match '^120\.\S+\s+(.+)$'
            ) {

                if ($null -ne $current) {

                    $current.BakiyeSigned =
                        [math]::Round(
                            [decimal]$current.BorcToplami -
                            [decimal]$current.AlacakToplami,
                            2
                        )

                    $current.BakiyeYonu =
                        if ($current.BakiyeSigned -gt 0.005) {
                            "B"
                        }
                        elseif ($current.BakiyeSigned -lt -0.005) {
                            "A"
                        }
                        else {
                            ""
                        }

                    $current.Bakiye =
                        [math]::Abs(
                            [decimal]$current.BakiyeSigned
                        )

                    $customers.Add($current)
                }

                $hesapKodu = $aTrim.Split(" ")[0]
                $musteriAdi = $aTrim.Substring($hesapKodu.Length).Trim()

                $current =
                    [PSCustomObject]@{
                        HesapKodu      = $hesapKodu
                        MusteriAdi     = $musteriAdi
                        Transactions   = New-Object System.Collections.Generic.List[object]
                        BorcToplami    = [decimal]0
                        AlacakToplami  = [decimal]0
                        BakiyeSigned   = [decimal]0
                        Bakiye         = [decimal]0
                        BakiyeYonu     = ""
                    }

                continue
            }

            if ($null -eq $current) {
                continue
            }

            # ------------------------------------------------
            # BLOK SONU
            # ------------------------------------------------

            if (
                $aTrim -eq "Nakli Yekün Hariç :" -or
                $aTrim -eq "Genel Toplam :"
            ) {
                continue
            }

            # Başlık satırı
            if (
                $aTrim -eq "TARİH" -and
                $b.Trim() -eq "TİP"
            ) {
                continue
            }

            # ------------------------------------------------
            # HAREKET SATIRI
            # ------------------------------------------------

            $rawDate =
                Get-CellValueLocal -ExcelRow $excelRow -ExcelCol 1

            if (
                $null -eq $rawDate -or
                -not ($rawDate -is [ValueType])
            ) {
                continue
            }

            # Excel COM hücresi tarih olarak System.DateTime,
            # OLE Automation sayısı veya metin döndürebilir.
            # Önce veri tipini kontrol ederek güvenli şekilde tarihe çeviriyoruz.
            $transactionDate = $null

            try {
                if ($rawDate -is [datetime]) {
                    $transactionDate = [datetime]$rawDate
                }
                elseif ($rawDate -is [double] -or
                        $rawDate -is [float] -or
                        $rawDate -is [decimal] -or
                        $rawDate -is [int] -or
                        $rawDate -is [long] -or
                        $rawDate -is [short]) {
                    $transactionDate =
                        [datetime]::FromOADate(
                            [double]$rawDate
                        )
                }
                elseif ($rawDate -is [string]) {
                    $parsedDate = [datetime]::MinValue

                    $ok = [datetime]::TryParse(
                        $rawDate.Trim(),
                        [System.Globalization.CultureInfo]::GetCultureInfo('tr-TR'),
                        [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
                        [ref]$parsedDate
                    )

                    if ($ok) {
                        $transactionDate = $parsedDate
                    }
                }
            }
            catch {
                $transactionDate = $null
            }

            if ($null -eq $transactionDate) {
                continue
            }

            $borc = [decimal]0
            $alacak = [decimal]0

            if ($null -ne $eRaw -and "$eRaw".Trim() -ne "") {
                try {
                    $borc = [decimal]$eRaw
                }
                catch {
                    $borc = [decimal]0
                }
            }

            if ($null -ne $fRaw -and "$fRaw".Trim() -ne "") {
                try {
                    $alacak = [decimal]$fRaw
                }
                catch {
                    $alacak = [decimal]0
                }
            }

            $current.BorcToplami += $borc
            $current.AlacakToplami += $alacak

            $running =
                [decimal]$current.BakiyeSigned +
                $borc -
                $alacak

            $current.BakiyeSigned = $running

            $yon =
                if ($running -gt 0.005) {
                    "B"
                }
                elseif ($running -lt -0.005) {
                    "A"
                }
                else {
                    ""
                }

            $current.Transactions.Add(
                [PSCustomObject]@{
                    Tarih        = $transactionDate
                    Tip          = $b.Trim()
                    FisNo        = $c.Trim()
                    Aciklama     = $d.Trim()
                    Borc         = $borc
                    Alacak       = $alacak
                    Bakiye       = [math]::Abs($running)
                    BakiyeYonu   = $yon
                }
            )
        }

        # Son hesap bloğunu ekle.
        if ($null -ne $current) {

            $current.BakiyeSigned =
                [math]::Round(
                    [decimal]$current.BorcToplami -
                    [decimal]$current.AlacakToplami,
                    2
                )

            $current.BakiyeYonu =
                if ($current.BakiyeSigned -gt 0.005) {
                    "B"
                }
                elseif ($current.BakiyeSigned -lt -0.005) {
                    "A"
                }
                else {
                    ""
                }

            $current.Bakiye =
                [math]::Abs(
                    [decimal]$current.BakiyeSigned
                )

            $customers.Add($current)
        }

        return $customers
    }
    finally {

        if ($usedRange -ne $null) {
            Release-ComObject $usedRange
            $usedRange = $null
        }

        if ($worksheet -ne $null) {
            Release-ComObject $worksheet
            $worksheet = $null
        }

        if ($workbook -ne $null) {
            try {
                $workbook.Close($false)
            }
            catch {}

            Release-ComObject $workbook
            $workbook = $null
        }

        if ($excel -ne $null) {
            try {
                $excel.Quit()
            }
            catch {}

            Release-ComObject $excel
            $excel = $null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ============================================================
# MESAJ OLUŞTUR
# ============================================================

function Test-MessageTemplate {

    if (
        [string]::IsNullOrWhiteSpace(
            $settings.MesajSablonu
        )
    ) {

        [System.Windows.Forms.MessageBox]::Show(
            "Henüz mesaj şablonu oluşturulmamış.`n`n" +
            "Önce [2] Mesaj Şablonunu Düzenle bölümünden " +
            "mesaj şablonunuzu oluşturun.",
            "Mesaj Şablonu Eksik",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return $false
    }

    return $true
}

function Build-Message {

    param(
        [Parameter(Mandatory = $true)]
        [string]$MusteriAdi,

        [Parameter(Mandatory = $true)]
        [decimal]$Tutar
    )

    $mesaj = [string]$settings.MesajSablonu

    $formattedAmount =
        Format-TurkishAmount -Value $Tutar

    $mesaj =
        $mesaj.Replace(
            "{MUSTERI_ADI}",
            $MusteriAdi
        )

    $mesaj =
        $mesaj.Replace(
            "{TUTAR}",
            $formattedAmount
        )

    return $mesaj
}

# ============================================================
# ÖNİZLEME
# ============================================================

function Show-Preview {

    if (-not (Test-MessageTemplate)) {
        return
    }

    if (
        [string]::IsNullOrWhiteSpace(
            $script:muavinPath
        )
    ) {

        [System.Windows.Forms.MessageBox]::Show(
            "Önce [1] Muavin Excel Dosyası Seç bölümünden " +
            "bir muavin dosyası seçmelisiniz.",
            "Muavin Dosyası Eksik",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return
    }

    try {

        $customers = @(
            Get-MuavinCustomerData -ExcelPath $script:muavinPath
        )

        $firstCustomer =
            $customers |
            Where-Object {
                $_.BakiyeSigned -gt 0.005
            } |
            Select-Object -First 1

        if ($null -eq $firstCustomer) {

            [System.Windows.Forms.MessageBox]::Show(
                "Borç bakiyesi bulunan müşteri bulunamadı.",
                "Bilgi",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            return
        }

        $preview =
            Build-Message `
                -MusteriAdi $firstCustomer.MusteriAdi `
                -Tutar $firstCustomer.Bakiye

        $form =
            New-Object System.Windows.Forms.Form

        $form.Text =
            "Mesaj Önizleme - $($firstCustomer.MusteriAdi)"

        $form.Size =
            New-Object System.Drawing.Size(850, 650)

        $form.StartPosition =
            "CenterScreen"

        $label =
            New-Object System.Windows.Forms.Label

        $label.Text =
            "Örnek müşteri: $($firstCustomer.MusteriAdi)  |  " +
            "Borç bakiyesi: $(Format-TurkishAmount $firstCustomer.Bakiye) TL"

        $label.Location =
            New-Object System.Drawing.Point(15, 15)

        $label.AutoSize = $true

        $box =
            New-Object System.Windows.Forms.TextBox

        $box.Multiline = $true
        $box.ReadOnly = $true
        $box.ScrollBars = "Both"
        $box.WordWrap = $true

        $box.Font =
            New-Object System.Drawing.Font(
                "Consolas",
                11
            )

        $box.Location =
            New-Object System.Drawing.Point(15, 45)

        $box.Size =
            New-Object System.Drawing.Size(805, 500)

        $box.Text =
            $preview

        $button =
            New-Object System.Windows.Forms.Button

        $button.Text = "Kapat"
        $button.Size = New-Object System.Drawing.Size(100, 35)
        $button.Location = New-Object System.Drawing.Point(720, 560)
        $button.DialogResult =
            [System.Windows.Forms.DialogResult]::OK

        $form.Controls.Add($label)
        $form.Controls.Add($box)
        $form.Controls.Add($button)

        $form.AcceptButton = $button

        [void]$form.ShowDialog()
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "Önizleme sırasında hata oluştu.`n`n" +
            "$($_.Exception.Message)",
            "Hata",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# ============================================================
# HTML - DOĞRUDAN PDF
# ============================================================

# ============================================================
# PDF DAMGASI
# ============================================================
# PDF'nin sağ üstteki boş alanına kırmızı, hafif eğik
# "CARİ HESAP BİLDİRİMİ / GÖNDERİLDİ" damgası ekler.
# Akışa dahil olmadığı için metin ve tablo düzenini bozmaz.


# ============================================================
# PDF DAMGASI + BELGE DÜZENİ
# ============================================================

function Get-EmbeddedStampPath {

    if ($script:EmbeddedStampPath -and (Test-Path -LiteralPath $script:EmbeddedStampPath)) {
        return $script:EmbeddedStampPath
    }

    $stampPath =
        Join-Path `
            -Path (Get-MFTModuleTempPath -ModuleId 'muavindenmesaj') `
            -ChildPath "MusteriCari_Gonderildi_Stamp.png"

    if (-not (Test-Path -LiteralPath $stampPath)) {

        $stampBase64 = @'
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
'@

        try {
            $stampBytes =
                [System.Convert]::FromBase64String(
                    ($stampBase64 -replace "\s", "")
                )

            [System.IO.File]::WriteAllBytes(
                $stampPath,
                $stampBytes
            )
        }
        catch {
            throw "Gönderildi mührü hazırlanamadı: $($_.Exception.Message)"
        }
    }

    $script:EmbeddedStampPath = $stampPath

    return $stampPath
}

# HTML tabanlı PDF üretim yardımcıları

function ConvertTo-HtmlSafe {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Resolve-PdfBrowserPath {
    $candidates = New-Object System.Collections.Generic.List[string]

    if ($env:ProgramFiles) {
        $candidates.Add((Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"))
        $candidates.Add((Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"))
    }

    $programFilesX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
    if ($programFilesX86) {
        $candidates.Add((Join-Path $programFilesX86 "Microsoft\Edge\Application\msedge.exe"))
        $candidates.Add((Join-Path $programFilesX86 "Google\Chrome\Application\chrome.exe"))
    }

    if ($env:LOCALAPPDATA) {
        $candidates.Add((Join-Path $env:LOCALAPPDATA "Microsoft\Edge\Application\msedge.exe"))
        $candidates.Add((Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe"))
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    foreach ($commandName in @("msedge.exe", "chrome.exe")) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue
        if ($null -ne $command -and $command.Source) {
            return $command.Source
        }
    }

    throw "PDF üretmek için Microsoft Edge veya Google Chrome bulunamadı. Tarayıcı kurulumunu kontrol edin."
}

function Convert-HtmlToPdf {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $finalPath = [System.IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($finalPath)

    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force -ErrorAction Stop | Out-Null
    }

    $browserPath = Resolve-PdfBrowserPath
    $workDirectory = Join-Path (Get-MFTModuleTempPath -ModuleId 'muavindenmesaj') ("CariBildirim_" + [Guid]::NewGuid().ToString("N"))
    $profileDirectory = Join-Path $workDirectory "BrowserProfile"
    $htmlPath = Join-Path $workDirectory "CariBildirim.html"
    $tempPdfPath = Join-Path $outputDirectory (
        "." + [System.IO.Path]::GetFileNameWithoutExtension($finalPath) + "." +
        [Guid]::NewGuid().ToString("N") + ".tmp.pdf"
    )

    try {
        New-Item -ItemType Directory -Path $profileDirectory -Force -ErrorAction Stop | Out-Null
        [System.IO.File]::WriteAllText(
            $htmlPath,
            $Html,
            (New-Object System.Text.UTF8Encoding($false))
        )

        $fileUri = ([System.Uri]$htmlPath).AbsoluteUri
        $arguments = @(
            "--headless=new",
            "--disable-gpu",
            "--disable-extensions",
            "--no-first-run",
            "--no-default-browser-check",
            "--no-pdf-header-footer",
            ("--user-data-dir=`"{0}`"" -f $profileDirectory),
            ("--print-to-pdf=`"{0}`"" -f $tempPdfPath),
            ("`"{0}`"" -f $fileUri)
        ) -join " "

        $process = Start-Process `
            -FilePath $browserPath `
            -ArgumentList $arguments `
            -PassThru `
            -Wait `
            -WindowStyle Hidden

        if (-not (Test-Path -LiteralPath $tempPdfPath -PathType Leaf)) {
            throw "Tarayıcı PDF dosyasını oluşturmadı. Tarayıcı çıkış kodu: $($process.ExitCode)"
        }

        $pdfInfo = Get-Item -LiteralPath $tempPdfPath -ErrorAction Stop
        if ($pdfInfo.Length -le 0) {
            throw "Tarayıcı boş bir PDF dosyası oluşturdu."
        }

        Move-Item -LiteralPath $tempPdfPath -Destination $finalPath -Force -ErrorAction Stop
        return $true
    }
    finally {
        if (Test-Path -LiteralPath $tempPdfPath) {
            Remove-Item -LiteralPath $tempPdfPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $workDirectory) {
            Remove-Item -LiteralPath $workDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-CustomerMessageParts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$MusteriAdi
    )

    $messageLines = $Message.Replace("`r`n", "`r").Split(
        @("`r"),
        [System.StringSplitOptions]::None
    )

    $bodyLines = New-Object System.Collections.Generic.List[string]
    $bankRows = New-Object System.Collections.Generic.List[object]
    $dateValue = ""
    $recipientValue = ""
    $bankHeaderFound = $false
    $inBankSection = $false
    $inBody = $false

    for ($i = 2; $i -lt $messageLines.Count; $i++) {
        $line = [string]$messageLines[$i]
        $trimmed = $line.Trim()

        if ($trimmed -match '^Düzenlenme Tarihi\s*[:\t]+(.*)$') {
            $dateValue = $Matches[1].Trim()
            continue
        }

        if ($trimmed -match '^Muhatap\s*(?:[:\t]+)\s*(.*)$') {
            $recipientValue = $Matches[1].Trim()
            continue
        }

        if ($trimmed -match '^HESAP BİLGİLERİ\s*:?$') {
            $bankHeaderFound = $true
            $inBankSection = $true
            continue
        }

        if ($inBankSection) {
            if ($trimmed -match '^-{3,}$') { continue }

            if ($trimmed -match '^Banka\s*(?:[:\t]+)\s*(.*)$') {
                [void]$bankRows.Add([PSCustomObject]@{ Label = "Banka:"; Value = $Matches[1].Trim() })
                continue
            }
            if ($trimmed -match '^Hesap Sahibi\s*(?:[:\t]+)\s*(.*)$') {
                [void]$bankRows.Add([PSCustomObject]@{ Label = "Hesap Sahibi:"; Value = $Matches[1].Trim() })
                continue
            }
            if ($trimmed -match '^IBAN\s*(?:[:\t]+)\s*(.*)$') {
                [void]$bankRows.Add([PSCustomObject]@{ Label = "IBAN:"; Value = $Matches[1].Trim() })
                continue
            }
            continue
        }

        if ($trimmed -ne "" -or $inBody) {
            [void]$bodyLines.Add($line)
            $inBody = $true
        }
    }

    while ($bodyLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($bodyLines[$bodyLines.Count - 1])) {
        $bodyLines.RemoveAt($bodyLines.Count - 1)
    }

    return [PSCustomObject]@{
        Heading          = if ($messageLines.Count -ge 1) { [string]$messageLines[0] } else { "SMMM MUSTAFA ORAKCI" }
        Subheading       = if ($messageLines.Count -ge 2) { [string]$messageLines[1] } else { "CARİ HESAP BORÇ BİLDİRİMİ" }
        DateValue        = $dateValue
        RecipientValue   = if ([string]::IsNullOrWhiteSpace($recipientValue)) { $MusteriAdi } else { $recipientValue }
        BodyLines        = @($bodyLines.ToArray())
        BankHeaderFound  = $bankHeaderFound
        BankRows         = @($bankRows.ToArray())
    }
}

# PDF imza ve altbilgi metinleri (mesaj şablonundan bağımsız, sabit).
# Neden: Şablonun ilk iki satırı büyük harfli başlık için kullanılır; imza adı/unvanı ve
# altbilgi başlığı ise her müşteri için aynı ve okunaklı yazımla görünmelidir.
$script:PdfSignatureName  = 'Mustafa Orakcı'
$script:PdfSignatureTitle = 'Serbest Muhasebeci Mali Müşavir'
$script:PdfFooterTitle    = 'Cari hesap borç bildirimi'

function ConvertTo-CssString {
    <#
    .SYNOPSIS
        CSS content özelliği için çift tırnaklı, kaçışlı metin üretir.
    .DESCRIPTION
        Müşteri adı sayfa altbilgisine CSS metni olarak yazıldığı için tırnak, ters bölü ve
        "<" karakterleri kaçışlanır. Neden: "</style>" içeren bir ad stil bloğunu bozmamalıdır.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $escaped = $Text -replace '[\r\n]+', ' '
    $escaped = $escaped.Replace('\', '\\').Replace('"', '\"').Replace('<', '\3C ')

    return '"' + $escaped + '"'
}

function ConvertTo-IbanText {
    <#
    .SYNOPSIS
        IBAN'ı dörtlü gruplara bölerek okunaklı hale getirir.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    $compact = ($Value -replace '\s', '').ToUpperInvariant()

    # Neden: IBAN biçimine uymayan bir değer (örn. hesap numarası) bozulmadan olduğu gibi gösterilir.
    if ($compact -notmatch '^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$') {
        return $Value
    }

    return ($compact -replace '(.{4})(?=.)', '$1 ')
}

function ConvertTo-AmountEmphasisHtml {
    <#
    .SYNOPSIS
        HTML-kodlanmış mesaj satırındaki bakiye tutarını (ve varsa ardındaki "TL" ifadesini) kalın yapar.
    .DESCRIPTION
        Tutarın başka bir sayının parçası olarak (örn. 250,00 içinde 1.250,00) eşleşmemesi için
        önünde ve ardında rakam/nokta/virgül kontrolü yapılır.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$AmountText
    )

    if ([string]::IsNullOrWhiteSpace($AmountText)) {
        return $Html
    }

    $pattern = '(?<![\d.,])' + [regex]::Escape($AmountText) + '(?!\d)(?:\s*TL\b)?'

    return [regex]::Replace($Html, $pattern, '<b class="amount">$0</b>')
}

function Get-CustomerPdfStyleBlock {
    <#
    .SYNOPSIS
        Cari hesap bildirimi PDF'inin stil bloğunu döndürür.
    .DESCRIPTION
        Renkler :root değişkenlerinde toplanmıştır; temayı değiştirmek için yalnızca oradaki
        değerler düzenlenir. Sayfa altbilgisi (müşteri adı ve "Sayfa x / y") @page kenar
        kutularıyla üretilir; bu özelliği desteklemeyen eski tarayıcılar altbilgiyi yok sayar.
    .PARAMETER FooterLeftText
        Her sayfanın sol altında gösterilecek metin.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FooterLeftText
    )

    $style = @'
<style>
:root { --accent: #ea580c; --accent-strong: #c2410c; --accent-dark: #7c2d12; --tint: #fff7ed; --tint-line: #fed7aa; --zebra: #fffaf5; --line: #ece5df; --ink: #2b2a28; --muted: #78716c; --faint: #a8a29e; --credit-bg: #edf8f0; --credit-ink: #16803c; }
@page { size: A4 portrait; margin: 11mm 9mm 14mm 9mm;
  @bottom-left { content: __FOOTER_LEFT__; width: 75%; font: 7.5pt "Segoe UI", Arial, Helvetica, sans-serif; color: #a8a29e; vertical-align: top; border-top: 0.5pt solid #ece5df; padding-top: 1.8mm; }
  @bottom-right { content: "Sayfa " counter(page) " / " counter(pages); width: 25%; font: 7.5pt "Segoe UI", Arial, Helvetica, sans-serif; color: #a8a29e; text-align: right; vertical-align: top; border-top: 0.5pt solid #ece5df; padding-top: 1.8mm; } }
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; font-family: "Segoe UI", Arial, Helvetica, sans-serif; color: var(--ink); font-size: 10pt; }
.top { display: flex; justify-content: space-between; align-items: flex-start; gap: 8mm; border-bottom: 2pt solid var(--accent); padding-bottom: 3mm; margin-bottom: 4mm; }
.heading-wrap { flex: 1; min-width: 0; }
.heading { font-size: 9.5pt; font-weight: 700; letter-spacing: 0.06em; color: var(--muted); line-height: 1.25; overflow-wrap: anywhere; }
.subheading { margin-top: 1.2mm; font-size: 15pt; font-weight: 700; color: var(--accent-dark); line-height: 1.25; overflow-wrap: anywhere; }
.stamp { display: block; width: 30mm; height: 30mm; object-fit: contain; transform: rotate(-8deg); flex: 0 0 30mm; }
.meta { width: 100%; border-collapse: collapse; margin: 1mm 0 4mm; font-size: 9.5pt; }
.meta td { padding: 1.2mm 1mm; vertical-align: top; }
.meta td:first-child { padding-left: 0; }
.meta .label { width: 34mm; font-weight: 700; white-space: nowrap; }
.meta .colon { width: 3mm; color: var(--faint); text-align: center; padding: 1.2mm 0; }
.tiles { display: flex; gap: 2.5mm; margin: 0 0 4mm; }
.tile { flex: 1; background: var(--tint); border: 0.5pt solid var(--tint-line); border-radius: 1mm; padding: 2mm 3mm; break-inside: avoid; page-break-inside: avoid; }
.tile .tile-label { font-size: 7.5pt; color: var(--muted); }
.tile .tile-value { font-size: 12pt; font-weight: 700; color: var(--accent-dark); white-space: nowrap; }
.tile.key { background: var(--accent-strong); border-color: var(--accent-strong); }
.tile.key .tile-label { color: #ffedd5; }
.tile.key .tile-value { color: #ffffff; }
.message { margin: 2mm 0 4mm; font-size: 10pt; line-height: 1.38; overflow-wrap: anywhere; }
.message-line { min-height: 1em; }
.message-gap { height: 2.2mm; }
.message b.amount { font-weight: 700; color: var(--accent-strong); }
.bank-card { background: var(--tint); border-left: 3pt solid var(--accent); padding: 2mm 3.5mm; margin: 2mm 0 4mm; break-inside: avoid; page-break-inside: avoid; }
.bank-title { font-size: 8pt; font-weight: 700; letter-spacing: 0.06em; color: var(--accent-dark); padding-bottom: 1mm; }
.bank { width: 100%; border-collapse: collapse; font-size: 9.5pt; }
.bank td { padding: 0.7mm 1mm; vertical-align: top; overflow-wrap: anywhere; }
.bank .label { width: 28mm; white-space: nowrap; }
.bank .colon { width: 3mm; color: var(--faint); text-align: center; padding: 0.7mm 0; }
.bank .iban { font-family: Consolas, "Courier New", monospace; letter-spacing: 0.02em; }
.statement-title { text-align: left; font-weight: 700; font-size: 10.5pt; color: var(--accent-dark); border-left: 3pt solid var(--accent); padding: 1mm 0 1mm 2mm; margin: 4mm 0 2mm; }
table.transactions { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 7.8pt; color: var(--ink); font-variant-numeric: tabular-nums; }
.transactions col.date { width: 17mm; } .transactions col.type { width: 14mm; } .transactions col.voucher { width: 19mm; } .transactions col.side { width: 7mm; } .transactions col.description { width: 63mm; } .transactions col.money { width: 24mm; }
.transactions th, .transactions td { border: 0; border-bottom: 0.4pt solid var(--line); padding: 1.2mm 0.9mm; vertical-align: top; overflow-wrap: anywhere; }
.transactions th { background: var(--accent-strong); color: #ffffff; text-align: center; font-weight: 700; vertical-align: middle; border-bottom: 0; padding: 1.6mm 0.9mm; }
.transactions tbody tr:nth-child(even):not(.total):not(.credit) td { background: var(--zebra); }
.transactions tr.credit td { background: var(--credit-bg); }
.transactions tr.credit td.money.credit-amount { color: var(--credit-ink); font-weight: 700; }
.transactions td.center { text-align: center; }
.transactions td.money { text-align: right; white-space: nowrap; }
.transactions td.balance { font-weight: 700; }
.transactions td.description { white-space: normal; }
.transactions thead { display: table-header-group; }
.transactions tr { break-inside: avoid; page-break-inside: avoid; }
.transactions .total td { background: #ffedd5; color: var(--accent-dark); font-weight: 700; border-top: 1.2pt solid var(--accent-strong); border-bottom: 0; padding: 1.8mm 0.9mm; }
.closing-block { margin-top: 5mm; break-inside: avoid; page-break-inside: avoid; }
.closing { font-size: 9.5pt; line-height: 1.3; }
.signature { margin-top: 9mm; font-size: 9.5pt; line-height: 1.35; }
.signature .sign-line { display: inline-block; min-width: 55mm; border-top: 0.5pt solid var(--faint); padding-top: 1mm; }
.signature .name { font-weight: 700; }
.signature .title { font-size: 8.5pt; color: var(--muted); }
</style>
'@

    return $style.Replace('__FOOTER_LEFT__', (ConvertTo-CssString -Text $FooterLeftText))
}

function Export-CustomerPdf {
    param(
        [Parameter(Mandatory = $true)]
        $Customer,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $message = Build-Message `
        -MusteriAdi $Customer.MusteriAdi `
        -Tutar $Customer.Bakiye

    # PDF, Excel çalışma sayfası üzerinden değil HTML metin ve tablo
    # düzeninden üretilir. Böylece açıklama satırları otomatik kayar;
    # hücre birleştirmesi veya sabit satır yüksekliği gerekmez.
    $parts = Get-CustomerMessageParts -Message $message -MusteriAdi $Customer.MusteriAdi
    $builder = New-Object System.Text.StringBuilder

    $stampMarkup = ""
    $stampPath = Get-EmbeddedStampPath
    if (Test-Path -LiteralPath $stampPath -PathType Leaf) {
        $stampBytes = [System.IO.File]::ReadAllBytes($stampPath)
        $stampBase64 = [System.Convert]::ToBase64String($stampBytes)
        $stampMarkup = "<img class='stamp' src='data:image/png;base64,$stampBase64' alt='Gönderildi' />"
    }

    $amountText = Format-TurkishAmount -Value $Customer.Bakiye
    $footerText = '{0} · {1}' -f $script:PdfFooterTitle, $Customer.MusteriAdi

    [void]$builder.AppendLine('<!DOCTYPE html>')
    [void]$builder.AppendLine('<html lang="tr"><head><meta charset="utf-8">')
    [void]$builder.AppendLine('<title>Cari Hesap Borç Bildirimi</title>')
    [void]$builder.AppendLine((Get-CustomerPdfStyleBlock -FooterLeftText $footerText))
    [void]$builder.AppendLine('</head><body>')

    [void]$builder.AppendLine('<div class="top"><div class="heading-wrap">')
    [void]$builder.AppendLine(("<div class='heading'>{0}</div>" -f (ConvertTo-HtmlSafe $parts.Heading)))
    [void]$builder.AppendLine(("<div class='subheading'>{0}</div>" -f (ConvertTo-HtmlSafe $parts.Subheading)))
    [void]$builder.AppendLine('</div>')
    if ($stampMarkup) { [void]$builder.AppendLine($stampMarkup) }
    [void]$builder.AppendLine('</div>')

    # İki nokta ayrı bir sütunda tutulur; böylece tüm değerler aynı hizada başlar.
    [void]$builder.AppendLine('<table class="meta"><tbody>')
    [void]$builder.AppendLine(("<tr><td class='label'>Düzenlenme Tarihi</td><td class='colon'>:</td><td>{0}</td></tr>" -f (ConvertTo-HtmlSafe $parts.DateValue)))
    [void]$builder.AppendLine(("<tr><td class='label'>Muhatap</td><td class='colon'>:</td><td>{0}</td></tr>" -f (ConvertTo-HtmlSafe $parts.RecipientValue)))
    [void]$builder.AppendLine('</tbody></table>')

    $summaryDebit = ConvertTo-HtmlSafe (Format-TurkishAmount -Value $Customer.BorcToplami)
    $summaryCredit = ConvertTo-HtmlSafe (Format-TurkishAmount -Value $Customer.AlacakToplami)
    $summaryBalance = ConvertTo-HtmlSafe $amountText
    [void]$builder.AppendLine('<div class="tiles">')
    [void]$builder.AppendLine(("<div class='tile'><div class='tile-label'>Toplam Borç</div><div class='tile-value'>{0}</div></div>" -f $summaryDebit))
    [void]$builder.AppendLine(("<div class='tile'><div class='tile-label'>Toplam Alacak</div><div class='tile-value'>{0}</div></div>" -f $summaryCredit))
    [void]$builder.AppendLine(("<div class='tile key'><div class='tile-label'>Ödenecek Bakiye</div><div class='tile-value'>{0} TL</div></div>" -f $summaryBalance))
    [void]$builder.AppendLine('</div>')

    [void]$builder.AppendLine('<div class="message">')
    foreach ($line in $parts.BodyLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            [void]$builder.AppendLine('<div class="message-gap"></div>')
        }
        else {
            $lineHtml = ConvertTo-AmountEmphasisHtml -Html (ConvertTo-HtmlSafe $line.Trim()) -AmountText $amountText
            [void]$builder.AppendLine(("<div class='message-line'>{0}</div>" -f $lineHtml))
        }
    }
    [void]$builder.AppendLine('</div>')

    if ($parts.BankHeaderFound) {
        [void]$builder.AppendLine('<div class="bank-card"><div class="bank-title">HESAP BİLGİLERİ</div>')
        [void]$builder.AppendLine('<table class="bank"><tbody>')
        foreach ($bankRow in $parts.BankRows) {
            $bankLabelText = ([string]$bankRow.Label).TrimEnd(':').Trim()
            $bankValueText = [string]$bankRow.Value
            $valueClass = ''
            if ($bankLabelText -eq 'IBAN') {
                $bankValueText = ConvertTo-IbanText -Value $bankValueText
                $valueClass = " class='iban'"
            }
            $bankLabel = ConvertTo-HtmlSafe $bankLabelText
            $bankValue = ConvertTo-HtmlSafe $bankValueText
            [void]$builder.AppendLine("<tr><td class='label'>$bankLabel</td><td class='colon'>:</td><td$valueClass>$bankValue</td></tr>")
        }
        [void]$builder.AppendLine('</tbody></table></div>')
    }

    [void]$builder.AppendLine('<div class="statement-title">CARİ HESAP DÖKÜMÜ</div>')
    [void]$builder.AppendLine('<table class="transactions"><colgroup><col class="date"><col class="type"><col class="voucher"><col class="side"><col class="description"><col class="money"><col class="money"><col class="money"></colgroup>')
    [void]$builder.AppendLine('<thead><tr><th>TARİH</th><th>TİP</th><th>FİŞ NO</th><th>B/A</th><th>AÇIKLAMA</th><th>BORÇ</th><th>ALACAK</th><th>BAKİYE</th></tr></thead><tbody>')

    foreach ($tx in $Customer.Transactions) {
        $dateText = if ($null -ne $tx.Tarih) { $tx.Tarih.ToString('dd.MM.yyyy') } else { '' }
        $borcText = if ([decimal]$tx.Borc -ne 0) { Format-TurkishAmount -Value $tx.Borc } else { '' }
        $alacakText = if ([decimal]$tx.Alacak -ne 0) { Format-TurkishAmount -Value $tx.Alacak } else { '' }
        $bakiyeText = Format-TurkishAmount -Value $tx.Bakiye

        $rowClass = if ([decimal]$tx.Alacak -ne 0) { " class='credit'" } else { '' }
        [void]$builder.AppendLine("<tr$rowClass>")
        [void]$builder.AppendLine(("<td class='center'>{0}</td>" -f (ConvertTo-HtmlSafe $dateText)))
        [void]$builder.AppendLine(("<td class='center'>{0}</td>" -f (ConvertTo-HtmlSafe $tx.Tip)))
        [void]$builder.AppendLine(("<td class='center'>{0}</td>" -f (ConvertTo-HtmlSafe $tx.FisNo)))
        [void]$builder.AppendLine(("<td class='center'>{0}</td>" -f (ConvertTo-HtmlSafe $tx.BakiyeYonu)))
        [void]$builder.AppendLine(("<td class='description'>{0}</td>" -f (ConvertTo-HtmlSafe $tx.Aciklama)))
        [void]$builder.AppendLine(("<td class='money'>{0}</td>" -f (ConvertTo-HtmlSafe $borcText)))
        $creditClass = if ([decimal]$tx.Alacak -ne 0) { 'money credit-amount' } else { 'money' }
        [void]$builder.AppendLine(("<td class='{0}'>{1}</td>" -f $creditClass, (ConvertTo-HtmlSafe $alacakText)))
        [void]$builder.AppendLine(("<td class='money balance'>{0}</td>" -f (ConvertTo-HtmlSafe $bakiyeText)))
        [void]$builder.AppendLine('</tr>')
    }

    $totalDebit = ConvertTo-HtmlSafe (Format-TurkishAmount -Value $Customer.BorcToplami)
    $totalCredit = ConvertTo-HtmlSafe (Format-TurkishAmount -Value $Customer.AlacakToplami)
    $totalBalance = ConvertTo-HtmlSafe (Format-TurkishAmount -Value $Customer.Bakiye)
    $totalRow = "<tr class='total'><td colspan='5'>TOPLAM</td><td class='money'>$totalDebit</td><td class='money'>$totalCredit</td><td class='money'>$totalBalance</td></tr>"
    [void]$builder.AppendLine($totalRow)
    [void]$builder.AppendLine('</tbody></table>')

    [void]$builder.AppendLine('<div class="closing-block"><div class="closing">Bilgilerinize sunar, iyi çalışmalar dileriz.</div>')
    [void]$builder.AppendLine(("<div class='signature'><div class='sign-line'><div class='name'>{0}</div><div class='title'>{1}</div></div></div></div>" -f (ConvertTo-HtmlSafe $script:PdfSignatureName), (ConvertTo-HtmlSafe $script:PdfSignatureTitle)))
    [void]$builder.AppendLine('</body></html>')

    $html = $builder.ToString()
    return (Convert-HtmlToPdf -Html $html -OutputPath $OutputPath)
}

# ============================================================
# PDF'LERİ OLUŞTUR
# ============================================================

# ============================================================
# PDF'LERİ OLUŞTUR
# ============================================================

function Create-CustomerPdfs {

    if (-not (Test-MessageTemplate)) {
        return
    }

    if (
        [string]::IsNullOrWhiteSpace(
            $script:muavinPath
        )
    ) {

        [System.Windows.Forms.MessageBox]::Show(
            "Önce [1] Muavin Excel Dosyası Seç bölümünden " +
            "bir muavin dosyası seçmelisiniz.",
            "Muavin Dosyası Eksik",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return
    }

    try {

        $customers = @(
            Get-MuavinCustomerData `
                -ExcelPath $script:muavinPath
        )

        if ($customers.Count -eq 0) {

            [System.Windows.Forms.MessageBox]::Show(
                "Muavin dosyasında 120.xx cari hesabı bulunamadı.",
                "Bilgi",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            return
        }

        # Sadece borçlu cari hesaplar PDF'e dönüştürülür.
        # Böylece 0 veya alacak bakiyesi olan kişilere
        # "borç bildirimi" gönderilmez.
        $debtors =
            @(
                $customers |
                Where-Object {
                    $_.BakiyeSigned -gt 0.005
                }
            )

        $zeroCount =
            @(
                $customers |
                Where-Object {
                    [math]::Abs($_.BakiyeSigned) -le 0.005
                }
            ).Count

        $creditCount =
            @(
                $customers |
                Where-Object {
                    $_.BakiyeSigned -lt -0.005
                }
            ).Count

        if ($debtors.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Borç bakiyesi bulunan müşteri yok. PDF oluşturulmadı.",
                "Bilgi",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $confirmMessage = @"
Muavin dosyasında:

Toplam cari hesap : $($customers.Count)
Borçlu            : $($debtors.Count)
Sıfır bakiye      : $zeroCount
Alacaklı          : $creditCount

Borçlu müşterilerin her biri için ayrı PDF oluşturulacak.

PDF içeriği:
1. Mesaj şablonu
2. Müşteriye özel borç bakiyesi
3. Cari hesap hareket dökümü
4. Genel toplam

Mevcut aynı isimli PDF'ler varsa üzerine yazılacaktır.

Devam etmek istiyor musunuz?
"@

        $confirm =
            [System.Windows.Forms.MessageBox]::Show(
                $confirmMessage,
                "Müşteri Bazlı PDF Oluştur",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

        if (
            $confirm -ne
            [System.Windows.Forms.DialogResult]::Yes
        ) {
            return
        }

        $sourceDir =
            Split-Path -Path $script:muavinPath -Parent

        $outputDir =
            Join-Path `
                -Path $sourceDir `
                -ChildPath "Cari_Hesap_PDF"

        if (-not (Test-Path -LiteralPath $outputDir)) {

            New-Item `
                -ItemType Directory `
                -Path $outputDir `
                -Force `
                -ErrorAction Stop | Out-Null
        }

        $successCount = 0
        $errorCount = 0

        Write-Host ""
        Write-Host "PDF'ler oluşturuluyor..." -ForegroundColor Cyan
        Write-Host ""

        foreach ($customer in $debtors) {

            $safeName =
                Get-SafeFileName `
                    -Name "$($customer.HesapKodu)_$($customer.MusteriAdi)"

            $pdfPath =
                Join-Path `
                    -Path $outputDir `
                    -ChildPath "$safeName.pdf"

            Write-Host `
                "[$($successCount + $errorCount + 1)/$($debtors.Count)] " `
                -NoNewline

            Write-Host `
                $customer.MusteriAdi `
                -ForegroundColor White

            try {

                $ok =
                    Export-CustomerPdf `
                        -Customer $customer `
                        -OutputPath $pdfPath

                if ($ok) {

                    $successCount++

                    Write-Host `
                        "  [OK] $(Format-TurkishAmount $customer.Bakiye) TL B" `
                        -ForegroundColor Green
                }
            }
            catch {

                $errorCount++

                $errorDetail =
                    "$($_.Exception.Message) | Satır: $($_.InvocationInfo.ScriptLineNumber)"

                Write-Host `
                    "  [HATA] $errorDetail" `
                    -ForegroundColor Red
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "PDF oluşturma tamamlandı.`n`n" +
            "Toplam cari hesap: $($customers.Count)`n" +
            "Borçlu müşteri: $($debtors.Count)`n" +
            "Başarılı PDF: $successCount`n" +
            "Hatalı PDF: $errorCount`n`n" +
            "PDF klasörü:`n$outputDir",
            "İşlem Tamamlandı",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "PDF oluşturma sırasında hata oluştu.`n`n" +
            "$($_.Exception.Message)",
            "Hata",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# ============================================================
# TXT MESAJLARI OLUŞTUR
# ============================================================

function Create-MessagesTxt {

    if (-not (Test-MessageTemplate)) {
        return
    }

    if (
        [string]::IsNullOrWhiteSpace(
            $script:muavinPath
        )
    ) {

        [System.Windows.Forms.MessageBox]::Show(
            "Önce [1] Muavin Excel Dosyası Seç bölümünden " +
            "bir muavin dosyası seçmelisiniz.",
            "Muavin Dosyası Eksik",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        return
    }

    try {

        $customers = @(
            Get-MuavinCustomerData `
                -ExcelPath $script:muavinPath
        )

        $debtors =
            @(
                $customers |
                Where-Object {
                    $_.BakiyeSigned -gt 0.005
                }
            )

        if ($debtors.Count -eq 0) {

            [System.Windows.Forms.MessageBox]::Show(
                "Borç bakiyesi bulunan müşteri yok.",
                "Bilgi",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            return
        }

        $messages =
            New-Object System.Collections.Generic.List[string]

        foreach ($customer in $debtors) {

            $message =
                Build-Message `
                    -MusteriAdi $customer.MusteriAdi `
                    -Tutar $customer.Bakiye

            [void]$messages.Add($message)
        }

        $sourceDir =
            Split-Path -Path $script:muavinPath -Parent

        $outputPath =
            Join-Path `
                -Path $sourceDir `
                -ChildPath "Musteri_Mesajlari.txt"

        $outputText =
            $messages -join "`r`n`r`n"

        [System.IO.File]::WriteAllText(
            $outputPath,
            $outputText,
            (New-Object System.Text.UTF8Encoding($false))
        )

        [System.Windows.Forms.MessageBox]::Show(
            "Mesaj TXT dosyası oluşturuldu.`n`n" +
            "Borçlu müşteri sayısı: $($debtors.Count)`n`n" +
            "Dosya:`n$outputPath",
            "İşlem Tamamlandı",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            "TXT oluşturma sırasında hata oluştu.`n`n" +
            "$($_.Exception.Message)",
            "Hata",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# ============================================================
# MENÜ
# ============================================================

function Show-Menu {

    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "✉️ MÜŞTERİ MESAJ + CARİ PDF" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""

    Write-Host "VERİ KAYNAĞI:" -ForegroundColor Yellow
    Write-Host "Her çalıştırmada Excel dosyası seçilir."
    Write-Host ""

    if (
        [string]::IsNullOrWhiteSpace(
            $script:muavinPath
        )
    ) {

        Write-Host `
            "Muavin Excel   : SEÇİLMEDİ" `
            -ForegroundColor Yellow
    }
    else {

        $fileName =
            Split-Path `
                $script:muavinPath `
                -Leaf

        Write-Host `
            "Muavin Excel   : $fileName" `
            -ForegroundColor Green

        Write-Host `
            $script:muavinPath `
            -ForegroundColor DarkGray
    }

    Write-Host ""

    if (
        [string]::IsNullOrWhiteSpace(
            $settings.MesajSablonu
        )
    ) {

        Write-Host `
            "Mesaj Şablonu  : OLUŞTURULMADI" `
            -ForegroundColor Yellow
    }
    else {

        Write-Host `
            "Mesaj Şablonu  : HAZIR" `
            -ForegroundColor Green
    }

    Write-Host ""

    if (Test-Path -LiteralPath $settingsPath) {

        Write-Host `
            "Ayar Dosyası   : MEVCUT" `
            -ForegroundColor Green
    }
    else {

        Write-Host `
            "Ayar Dosyası   : HENÜZ OLUŞTURULMADI" `
            -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host "  1. 📂  Muavin Excel Dosyası Seç"
    Write-Host "  2. ✏️  Mesaj Şablonunu Düzenle"
    Write-Host "  3. 👁️  Önizleme"
    Write-Host "  4. 📄  Müşteri Bazlı PDF'leri Oluştur"
    Write-Host "  5. 💬  TXT Mesajları Oluştur"
    Write-Host "  Q  🚪  Çıkış"
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host ""
}

function Pause-Menu {

    Write-Host ""

    Write-Host `
        "Devam etmek için bir tuşa basın..." `
        -ForegroundColor DarkGray

    [void][System.Console]::ReadKey($true)
}

# ============================================================
# ANA PROGRAM
# ============================================================

$script:muavinPath = $null

$programdanCik = $false

do {

    Show-Menu

    $choice =
        Read-Host "Seçiminiz"

    switch ($choice) {

        "1" {

            $selected =
                Select-MuavinFile

            if ($selected) {

                $script:muavinPath =
                    $selected

                Write-Host ""

                Write-Host `
                    "Muavin Excel dosyası seçildi:" `
                    -ForegroundColor Green

                Write-Host `
                    $script:muavinPath `
                    -ForegroundColor White

                Pause-Menu
            }
        }

        "2" {
            Edit-MessageTemplate
        }

        "3" {
            Show-Preview
        }

        "4" {
            Create-CustomerPdfs
        }

        "5" {
            Create-MessagesTxt
        }

        "Q" {
            Clear-Host

            Write-Host ""
            Write-Host `
                "Program kapatılıyor..." `
                -ForegroundColor Yellow

            Write-Host ""

            $programdanCik = $true
        }

        "q" {
            Clear-Host

            Write-Host ""
            Write-Host `
                "Program kapatılıyor..." `
                -ForegroundColor Yellow

            Write-Host ""

            $programdanCik = $true
        }

        default {

            Write-Host ""

            Write-Host `
                "Geçersiz seçim. Lütfen 1-5 veya Q seçin." `
                -ForegroundColor Red

            Pause-Menu
        }
    }

} while (-not $programdanCik)


