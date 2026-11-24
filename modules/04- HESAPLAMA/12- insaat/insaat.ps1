# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

# ============================================================
# KAT KARŞILIĞI İNŞAAT PROJE YÖNETİMİ VE RAPORLAMA ARACI
# REVİZYON 2
#
# PowerShell 5 uyumludur.
#
# REVİZYON 2
# ------------------------------------------------------------
# 1. Hisseli arsa / çoklu fatura alıcısı desteği
# 2. Her fatura alıcısı için ayrı KDV mükellefiyeti
# 3. %5 / %10 otomatik maliyet bedeli hesabı
# 4. Bağımsız bölüm sayısı Integer
# 5. Türkçe sayı formatı desteği
# 6. Güvenli proje dosya adı
# 7. Aynı proje / çakışan dosya adı kontrolü
# 8. Müteahhit bölümünde matrah 0 ise K/Z hesaplanmaz
# 9. Excel'e gerçek sayısal değerler gönderilir
# 10. Excel KDV sütunu gerçek yüzde formatında tutulur
# 11. Eski JSON yapıları mümkün olduğunca Rev.1'e uyarlanır
# 12. Eski projede bilinmeyen KDV durumu otomatik %10 kabul edilmez
# 13. Bölüm/fatura payı toplamı %100 kontrol edilir
# 14. KDV oranı 0-100 arasında kontrol edilir
# 15. Arsa sahibi bölümünde hisseli/hissesiz yapı açıkça girilir
# 16. Ana rapor bağımsız bölüm bazında tek satırdır
# 17. Hisse/fatura alt kırılımı ayrı tabloda gösterilir
# 18. Excel ana rapor + hisse detayı ayrı sayfalara yazılır ve kaydedilir
# ============================================================

# ------------------------------------------------------------
# GENEL AYARLAR
# ------------------------------------------------------------

$GLOBAL:ProjectDataFolder = Get-MFTModuleDataPath -ModuleId 'insaat' -Category Projects
# Neden: Proje verisi (JSON) ile üretilen Excel raporları ayrı klasörlerde tutulur; Projects yalnızca proje dosyalarını içerir.
$GLOBAL:ProjectReportFolder = Get-MFTModuleDataPath -ModuleId 'insaat' -Category Reports
$GLOBAL:ProgramVersion    = "inşaat.ps1 Revizyon 2"

# ------------------------------------------------------------
# KLASÖR
# ------------------------------------------------------------

function Ensure-DataFolder {
    if (-not (Test-Path -LiteralPath $GLOBAL:ProjectDataFolder)) {
        New-Item -ItemType Directory -Path $GLOBAL:ProjectDataFolder -Force | Out-Null
    }
}

# ------------------------------------------------------------
# GENEL YARDIMCI FONKSİYONLAR
# ------------------------------------------------------------

function Pause-Action {
    Read-Host "`nDevam etmek için ENTER tuşuna basın" | Out-Null
}

function ConvertTo-SafeFileName {
    param(
        [string]$Name
    )

    $safeName = $Name.Trim()
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace([string]$char, "_")
    }

    $safeName = $safeName.TrimEnd([char[]](' ', '.'))

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Proje_" + (Get-Date).ToString("yyyyMMdd_HHmmss")
    }

    # Windows ayrılmış dosya adları
    $reservedNames = @(
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    )

    if ($reservedNames -contains $safeName.ToUpper()) {
        $safeName = "Proje_" + $safeName
    }

    return $safeName
}

function ConvertTo-DecimalValue {
    param(
        [string]$Text
    )

    $value = $Text.Trim()

    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Boş sayı."
    }

    $value = $value.Replace("₺", "")
    $value = $value.Replace("%", "")
    $value = $value.Replace(" ", "")

    # Türkçe binlik + ondalık: 1.500.000,50
    if ($value -match '^\d{1,3}(\.\d{3})+(,\d+)?$') {
        $value = $value.Replace(".", "")
        $value = $value.Replace(",", ".")

        return [decimal]::Parse(
            $value,
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    # Türkçe sayı biçimi: binlik nokta, ondalık virgül.
    if ($value -match '^\d+(,\d+)?$') {
        $value = $value.Replace(",", ".")

        return [decimal]::Parse(
            $value,
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    # Binlik ayıracı olmayan noktalı ondalığı da kabul et (1500.50).
    # 1.500 önceki binlik gruplama kuralına uyar ve 1500 olarak okunur.
    if ($value -match '^\d+\.\d+$') {
        return [decimal]::Parse(
            $value,
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    throw "Geçersiz sayı formatı."
}

function Read-DecimalInput {
    param(
        [string]$Prompt,
        [decimal]$Minimum = 0,
        [Nullable[decimal]]$DefaultValue = $null,
        [Nullable[decimal]]$Maximum = $null
    )

    while ($true) {
        if ($null -ne $DefaultValue) {
            $displayPrompt = "{0,-38}: [{1:N2}]" -f $Prompt, $DefaultValue
        }
        else {
            $displayPrompt = "{0,-38}: " -f $Prompt
        }

        $inputValue = Read-Host $displayPrompt

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue) -and $null -ne $DefaultValue) {
            return [decimal]$DefaultValue
        }

        try {
            $number = ConvertTo-DecimalValue $inputValue

            if ($number -lt $Minimum) {
                Write-Host "❌ Değer $Minimum değerinden küçük olamaz." -ForegroundColor Red
                continue
            }

            if ($null -ne $Maximum -and $number -gt $Maximum) {
                Write-Host "❌ Değer $Maximum değerinden büyük olamaz." -ForegroundColor Red
                continue
            }

            return $number
        }
        catch {
            Write-Host "❌ Geçersiz sayı. Örnek: 1.500.000,50" -ForegroundColor Yellow
        }
    }
}

function Read-IntegerInput {
    param(
        [string]$Prompt,
        [int]$Minimum = 1,
        [Nullable[int]]$DefaultValue = $null
    )

    while ($true) {
        if ($null -ne $DefaultValue) {
            $displayPrompt = "{0,-38}: [{1}]" -f $Prompt, $DefaultValue
        }
        else {
            $displayPrompt = "{0,-38}: " -f $Prompt
        }

        $inputValue = Read-Host $displayPrompt

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue) -and $null -ne $DefaultValue) {
            return [int]$DefaultValue
        }

        $number = 0
        if ([int]::TryParse(
            $inputValue.Trim(),
            [System.Globalization.NumberStyles]::None,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$number
        )) {

            if ($number -ge $Minimum) {
                return $number
            }

            Write-Host "❌ Değer en az $Minimum olmalıdır." -ForegroundColor Red
            continue
        }

        Write-Host "❌ Geçersiz tam sayı." -ForegroundColor Yellow
    }
}

function Read-StringInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue = $null
    )

    while ($true) {
        if ($null -ne $DefaultValue) {
            $displayPrompt = "{0,-38}: [{1}]" -f $Prompt, $DefaultValue
        }
        else {
            $displayPrompt = "{0,-38}: " -f $Prompt
        }

        $inputValue = Read-Host $displayPrompt

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue) -and $null -ne $DefaultValue) {
            return $DefaultValue
        }

        if (-not [string]::IsNullOrWhiteSpace($inputValue)) {
            return $inputValue.Trim()
        }

        Write-Host "❌ Bu alan boş bırakılamaz." -ForegroundColor Yellow
    }
}

function Read-DateInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue = $null
    )

    while ($true) {
        if ($null -ne $DefaultValue) {
            $displayPrompt = "{0,-38}: [{1}]" -f $Prompt, $DefaultValue
        }
        else {
            $displayPrompt = "{0,-38}: " -f $Prompt
        }

        $inputValue = Read-Host $displayPrompt

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue) -and $null -ne $DefaultValue) {
            return $DefaultValue
        }

        $dt = [datetime]::MinValue
        $ref = [ref]$dt

        $success = [datetime]::TryParseExact(
            $inputValue.Trim(),
            "dd.MM.yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            $ref
        )

        if ($success) {
            return $ref.Value.ToString("dd.MM.yyyy")
        }

        Write-Host "❌ Geçersiz tarih. Format: gg.aa.yyyy" -ForegroundColor Yellow
    }
}

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$DefaultValue = $false
    )

    while ($true) {
        if ($DefaultValue) {
            $defaultText = "E"
        }
        else {
            $defaultText = "H"
        }

        $inputValue = Read-Host "$Prompt [E/H - ENTER: $defaultText]"

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            return $DefaultValue
        }

        switch ($inputValue.Trim().ToUpperInvariant()) {
            "E" { return $true }
            "H" { return $false }
            default {
                Write-Host "❌ Lütfen E veya H girin." -ForegroundColor Yellow
            }
        }
    }
}

function Read-OwnerType {
    param(
        [string]$CurrentValue = $null
    )

    while ($true) {
        Write-Host ""
        Write-Host "1. Arsa Sahibi"
        Write-Host "2. Müteahhit"

        if ($null -ne $CurrentValue) {
            $inputValue = Read-Host "Yeni seçim [ENTER = $CurrentValue]"

            if ([string]::IsNullOrWhiteSpace($inputValue)) {
                return $CurrentValue
            }
        }
        else {
            $inputValue = Read-Host "Seçiminiz"
        }

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        switch ($inputValue) {
            "1" { return "Arsa Sahibi" }
            "2" { return "Müteahhit" }
            default {
                Write-Host "❌ Geçersiz seçim." -ForegroundColor Red
            }
        }
    }
}

function Read-HakPayiPercent {
    param(
        [string]$Prompt
    )

    while ($true) {
        $inputValue = Read-Host $Prompt

        if ($inputValue.ToUpper() -eq "Q") {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            Write-Host "❌ Hak payı boş bırakılamaz." -ForegroundColor Yellow
            continue
        }

        # 1/2, 1/4 gibi girişler
        if ($inputValue -match '^\s*(\d+)\s*/\s*(\d+)\s*$') {
            try {
                $pay = [decimal]::Parse($Matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
                $payda = [decimal]::Parse($Matches[2], [System.Globalization.CultureInfo]::InvariantCulture)

                if ($payda -gt 0 -and $pay -gt 0 -and $pay -le $payda) {
                    return [math]::Round(($pay / $payda) * 100, 6)
                }
            }
            catch {
                # Geçersiz veya decimal aralığını aşan kesir aşağıda reddedilir.
            }
        }

        try {
            $percent = ConvertTo-DecimalValue $inputValue

            if ($percent -gt 0 -and $percent -le 100) {
                return $percent
            }
        }
        catch {
        }

        Write-Host "❌ Geçersiz hak payı. Örnek: 50 veya 1/2" -ForegroundColor Yellow
    }
}

# ------------------------------------------------------------
# ESKİ PROJEYİ REV.1'E UYARLAMA
# ------------------------------------------------------------

function Convert-ProjectToCurrent {
    param(
        $Project
    )

    if ($null -eq $Project) {
        return $null
    }

    if (-not ($Project.PSObject.Properties.Name -contains "Bolumler")) {
        $Project | Add-Member -MemberType NoteProperty -Name "Bolumler" -Value @()
    }

    foreach ($b in @($Project.Bolumler)) {

        # Eski Sahip -> Taraf
        if (-not ($b.PSObject.Properties.Name -contains "Taraf")) {
            $owner = "Müteahhit"

            if ($b.PSObject.Properties.Name -contains "Sahip") {
                if ([string]$b.Sahip -match "arsa") {
                    $owner = "Arsa Sahibi"
                }
            }

            $b | Add-Member -MemberType NoteProperty -Name "Taraf" -Value $owner
        }

        # KDV oranı yoksa bilinmiyor/0
        if (-not ($b.PSObject.Properties.Name -contains "Kdv")) {
            $b | Add-Member -MemberType NoteProperty -Name "Kdv" -Value ([decimal]0)
        }

        if (-not ($b.PSObject.Properties.Name -contains "SatisMatrahi")) {
            $oldMatrah = [decimal]0

            if ($b.PSObject.Properties.Name -contains "Matrah") {
                try {
                    $oldMatrah = [decimal]$b.Matrah
                }
                catch {
                    $oldMatrah = [decimal]0
                }
            }

            $b | Add-Member -MemberType NoteProperty -Name "SatisMatrahi" -Value $oldMatrah
        }

        if (-not ($b.PSObject.Properties.Name -contains "FaturaAlcilari")) {
            $b | Add-Member -MemberType NoteProperty -Name "FaturaAlcilari" -Value @()
        }

        if ($b.Taraf -eq "Arsa Sahibi") {

            if (@($b.FaturaAlcilari).Count -eq 0) {
                $oldMatrah = [decimal]0

                if ($b.PSObject.Properties.Name -contains "Matrah") {
                    try {
                        $oldMatrah = [decimal]$b.Matrah
                    }
                    catch {
                        $oldMatrah = [decimal]0
                    }
                }

                # Eski projede KDV mükellefiyeti bilinmiyor.
                $legacyRecipient = [PSCustomObject]@{
                    AdUnvan      = "Arsa Sahibi"
                    KdvMukellefi = $null
                    HakPayi      = [decimal]100
                    Matrah       = $oldMatrah
                }

                $b.FaturaAlcilari = @($legacyRecipient)
            }
            else {
                $convertedRecipients = @()

                foreach ($recipient in @($b.FaturaAlcilari)) {
                    if (-not ($recipient.PSObject.Properties.Name -contains "KdvMukellefi")) {
                        $recipient | Add-Member -MemberType NoteProperty -Name "KdvMukellefi" -Value $null
                    }

                    if (-not ($recipient.PSObject.Properties.Name -contains "HakPayi")) {
                        $recipient | Add-Member -MemberType NoteProperty -Name "HakPayi" -Value ([decimal]100)
                    }

                    if (-not ($recipient.PSObject.Properties.Name -contains "Matrah")) {
                        $recipient | Add-Member -MemberType NoteProperty -Name "Matrah" -Value ([decimal]0)
                    }

                    $convertedRecipients += $recipient
                }

                $b.FaturaAlcilari = @($convertedRecipients)
            }

            # Rev.2: Hisseli alanı yoksa alıcı sayısından türet.
            if (-not ($b.PSObject.Properties.Name -contains "Hisseli")) {
                $b | Add-Member -MemberType NoteProperty -Name "Hisseli" -Value (@($b.FaturaAlcilari).Count -gt 1)
            }
            elseif ((@($b.FaturaAlcilari).Count -gt 1) -and (-not [bool]$b.Hisseli)) {
                # Birden fazla alıcı mevcutsa yapı fiilen hisselidir.
                $b.Hisseli = $true
            }

            $b.SatisMatrahi = [decimal]0

            if ($b.PSObject.Properties.Name -contains "Matrah") {
                $b.Matrah = [decimal]0
            }
        }
        else {
            if (-not ($b.PSObject.Properties.Name -contains "Hisseli")) {
                $b | Add-Member -MemberType NoteProperty -Name "Hisseli" -Value $false
            }
            else {
                $b.Hisseli = $false
            }

            $b.FaturaAlcilari = @()
        }
    }

    return $Project
}

# ------------------------------------------------------------
# PROJE KAYDET / YÜKLE / LİSTELE
# ------------------------------------------------------------

function Save-ProjectJson {
    param(
        $Data,
        [string]$ProjectPath = $null
    )

    Ensure-DataFolder

    $Data.ProgramSurumu = $GLOBAL:ProgramVersion

    if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
        $safeName = ConvertTo-SafeFileName $Data.ProjeAdi
        $ProjectPath = Join-Path $GLOBAL:ProjectDataFolder "$safeName.json"
    }

    $json = $Data | ConvertTo-Json -Depth 30

    try {
        [System.IO.File]::WriteAllText(
            $ProjectPath,
            $json,
            [System.Text.Encoding]::UTF8
        )
    }
    catch {
        throw "Proje dosyası kaydedilemedi ('$ProjectPath'): $($_.Exception.Message)"
    }
}

function Load-Project {
    param(
        $ProjectItem
    )

    $path = $null

    if ($ProjectItem -is [string]) {
        $safeName = ConvertTo-SafeFileName $ProjectItem
        $path = Join-Path $GLOBAL:ProjectDataFolder "$safeName.json"
    }
    else {
        $path = $ProjectItem.Path
    }

    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host "❌ Proje dosyası bulunamadı: $path" -ForegroundColor Red
        return $null
    }

    try {
        $jsonContent = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $project = $jsonContent | ConvertFrom-Json -ErrorAction Stop
        return (Convert-ProjectToCurrent $project)
    }
    catch {
        Write-Host "❌ HATA: JSON dosyası okunamadı." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return $null
    }
}

function Load-ProjectList {
    Ensure-DataFolder

    $files = @(
        Get-ChildItem `
            -LiteralPath $GLOBAL:ProjectDataFolder `
            -Filter "*.json" `
            -File `
            -ErrorAction SilentlyContinue
    )

    $result = @()

    foreach ($file in $files) {
        try {
            $json = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
            $project = $json | ConvertFrom-Json -ErrorAction Stop

            if (-not [string]::IsNullOrWhiteSpace([string]$project.ProjeAdi)) {
                $result += [PSCustomObject]@{
                    ProjeAdi        = [string]$project.ProjeAdi
                    OlusturmaTarihi = [string]$project.OlusturmaTarihi
                    Path            = $file.FullName
                    FileName        = $file.Name
                }
            }
        }
        catch {
            Write-Host "⚠️ Bozuk veya okunamayan proje dosyası: $($file.Name)" -ForegroundColor Yellow
        }
    }

    return @($result | Sort-Object ProjeAdi)
}

function Test-ProjectNameExists {
    param(
        [string]$ProjectName
    )

    $projects = @(Load-ProjectList)
    $safeName = ConvertTo-SafeFileName $ProjectName
    $safePath = Join-Path $GLOBAL:ProjectDataFolder "$safeName.json"

    foreach ($project in $projects) {
        if ([string]::Equals(
            $project.ProjeAdi,
            $ProjectName,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            return $true
        }

        if ([string]::Equals(
            $project.FileName,
            "$safeName.json",
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            return $true
        }
    }

    if (Test-Path -LiteralPath $safePath) {
        return $true
    }

    return $false
}

# ------------------------------------------------------------
# PROJE SEÇİMİ
# ------------------------------------------------------------

function Select-Project {
    param(
        [string]$Prompt = "Lütfen bir proje seçin:"
    )

    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "📂 PROJE SEÇİMİ" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""

    $projects = @(Load-ProjectList)

    if ($projects.Count -eq 0) {
        Write-Host "Kayıtlı proje bulunamadı." -ForegroundColor Yellow
        Pause-Action
        return $null
    }

    Write-Host $Prompt
    Write-Host ""

    for ($i = 0; $i -lt $projects.Count; $i++) {
        Write-Host (
            "{0}. {1}    [{2}]" -f
            ($i + 1),
            $projects[$i].ProjeAdi,
            $projects[$i].OlusturmaTarihi
        )
    }

    Write-Host ""

    while ($true) {
        $selection = Read-Host "Seçim (1-$($projects.Count)) veya Q"

        if ($selection.ToUpper() -eq "Q") {
            return $null
        }

        if ($selection -notmatch '^\d+$') {
            Write-Host "❌ Geçersiz seçim." -ForegroundColor Red
            continue
        }

        $index = [int]$selection - 1

        if ($index -lt 0 -or $index -ge $projects.Count) {
            Write-Host "❌ Geçersiz aralık." -ForegroundColor Red
            continue
        }

        return $projects[$index]
    }
}

# ------------------------------------------------------------
# HİSSELİ ARSA / FATURA ALICILARI
# ------------------------------------------------------------

function Read-ArsaFaturaAlcilari {
    param(
        [bool]$Hisseli = $true
    )

    while ($true) {
        Clear-Host

        Write-Host ("-" * 72) -ForegroundColor Magenta
        if ($Hisseli) {
            Write-Host "👥 HİSSELİ ARSA SAHİBİ / AYRI FATURA ALICILARI" -ForegroundColor Magenta
        }
        else {
            Write-Host "👤 ARSA SAHİBİ / FATURA ALICISI" -ForegroundColor Magenta
        }
        Write-Host ("-" * 72) -ForegroundColor Magenta
        Write-Host ""

        if ($Hisseli) {
            Write-Host "Her hissedarı ayrı girin. Hak payları toplamı %100 olmalıdır."
            Write-Host "KDV mükellefiyeti her alıcı için ayrı belirlenir."
        }
        else {
            Write-Host "Hissesiz bölümde tek fatura alıcısı %100 pay ile kaydedilir."
            Write-Host "KDV mükellefiyeti bu alıcı için belirlenir."
        }

        Write-Host "0 matrah = VUK 267 maliyet bedeli esasına göre otomatik hesap." -ForegroundColor DarkYellow
        Write-Host ""

        if ($Hisseli) {
            $count = Read-IntegerInput "Kaç adet hissedar / fatura alıcısı var?" 2
        }
        else {
            $count = 1
        }

        if ($null -eq $count) {
            return $null
        }

        $recipients = @()

        for ($i = 1; $i -le $count; $i++) {
            Write-Host ""
            if ($Hisseli) {
                Write-Host "▶ HİSSEDAR / FATURA ALICISI $i" -ForegroundColor Green
            }
            else {
                Write-Host "▶ FATURA ALICISI" -ForegroundColor Green
            }

            $name = Read-StringInput "Ad / Unvan"
            if ($null -eq $name) {
                return $null
            }

            $isKdvMukellefi = Read-YesNo "KDV mükellefi mi?" $false
            if ($null -eq $isKdvMukellefi) {
                return $null
            }

            if ($Hisseli) {
                $share = Read-HakPayiPercent "Hak Payı (% veya 1/2 gibi)"
                if ($null -eq $share) {
                    return $null
                }
            }
            else {
                $share = [decimal]100
            }

            $manualMatrah = Read-DecimalInput "Fatura Matrahı (0 = otomatik)" 0
            if ($null -eq $manualMatrah) {
                return $null
            }

            $recipients += [PSCustomObject]@{
                AdUnvan      = $name
                KdvMukellefi = [bool]$isKdvMukellefi
                HakPayi      = [decimal]$share
                Matrah       = [decimal]$manualMatrah
            }
        }

        $totalShare = [decimal]0

        foreach ($recipient in $recipients) {
            $totalShare += [decimal]$recipient.HakPayi
        }

        $totalShare = [math]::Round($totalShare, 4)

        if ($totalShare -ne 100) {
            Write-Host ""
            Write-Host "❌ Hak payı toplamı %100 değil. Mevcut toplam: %$totalShare" -ForegroundColor Red
            Pause-Action
            continue
        }

        return @($recipients)
    }
}

# ------------------------------------------------------------
# BAĞIMSIZ BÖLÜM OLUŞTURMA
# ------------------------------------------------------------

function New-Bolum {
    param(
        [int]$No
    )

    Write-Host ""
    Write-Host "════════ BÖLÜM $No ════════" -ForegroundColor Green

    $area = Read-DecimalInput "Alan (m²)" 0.01
    if ($null -eq $area) {
        return $null
    }

    $owner = Read-OwnerType
    if ($null -eq $owner) {
        return $null
    }

    $hisseli = $false

    if ($owner -eq "Arsa Sahibi") {
        $hisseli = Read-YesNo "Bu bağımsız bölüm hisseli mi?" $false
        if ($null -eq $hisseli) {
            return $null
        }
    }

    $kdvRate = Read-DecimalInput "KDV Oranı (%)" 0 0 100
    if ($null -eq $kdvRate) {
        return $null
    }

    $section = [PSCustomObject]@{
        No             = "$No"
        Alan           = [decimal]$area
        Taraf          = $owner
        Hisseli        = [bool]$hisseli
        Kdv            = [decimal]$kdvRate
        FaturaAlcilari = @()
        SatisMatrahi   = [decimal]0
        Matrah         = [decimal]0
    }

    if ($owner -eq "Arsa Sahibi") {
        $recipients = Read-ArsaFaturaAlcilari -Hisseli $hisseli

        if ($null -eq $recipients) {
            return $null
        }

        $section.FaturaAlcilari = @($recipients)
    }
    else {
        $matrah = Read-DecimalInput "Satış / Fatura Matrahı (0 = henüz satılmadı)" 0

        if ($null -eq $matrah) {
            return $null
        }

        $section.SatisMatrahi = [decimal]$matrah
        $section.Matrah = [decimal]$matrah
    }

    return $section
}

# ------------------------------------------------------------
# YENİ PROJE
# ------------------------------------------------------------

function Yeni-Proje {
    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "🆕 YENİ KAT KARŞILIĞI İNŞAAT PROJESİ" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "İstediğiniz zaman Q ile işlemi iptal edebilirsiniz." -ForegroundColor DarkYellow
    Write-Host ""

    $projectName = Read-StringInput "Proje Adı"
    if ($null -eq $projectName) {
        return $false
    }

    if (Test-ProjectNameExists $projectName) {
        Write-Host ""
        Write-Host "❌ Bu isimde veya aynı güvenli dosya adına sahip bir proje zaten mevcut." -ForegroundColor Red
        Pause-Action
        return $false
    }

    $parcel = Read-StringInput "Parsel / Ada"
    if ($null -eq $parcel) {
        return $false
    }

    $licenseDate = Read-DateInput "Ruhsat Tarihi (gg.aa.yyyy)"
    if ($null -eq $licenseDate) {
        return $false
    }

    $totalArea = Read-DecimalInput "Toplam İnşaat Alanı (m²)" 0.01
    if ($null -eq $totalArea) {
        return $false
    }

    $totalCost = Read-DecimalInput "Toplam İnşaat Maliyeti (₺)" 0
    if ($null -eq $totalCost) {
        return $false
    }

    $unitCost = [decimal]0

    if ($totalArea -gt 0) {
        $unitCost = [math]::Round($totalCost / $totalArea, 2)
    }

    $unitCostText = "Birim Maliyet: {0:N2} ₺/m²" -f $unitCost
    Write-Host ""
    Write-Host $unitCostText -ForegroundColor Green
    Write-Host ""

    $sectionCount = Read-IntegerInput "Kaç adet bağımsız bölüm var?" 1
    if ($null -eq $sectionCount) {
        return $false
    }

    $sections = @()

    for ($i = 1; $i -le $sectionCount; $i++) {
        $section = New-Bolum -No $i

        if ($null -eq $section) {
            return $false
        }

        $sections += $section
    }

    $project = [PSCustomObject]@{
        ProgramSurumu   = $GLOBAL:ProgramVersion
        ProjeAdi        = $projectName
        OlusturmaTarihi = (Get-Date).ToString("dd.MM.yyyy")
        Parsel          = $parcel
        RuhsatTarihi    = $licenseDate
        ToplamAlan      = [decimal]$totalArea
        ToplamMaliyet   = [decimal]$totalCost
        BirimMaliyet    = [decimal]$unitCost
        Bolumler        = @($sections)
    }

    try {
        Save-ProjectJson $project
    }
    catch {
        Write-Host "❌ Proje kaydedilemedi: $_" -ForegroundColor Red
        return $false
    }

    Write-Host ""
    Write-Host "✅ Proje başarıyla kaydedildi." -ForegroundColor Green

    return $true
}

# ------------------------------------------------------------
# PROJE LİSTELE
# ------------------------------------------------------------

function Listele-Projeler {
    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "📂 KAYITLI PROJELER" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""

    $projects = @(Load-ProjectList)

    if ($projects.Count -eq 0) {
        Write-Host "Kayıtlı proje bulunamadı." -ForegroundColor Yellow
    }
    else {
        foreach ($project in $projects) {
            Write-Host ("• {0}   [{1}]" -f $project.ProjeAdi, $project.OlusturmaTarihi)
        }
    }

    return $true
}

# ------------------------------------------------------------
# PROJE SİL
# ------------------------------------------------------------

function Sil-Proje {
    $projectItem = Select-Project "Silmek istediğiniz projeyi seçin:"

    if ($null -eq $projectItem) {
        return $false
    }

    Write-Host ""
    Write-Host ("Seçilen proje: {0}" -f $projectItem.ProjeAdi) -ForegroundColor Cyan

    $confirm = Read-YesNo "Bu proje silinsin mi?" $false

    if ($null -eq $confirm) {
        return $false
    }

    if (-not $confirm) {
        Write-Host "Silme işlemi iptal edildi." -ForegroundColor Yellow
        return $false
    }

    try {
        Remove-Item -LiteralPath $projectItem.Path -Force -ErrorAction Stop

        Write-Host "✅ Proje başarıyla silindi." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Proje silinemedi: $_" -ForegroundColor Red
        return $false
    }
}

# ------------------------------------------------------------
# PROJE GÜNCELLE
# ------------------------------------------------------------

function Update-Proje {
    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "🛠️ PROJE GÜNCELLEME" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta

    $projectItem = Select-Project "Güncellenecek projeyi seçin:"

    if ($null -eq $projectItem) {
        return $false
    }

    $project = Load-Project $projectItem

    if ($null -eq $project) {
        return $false
    }

    Write-Host ""
    Write-Host ("Proje: {0}" -f $project.ProjeAdi) -ForegroundColor Cyan
    Write-Host "ENTER = mevcut değeri koru | Q = işlemi iptal et" -ForegroundColor DarkYellow
    Write-Host ""

    $project.Parsel = Read-StringInput "Parsel / Ada" $project.Parsel
    if ($null -eq $project.Parsel) {
        return $false
    }

    $project.RuhsatTarihi = Read-DateInput "Ruhsat Tarihi" $project.RuhsatTarihi
    if ($null -eq $project.RuhsatTarihi) {
        return $false
    }

    $project.ToplamAlan = Read-DecimalInput "Toplam İnşaat Alanı (m²)" 0.01 ([decimal]$project.ToplamAlan)
    if ($null -eq $project.ToplamAlan) {
        return $false
    }

    $project.ToplamMaliyet = Read-DecimalInput "Toplam İnşaat Maliyeti (₺)" 0 ([decimal]$project.ToplamMaliyet)
    if ($null -eq $project.ToplamMaliyet) {
        return $false
    }

    if ($project.ToplamAlan -gt 0) {
        $project.BirimMaliyet = [math]::Round($project.ToplamMaliyet / $project.ToplamAlan, 2)
    }
    else {
        $project.BirimMaliyet = 0
    }

    $unitCostText = "Birim Maliyet: {0:N2} ₺/m²" -f $project.BirimMaliyet
    Write-Host ""
    Write-Host $unitCostText -ForegroundColor Green

    for ($i = 0; $i -lt @($project.Bolumler).Count; $i++) {
        $b = $project.Bolumler[$i]

        Write-Host ""
        Write-Host "════════ BÖLÜM $($b.No) ════════" -ForegroundColor Green

        $b.Alan = Read-DecimalInput "- Alan (m²)" 0.01 ([decimal]$b.Alan)
        if ($null -eq $b.Alan) {
            return $false
        }

        $oldOwner = [string]$b.Taraf
        $b.Taraf = Read-OwnerType $oldOwner

        if ($null -eq $b.Taraf) {
            return $false
        }

        $oldHisseli = $false
        if ($b.PSObject.Properties.Name -contains "Hisseli") {
            $oldHisseli = [bool]$b.Hisseli
        }
        elseif ($b.Taraf -eq "Arsa Sahibi") {
            $oldHisseli = (@($b.FaturaAlcilari).Count -gt 1)
        }

        $newHisseli = $false
        if ($b.Taraf -eq "Arsa Sahibi") {
            $newHisseli = Read-YesNo "Bu bağımsız bölüm hisseli mi?" $oldHisseli
            if ($null -eq $newHisseli) {
                return $false
            }
        }

        $b.Hisseli = [bool]$newHisseli

        $b.Kdv = Read-DecimalInput "- KDV Oranı (%)" 0 ([decimal]$b.Kdv) 100
        if ($null -eq $b.Kdv) {
            return $false
        }

        if ($b.Taraf -eq "Arsa Sahibi") {
            $recipientList = @($b.FaturaAlcilari)
            $needsRecipientEntry = ($oldOwner -ne "Arsa Sahibi" -or $recipientList.Count -eq 0 -or $oldHisseli -ne $newHisseli)

            # Eski projede KDV durumu bilinmiyorsa yeniden giriş zorunlu.
            if (-not $needsRecipientEntry) {
                foreach ($recipient in $recipientList) {
                    if (-not ($recipient.PSObject.Properties.Name -contains "KdvMukellefi")) {
                        $needsRecipientEntry = $true
                        break
                    }

                    if ($null -eq $recipient.KdvMukellefi) {
                        $needsRecipientEntry = $true
                        break
                    }
                }
            }

            if ($needsRecipientEntry) {
                Write-Host ""
                Write-Host "⚠️ Fatura alıcısı / hisse bilgileri yeniden girilmelidir." -ForegroundColor Yellow

                $b.FaturaAlcilari = Read-ArsaFaturaAlcilari -Hisseli $newHisseli

                if ($null -eq $b.FaturaAlcilari) {
                    return $false
                }
            }
            else {
                $editRecipients = Read-YesNo "Fatura alıcılarını / hisseleri yeniden düzenlemek ister misiniz?" $false

                if ($null -eq $editRecipients) {
                    return $false
                }

                if ($editRecipients) {
                    $b.FaturaAlcilari = Read-ArsaFaturaAlcilari -Hisseli $newHisseli

                    if ($null -eq $b.FaturaAlcilari) {
                        return $false
                    }
                }
            }

            $b.SatisMatrahi = 0
            $b.Matrah = 0
        }
        else {
            $b.Hisseli = $false
            $b.FaturaAlcilari = @()

            $currentMatrah = [decimal]0

            if ($b.PSObject.Properties.Name -contains "SatisMatrahi") {
                $currentMatrah = [decimal]$b.SatisMatrahi
            }
            elseif ($b.PSObject.Properties.Name -contains "Matrah") {
                $currentMatrah = [decimal]$b.Matrah
            }

            $b.SatisMatrahi = Read-DecimalInput "- Satış / Fatura Matrahı (0 = satılmadı)" 0 $currentMatrah

            if ($null -eq $b.SatisMatrahi) {
                return $false
            }

            $b.Matrah = $b.SatisMatrahi
        }
    }

    try {
        Save-ProjectJson $project $projectItem.Path
    }
    catch {
        Write-Host "❌ Proje güncellenemedi: $_" -ForegroundColor Red
        return $false
    }

    Write-Host ""
    Write-Host "✅ Proje başarıyla güncellendi." -ForegroundColor Green

    return $true
}

# ------------------------------------------------------------
# RAPOR HESAPLAMA MOTORU
# ------------------------------------------------------------

function Get-ProjectReport {
    param(
        $Project
    )

    $arsaToplamAlan = [decimal]0
    $mteToplamAlan = [decimal]0
    $bolumToplamAlan = [decimal]0

    foreach ($b in @($Project.Bolumler)) {
        $bolumToplamAlan += [decimal]$b.Alan

        if ($b.Taraf -eq "Arsa Sahibi") {
            $arsaToplamAlan += [decimal]$b.Alan
        }
        else {
            $mteToplamAlan += [decimal]$b.Alan
        }
    }

    $arsaToplamMaliyet = [math]::Round(
        $arsaToplamAlan * [decimal]$Project.BirimMaliyet,
        2
    )

    if ($mteToplamAlan -gt 0) {
        $birimArsaMaliyet = [math]::Round(
            $arsaToplamMaliyet / $mteToplamAlan,
            2
        )
    }
    else {
        $birimArsaMaliyet = [decimal]0
    }

    $rows = @()
    $recipientRows = @()

    $totalAlan = [decimal]0
    $totalBolumMaliyet = [decimal]0
    $totalArsaMaliyet = [decimal]0
    $totalToplamMaliyet = [decimal]0
    $totalMatrah = [decimal]0
    $totalKdvTutar = [decimal]0
    $totalGenelToplam = [decimal]0
    $totalKarZarar = [decimal]0
    $karZararHesaplanan = 0

    foreach ($b in @($Project.Bolumler)) {
        $bolumMaliyet = [math]::Round(
            [decimal]$b.Alan * [decimal]$Project.BirimMaliyet,
            2
        )

        $totalAlan += [decimal]$b.Alan

        if ($b.Taraf -eq "Arsa Sahibi") {
            $recipients = @($b.FaturaAlcilari)

            if ($recipients.Count -eq 0) {
                throw "Bölüm $($b.No) için arsa sahibi fatura alıcısı bulunmuyor."
            }

            if (-not ($b.PSObject.Properties.Name -contains "Hisseli")) {
                $b | Add-Member -MemberType NoteProperty -Name "Hisseli" -Value ($recipients.Count -gt 1)
            }

            $totalShareCheck = [decimal]0
            foreach ($recipient in $recipients) {
                $totalShareCheck += [decimal]$recipient.HakPayi
            }
            $totalShareCheck = [math]::Round($totalShareCheck, 4)

            if ($totalShareCheck -ne 100) {
                throw "Bölüm $($b.No) için hak payı toplamı %100 değil: %$totalShareCheck"
            }

            $sectionMatrah = [decimal]0
            $sectionKdv = [decimal]0
            $sectionGenelToplam = [decimal]0
            $sectionKarZarar = [decimal]0
            $sectionMatrahSources = @()
            $sectionKdvStatus = @()
            $sectionMarkup = [decimal]0
            $sectionMarkupCount = 0
            $sectionShareCostUsed = [decimal]0

            # Her bölüm için maliyet toplamı ana tabloda bir kez yazılır.
            $totalBolumMaliyet += $bolumMaliyet

            for ($ri = 0; $ri -lt $recipients.Count; $ri++) {
                $recipient = $recipients[$ri]

                if (-not ($recipient.PSObject.Properties.Name -contains "KdvMukellefi")) {
                    throw "Bölüm $($b.No) / $($recipient.AdUnvan) için KDV mükellefiyeti belirtilmemiş."
                }

                if ($null -eq $recipient.KdvMukellefi) {
                    throw "Bölüm $($b.No) / $($recipient.AdUnvan) için KDV mükellefiyeti belirtilmemiş."
                }

                $share = [decimal]$recipient.HakPayi

                if ($share -le 0 -or $share -gt 100) {
                    throw "Bölüm $($b.No) / $($recipient.AdUnvan) hak payı geçersiz."
                }

                # Son hissedara kalan tutarı vererek kuruş farkını kapat.
                if ($ri -eq ($recipients.Count - 1)) {
                    $shareCost = [math]::Round($bolumMaliyet - $sectionShareCostUsed, 2)
                }
                else {
                    $shareCost = [math]::Round(
                        $bolumMaliyet * ($share / 100),
                        2
                    )
                }

                $sectionShareCostUsed += $shareCost

                if ([bool]$recipient.KdvMukellefi) {
                    $markupRate = [decimal]0.05
                    $kdvStatus = "Evet"
                }
                else {
                    $markupRate = [decimal]0.10
                    $kdvStatus = "Hayır"
                }

                if ([decimal]$recipient.Matrah -gt 0) {
                    $matrah = [decimal]$recipient.Matrah
                    $matrahKaynak = "Manuel"
                }
                else {
                    $matrah = [math]::Round(
                        $shareCost * (1 + $markupRate),
                        2
                    )
                    $matrahKaynak = "Otomatik"
                }

                $kdvTutar = [math]::Round(
                    $matrah * ([decimal]$b.Kdv / 100),
                    2
                )

                $genelToplam = [math]::Round(
                    $matrah + $kdvTutar,
                    2
                )

                $karZarar = [math]::Round(
                    $matrah - $shareCost,
                    2
                )


                $shareArea = [math]::Round(
                    [decimal]$b.Alan * ($share / 100),
                    2
                )

                $recipientRows += [PSCustomObject]@{
                    BolumNo             = [string]$b.No
                    Alan                = [decimal]$b.Alan
                    AlanPayi            = [decimal]$shareArea
                    Taraf               = "Arsa Sahibi"
                    FaturaAlıcısı       = [string]$recipient.AdUnvan
                    HakPayi             = [decimal]$share
                    KdvMukellefi        = $kdvStatus
                    BolumMaliyeti       = [decimal]$shareCost
                    KdvOrani            = [decimal]$b.Kdv
                    Matrah              = [decimal]$matrah
                    MatrahKaynak        = $matrahKaynak
                    KdvTutar            = [decimal]$kdvTutar
                    GenelToplam         = [decimal]$genelToplam
                    KarZarar            = [decimal]$karZarar
                    MarkupOrani         = [decimal]($markupRate * 100)
                }

                $sectionMatrah += $matrah
                $sectionKdv += $kdvTutar
                $sectionGenelToplam += $genelToplam
                $sectionKarZarar += $karZarar
                $sectionMatrahSources += $matrahKaynak
                $sectionKdvStatus += $kdvStatus
                $sectionMarkup += ($markupRate * 100)
                $sectionMarkupCount++
            }

            $sectionMatrah = [math]::Round($sectionMatrah, 2)
            $sectionKdv = [math]::Round($sectionKdv, 2)
            $sectionGenelToplam = [math]::Round($sectionGenelToplam, 2)
            $sectionKarZarar = [math]::Round($sectionKarZarar, 2)

            $totalMatrah += $sectionMatrah
            $totalKdvTutar += $sectionKdv
            $totalGenelToplam += $sectionGenelToplam
            $totalKarZarar += $sectionKarZarar
            $karZararHesaplanan++

            $isKdvEvet = ($sectionKdvStatus | Where-Object { $_ -eq "Evet" }).Count
            $isKdvHayir = ($sectionKdvStatus | Where-Object { $_ -eq "Hayır" }).Count

            if ($isKdvEvet -gt 0 -and $isKdvHayir -gt 0) {
                $sectionKdvMukellefi = "Karışık"
            }
            elseif ($isKdvEvet -gt 0) {
                $sectionKdvMukellefi = "Evet"
            }
            else {
                $sectionKdvMukellefi = "Hayır"
            }

            if ($sectionMatrahSources -contains "Manuel" -and $sectionMatrahSources -contains "Otomatik") {
                $sectionMatrahKaynak = "Karışık"
            }
            elseif ($sectionMatrahSources -contains "Manuel") {
                $sectionMatrahKaynak = "Manuel"
            }
            else {
                $sectionMatrahKaynak = "Otomatik"
            }

            if ([bool]$b.Hisseli -and $recipients.Count -gt 1) {
                $sectionRecipientText = "Hisseli ($($recipients.Count) kişi)"
            }
            else {
                $sectionRecipientText = [string]$recipients[0].AdUnvan
            }

            if ($sectionMarkupCount -gt 0) {
                $sectionMarkup = [math]::Round($sectionMarkup / $sectionMarkupCount, 2)
            }

            $rows += [PSCustomObject]@{
                BolumNo             = [string]$b.No
                Alan                = [decimal]$b.Alan
                Taraf               = "Arsa Sahibi"
                Hisseli             = [bool]$b.Hisseli
                FaturaAlıcısı       = $sectionRecipientText
                HakPayi             = [decimal]100
                KdvMukellefi        = $sectionKdvMukellefi
                BolumMaliyet        = [decimal]$bolumMaliyet
                ArsaMaliyet         = [decimal]0
                ToplamMaliyet       = [decimal]0
                KdvOrani            = [decimal]$b.Kdv
                Matrah              = [decimal]$sectionMatrah
                MatrahKaynak        = $sectionMatrahKaynak
                KdvTutar            = [decimal]$sectionKdv
                GenelToplam         = [decimal]$sectionGenelToplam
                KarZarar            = [decimal]$sectionKarZarar
                KarZararHesaplandi  = $true
                MarkupOrani         = [decimal]$sectionMarkup
            }
        }
        else {
            if ($mteToplamAlan -gt 0) {
                $arsaMaliyet = [math]::Round(
                    [decimal]$b.Alan * $birimArsaMaliyet,
                    2
                )
            }
            else {
                $arsaMaliyet = [decimal]0
            }

            $toplamMaliyet = [math]::Round(
                $bolumMaliyet + $arsaMaliyet,
                2
            )

            $matrah = [decimal]$b.SatisMatrahi

            if ($matrah -gt 0) {
                $kdvTutar = [math]::Round(
                    $matrah * ([decimal]$b.Kdv / 100),
                    2
                )

                $genelToplam = [math]::Round(
                    $matrah + $kdvTutar,
                    2
                )

                $karZarar = [math]::Round(
                    $matrah - $toplamMaliyet,
                    2
                )
            }
            else {
                $kdvTutar = [decimal]0
                $genelToplam = [decimal]0
                $karZarar = $null
            }

            $totalBolumMaliyet += $bolumMaliyet
            $totalArsaMaliyet += $arsaMaliyet
            $totalToplamMaliyet += $toplamMaliyet

            if ($matrah -gt 0) {
                $totalMatrah += $matrah
                $totalKdvTutar += $kdvTutar
                $totalGenelToplam += $genelToplam
                $totalKarZarar += $karZarar
                $karZararHesaplanan++
                $matrahKaynak = "Manuel"
            }
            else {
                $matrahKaynak = "Satılmadı"
            }

            $rows += [PSCustomObject]@{
                BolumNo             = [string]$b.No
                Alan                = [decimal]$b.Alan
                Taraf               = "Müteahhit"
                Hisseli             = $false
                FaturaAlıcısı       = "Müteahhit Satışı"
                HakPayi             = [decimal]100
                KdvMukellefi        = ""
                BolumMaliyet        = [decimal]$bolumMaliyet
                ArsaMaliyet         = [decimal]$arsaMaliyet
                ToplamMaliyet       = [decimal]$toplamMaliyet
                KdvOrani            = [decimal]$b.Kdv
                Matrah              = [decimal]$matrah
                MatrahKaynak        = $matrahKaynak
                KdvTutar            = [decimal]$kdvTutar
                GenelToplam         = [decimal]$genelToplam
                KarZarar            = $karZarar
                KarZararHesaplandi  = ($null -ne $karZarar)
                MarkupOrani         = [decimal]0
            }
        }
    }

    # Bölüm alanı ile proje alanı farkını raporlamada göstermek için saklıyoruz.
    $alanFarki = [math]::Round(
        [decimal]$Project.ToplamAlan - $bolumToplamAlan,
        2
    )

    return [PSCustomObject]@{
        ArsaToplamAlan         = [decimal]$arsaToplamAlan
        MteToplamAlan          = [decimal]$mteToplamAlan
        BolumToplamAlan        = [decimal]$bolumToplamAlan
        ProjeAlanFarki         = [decimal]$alanFarki
        ArsaToplamMaliyet      = [decimal]$arsaToplamMaliyet
        BirimArsaMaliyet       = [decimal]$birimArsaMaliyet
        TotalAlan              = [decimal]$totalAlan
        TotalBolumMaliyet      = [decimal]$totalBolumMaliyet
        TotalArsaMaliyet       = [decimal]$totalArsaMaliyet
        TotalToplamMaliyet     = [decimal]$totalToplamMaliyet
        TotalMatrah            = [decimal]$totalMatrah
        TotalKdvTutar          = [decimal]$totalKdvTutar
        TotalGenelToplam       = [decimal]$totalGenelToplam
        TotalKarZarar          = [decimal]$totalKarZarar
        KarZararHesaplanan     = $karZararHesaplanan
        Rows                   = @($rows)
        RecipientRows          = @($recipientRows)
    }
}

# ------------------------------------------------------------
# EKRAN RAPORU
# ------------------------------------------------------------

function Show-Signature {
    $title = "Hazırlayan: "
    $name = "SMMM Mustafa Orakcı"

    foreach ($char in $title.ToCharArray()) {
        Write-Host -NoNewline $char -ForegroundColor Red
        Start-Sleep -Milliseconds 30
    }

    foreach ($char in $name.ToCharArray()) {
        Write-Host -NoNewline $char -ForegroundColor Cyan
        Start-Sleep -Milliseconds 30
    }

    Write-Host ""
}

function Show-ProjectReport {
    param(
        $Project,
        $ReportData
    )

    Clear-Host

    $tableWidth = 190

    Write-Host ("=" * $tableWidth)
    Write-Host "🏗️ KAT KARŞILIĞI İNŞAAT PROJE RAPORU" -ForegroundColor Cyan
    Write-Host ("=" * $tableWidth)

    Write-Host ""

    $text = "Proje Adı".PadRight(35) + ": " + [string]$Project.ProjeAdi
    Write-Host $text

    $text = "Oluşturma Tarihi".PadRight(35) + ": " + [string]$Project.OlusturmaTarihi
    Write-Host $text

    $text = "Parsel / Ada".PadRight(35) + ": " + [string]$Project.Parsel
    Write-Host $text

    $text = "Ruhsat Tarihi".PadRight(35) + ": " + [string]$Project.RuhsatTarihi
    Write-Host $text

    Write-Host ""
    Write-Host "📜 PROJE / MALİYET BİLGİLERİ" -ForegroundColor Cyan

    $text = "Toplam İnşaat Alanı".PadRight(35) + ": {0:N2} m²" -f $Project.ToplamAlan
    Write-Host $text

    $text = "Bölümlerin Toplam Alanı".PadRight(35) + ": {0:N2} m²" -f $ReportData.BolumToplamAlan
    Write-Host $text

    if ($ReportData.ProjeAlanFarki -ne 0) {
        $text = "Alan Farkı".PadRight(35) + ": {0:N2} m²" -f $ReportData.ProjeAlanFarki
        Write-Host $text -ForegroundColor Yellow
    }

    $text = "Toplam İnşaat Maliyeti".PadRight(35) + ": {0:N2} ₺" -f $Project.ToplamMaliyet
    Write-Host $text

    $text = "Birim İnşaat Maliyeti".PadRight(35) + ": {0:N2} ₺/m²" -f $Project.BirimMaliyet
    Write-Host $text

    $text = "Arsa Sahibi Toplam Alan".PadRight(35) + ": {0:N2} m²" -f $ReportData.ArsaToplamAlan
    Write-Host $text

    $text = "Müteahhit Toplam Alan".PadRight(35) + ": {0:N2} m²" -f $ReportData.MteToplamAlan
    Write-Host $text

    $text = "Arsa Sahibi Bölümlerinin Maliyeti".PadRight(35) + ": {0:N2} ₺" -f $ReportData.ArsaToplamMaliyet
    Write-Host $text

    $text = "Dağıtılacak Birim Arsa Maliyeti".PadRight(35) + ": {0:N2} ₺/m²" -f $ReportData.BirimArsaMaliyet
    Write-Host $text

    Write-Host ""
    Write-Host "📋 BAĞIMSIZ BÖLÜM ANA RAPORU" -ForegroundColor Cyan
    Write-Host ""

    $header = "{0,-5} {1,9} {2,-13} {3,-24} {4,7} {5,-10} {6,14} {7,13} {8,14} {9,7} {10,14} {11,13} {12,14} {13,14}" -f `
        "No", "Alan", "Taraf", "Fatura / Dağılım", "Pay %", "KDV Mük.", `
        "İnş. ML.", "Arsa ML.", "Toplam ML.", "KDV %", `
        "Matrah", "KDV", "Gen. Top.", "K/Z"

    Write-Host $header -ForegroundColor Gray
    Write-Host ("-" * $tableWidth)

    foreach ($row in $ReportData.Rows) {

        if ($null -eq $row.KarZarar) {
            $profitText = "—"
        }
        else {
            $profitText = "{0:N2}" -f $row.KarZarar
        }

        if ($row.Taraf -eq "Arsa Sahibi") {
            $toplamMaliyetText = "—"
        }
        else {
            $toplamMaliyetText = "{0:N2}" -f $row.ToplamMaliyet
        }

        $line = "{0,-5} {1,9:N2} {2,-13} {3,-24} {4,7:N2} {5,-10} {6,14:N2} {7,13:N2} {8,14} {9,7:N2} {10,14:N2} {11,13:N2} {12,14:N2} {13,14}" -f `
            $row.BolumNo,
            $row.Alan,
            $row.Taraf,
            $row.FaturaAlıcısı,
            $row.HakPayi,
            $row.KdvMukellefi,
            $row.BolumMaliyet,
            $row.ArsaMaliyet,
            $toplamMaliyetText,
            $row.KdvOrani,
            $row.Matrah,
            $row.KdvTutar,
            $row.GenelToplam,
            $profitText

        Write-Host $line
    }

    Write-Host ("-" * $tableWidth)

    $totalToplamMaliyetText = "{0:N2}" -f $ReportData.TotalToplamMaliyet
    $totalText = "{0,-5} {1,9:N2} {2,-13} {3,-24} {4,7} {5,-10} {6,14:N2} {7,13:N2} {8,14} {9,7} {10,14:N2} {11,13:N2} {12,14:N2} {13,14:N2}" -f `
        "TOP.",
        $ReportData.TotalAlan,
        "",
        "TOPLAM",
        "",
        "",
        $ReportData.TotalBolumMaliyet,
        $ReportData.TotalArsaMaliyet,
        $totalToplamMaliyetText,
        "",
        $ReportData.TotalMatrah,
        $ReportData.TotalKdvTutar,
        $ReportData.TotalGenelToplam,
        $ReportData.TotalKarZarar

    Write-Host $totalText -ForegroundColor Green
    Write-Host ("=" * $tableWidth)

    if (@($ReportData.RecipientRows).Count -gt 0) {
        Write-Host ""
        Write-Host "📌 ARSA SAHİBİ HİSSE / FATURA ALT KIRILIMI" -ForegroundColor Cyan
        Write-Host ""

        $detailWidth = 150
        $detailHeader = "{0,-5} {1,9} {2,10} {3,-20} {4,8} {5,10} {6,-10} {7,14} {8,7} {9,13} {10,12} {11,14} {12,13}" -f `
            "Böl.", "Bölüm Alan", "Alan Payı", "Fatura Alıcısı", "Pay %", "KDV Mük.", `
            "KDV %", "Maliyet Payı", "Matrah", "KDV", "Gen. Top.", "K/Z", "Matrah Kay."

        Write-Host $detailHeader -ForegroundColor Gray
        Write-Host ("-" * $detailWidth)

        foreach ($detail in $ReportData.RecipientRows) {
            $detailLine = "{0,-5} {1,9:N2} {2,10:N2} {3,-20} {4,8:N2} {5,10} {6,14:N2} {7,14:N2} {8,13:N2} {9,12:N2} {10,14:N2} {11,13:N2} {12,-13}" -f `
                $detail.BolumNo,
                $detail.Alan,
                $detail.AlanPayi,
                $detail.FaturaAlıcısı,
                $detail.HakPayi,
                $detail.KdvMukellefi,
                $detail.KdvOrani,
                $detail.BolumMaliyeti,
                $detail.Matrah,
                $detail.KdvTutar,
                $detail.GenelToplam,
                $detail.KarZarar,
                $detail.MatrahKaynak

            Write-Host $detailLine
        }

        Write-Host ("-" * $detailWidth)

        $detailTotalCost = [decimal]0
        $detailTotalMatrah = [decimal]0
        $detailTotalKdv = [decimal]0
        $detailTotalGenel = [decimal]0
        $detailTotalKz = [decimal]0

        foreach ($detail in $ReportData.RecipientRows) {
            $detailTotalCost += [decimal]$detail.BolumMaliyeti
            $detailTotalMatrah += [decimal]$detail.Matrah
            $detailTotalKdv += [decimal]$detail.KdvTutar
            $detailTotalGenel += [decimal]$detail.GenelToplam
            $detailTotalKz += [decimal]$detail.KarZarar
        }

        $detailTotalLine = "{0,-5} {1,9} {2,10:N2} {3,-20} {4,8:N2} {5,10} {6,14} {7,14:N2} {8,13:N2} {9,12:N2} {10,14:N2} {11,13:N2} {12,-13}" -f `
            "TOP.",
            "",
            $ReportData.ArsaToplamAlan,
            "TOPLAM",
            100,
            "",
            "",
            $detailTotalCost,
            $detailTotalMatrah,
            $detailTotalKdv,
            $detailTotalGenel,
            $detailTotalKz,
            ""

        Write-Host $detailTotalLine -ForegroundColor Green
        Write-Host ("=" * $detailWidth)
    }

    Write-Host ""
    Write-Host "Not: Ana tablo her bağımsız bölümü yalnızca bir kez gösterir; hisseli bölümlerin hissedarları alt tabloda açılır." -ForegroundColor DarkYellow
    Write-Host "Not: Arsa sahibi bölümünde %5/%10 maliyet bedeli ilavesi fatura alıcısının KDV mükellefiyetine göre alt tabloda hesaplanır." -ForegroundColor DarkYellow
    Write-Host "Not: Müteahhit bölümünde satış/fatura matrahı 0 ise K/Z hesaplanmaz ve '—' gösterilir." -ForegroundColor DarkYellow

    Write-Host ""
    Show-Signature
}

# ------------------------------------------------------------
# EXCEL RAPORU
# ------------------------------------------------------------

function Create-ExcelReport {
    param(
        $Project,
        $ReportData
    )

    Write-Host ""
    Write-Host "📊 Excel raporu oluşturuluyor..." -ForegroundColor Green

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $detailWorksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Sheets.Item(1)
        $worksheet.Name = "Ana Rapor"

        $detailWorksheet = $workbook.Sheets.Add()
        $detailWorksheet.Name = "Hisse Detayı"

        # ------------------------------------------------------------
        # ANA RAPOR SAYFASI
        # ------------------------------------------------------------

        $row = 1

        $worksheet.Cells.Item($row, 1).Value2 = "KAT KARŞILIĞI İNŞAAT PROJE RAPORU"
        $worksheet.Cells.Item($row, 1).Font.Bold = $true
        $worksheet.Cells.Item($row, 1).Font.Size = 14

        $worksheet.Range(
            $worksheet.Cells.Item($row, 1),
            $worksheet.Cells.Item($row, 15)
        ).Merge() | Out-Null

        $row += 2

        $worksheet.Cells.Item($row, 1).Value2 = "Proje Adı"
        $worksheet.Cells.Item($row, 2).Value2 = $Project.ProjeAdi
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Oluşturma Tarihi"
        $worksheet.Cells.Item($row, 2).Value2 = $Project.OlusturmaTarihi
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Parsel / Ada"
        $worksheet.Cells.Item($row, 2).Value2 = $Project.Parsel
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Ruhsat Tarihi"
        $worksheet.Cells.Item($row, 2).Value2 = $Project.RuhsatTarihi
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Toplam İnşaat Alanı (m²)"
        $worksheet.Cells.Item($row, 2).Value2 = [double]$Project.ToplamAlan
        $worksheet.Cells.Item($row, 2).NumberFormat = "#,##0.00"
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Toplam İnşaat Maliyeti (₺)"
        $worksheet.Cells.Item($row, 2).Value2 = [double]$Project.ToplamMaliyet
        $worksheet.Cells.Item($row, 2).NumberFormat = "#,##0.00 ""₺"""
        $row++

        $worksheet.Cells.Item($row, 1).Value2 = "Birim İnşaat Maliyeti (₺/m²)"
        $worksheet.Cells.Item($row, 2).Value2 = [double]$Project.BirimMaliyet
        $worksheet.Cells.Item($row, 2).NumberFormat = "#,##0.00 ""₺"""
        $row += 2

        $worksheet.Cells.Item($row, 1).Value2 = "ÖZET"
        $worksheet.Cells.Item($row, 1).Font.Bold = $true
        $row++

        $summaryRows = @(
            [PSCustomObject]@{ Etiket = "Arsa Sahibi Toplam Alan (m²)"; Deger = [double]$ReportData.ArsaToplamAlan; Format = "#,##0.00" },
            [PSCustomObject]@{ Etiket = "Müteahhit Toplam Alan (m²)"; Deger = [double]$ReportData.MteToplamAlan; Format = "#,##0.00" },
            [PSCustomObject]@{ Etiket = "Bölümlerin Toplam Alanı (m²)"; Deger = [double]$ReportData.BolumToplamAlan; Format = "#,##0.00" },
            [PSCustomObject]@{ Etiket = "Proje Alan Farkı (m²)"; Deger = [double]$ReportData.ProjeAlanFarki; Format = "#,##0.00" },
            [PSCustomObject]@{ Etiket = "Arsa Sahibi Bölümleri Maliyeti (₺)"; Deger = [double]$ReportData.ArsaToplamMaliyet; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Birim Arsa Maliyeti (₺/m²)"; Deger = [double]$ReportData.BirimArsaMaliyet; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam Bölüm İnşaat Maliyeti (₺)"; Deger = [double]$ReportData.TotalBolumMaliyet; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam Arsa Maliyeti (₺)"; Deger = [double]$ReportData.TotalArsaMaliyet; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Müteahhit Toplam Maliyeti (₺)"; Deger = [double]$ReportData.TotalToplamMaliyet; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam Fatura Matrahı (₺)"; Deger = [double]$ReportData.TotalMatrah; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam KDV (₺)"; Deger = [double]$ReportData.TotalKdvTutar; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam Genel Toplam (₺)"; Deger = [double]$ReportData.TotalGenelToplam; Format = "#,##0.00 ""₺""" },
            [PSCustomObject]@{ Etiket = "Toplam K/Z (₺)"; Deger = [double]$ReportData.TotalKarZarar; Format = "#,##0.00 ""₺""" }
        )

        foreach ($summary in $summaryRows) {
            $worksheet.Cells.Item($row, 1).Value2 = [string]$summary.Etiket
            $worksheet.Cells.Item($row, 2).Value2 = [double]$summary.Deger
            $worksheet.Cells.Item($row, 2).NumberFormat = [string]$summary.Format
            $row++
        }

        $row += 2

        $headers = @(
            "Bölüm",
            "Alan (m²)",
            "Taraf",
            "Fatura / Dağılım",
            "Bölüm Pay %",
            "KDV Mükellefi",
            "Bölüm İnşaat Maliyeti",
            "Arsa Maliyeti",
            "Toplam Maliyet",
            "KDV %",
            "Matrah",
            "KDV",
            "Genel Toplam",
            "K/Z",
            "Matrah Kaynağı"
        )

        for ($i = 0; $i -lt $headers.Count; $i++) {
            $cell = $worksheet.Cells.Item($row, $i + 1)
            $cell.Value2 = $headers[$i]
            $cell.Font.Bold = $true
            $cell.Borders.Weight = 2
        }

        $row++

        foreach ($item in $ReportData.Rows) {
            $worksheet.Cells.Item($row, 1).Value2 = $item.BolumNo
            $worksheet.Cells.Item($row, 2).Value2 = [double]$item.Alan
            $worksheet.Cells.Item($row, 3).Value2 = $item.Taraf
            $worksheet.Cells.Item($row, 4).Value2 = $item.FaturaAlıcısı
            $worksheet.Cells.Item($row, 5).Value2 = [double]($item.HakPayi / 100)
            $worksheet.Cells.Item($row, 6).Value2 = $item.KdvMukellefi
            $worksheet.Cells.Item($row, 7).Value2 = [double]$item.BolumMaliyet
            $worksheet.Cells.Item($row, 8).Value2 = [double]$item.ArsaMaliyet
            $worksheet.Cells.Item($row, 9).Value2 = [double]$item.ToplamMaliyet
            $worksheet.Cells.Item($row, 10).Value2 = [double]($item.KdvOrani / 100)
            $worksheet.Cells.Item($row, 11).Value2 = [double]$item.Matrah
            $worksheet.Cells.Item($row, 12).Value2 = [double]$item.KdvTutar
            $worksheet.Cells.Item($row, 13).Value2 = [double]$item.GenelToplam

            if ($null -ne $item.KarZarar) {
                $worksheet.Cells.Item($row, 14).Value2 = [double]$item.KarZarar
            }
            else {
                $worksheet.Cells.Item($row, 14).Value2 = ""
            }

            $worksheet.Cells.Item($row, 15).Value2 = $item.MatrahKaynak

            $worksheet.Cells.Item($row, 2).NumberFormat = "#,##0.00"
            $worksheet.Cells.Item($row, 5).NumberFormat = "0.00%"
            $worksheet.Cells.Item($row, 7).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 8).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 9).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 10).NumberFormat = "0.00%"
            $worksheet.Cells.Item($row, 11).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 12).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 13).NumberFormat = "#,##0.00 ""₺"""
            $worksheet.Cells.Item($row, 14).NumberFormat = "#,##0.00 ""₺"""

            $row++
        }

        for ($i = 1; $i -le $headers.Count; $i++) {
            $worksheet.Cells.Item($row, $i).Font.Bold = $true
            $worksheet.Cells.Item($row, $i).Borders.Weight = 2
        }

        $worksheet.Cells.Item($row, 1).Value2 = "TOPLAM"
        $worksheet.Cells.Item($row, 2).Value2 = [double]$ReportData.TotalAlan
        $worksheet.Cells.Item($row, 7).Value2 = [double]$ReportData.TotalBolumMaliyet
        $worksheet.Cells.Item($row, 8).Value2 = [double]$ReportData.TotalArsaMaliyet
        $worksheet.Cells.Item($row, 9).Value2 = [double]$ReportData.TotalToplamMaliyet
        $worksheet.Cells.Item($row, 11).Value2 = [double]$ReportData.TotalMatrah
        $worksheet.Cells.Item($row, 12).Value2 = [double]$ReportData.TotalKdvTutar
        $worksheet.Cells.Item($row, 13).Value2 = [double]$ReportData.TotalGenelToplam
        $worksheet.Cells.Item($row, 14).Value2 = [double]$ReportData.TotalKarZarar

        $worksheet.Cells.Item($row, 2).NumberFormat = "#,##0.00"
        $worksheet.Cells.Item($row, 7).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 8).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 9).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 11).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 12).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 13).NumberFormat = "#,##0.00 ""₺"""
        $worksheet.Cells.Item($row, 14).NumberFormat = "#,##0.00 ""₺"""

        $worksheet.Columns.AutoFit() | Out-Null

        # ------------------------------------------------------------
        # HİSSE DETAYI SAYFASI
        # ------------------------------------------------------------

        $detailRow = 1
        $detailWorksheet.Cells.Item($detailRow, 1).Value2 = "ARSA SAHİBİ HİSSE / FATURA ALT KIRILIMI"
        $detailWorksheet.Cells.Item($detailRow, 1).Font.Bold = $true
        $detailWorksheet.Cells.Item($detailRow, 1).Font.Size = 14
        $detailWorksheet.Range(
            $detailWorksheet.Cells.Item($detailRow, 1),
            $detailWorksheet.Cells.Item($detailRow, 14)
        ).Merge() | Out-Null
        $detailRow += 2

        $detailHeaders = @(
            "Bölüm",
            "Bölüm Alanı (m²)",
            "Alan Payı (m²)",
            "Fatura Alıcısı",
            "Hak Payı %",
            "KDV Mükellefi",
            "Maliyet Payı",
            "KDV %",
            "Matrah",
            "KDV",
            "Genel Toplam",
            "K/Z",
            "Matrah Kaynağı",
            "Markup %"
        )

        for ($i = 0; $i -lt $detailHeaders.Count; $i++) {
            $cell = $detailWorksheet.Cells.Item($detailRow, $i + 1)
            $cell.Value2 = $detailHeaders[$i]
            $cell.Font.Bold = $true
            $cell.Borders.Weight = 2
        }
        $detailRow++

        foreach ($item in $ReportData.RecipientRows) {
            $detailWorksheet.Cells.Item($detailRow, 1).Value2 = $item.BolumNo
            $detailWorksheet.Cells.Item($detailRow, 2).Value2 = [double]$item.Alan
            $detailWorksheet.Cells.Item($detailRow, 3).Value2 = [double]$item.AlanPayi
            $detailWorksheet.Cells.Item($detailRow, 4).Value2 = $item.FaturaAlıcısı
            $detailWorksheet.Cells.Item($detailRow, 5).Value2 = [double]($item.HakPayi / 100)
            $detailWorksheet.Cells.Item($detailRow, 6).Value2 = $item.KdvMukellefi
            $detailWorksheet.Cells.Item($detailRow, 7).Value2 = [double]$item.BolumMaliyeti
            $detailWorksheet.Cells.Item($detailRow, 8).Value2 = [double]($item.KdvOrani / 100)
            $detailWorksheet.Cells.Item($detailRow, 9).Value2 = [double]$item.Matrah
            $detailWorksheet.Cells.Item($detailRow, 10).Value2 = [double]$item.KdvTutar
            $detailWorksheet.Cells.Item($detailRow, 11).Value2 = [double]$item.GenelToplam
            $detailWorksheet.Cells.Item($detailRow, 12).Value2 = [double]$item.KarZarar
            $detailWorksheet.Cells.Item($detailRow, 13).Value2 = $item.MatrahKaynak
            $detailWorksheet.Cells.Item($detailRow, 14).Value2 = [double]($item.MarkupOrani / 100)

            $detailWorksheet.Cells.Item($detailRow, 2).NumberFormat = "#,##0.00"
            $detailWorksheet.Cells.Item($detailRow, 3).NumberFormat = "#,##0.00"
            $detailWorksheet.Cells.Item($detailRow, 5).NumberFormat = "0.00%"
            $detailWorksheet.Cells.Item($detailRow, 7).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 8).NumberFormat = "0.00%"
            $detailWorksheet.Cells.Item($detailRow, 9).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 10).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 11).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 12).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 14).NumberFormat = "0.00%"

            $detailRow++
        }

        if (@($ReportData.RecipientRows).Count -gt 0) {
            for ($i = 1; $i -le $detailHeaders.Count; $i++) {
                $detailWorksheet.Cells.Item($detailRow, $i).Font.Bold = $true
                $detailWorksheet.Cells.Item($detailRow, $i).Borders.Weight = 2
            }

            $detailWorksheet.Cells.Item($detailRow, 1).Value2 = "TOPLAM"
            $detailWorksheet.Cells.Item($detailRow, 3).Value2 = [double]$ReportData.ArsaToplamAlan
            $detailWorksheet.Cells.Item($detailRow, 5).Value2 = 1

            $detailCost = [decimal]0
            $detailMatrah = [decimal]0
            $detailKdv = [decimal]0
            $detailGenel = [decimal]0
            $detailKz = [decimal]0

            foreach ($item in $ReportData.RecipientRows) {
                $detailCost += [decimal]$item.BolumMaliyeti
                $detailMatrah += [decimal]$item.Matrah
                $detailKdv += [decimal]$item.KdvTutar
                $detailGenel += [decimal]$item.GenelToplam
                $detailKz += [decimal]$item.KarZarar
            }

            $detailWorksheet.Cells.Item($detailRow, 7).Value2 = [double]$detailCost
            $detailWorksheet.Cells.Item($detailRow, 9).Value2 = [double]$detailMatrah
            $detailWorksheet.Cells.Item($detailRow, 10).Value2 = [double]$detailKdv
            $detailWorksheet.Cells.Item($detailRow, 11).Value2 = [double]$detailGenel
            $detailWorksheet.Cells.Item($detailRow, 12).Value2 = [double]$detailKz

            $detailWorksheet.Cells.Item($detailRow, 3).NumberFormat = "#,##0.00"
            $detailWorksheet.Cells.Item($detailRow, 5).NumberFormat = "0.00%"
            $detailWorksheet.Cells.Item($detailRow, 7).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 9).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 10).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 11).NumberFormat = "#,##0.00 ""₺"""
            $detailWorksheet.Cells.Item($detailRow, 12).NumberFormat = "#,##0.00 ""₺"""
        }

        $detailWorksheet.Columns.AutoFit() | Out-Null

        # Excel çıktısını rapor klasörüne kaydet (Data\Modules\insaat\Reports).
        $safeName = ConvertTo-SafeFileName $Project.ProjeAdi
        $excelPath = Join-Path $GLOBAL:ProjectReportFolder ("{0}_Rapor_{1}.xlsx" -f $safeName, (Get-Date).ToString("yyyyMMdd_HHmmss"))

        $workbook.SaveAs($excelPath)
        Write-Host "✅ Excel raporu kaydedildi: $excelPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Excel raporu oluşturulurken hata oluştu:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return $false
    }
    finally {
        if ($detailWorksheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($detailWorksheet) | Out-Null
        }

        if ($worksheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        }

        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        if ($excel) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ------------------------------------------------------------
# RAPORLA
# ------------------------------------------------------------

function Raporla-Proje {
    $projectItem = Select-Project "Raporlamak istediğiniz projeyi seçin:"

    if ($null -eq $projectItem) {
        return $false
    }

    $project = Load-Project $projectItem

    if ($null -eq $project) {
        return $false
    }

    try {
        $reportData = Get-ProjectReport $project
    }
    catch {
        Write-Host ""
        Write-Host "❌ RAPOR HESAPLAMA HATASI" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host ""
        return $false
    }

    Show-ProjectReport $project $reportData

    Write-Host ""
    $excelChoice = Read-Host "📁 Excel çıktısı oluşturulsun mu? [E/H]"

    if ($excelChoice.ToUpper() -eq "E") {
        Create-ExcelReport $project $reportData | Out-Null
    }

    return $true
}

# ------------------------------------------------------------
# ANA MENÜ
# ------------------------------------------------------------

function Show-Menu {
    Clear-Host

    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "🏗️ --- KAT KARŞILIĞI İNŞAAT RAPORLAMA ARACI ---" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host "  1. 🆕  Yeni Proje Oluştur"
    Write-Host "  2. 📄  Projeleri Listele"
    Write-Host "  3. 🗑️  Proje Sil"
    Write-Host "  4. 🔄  Proje Güncelle"
    Write-Host "  5. 📊  Rapor Al"
    Write-Host "  Q  🚪  Çıkış"
    Write-Host ("-" * 72) -ForegroundColor DarkGray
}

# ------------------------------------------------------------
# ÇALIŞTIR
# ------------------------------------------------------------

Ensure-DataFolder

$selection = ""

while ($true) {
    Show-Menu

    Write-Host ""
    $selection = Read-Host "Seçiminiz"

    switch ($selection.ToUpper()) {
        "1" {
            $success = Yeni-Proje
            if ($success) {
                Pause-Action
            }
        }

        "2" {
            Listele-Projeler | Out-Null
            Pause-Action
        }

        "3" {
            $success = Sil-Proje
            if ($success) {
                Pause-Action
            }
        }

        "4" {
            $success = Update-Proje
            if ($success) {
                Pause-Action
            }
        }

        "5" {
            $success = Raporla-Proje
            Pause-Action
        }

        "Q" {
            Write-Host ""
            Write-Host "Çıkış yapılıyor..." -ForegroundColor Yellow
            break
        }

        default {
            Write-Host ""
            Write-Host "❌ Geçersiz seçim!" -ForegroundColor Red
            Start-Sleep -Milliseconds 700
        }
    }

    if ($selection.ToUpper() -eq "Q") {
        break
    }
}

