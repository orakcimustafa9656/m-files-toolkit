#requires -Version 5.1

# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

function Read-ToolkitLine {
    param([string]$Prompt)

    if ($Prompt) { Write-Host $Prompt -NoNewline }
    $buffer = [System.Text.StringBuilder]::new()
    $cursor = 0
    $lineStartX = [System.Console]::CursorLeft
    $lineStartY = [System.Console]::CursorTop
    $width = [System.Console]::BufferWidth
    if ($width -lt 1) { $width = 80 }
    $renderedLength = 0

    while ($true) {
        $key = [System.Console]::ReadKey($true)
        $changed = $false
        switch ($key.Key) {
            Enter {
                $absolute = $lineStartX + $buffer.Length
                [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
                [System.Console]::WriteLine()
                return $buffer.ToString()
            }
            Escape {
                $absolute = $lineStartX + $buffer.Length
                [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
                [System.Console]::WriteLine()
                throw [System.OperationCanceledException]::new('Kullanıcı Esc ile işlemi iptal etti.')
            }
            Left {
                if ($cursor -gt 0) { $cursor--; $changed = $true }
            }
            Right {
                if ($cursor -lt $buffer.Length) { $cursor++; $changed = $true }
            }
            Home {
                if ($cursor -gt 0) { $cursor = 0; $changed = $true }
            }
            End {
                if ($cursor -lt $buffer.Length) { $cursor = $buffer.Length; $changed = $true }
            }
            Backspace {
                if ($cursor -gt 0) {
                    [void]$buffer.Remove($cursor - 1, 1)
                    $cursor--
                    $changed = $true
                }
            }
            Delete {
                if ($cursor -lt $buffer.Length) {
                    [void]$buffer.Remove($cursor, 1)
                    $changed = $true
                }
            }
            default {
                if (-not [char]::IsControl($key.KeyChar)) {
                    [void]$buffer.Insert($cursor, $key.KeyChar)
                    $cursor++
                    $changed = $true
                }
            }
        }

        if ($changed) {
            [System.Console]::SetCursorPosition($lineStartX, $lineStartY)
            [System.Console]::Write($buffer.ToString())
            $clearCount = $renderedLength - $buffer.Length
            if ($clearCount -gt 0) { [System.Console]::Write((' ' * $clearCount)) }
            $absolute = $lineStartX + $cursor
            [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
            $renderedLength = $buffer.Length
        }
    }
}


# --------------------------------------------
# 🚀  Yardımcı Gösterim Fonksiyonları
# --------------------------------------------
function Show-Title     { Write-Host "`nEsc - İşlemi İptal Et" -ForegroundColor DarkRed }
function Show-Section   { param($text) Write-Host "`n📌 $text" -ForegroundColor Yellow; Write-Host ("─" * 40) -ForegroundColor DarkGray }
function Show-Error     { param($msg) Write-Host "❌ $msg" -ForegroundColor Red }
function Show-Warning   { param($msg) Write-Host "⚠️ $msg" -ForegroundColor Yellow }
function Show-Success   { param($msg) Write-Host "✅ $msg" -ForegroundColor Green }
function Show-Info      { param($msg) Write-Host "ℹ️ $msg" -ForegroundColor Cyan }

# --------------------------------------------
# 🚀  Giriş Kontrol Fonksiyonları
# --------------------------------------------
function Get-ValidDateInput {
    param([string]$prompt = "Tarih girin (yyyy-MM-dd)")
    while ($true) {
        $input = Read-ToolkitLine -Prompt ($prompt + ": ")
        try {
            $dt = [datetime]::ParseExact($input, 'yyyy-MM-dd', $null)
            return $dt.ToString('yyyy-MM-dd')
        } catch {
            Show-Warning "Tarih formatı geçersiz. yyyy-MM-dd formatında giriniz."
        }
    }
}

# --------------------------------------------
# 🚀  Sayısal Giriş Alan ve Kontrol Eden Fonksiyon
# --------------------------------------------
function Get-ValidDoubleInput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt,

        [Nullable[double]]$Min = $null,

        [Nullable[double]]$Max = $null,

        [string]$Icon = $null,

        [ConsoleColor]$PromptColor = [ConsoleColor]::Cyan,

        [ConsoleColor]$ErrorColor = [ConsoleColor]::Red
    )
    if (-not $Icon) { $Icon = "🔹" }
    while ($true) {
        [double]$number = 0
        Write-Host "${Icon} ${Prompt}: " -ForegroundColor $PromptColor -NoNewline
        $input = Read-ToolkitLine
        if ([string]::IsNullOrWhiteSpace($input)) {
            Write-Host "❌ Giriş boş bırakılamaz." -ForegroundColor $ErrorColor
            continue
        }
        # Kültür seçimi virgül veya nokta için
        $culture = if ($input.Contains(',')) { 
            [System.Globalization.CultureInfo]::GetCultureInfo("tr-TR") 
        } else { 
            [System.Globalization.CultureInfo]::InvariantCulture 
        }
        # Geçersiz sayı kontrolü
        $success = [double]::TryParse($input, [System.Globalization.NumberStyles]::Number, $culture, [ref]$number)
        if (-not $success) {
            Write-Host "❌ Geçersiz sayı formatı veya sayı çok büyük/küçük. Örnek: 1234,56 veya 1234.56" -ForegroundColor $ErrorColor
            continue
        }
        # Min kontrolü
        if ($Min -ne $null -and $number -lt $Min) {
            Write-Host "❌ Girilen değer minimum $Min olmalı." -ForegroundColor $ErrorColor
            continue
        }
        # Max kontrolü
        if ($Max -ne $null -and $number -gt $Max) {
            Write-Host "❌ Girilen değer maksimum $Max olmalı." -ForegroundColor $ErrorColor
            continue
        }
        return [math]::Round($number, 2)
    }
}

# --------------------------------------------
# 🚀  TCMB Kur Çekme
# --------------------------------------------
function Get-TcmbKur {
    param($Tarih, $DovizKod)
    try {
        $dt = [datetime]::ParseExact($Tarih, 'yyyy-MM-dd', $null)
        $url = "https://www.tcmb.gov.tr/kurlar/$($dt.ToString('yyyyMM'))/$($dt.ToString('ddMMyyyy')).xml"
        $xml = [xml](Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop).Content
        $kurNode = $xml.Tarih_Date.Currency | Where-Object { $_.CurrencyCode -eq $DovizKod }
        if ($kurNode.ForexBuying -ne '') {
            return [decimal]$kurNode.ForexBuying
        } else {
            Show-Warning "Kur verisi bulunamadı."
            return $null
        }
    } catch {
        Show-Warning "Kur verisi alınamadı (hafta sonu veya veri yok)."
        return $null
    }
}

# --------------------------------------------
# 🚀  CSV Export
# --------------------------------------------
function Export-Rapor {
    param($data)
    if ($data.Count -eq 0) {
        Show-Warning "Rapor için kayıtlı işlem yok."
        return $false
    }
    $reportFolder = Get-MFTModuleDataPath -ModuleId 'kurfarki' -Category Reports
    $fileName = "kurfarki_rapor_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $reportPath = Join-Path $reportFolder $fileName
    $data | Export-Csv -Path $reportPath -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
    Show-Success "Rapor oluşturuldu: $reportPath"
    return $true
}

# --------------------------------------------
# 🚀  Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    $allFirmalar = @()
    while ($true) {
        Show-Title
        Show-Section "Yeni İşlem Girişi"

        $tarih = Get-ValidDateInput -prompt "📅 Kur Tarihi (yyyy-MM-dd)  "

        while ($true) {
            Write-Host "💰 Döviz Türü (USD/EUR)     : " -NoNewline -ForegroundColor Magenta
            $doviz = (Read-ToolkitLine).ToUpper()
            if ($doviz -in @('USD','EUR')) { break }
            Show-Warning "Sadece USD veya EUR girilebilir."
        }

        Show-Info "TCMB kuru sorgulanıyor..."
        $kur = Get-TcmbKur -Tarih $tarih -DovizKod $doviz
        if (-not $kur) { continue }
        Show-Success "Kur alındı               : $kur"

        $dovizBakiye = Get-ValidDoubleInput -Prompt "💸 Döviz Bakiyesi        " -Min 0
        $tlBakiye    = Get-ValidDoubleInput -Prompt "📈 Cari TL Bakiyesi      " -Min 0

        Write-Host "   💬 Firma Açıklaması Girin: " -NoNewline -ForegroundColor Cyan
        $aciklama = Read-ToolkitLine

        $hesaplananTL = [math]::Round($dovizBakiye * $kur, 2)
        $kurFarki     = [math]::Round($tlBakiye - $hesaplananTL, 2)

        Show-Section "Hesap Sonucu"
        Write-Host ("🔀 Hesaplanan TL Karşılığı  : {0,12}" -f ("{0:N2} TL" -f $hesaplananTL)) -ForegroundColor Cyan
        Write-Host ("📉 Cari Hesap TL Bakiyesi   : {0,12}" -f ("{0:N2} TL" -f $tlBakiye)) -ForegroundColor Cyan

        $kurEmoji = if ($kurFarki -lt 0) { "📉" } else { "📈" }
        $kurRenk  = if ($kurFarki -lt 0) { "Red" } else { "Green" }
        Write-Host ("$kurEmoji Kur Farkı                : {0,12}" -f ("{0:N2} TL" -f $kurFarki)) -ForegroundColor $kurRenk

        Write-Host "
ℹ️ Not: Alıcı/satıcı durumuna göre kur farkı gelir/gider olabilir." -ForegroundColor DarkYellow

        $allFirmalar += [pscustomobject]@{
            Tarih = $tarih; Doviz = $doviz; Kur = $kur
            DovizBakiyesi = $dovizBakiye; CariTLBakiyesi = $tlBakiye
            HesaplananTL = $hesaplananTL; KurFarki = $kurFarki; Aciklama = $aciklama
        }

        while ($true) {
            Write-Host "➡️ Yeni firma için devam edilsin mi? (E/Q): " -NoNewline -ForegroundColor Yellow
            $devam = Read-ToolkitLine
            if ([string]::IsNullOrWhiteSpace($devam)) { continue }
            $devam = $devam.ToUpper()

            if ($devam -eq 'E') { break }
            elseif ($devam -eq 'Q') {
                Write-Host "📂 Rapor oluşturulsun mu? (E/H): " -NoNewline -ForegroundColor Yellow
                $rapor = Read-ToolkitLine
                if ($rapor.ToUpper() -eq 'E') { Export-Rapor -data $allFirmalar }
                return
            }
        }
    }
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  💱 Kur Farkı Hesaplama Aracı' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Döviz bakiyesi veren hesapların kur farkını hesaplayabilirsiniz." -ForegroundColor Yellow
    Write-Host "📌 İşlem sonunda Csv raporu alabilirsiniz." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    try {
        StartProgram
    }
    catch [System.OperationCanceledException] {
        Write-Host '⚠️ İşlem iptal edildi.' -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host ("-" * 52) -ForegroundColor DarkGray
    Write-Host "  🔚 [Q/Esc] Ana menüye dön  |  [D] Yeniden çalıştır" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor DarkGray
    Write-Host ""
    do {
        $keyInfo = [System.Console]::ReadKey($true)
        $next = $keyInfo.Key.ToString().ToUpper()
    } while ($next -ne 'Q' -and $next -ne 'D' -and $next -ne 'ESCAPE')
    if ($next -eq 'Q' -or $next -eq 'ESCAPE') { break }
} while ($true)
