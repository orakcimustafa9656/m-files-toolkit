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
# 🚀 Sayısal Giriş Alan ve Kontrol Eden Fonksiyon
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
# 🚀 Geçerli yüzde girişini alır ve doğrular
#    Minimum 0, maksimum 100 aralığında ondalıklı sayı kabul eder
# --------------------------------------------
function Get-ValidPercentageInput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        [double]$Default = 70
    )
    while ($true) {
        $input = Read-ToolkitLine -Prompt ("{0,-17} (Mevcut %$Default): " -f $Prompt)
        if ([string]::IsNullOrWhiteSpace($input)) {
            return $Default / 100
        }
        if ($input -notmatch '^[0-9.,]+$') {
            Write-Host "❌ Lütfen sadece 1-100 arası sayı ve tek nokta (.) veya virgül (,) girin." -ForegroundColor Red
            continue
        }
        if ($input.Contains('.') -and $input.Contains(',')) {
            Write-Host "❌ Hem nokta hem virgül kullanılamaz." -ForegroundColor Red
            continue
        }
        $dotCount = ($input.ToCharArray() | Where-Object { $_ -eq '.' }).Count
        $commaCount = ($input.ToCharArray() | Where-Object { $_ -eq ',' }).Count
        if ($dotCount -gt 1 -or $commaCount -gt 1) {
            Write-Host "❌ Birden fazla ondalık ayırıcı kullanılamaz." -ForegroundColor Red
            continue
        }
        if ($commaCount -eq 1) {
            $input = $input -replace ',', '.'
        }
        try {
            $value = [double]$input
            if ($value -ge 1 -and $value -le 100) {
                return $value / 100
            }
            else {
                Write-Host "❌ Lütfen 1 ile 100 arasında bir değer girin." -ForegroundColor Red
            }
        }
        catch {
            Write-Host "❌ Geçersiz sayı formatı." -ForegroundColor Red
        }
    }
}

# --------------------------------------------
# 🚀 Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    $kisitlamaOrani = Get-ValidPercentageInput -Prompt "Kısıtlama Oranı" -Default 70
    $genelToplam = Get-ValidDoubleInput -Prompt "Genel Toplam               " -Min 0
    $kdv = Get-ValidDoubleInput -Prompt "KDV                        " -Min 0

    $matrah = $genelToplam - $kdv
    $giderYazilabilecekTutar = [math]::Round($matrah * $kisitlamaOrani, 2)
    $kkgGider = [math]::Round($matrah - $giderYazilabilecekTutar, 2)
    $indirilebilecekKDV = [math]::Round($kdv * $kisitlamaOrani, 2)
    $kkgKDV = [math]::Round($kdv - $indirilebilecekKDV, 2)

    function Write-Entry ($label, $borc, $alacak, $labelColor = "Green") {
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $borcText = if ([string]::IsNullOrEmpty($borc)) { "".PadRight(15) } else { [string]::Format($culture, "{0,-15:F2}", $borc) }
        $alacakText = if ([string]::IsNullOrEmpty($alacak)) { "".PadRight(14) } else { [string]::Format($culture, "{0,-14:F2}", $alacak) }
        $line = "{0,-30}: {1}{2}" -f $label, $borcText, $alacakText
        Write-Host $line -ForegroundColor $labelColor
    }

    Write-Entry "Kısıtlama Oranı" ("{0}%" -f ([math]::Round($kisitlamaOrani * 100, 2))) ""
    Write-Entry "Matrah" $matrah ""
    Write-Host ""
    Write-Host ("{0,-31} {1,-15} {2,-14}" -f "", "Borç Tutar", "Alacak Tutar") -ForegroundColor Cyan
    Write-Host ("{0,-30} {1,-45}" -f "", ("-" * 45)) -ForegroundColor DarkGray
    Write-Entry "Gider Yazılabilecek Tutar" $giderYazilabilecekTutar ""
    Write-Entry "KKG Gider" $kkgGider ""
    Write-Entry "İndirilebilecek KDV" $indirilebilecekKDV ""
    Write-Entry "KKG KDV" $kkgKDV ""

    $borcToplam = [math]::Round($giderYazilabilecekTutar + $kkgGider + $indirilebilecekKDV + $kkgKDV, 2)
    $alacakToplam = [math]::Round($genelToplam, 2)

    Write-Host ("{0,-30} {1,-45}" -f "", ("-" * 45)) -ForegroundColor DarkGray
    Write-Entry "Genel Toplam" $borcToplam (" " + $alacakToplam)
}

# --------------------------------------------
# 🚀 Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  ⛽ Binek Otomobil Gider Kısıtlaması Hesaplama" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Bu araç, gider olarak yazılabilecek tutarları ve KKEG (Kısıtlanmış Giderler)" -ForegroundColor Yellow
    Write-Host "📌 hesaplar. Kısıtlama oranını kendiniz belirleyebilirsiniz (varsayılan %70)." -ForegroundColor Yellow
    Write-Host "📌 Lütfen aşağıdaki değerleri giriniz:" -ForegroundColor Yellow
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
