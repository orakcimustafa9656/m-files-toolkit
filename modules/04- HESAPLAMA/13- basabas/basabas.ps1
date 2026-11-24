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
        $culture = if ($input.Contains(',')) { 
            [System.Globalization.CultureInfo]::GetCultureInfo("tr-TR") 
        } else { 
            [System.Globalization.CultureInfo]::InvariantCulture 
        }
        $success = [double]::TryParse($input, [System.Globalization.NumberStyles]::Number, $culture, [ref]$number)
        if (-not $success) {
            Write-Host "❌ Geçersiz sayı formatı. Örnek: 1234,56 veya 1234.56" -ForegroundColor $ErrorColor
            continue
        }
        if ($Min -ne $null -and $number -lt $Min) {
            Write-Host "❌ Girilen değer minimum $Min olmalı." -ForegroundColor $ErrorColor
            continue
        }
        if ($Max -ne $null -and $number -gt $Max) {
            Write-Host "❌ Girilen değer maksimum $Max olmalı." -ForegroundColor $ErrorColor
            continue
        }
        return [math]::Round($number, 2)
    }
}

# --------------------------------------------
# 🚀 Program Döngüsü
# --------------------------------------------
function StartProgram {
    $sabit    = Get-ValidDoubleInput -Prompt "Sabit Maliyetleri Giriniz (TL)      "         -Min 0 -Icon "🏗️"
    $degisken = Get-ValidDoubleInput -Prompt "Birim Değişken Maliyeti Giriniz (TL)"   -Min 0 -Icon "🔧"
    $satis    = Get-ValidDoubleInput -Prompt "Birim Satış Fiyatını Giriniz (TL)   "      -Min 0 -Icon "💵"

    if ($satis -le $degisken) {
        Write-Host "`n❗ HATA: Satış fiyatı, değişken maliyetten büyük olmalıdır. Bu fiyatla kâr edemezsiniz." -ForegroundColor Red
        return
    }
    $miktar = [math]::Round($sabit / ($satis - $degisken), 2)
    $ciro   = [math]::Round($miktar * $satis, 2)
    Write-Host "`n📊 Başabaş Noktası Sonuçları" -ForegroundColor Magenta
    Write-Host "-------------------------------------" -ForegroundColor Magenta
    Write-Host "🎯 Satış Miktarı (Adet)   : $miktar" -ForegroundColor Green
    Write-Host "💰 Gerekli Ciro (TL)      : $ciro" -ForegroundColor Green
}

# --------------------------------------------
# 🚀 Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  📊 Başabaş Noktası Hesaplama Aracı' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Sabit ve değişken maliyeti olan ürünün Başabaş noktası miktar ve tutarını hesaplar." -ForegroundColor Yellow
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
