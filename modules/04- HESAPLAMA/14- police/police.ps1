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
# 🚀 Tarih girdisi alma ve doğrulama fonksiyonu
# --------------------------------------------
function Read-Date {
    param(
        [string]$Prompt
    )
    $date = [datetime]::MinValue
    do {
        $input = Read-ToolkitLine -Prompt ($Prompt + ": ")
        $refDate = [ref]$date
        $valid = [datetime]::TryParseExact($input, "dd.MM.yyyy", $null, [System.Globalization.DateTimeStyles]::None, $refDate)
        $date = $refDate.Value
        if (-not $valid) {
            Write-Host "Geçersiz tarih! Lütfen GG.AA.YYYY formatında girin." -ForegroundColor Red
        } else {
            return $date
        }
    } until ($valid)
}

# --------------------------------------------
# 🚀 Ondalık sayı girdisi alma fonksiyonu
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
            Write-Host "❌ Geçersiz sayı formatı veya sayı çok büyük/küçük. Örnek: 1234,56 veya 1234.56" -ForegroundColor $ErrorColor
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
# 🚀 Poliçe muhasebeleştirme hesaplama ve çıktı fonksiyonu
# --------------------------------------------
function Calculate-PoliçeMuhasebeleştirme {
    param (
        [datetime]$BaslangicTarihi,
        [datetime]$BitisTarihi,
        [double]$Tutar
    )

    $trCulture = [System.Globalization.CultureInfo]::GetCultureInfo("tr-TR")

    $baslangicYili = $BaslangicTarihi.Year
    $bitisYili = $BitisTarihi.Year
    $toplamGun = ($BitisTarihi - $BaslangicTarihi).Days + 1
    $sonGunIlkYil = [datetime]::new($baslangicYili, 12, 31)
    $ilkYilBitisTarihi = if ($BitisTarihi -lt $sonGunIlkYil) { $BitisTarihi } else { $sonGunIlkYil }
    $gunIlkYil = ($ilkYilBitisTarihi - $BaslangicTarihi).Days + 1
    $gunErtesiYil = $toplamGun - $gunIlkYil

    $gunlukTutar = $Tutar / $toplamGun
    # Yıllık toplamları kuruş bazında hesapla; son yıl kalan tutarı alır.
    $tutarIlkYil = [math]::Round($Tutar * $gunIlkYil / $toplamGun, 2, [MidpointRounding]::AwayFromZero)
    $tutarErtesiYil = [math]::Round($Tutar - $tutarIlkYil, 2, [MidpointRounding]::AwayFromZero)

    Write-Host "İlk Yıl Gün Sayısı                      : $gunIlkYil gün"
    Write-Host "Ertesi Yıl Gün Sayısı                   : $gunErtesiYil gün"
    Write-Host "İlk Yıl Gider Yazılabilecek Tutar       : $tutarIlkYil"
    Write-Host "Ertesi Yıl Gider Yazılabilecek Tutar    : $tutarErtesiYil"
    Write-Host ""
    Write-Host "Ay Bazında Gider Dağılımı:" -ForegroundColor Magenta

    # Poliçe dönemi boyunca her ayı kapsa ve ilk yıl/ertesi yıl gruplarını
    # ayrı ayrı yuvarla. Her grubun son ayı kuruş farkını kapatır.
    $ayliklar = [System.Collections.Generic.List[object]]::new()
    $ayBaslangici = [datetime]::new($BaslangicTarihi.Year, $BaslangicTarihi.Month, 1)
    while ($ayBaslangici -le $BitisTarihi) {
        $aySonu = $ayBaslangici.AddMonths(1).AddDays(-1)
        $hesapBaslangic = if ($BaslangicTarihi -gt $ayBaslangici) { $BaslangicTarihi } else { $ayBaslangici }
        $hesapBitis = if ($BitisTarihi -lt $aySonu) { $BitisTarihi } else { $aySonu }
        $gunSayisi = ($hesapBitis - $hesapBaslangic).Days + 1
        $ayTutar = [math]::Round($gunSayisi * $gunlukTutar, 2, [MidpointRounding]::AwayFromZero)
        [void]$ayliklar.Add([pscustomobject]@{
            Baslangic = $ayBaslangici
            Gun = $gunSayisi
            Tutar = $ayTutar
            Grup = if ($ayBaslangici.Year -eq $baslangicYili) { 'ilk' } else { 'ertesi' }
        })
        $ayBaslangici = $ayBaslangici.AddMonths(1)
    }

    foreach ($grup in @('ilk', 'ertesi')) {
        $grupAylar = @($ayliklar | Where-Object { $_.Grup -eq $grup })
        if ($grupAylar.Count -eq 0) { continue }
        $grupToplam = if ($grup -eq 'ilk') { $tutarIlkYil } else { $tutarErtesiYil }
        $yuvarlanmisToplam = ($grupAylar | Measure-Object -Property Tutar -Sum).Sum
        $grupAylar[-1].Tutar = [math]::Round($grupAylar[-1].Tutar + $grupToplam - $yuvarlanmisToplam, 2, [MidpointRounding]::AwayFromZero)
    }

    foreach ($grup in @('ilk', 'ertesi')) {
        $grupAylar = @($ayliklar | Where-Object { $_.Grup -eq $grup })
        if ($grupAylar.Count -eq 0) { continue }
        if ($grup -eq 'ilk') {
            Write-Host "İlk Yıl ($baslangicYili):" -ForegroundColor Cyan
        } else {
            Write-Host "Ertesi Yıl ($($baslangicYili + 1)) :" -ForegroundColor Cyan
        }
        foreach ($aylik in $grupAylar) {
            $ayAd = $aylik.Baslangic.ToString('MMMM', $trCulture)
            $etiket = ("{0} {1}" -f $ayAd, $aylik.Baslangic.Year).PadRight(40)
            $tutarMetni = $aylik.Tutar.ToString('N2', $trCulture)
            Write-Host "${etiket}: $tutarMetni ( $($aylik.Gun) gün )"
        }
    }
    Write-Host ""
}

# --------------------------------------------
# 🚀  Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    # Tarihleri oku
    $baslangicTarihi = Read-Date -Prompt "Başlangıç Tarihi (GG.AA.YYYY formatında, örn: 04.06.2025)"
    do {
        $bitisTarihi = Read-Date -Prompt "Bitiş Tarihi (GG.AA.YYYY formatında, örn: 31.12.2026)    "
        if ($bitisTarihi -lt $baslangicTarihi) {
            Write-Host "Bitiş tarihi başlangıç tarihinden küçük olamaz! Tekrar girin." -ForegroundColor Red
        }
    } until ($bitisTarihi -ge $baslangicTarihi)

    # Tutarı oku
    $parsedTutar = Get-ValidDoubleInput -Prompt "Poliçe Tutarı (örn: 10000,50 veya 10000.50)           " -Min 0

    Write-Host "Hesaplanıyor..." -ForegroundColor Magenta
    Start-Sleep -Milliseconds 1000

    # Hesaplamayı yap
    Calculate-PoliçeMuhasebeleştirme -BaslangicTarihi $baslangicTarihi -BitisTarihi $bitisTarihi -Tutar $parsedTutar
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  🗓️ Sigorta Poliçesi Muhasebeleştirme" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Poliçe tutarını ilk yıl için aylara dağıtır." -ForegroundColor Yellow
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
