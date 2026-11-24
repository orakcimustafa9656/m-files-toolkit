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
# 🚀  Kullanıcıdan maksimum uzunluk kontrolü ile veri alma fonksiyonu
# --------------------------------------------
function Get-Input {
    param(
        [string]$Label,
        [int]   $MaxLength
    )

    $PadLength   = 30
    $LabelPadded = "{0,-$PadLength}" -f $Label

    Write-Host $LabelPadded -ForegroundColor Magenta -NoNewline
    Write-Host ":" -ForegroundColor Magenta -NoNewline

    $Input = Read-ToolkitLine

    if ([string]::IsNullOrWhiteSpace($Input)) {
        return '.' * $MaxLength
    }
    elseif ($Input.Length -gt $MaxLength) {
        return $Input.Substring(0, $MaxLength)
    }
    else {
        return $Input
    }
}

# --------------------------------------------
# 🚀 Segment Yazma Fonksiyonu
# --------------------------------------------
function Write-Segment {
    param(
        [string]$Text,
        [switch]$IsVariable
    )
    if ($IsVariable) {
        # Mor renk RGB(128,0,128)
        $Selection.Font.TextColor.RGB = 8388736
    }
    else {
        # Siyah
        $Selection.Font.TextColor.RGB = 0
    }
    $Selection.TypeText($Text)
}

# --------------------------------------------
# 🚀  Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    # Maksimum uzunlukları belirle
    $MaxLengths = @{ 
        odemeGunu    = 10
        tutarTL      = 10
        tutarKurus   = 2
        no           = 9
        tarih        = 10
        yer          = 20
        sayin        = 42
        yaziIleTutar = 43
        yaziIleKurus = 17
        isim         = 50
        adres        = 50
        tcBorclu     = 11
        kefil        = 50
        tcKefil      = 11
    }

    # Veri Girişi
    $odemeGunu    = Get-Input "Ödeme Günü (gg.aa.yyyy)" $MaxLengths.odemeGunu
    $tutarTL      = Get-Input "Türk Lirası" $MaxLengths.tutarTL
    $tutarKurus   = Get-Input "Kuruş" $MaxLengths.tutarKurus
    $no           = Get-Input "Senet No" $MaxLengths.no
    $tarih        = Get-Input "Düzenleme Tarihi (gg.aa.yyyy)" $MaxLengths.tarih
    $yer          = Get-Input "Düzenleme Yeri" $MaxLengths.yer
    $sayin        = Get-Input "Sayın (Alacaklı)" $MaxLengths.sayin
    $yaziIleTutar = Get-Input "Tutar (Yazıyla)" $MaxLengths.yaziIleTutar
    $yaziIleKurus = Get-Input "Kuruş (Yazıyla)" $MaxLengths.yaziIleKurus
    $isim         = Get-Input "Borçlu İsmi / Unvanı" $MaxLengths.isim
    $adres        = Get-Input "Adres" $MaxLengths.adres
    $tcBorclu     = Get-Input "Borçlu TC/V.D. No" $MaxLengths.tcBorclu
    $kefil        = Get-Input "Kefil Adı" $MaxLengths.kefil
    $tcKefil      = Get-Input "Kefil TC/V.D. No" $MaxLengths.tcKefil

    Clear-Host

    # Konsolda Senet Metni
    Write-Host "`n                            *** BONO (SENET) ***`n"
    Write-Host ("{0,-12} {1,-10} {2,-8} {3,-15}" -f "Ödeme Günü", "TL", "Kuruş", "Senet No")
    Write-Host ("{0,-12} {1,-10} {2,-8} {3,-15}" -f $odemeGunu, $tutarTL, $tutarKurus, $no) -ForegroundColor Magenta
    Write-Host ""

    Write-Host -NoNewline "İş bu emre muharrer senedim... mukabilinde "
    Write-Host $odemeGunu -ForegroundColor Magenta -NoNewline
    Write-Host " tarihinde"
    Write-Host -NoNewline "Sayın "
    Write-Host $sayin -ForegroundColor Magenta -NoNewline
    Write-Host " veyahut emruhavale.......... ,"
    Write-Host -NoNewline "Yukarıda yazılı yalnız "
    Write-Host $yaziIleTutar -ForegroundColor Magenta -NoNewline
    Write-Host " Türk Lirası"
    Write-Host $yaziIleKurus -ForegroundColor Magenta -NoNewline
    Write-Host " Kuruş ödeyeceği ..... Bedeli .......... olarak Ahzolunmuştur."
    Write-Host ""

    Write-Host "İş Bu Bono vadesinde ödenmediği takdirde müteakip bonoların muacceliyet"
    Write-Host "kazanacağını, İhtilaf halinde  " -NoNewline
    Write-Host $sayin -ForegroundColor Magenta
    Write-Host "yerleşim yerindeki mahkemelerin yetkili olduğunu kabul ederim... Okudu..."
    Write-Host ""

    Write-Host "Borçlu Bilgileri"
    Write-Host -NoNewline "İsim / Unvan      : "
    Write-Host $isim -ForegroundColor Magenta
    Write-Host -NoNewline "Adres             : "
    Write-Host $adres -ForegroundColor Magenta
    Write-Host -NoNewline "TC/V.D. No        : "
    Write-Host $tcBorclu -ForegroundColor Magenta
    Write-Host -NoNewline "Düzenleme Tarihi  : "
    Write-Host $tarih -ForegroundColor Magenta
    Write-Host -NoNewline "Düzenleme Yeri    : "
    Write-Host $yer -ForegroundColor Magenta
    Write-Host "İmza              : ___________________________`n"

    Write-Host "Kefil Bilgileri"
    Write-Host -NoNewline "İsim              : "
    Write-Host $kefil -ForegroundColor Magenta
    Write-Host -NoNewline "TC/V.D. No        : "
    Write-Host $tcKefil -ForegroundColor Magenta
    Write-Host "İmza              : ___________________________`n"

    # Word Aktarımı
    $Export = Read-ToolkitLine -Prompt "Word'e aktarılmasını istiyor musunuz? (E/H): "

    if ($Export -eq 'E') {
        $word      = New-Object -ComObject Word.Application
        $doc       = $word.Documents.Add()
        $selection = $word.Selection

        $selection.Font.Name = 'Consolas'
        $selection.Font.Size = 10

        # Write-Segment fonksiyonunda $selection global yapıldı
        function Write-Segment {
            param(
                [string]$Text,
                [switch]$IsVariable
            )
            if ($IsVariable) {
                $selection.Font.TextColor.RGB = 8388736
            }
            else {
                $selection.Font.TextColor.RGB = 0
            }
            $selection.TypeText($Text)
        }

        Write-Segment "                            *** BONO (SENET) ***`r`n`r`n"

        Write-Segment ("{0,-12} {1,-10} {2,-8} {3,-15}`r`n" -f "Ödeme Günü", "TL", "Kuruş", "Senet No")
        Write-Segment ("{0,-12} {1,-10} {2,-8} {3,-15}`r`n`r`n" -f $odemeGunu, $tutarTL, $tutarKurus, $no) -IsVariable

        Write-Segment "İş bu emre muharrer senedim... mukabilinde "
        Write-Segment "$odemeGunu" -IsVariable
        Write-Segment " tarihinde`r`n"
        Write-Segment "Sayın "
        Write-Segment "$sayin" -IsVariable
        Write-Segment " veyahut emruhavale.......... ,`r`n"

        Write-Segment "Yukarıda yazılı yalnız "
        Write-Segment "$yaziIleTutar" -IsVariable
        Write-Segment " Türk Lirası`r`n"
        Write-Segment "$yaziIleKurus" -IsVariable
        Write-Segment " Kuruş ödeyeceği ..... Bedeli .......... olarak Ahzolunmuştur.`r`n`r`n"

        Write-Segment "İş Bu Bono vadesinde ödenmediği takdirde müteakip bonoların muacceliyet `r`n"
        Write-Segment "kazanacağını, İhtilaf halinde  "
        Write-Segment "$sayin" -IsVariable
        Write-Segment "`r`nyerleşim yerindeki mahkemelerin yetkili olduğunu kabul ederim... Okudu...`r`n`r`n"

        Write-Segment "Borçlu Bilgileri`r`n"
        Write-Segment "İsim / Unvan      : "
        Write-Segment "$isim" -IsVariable
        Write-Segment "`r`nAdres             : "
        Write-Segment "$adres" -IsVariable
        Write-Segment "`r`nTC/V.D. No        : "
        Write-Segment "$tcBorclu" -IsVariable
        Write-Segment "`r`nDüzenleme Tarihi  : "
        Write-Segment "$tarih" -IsVariable
        Write-Segment "`r`nDüzenleme Yeri    : "
        Write-Segment "$yer" -IsVariable
        Write-Segment "`r`nİmza              : ___________________________`r`n`r`n"

        Write-Segment "Kefil Bilgileri`r`n"
        Write-Segment "İsim              : "
        Write-Segment "$kefil" -IsVariable
        Write-Segment "`r`nTC/V.D. No        : "
        Write-Segment "$tcKefil" -IsVariable
        Write-Segment "`r`nİmza              : ___________________________`r`n"

        $word.Visible = $true
    }
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  📜 Senet Düzenleme" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Boş Senet için alanları Enter ile boş bırakabilirsiniz." -ForegroundColor Yellow
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
