# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

# 🚗 Araç Alımı için KDV, ÖTV, Amortisman ve KKEG Hesaplama Aracı
# Modüler, ayarları JSON dosyasında saklar

# region Tanımlar ve Ayarlar
$ayarDosyasi = Join-Path (Get-MFTModuleDataPath -ModuleId 'binekotoalim' -Category Config) 'binekotoalım_config.json'

# Varsayılan ayarlar (2025 yılı için)
$varsayilanAyarlar = @{
    AmortismanSiniri_Maliyetli = 2100000 
    AmortismanSiniri_Net = 1100000      
    OtvKdvGiderSiniri = 990000         
    AmortismanOrani = 0.20             
    YilSayisi = 5                      
}
#endregion

# --------------------------------------------
# 🚀 region Ayar Fonksiyonları
# --------------------------------------------
function Save-Settings {
    param ([PSCustomObject]$Ayarlar)
    try {
        $Ayarlar | ConvertTo-Json -Depth 3 | Set-Content -Encoding UTF8 $ayarDosyasi
        Write-Host "Ayarlar başarıyla kaydedildi: $($ayarDosyasi)" -ForegroundColor Green
    }
    catch {
        Write-Error "Ayarlar kaydedilemedi: $_"
    }
}

# --------------------------------------------
# 🚀 Ayarları Yükleme Fonksiyonu
# --------------------------------------------
function Load-Settings {
    if (Test-Path $ayarDosyasi) {
        try {
            $json = Get-Content $ayarDosyasi -Raw -Encoding UTF8
            $loadedSettings = $json | ConvertFrom-Json
            
            $settingsModified = $false
            foreach ($key in $varsayilanAyarlar.Keys) {
                if (-not $loadedSettings.PSObject.Properties.Name.Contains($key)) {
                    $loadedSettings | Add-Member -MemberType NoteProperty -Name $key -Value $varsayilanAyarlar.$key -Force
                    $settingsModified = $true
                }
            }

            if ($settingsModified) {
                Save-Settings $loadedSettings
            }
            return $loadedSettings
        }
        catch {
            Write-Warning "Ayar dosyası bozuk. Varsayılan ayarlar yükleniyor."
            Save-Settings $varsayilanAyarlar
            return $varsayilanAyarlar
        }
    }
    else {
        Write-Host "Ayar dosyası bulunamadı, varsayılan ayarlar kullanılıyor." -ForegroundColor Yellow
        Save-Settings $varsayilanAyarlar
        return $varsayilanAyarlar
    }
}
#endregion

# --------------------------------------------
# 🚀 region Yardımcı Fonksiyonlar
# --------------------------------------------
function Format-Currency {
    param([decimal]$Value, [string]$ForegroundColor)
    $formatted = "{0:N2} ₺" -f $Value
    if ($ForegroundColor) {
        Write-Host $formatted -ForegroundColor $ForegroundColor -NoNewline
        ""
    } else {
        $formatted
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
        $input = Read-Host

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
#endregion

# --------------------------------------------
# 🚀 Hesaplama Fonksiyonları
# --------------------------------------------
function Calculate-WithCostsAdded {
    param (
        [decimal]$NetBedel,
        [decimal]$Otv,
        [decimal]$Kdv,
        [pscustomobject]$Ayarlar,
        [int]$IktisapAyi
    )

    $toplamMaliyet = $NetBedel + $Otv + $Kdv
    $limit = $Ayarlar.AmortismanSiniri_Maliyetli
    $amortismanOrani = $Ayarlar.AmortismanOrani
    $yilSayisi = $Ayarlar.YilSayisi
    
    # KKEG hesaplama (toplam maliyet - amortisman limiti)
    $kkegToplam = [math]::Max(0, $toplamMaliyet - $limit)
    
    # Amortisman hesaplama
    $amortismanlar = @()
    $kalanDeger = $toplamMaliyet
    $kalanLimit = $limit
    $currentYear = [DateTime]::Now.Year
    
    # İlk yıl kısmi amortisman
    $kalanAy = 13 - $IktisapAyi
    $ilkYilOrani = ($amortismanOrani / 12) * $kalanAy
    
    # İlk yıl
    $brutAmortisman = [math]::Round($kalanDeger * $ilkYilOrani, 2)
    $azamiAmortisman = [math]::Round($kalanLimit * $ilkYilOrani, 2)
    $kkegAmortisman = [math]::Max(0, $brutAmortisman - $azamiAmortisman)
    
    $amortismanlar += [pscustomobject]@{
        Yil = $currentYear
        BrutAmortisman = $brutAmortisman
        AzamiAmortisman = $azamiAmortisman
        KKEGAmortisman = $kkegAmortisman
    }
    
    $kalanDeger -= $brutAmortisman
    $kalanLimit -= $azamiAmortisman
    
    # Ara yıllar (normal amortisman)
    for ($i = 1; $i -lt $yilSayisi - 1; $i++) {
        $brutAmortisman = [math]::Round($toplamMaliyet * $amortismanOrani, 2)
        $azamiAmortisman = [math]::Round($limit * $amortismanOrani, 2)
        $kkegAmortisman = [math]::Max(0, $brutAmortisman - $azamiAmortisman)
        
        $amortismanlar += [pscustomobject]@{
            Yil = $currentYear + $i
            BrutAmortisman = $brutAmortisman
            AzamiAmortisman = $azamiAmortisman
            KKEGAmortisman = $kkegAmortisman
        }
        
        $kalanDeger -= $brutAmortisman
        $kalanLimit -= $azamiAmortisman
    }
    
    # Son yıl (kalan tutar)
    if ($yilSayisi -gt 1) {
        $brutAmortisman = $kalanDeger
        $azamiAmortisman = $kalanLimit
        $kkegAmortisman = [math]::Max(0, $brutAmortisman - $azamiAmortisman)
        
        $amortismanlar += [pscustomobject]@{
            Yil = $currentYear + $yilSayisi - 1
            BrutAmortisman = $brutAmortisman
            AzamiAmortisman = $azamiAmortisman
            KKEGAmortisman = $kkegAmortisman
        }
    }

    # Kümülatif değerleri hesapla
    $cumAzami = 0
    $cumBrut = 0
    $cumKKEG = 0
    
    foreach ($amort in $amortismanlar) {
        $cumAzami += $amort.AzamiAmortisman
        $cumBrut += $amort.BrutAmortisman
        $cumKKEG += $amort.KKEGAmortisman
        
        $amort | Add-Member -NotePropertyName CumAzami -NotePropertyValue $cumAzami
        $amort | Add-Member -NotePropertyName CumBrut -NotePropertyValue $cumBrut
        $amort | Add-Member -NotePropertyName CumKKEG -NotePropertyValue $cumKKEG
    }
    $ilkYilKKEG = $amortismanlar[0].KKEGAmortisman
    return [pscustomobject]@{
        ToplamAracBedeli = $toplamMaliyet
        AmortismanaTabiTutar = $limit
        KKEG_AmortismanSiniri = $kkegToplam
	ToplamKKEG_IlkYil = $ilkYilKKEG
        YillikAmortismanlar = $amortismanlar
    }
}

# --------------------------------------------
# 🚀 Calculate-WithCostsExpensed Fonksiyonu
# --------------------------------------------
function Calculate-WithCostsExpensed {
    param (
        [decimal]$NetBedel,
        [decimal]$Otv,
        [decimal]$Kdv,
        [pscustomobject]$Ayarlar,
        [int]$IktisapAyi
    )

    $giderToplami = $Otv + $Kdv
    $toplamMaliyet = $NetBedel
    $limit = $Ayarlar.AmortismanSiniri_Net
    $giderSiniri = $Ayarlar.OtvKdvGiderSiniri
    
    # KKEG hesaplamaları
    $kkegOtvKdv = [math]::Round([math]::Max(0, $giderToplami - $giderSiniri), 2)
    $kkegAmortismanToplam = [math]::Round([math]::Max(0, $toplamMaliyet - $limit), 2)
    
    # Amortisman hesaplama
    $amortismanlar = @()
    $currentYear = [DateTime]::Now.Year
    $amortismanOrani = $Ayarlar.AmortismanOrani
    
    # İlk yıl kısmi amortisman (12. ay için 1 ay)
    $kalanAy = 13 - $IktisapAyi
    $ilkYilOrani = ($amortismanOrani / 12) * $kalanAy
    
    # İlk yıl (2025)
    $brutAmortisman = [math]::Round($toplamMaliyet * $ilkYilOrani, 2)
    $azamiAmortisman = [math]::Round($limit * $ilkYilOrani, 2)
    $kkegAmortisman = [math]::Round($brutAmortisman - $azamiAmortisman, 2)
    
    $amortismanlar += [pscustomobject]@{
        Yil = $currentYear
        BrutAmortisman = $brutAmortisman
        AzamiAmortisman = $azamiAmortisman
        KKEGAmortisman = $kkegAmortisman
    }
    
    # Ara yıllar (2026-2028)
    for ($i = 1; $i -le 3; $i++) {
        $brutAmortisman = [math]::Round($toplamMaliyet * $amortismanOrani, 2)
        $azamiAmortisman = [math]::Round($limit * $amortismanOrani, 2)
        $kkegAmortisman = [math]::Round($brutAmortisman - $azamiAmortisman, 2)
        
        $amortismanlar += [pscustomobject]@{
            Yil = $currentYear + $i
            BrutAmortisman = $brutAmortisman
            AzamiAmortisman = $azamiAmortisman
            KKEGAmortisman = $kkegAmortisman
        }
    }
    
    # Son yıl (2029) - Kalan tutar
    $toplamBrut = ($amortismanlar | Measure-Object -Property BrutAmortisman -Sum).Sum
    $toplamAzami = ($amortismanlar | Measure-Object -Property AzamiAmortisman -Sum).Sum
    
    $brutAmortisman = [math]::Round($toplamMaliyet - $toplamBrut, 2)
    $azamiAmortisman = [math]::Round($limit - $toplamAzami, 2)
    $kkegAmortisman = [math]::Round($brutAmortisman - $azamiAmortisman, 2)
    
    $amortismanlar += [pscustomobject]@{
        Yil = $currentYear + 4
        BrutAmortisman = $brutAmortisman
        AzamiAmortisman = $azamiAmortisman
        KKEGAmortisman = $kkegAmortisman
    }

# Kümülatif değerleri hesapla
$cumAzami = 0
$cumBrut = 0
$cumKKEG = 0

foreach ($amort in $amortismanlar) {
    $cumAzami = [math]::Round($cumAzami + $amort.AzamiAmortisman, 2)
    $cumBrut = [math]::Round($cumBrut + $amort.BrutAmortisman, 2)
    $cumKKEG = [math]::Round($cumKKEG + $amort.KKEGAmortisman, 2)
    
    $amort | Add-Member -NotePropertyName CumAzami -NotePropertyValue $cumAzami
    $amort | Add-Member -NotePropertyName CumBrut -NotePropertyValue $cumBrut
    $amort | Add-Member -NotePropertyName CumKKEG -NotePropertyValue $cumKKEG
}

# İlk yıl KKEG değeri güvenli şekilde alınır
$ilkYilKKEG = if ($amortismanlar.Count -gt 0) { $amortismanlar[0].KKEGAmortisman } else { 0 }

return [pscustomobject]@{
    ToplamAracBedeli = [math]::Round($toplamMaliyet + $giderToplami, 2)
    AmortismanaTabiTutar = $limit
    OtvKdvGiderYazilabilir = [math]::Round([math]::Min($giderToplami, $giderSiniri), 2)
    OtvKdvKKEG = $kkegOtvKdv
    KKEG_AmortismanSiniri = $kkegAmortismanToplam
    ToplamKKEG = [math]::Round($kkegOtvKdv + $kkegAmortismanToplam, 2)
    ToplamKKEG_IlkYil = $ilkYilKKEG
    YillikAmortismanlar = $amortismanlar
}

}
#endregion

# --------------------------------------------
# 🚀 Görüntüleme Fonksiyonları
# --------------------------------------------
function Write-Title {
    Clear-Host
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host "🚗 --- Araç Alımı için KDV, ÖTV, Amortisman ve KKEG Hesaplama Aracı ---" -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Yasal limitler güncellenmiş olabilir." -ForegroundColor Yellow
    Write-Host "📌 Kullandığınız yıla ait limitleri kontrol ediniz!" -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Show-ResultSummary Fonksiyonu
# --------------------------------------------
function Show-ResultSummary {
    param ([PSCustomObject]$Result, [string]$Scenario)
    
    Write-Host "  Toplam Araç Bedeli (Nakit Çıkışı)    : $(Format-Currency $Result.ToplamAracBedeli)"
    
    if ($Scenario -eq "Maliyetli") {
        Write-Host "  Amortismana Esas Tutar Limit         : $(Format-Currency $Result.AmortismanaTabiTutar)"
        if ($Result.KKEG_AmortismanSiniri -gt 0) {
            Write-Host "  KKEG Olarak Dikkate Alınacak Tutar   : $(Format-Currency $Result.KKEG_AmortismanSiniri)" -ForegroundColor Red
        }
    } 
    else {
        Write-Host "  Amortismana Esas Tutar Limit         : $(Format-Currency $Result.AmortismanaTabiTutar)"
        Write-Host "  KDV+ÖTV'den Gider Yazılabilir Tutar  : $(Format-Currency $Result.OtvKdvGiderYazilabilir)"
        Write-Host "  KDV+ÖTV'den KKEG Olan Tutar          : $(Format-Currency $Result.OtvKdvKKEG)" -ForegroundColor Red
        if ($Result.KKEG_AmortismanSiniri -gt 0) {
            Write-Host "  Net Bedel Amortisman Sınırı Aşımı    : $(Format-Currency $Result.KKEG_AmortismanSiniri)" -ForegroundColor Red
        }
        Write-Host "  İlk Yıl Toplam KKEG                  : $(Format-Currency $Result.ToplamKKEG_IlkYil)" -ForegroundColor Red
    }
}

# --------------------------------------------
# 🚀 Amortisman Tablosu
# --------------------------------------------
function Show-AmortizationTable {
    param ([PSCustomObject]$Result)
    
    $table = foreach ($year in $Result.YillikAmortismanlar) {
        [PSCustomObject]@{
            'Yıl' = $year.Yil
            'Yıllık Amortisman' = Format-Currency $year.BrutAmortisman
            'Azami Amortisman' = Format-Currency $year.AzamiAmortisman
            'KKEG Amortisman' = if ($year.KKEGAmortisman -gt 0) { 
                Format-Currency $year.KKEGAmortisman
            } else { "-" }
            'Küm. Azami' = Format-Currency $year.CumAzami
            'Küm. Toplam' = Format-Currency $year.CumBrut
            'Küm. KKEG' = if ($year.CumKKEG -gt 0) { 
                Format-Currency $year.CumKKEG
            } else { "-" }
        }
    }
    
    # Toplam satırı ekle
    $lastYear = $Result.YillikAmortismanlar[-1]
    $table += [PSCustomObject]@{
        'Yıl' = "Toplam"
        'Yıllık Amortisman' = Format-Currency $lastYear.CumBrut
        'Azami Amortisman' = Format-Currency $lastYear.CumAzami
        'KKEG Amortisman' = if ($lastYear.CumKKEG -gt 0) { 
            Format-Currency $lastYear.CumKKEG
        } else { "-" }
        'Küm. Azami' = "-"
        'Küm. Toplam' = "-"
        'Küm. KKEG' = "-"
    }
    
    $table | Format-Table -AutoSize
}
#endregion

# --------------------------------------------
# 🚀 Show-CalculationMenu Fonksiyonu 
# --------------------------------------------
function Show-CalculationMenu {
    param ([PSCustomObject]$Ayarlar)
    
    Write-Title
    Write-Host "Hesaplama için araç bilgilerini girin."
    
    $netBedel = Get-ValidDoubleInput -Prompt "Net Bedel (Vergisiz Araç Bedeli - TL)  " -Min 0
    $otvTutar = Get-ValidDoubleInput -Prompt "ÖTV Tutarı (TL)                        " -Min 0
    $kdvTutar = Get-ValidDoubleInput -Prompt "KDV Tutarı (TL)                        " -Min 0
    
    [int]$iktisapAyi = 0
    do {
        Write-Host "🔹 İktisap Ayı (1-12)                     : " -ForegroundColor Cyan -NoNewline
        $input = Read-Host
    } while (-not ([int]::TryParse($input, [ref]$iktisapAyi) -and $iktisapAyi -ge 1 -and $iktisapAyi -le 12))

    $sonuclar = @{
        Maliyetli = Calculate-WithCostsAdded -NetBedel $netBedel -Otv $otvTutar -Kdv $kdvTutar -Ayarlar $Ayarlar -IktisapAyi $iktisapAyi
        Giderli   = Calculate-WithCostsExpensed -NetBedel $netBedel -Otv $otvTutar -Kdv $kdvTutar -Ayarlar $Ayarlar -IktisapAyi $iktisapAyi
    }

    Write-Host "`n------ HESAPLAMA SONUÇLARI ------" -ForegroundColor Green

    Write-Host "`n👉 KDV + ÖTV Maliyetlere Eklenmiş Durum (GVK md. 40/1)" -ForegroundColor Yellow
    Show-ResultSummary -Result $sonuclar.Maliyetli -Scenario "Maliyetli"
    Show-AmortizationTable -Result $sonuclar.Maliyetli

    Write-Host "`n👉 KDV + ÖTV Gider Yazılan Durum (GVK md. 40/7)" -ForegroundColor Yellow
    Show-ResultSummary -Result $sonuclar.Giderli -Scenario "Giderli"
    Show-AmortizationTable -Result $sonuclar.Giderli

    Read-Host "`nDevam etmek için ENTER'a basın..."
}

# --------------------------------------------
# 🚀 Ayarlar Menüsü
# --------------------------------------------
function Show-SettingsMenu {
    param ([PSCustomObject]$Ayarlar)
    
    do {
        Clear-Host
        Write-Host ("-" * 72) -ForegroundColor Magenta
        Write-Host "⚙️ --- AYARLAR ---" -ForegroundColor Magenta
        Write-Host ("-" * 72) -ForegroundColor Magenta

        Write-Host "Mevcut Ayarlar:"
        Write-Host ""
        Write-Host ("-" * 72) -ForegroundColor DarkGray
        Write-Host ("  1. 💰 {0,-14}  Amortisman sınırı (KDV+ÖTV maliyete eklendiyse)" -f (Format-Currency $Ayarlar.AmortismanSiniri_Maliyetli))
        Write-Host ("  2. 💰 {0,-14}  Amortisman sınırı (KDV+ÖTV gider yazıldıysa)" -f (Format-Currency $Ayarlar.AmortismanSiniri_Net))
        Write-Host ("  3. 🧾 {0,-14}  ÖTV ve KDV'den gider yazma sınırı" -f (Format-Currency $Ayarlar.OtvKdvGiderSiniri))
        Write-Host ("  4. 📊 {0,-14}  Amortisman oranı" -f ('{0}%' -f [math]::Round($Ayarlar.AmortismanOrani*100, 2)))
        Write-Host ("  5. 📅 {0,-14}  Amortisman süresi" -f ("$($Ayarlar.YilSayisi) yıl"))
        Write-Host "  Q  ↩️  Ana Menü"
        Write-Host ""
        Write-Host ("-" * 72) -ForegroundColor DarkGray
        
        $secim = Read-Host "Seçiminiz"
        
        switch ($secim) {
            "1" { $Ayarlar.AmortismanSiniri_Maliyetli = Get-ValidDoubleInput  "Yeni sınır (TL) " }
            "2" { $Ayarlar.AmortismanSiniri_Net = Get-ValidDoubleInput  "Yeni sınır (TL) " }
            "3" { $Ayarlar.OtvKdvGiderSiniri = Get-ValidDoubleInput  "Yeni sınır (TL) " }
            "4" { 
                $yeniOran = Get-ValidDoubleInput  "Yeni oran (%) "
                $Ayarlar.AmortismanOrani = $yeniOran / 100
            }
            "5" {
                do {
                $yeniYilStr = Read-Host "Yeni süre (yıl) (1-30 arası bir sayı girin)"
                if ($yeniYilStr -match '^\d+$') {
                $yeniYil = [int]$yeniYilStr
                } else {
                $yeniYil = 0
                }
                } while ($yeniYil -lt 1 -or $yeniYil -gt 30)
    
                $Ayarlar.YilSayisi = $yeniYil
            }

            "Q" { return }
            default { Write-Host "[!] Geçersiz seçim. Lütfen yukarıdaki seçeneklerden birini girin (1, 5 veya Q)." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
        
        if ($secim -match "^[1-5]$") {
            Save-Settings $Ayarlar
        }
    } while ($true)
}

# --------------------------------------------
# 🚀 Ana Menü
# --------------------------------------------
function Show-Menu {
    while ($true) {
        $ayarlar = Load-Settings
        
        Write-Title
        Write-Host ("-" * 72) -ForegroundColor DarkGray
        Write-Host "  1. 🧮  Hesaplama Yap"
        Write-Host "  2. ⚙️  Ayarlar"
        Write-Host "  Q  🚪  Çıkış"
        Write-Host ""
        Write-Host ("-" * 72) -ForegroundColor DarkGray
        
        $secim = Read-Host "Seçiminiz"
        
        switch ($secim) {
            "1" { Show-CalculationMenu -Ayarlar $ayarlar }
            "2" { Show-SettingsMenu -Ayarlar $ayarlar }
            "Q" { exit }
            default { Write-Host "[!] Geçersiz seçim. Lütfen yukarıdaki seçeneklerden birini girin (1, 2 veya Q)." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    }
}
#endregion

# Ana Program
Show-Menu
