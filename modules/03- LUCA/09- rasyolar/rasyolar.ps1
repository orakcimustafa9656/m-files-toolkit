# --------------------------------------------
# 🚀  Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
# --- Gerekli Assembly'leri Yükle ---
Add-Type -AssemblyName System.Windows.Forms
Try {
    # Dosya Seçimi için OpenFileDialog nesnesi oluştur
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel ve CSV Dosyaları (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv|Excel Dosyaları (*.xls;*.xlsx)|*.xls;*.xlsx|CSV Dosyaları (*.csv)|*.csv|Tüm Dosyalar (*.*)|*.*"
    $dialog.Title = "Bir Luca Mizan Excel veya CSV dosyası seçin"

    if ($dialog.ShowDialog() -ne 'OK') {
        Write-Host "İşlem iptal edildi." -ForegroundColor Yellow
        $script:selectionCancelled = $true
        return
    }

    $filePath = $dialog.FileName

    if (-not (Test-Path $filePath)) {
        Write-Host "Seçilen dosya sistemde bulunamadı." -ForegroundColor Red
        return
    }

    Write-Host "`nDosya: $filePath`n" -ForegroundColor Cyan
    Write-Host "Veriler okunuyor ve oranlar hesaplanıyor..." -ForegroundColor Green
}
catch {
    Write-Host "Bir hata oluştu: $_" -ForegroundColor Red
    return
}

# --------------------------------------------
# 🚀 Rapor için Renkli Yazdırma Fonksiyonları ---
# --------------------------------------------
function Write-ColoredLine {
    param(
        [string]$Text,
        [ConsoleColor]$Color = 'White',
        [ConsoleColor]$BackColor = 'Black', # Bu parametreyi kaldırmıyoruz, default olarak black kalıyor
        [int]$Width = 0 # Yeni parametre: Çıktının hedeflenen genişliği
    )
    $oldFG = $Host.UI.RawUI.ForegroundColor
    $oldBG = $Host.UI.RawUI.BackgroundColor
    $Host.UI.RawUI.ForegroundColor = $Color
    
    # Yeni: Eğer BackColor belirtilmediyse (yani varsayılan 'Black' olarak kaldıysa), default konsol arka planını kullan
    # Aksi takdirde belirtilen arka plan rengini kullan
    if ($PSBoundParameters.ContainsKey('BackColor')) {
        $Host.UI.RawUI.BackgroundColor = $BackColor
    } else {
        # Eğer BackColor parametresi hiç geçilmediyse, konsolun varsayılan arka planını kullan.
        # Bunun için oldBG'yi geri yüklüyoruz veya direkt ayarlama yapmıyoruz.
        # Basitçe, parametreyi hiç set etmeyince PowerShell default'u kullanacaktır.
        # Ancak fonksiyon içinde bu şekilde yapmak yerine, parametreyi atayarak bunu kontrol edebiliriz.
        # Bu durumda, eğer bu fonksiyon içinde BackColor'ı kullanmak istemiyorsak,
        # onu set etme kısmını kaldırmamız gerekir.
        # En basit yol, 'BackColor' parametresini fonksiyon çağrılarında hiç geçmemektir.
        # Veya, 'BackColor' parametresini kullanmayalım ve varsayılan arka plan oluşsun.
        # Şimdilik, fonksiyon tanımını bu haliyle bırakıp, çağırdığımız yerlerden BackColor'ı kaldırıyorum.
    }


    $outputTxt = $Text
    if ($Width -gt 0) {
        if ($Text.Length -lt $Width) {
            $outputTxt = $Text.PadRight($Width) # Belirtilen genişliğe kadar boşluk doldur
        } elseif ($Text.Length -gt $Width) {
            $outputTxt = $Text.Substring(0, $Width) # Genişlikten fazlaysa kırp
        }
    }

    Write-Host $outputTxt
    $Host.UI.RawUI.ForegroundColor = $oldFG
    $Host.UI.RawUI.BackgroundColor = $oldBG # Varsayılan arka plan rengine dön
}

# --------------------------------------------
# 🚀 Bölüm Başlığı Oluşturucu
# --------------------------------------------
function Write-SectionHeader {
    param([string]$Title)
    # Üstteki "=" çizgisi 100 karakter genişliğinde
    Write-ColoredLine ("`n" + ("=" * 100)) -Color DarkCyan

    $upperTitle = $Title.ToUpper()
    $totalWidth = 100 # Bu, başlığın hedeflenen genişliği
    $titleLength = $upperTitle.Length

    # Başlığı ortalamak için boşluk hesabı
    $padding = $totalWidth - $titleLength
    $leftPad = [int]($padding / 2) # Tam sayıya yuvarlayarak soldaki boşluk
    $rightPad = $padding - $leftPad # Kalan boşluk sağa

    # Ortalanmış metin dizisi
    $centeredTitleString = (" " * $leftPad) + $upperTitle + (" " * $rightPad)

    # Arka plan rengi artık belirtilmiyor, varsayılan konsol arka planı kullanılacak
    Write-ColoredLine $centeredTitleString -Color Cyan -Width $totalWidth

    # Alttaki "=" çizgisi 100 karakter genişliğinde
    Write-ColoredLine (($("=" * 100) + "`n")) -Color DarkCyan
}

# --------------------------------------------
# 🚀 CSV İşlemleri ve Write-Ratio Fonksiyonu
# --------------------------------------------
function Write-Ratio {
    param(
        [string]$Name,
        [double]$Value,
        [string]$Formula,
        [string]$Description,
        [switch]$IsPercentage = $false,
        [int]$Score = $null,
        [string]$Comment = $null,
        [string]$UnavailableReason = $null
    )
    if (-not [string]::IsNullOrWhiteSpace($UnavailableReason)) {
        Write-ColoredLine "$($Name.PadRight(30)) : Hesaplanamadı" -Color Yellow
        Write-ColoredLine "  Formül   : $Formula" -Color Gray
        Write-ColoredLine "  Açıklama : $Description" -Color DarkGray
        Write-ColoredLine "  Neden    : $UnavailableReason" -Color Red
        Write-ColoredLine ("-" * 100) -Color DarkGray -Width 100
        Write-Host ""
        return
    }
    $valStr = if ($IsPercentage) { "{0:N2}%" -f $Value } else { "{0:N2}" -f $Value }

    # Oran Adı ve Değerini daha belirgin ve hizalı yazdır
    Write-ColoredLine ("$($Name.PadRight(30)) : $valStr") -Color Yellow

    # Detayları hizala
    Write-ColoredLine ("  Formül   : $Formula") -Color Gray
    Write-ColoredLine ("  Açıklama : $Description") -Color DarkGray

    # Puan ve Yorumu göster, puana göre renk seç
    if ($Score -ne $null -and $Comment -ne $null) {
        $scoreColor = 'Red' # Zayıf
        if ($Score -eq 5) { $scoreColor = 'Green' } # Çok İyi
        elseif ($Score -eq 3) { $scoreColor = 'Yellow' } # Orta
        
        Write-ColoredLine ("  Puan     : $Score") -Color $scoreColor
        Write-ColoredLine ("  Yorum    : $Comment") -Color DarkGreen
    }
    # Oranlar arasına ayraç (100 karakter genişliğinde)
    Write-ColoredLine ("-" * 100) -Color DarkGray -Width 100
    Write-Host ""
}

# --------------------------------------------
# 🚀 Dosya Import ve Veri İşleme
# --------------------------------------------

$headers = @('HESAP KODU','HESAP ADI','Bos1','BORÇ','ALACAK','BORÇ BAKİYESİ','ALACAK BAKİYESİ')

try {
    $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
    if ($extension -eq '.csv') {
        # Header satırını otomatik olarak almamak için -Header ile manuel olarak belirlendi.
        # İlk 6 satırı atlamak için Select-Object -Skip 6 kullanıldı.
        $data = @(Import-Csv -Path $filePath -Delimiter ';' -Header $headers -Encoding Default | Select-Object -Skip 6)
    }
    elseif ($extension -eq '.xls' -or $extension -eq '.xlsx') {
        # Excel dosyaları Windows'ta kurulu Microsoft Excel üzerinden okunur.
        $excel = $null
        $workbook = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($filePath, 0, $true)
            $worksheet = $workbook.Worksheets.Item(1)
            $usedRange = $worksheet.UsedRange
            $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1
            $data = @()
            for ($row = 7; $row -le $lastRow; $row++) {
                $item = [ordered]@{}
                for ($column = 1; $column -le $headers.Count; $column++) {
                    $item[$headers[$column - 1]] = [string]$worksheet.Cells.Item($row, $column).Text
                }
                $data += [pscustomobject]$item
            }
        }
        finally {
            if ($workbook) { $workbook.Close($false) | Out-Null }
            if ($excel) { $excel.Quit() }
            if ($worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
            if ($usedRange) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) }
            if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
    else {
        throw "Desteklenmeyen dosya biçimi: $extension"
    }
}
catch {
    Write-Host "Dosya okunurken hata oluştu: $_" -ForegroundColor Red
    return
}

$kodlar = @{
    1 = 6; 10 = 6; 11 = 6; 12 = 6; 15 = 6; 2 = 6; # Aktif hesaplar (Borç Bakiyesi)
    3 = 7; 32 = 7; 4 = 7; 5 = 7; 60 = 7; 61 = 6; 62 = 6 # Pasif/Gider/Gelir hesaplar (Alacak Bakiyesi veya özel durumlar)
}

$basliklar = @(
    @{ Ad = "1- Dönen Varlıklar"; Kod = 1 },
    @{ Ad = "10- Hazır Değerler"; Kod = 10 },
    @{ Ad = "11- Menkul Kıymetler"; Kod = 11 },
    @{ Ad = "12- Ticari Alacaklar"; Kod = 12 },
    @{ Ad = "15- Stoklar"; Kod = 15 },
    @{ Ad = "2- Duran Varlıklar"; Kod = 2 },
    @{ Ad = "3- Kısa Vadeli YabancıKaynaklar"; Kod = 3 },
    @{ Ad = "32- Ticari Borçlar"; Kod = 32 },
    @{ Ad = "4- Uzun Vadeli Yabancı Kaynaklar"; Kod = 4 },
    @{ Ad = "5- Öz Kaynaklar"; Kod = 5 },
    @{ Ad = "60- Brüt Satışlar"; Kod = 60 },
    @{ Ad = "61- Satış İndirimleri"; Kod = 61 },
    @{ Ad = "62- Satışların Maliyeti"; Kod = 62 }
)

$tutarlar = @{}
$bulunanKodlar = @{}

foreach ($baslik in $basliklar) {
    $kod = $baslik.Kod
    # ToString() kullanarak Hesap Kodu sütunu ile eşleşmeyi sağla
    $satir = $data | Where-Object { $_.'HESAP KODU' -eq $kod.ToString() }
    if ($satir) {
        $bulunanKodlar[$kod.ToString()] = $true
        $sutunIndex = $kodlar[$kod] # 6: BORÇ BAKİYESİ, 7: ALACAK BAKİYESİ
        $columnName = if ($sutunIndex -eq 6) { 'BORÇ BAKİYESİ' } else { 'ALACAK BAKİYESİ' }
        $tutar = $satir.$columnName

        if ([string]::IsNullOrEmpty($tutar)) {
            $tutarNum = 0
        }
        else {
            # Nokta yerine virgül kullanılarak küsurat ayrımı yapılıyorsa düzelt
            $tutarNum = [double]($tutar -replace '\.', '' -replace ',', '.')
        }
        $tutarlar[$kod.ToString()] = $tutarNum
    }
    else {
        $tutarlar[$kod.ToString()] = 0
    }
}

# Gerekli ek toplamlar
$tutarlar['ToplamBorclar'] = $tutarlar['3'] + $tutarlar['4']
$tutarlar['ToplamAktifler'] = $tutarlar['1'] + $tutarlar['2']
$tutarlar['DevamliSermaye'] = $tutarlar['5'] + $tutarlar['4']


# --- ORANLARIN HESAPLANMASI ---

$cariOran = if ($tutarlar['3'] -ne 0) { $tutarlar['1'] / $tutarlar['3'] } else { 0 }
$asitTest = if ($tutarlar['3'] -ne 0) { ($tutarlar['1'] - $tutarlar['15']) / $tutarlar['3'] } else { 0 }
$nakitOrani = if ($tutarlar['3'] -ne 0) { ($tutarlar['10'] + $tutarlar['11']) / $tutarlar['3'] } else { 0 }
$stokBagimlilik = if ($tutarlar['1'] -ne 0) { $tutarlar['15'] / $tutarlar['1'] } else { 0 }

$kaldirac = if ($tutarlar['ToplamAktifler'] -ne 0) { $tutarlar['ToplamBorclar'] / $tutarlar['ToplamAktifler'] } else { 0 }
$finansman = if ($tutarlar['ToplamBorclar'] -ne 0) { $tutarlar['5'] / $tutarlar['ToplamBorclar'] } else { 0 }
$borcOzsermaye = if ($tutarlar['5'] -ne 0) { $tutarlar['ToplamBorclar'] / $tutarlar['5'] } else { 0 }
$otofinansman = if ($tutarlar['ToplamAktifler'] -ne 0) { $tutarlar['5'] / $tutarlar['ToplamAktifler'] } else { 0 }

$brutSatis = $tutarlar['60']
$satisIndirim = $tutarlar['61']
$satisMaliyeti = $tutarlar['62']
$brutKarMarji = 0
if ($brutSatis -ne 0) {
    # Satış indirimleri ve maliyeti borç bakiyesi, brüt satışlar alacak bakiyesi alınır.
    # Bu nedenle 61 ve 62 kodlu hesapların tutarlarını alırken borç bakiyesi alınmalıdır.
    # Buradaki hesaplama zaten doğru. ((60 - 61 - 62) / 60) * 100
    $brutKarMarji = (($brutSatis - $satisIndirim - $satisMaliyeti) / $brutSatis) * 100
}

# Temel hesap kalemleri yoksa veya bölme işleminin paydası sıfırsa oran hesaplanamaz.
# Alt hesap grupları (ör. 11 Menkul Kıymetler, 61 Satış İndirimleri) bulunmadığında sıfır kabul edilir.
$kodAdlari = @{
    '1' = '1 Dönen Varlıklar'; '3' = '3 Kısa Vadeli Yabancı Kaynaklar'
    '5' = '5 Öz Kaynaklar'; '10' = '10 Hazır Değerler'; '15' = '15 Stoklar'
    '60' = '60 Brüt Satışlar'; '62' = '62 Satışların Maliyeti'
}
function Get-RatioUnavailableReason {
    param([string[]]$RequiredCodes, [double]$Denominator)
    $missing = @($RequiredCodes | Where-Object { -not $bulunanKodlar.ContainsKey($_) })
    if ($missing.Count -gt 0) {
        $names = @($missing | ForEach-Object { $kodAdlari[$_] })
        return "Mizanda gerekli hesap kalemi bulunamadı: $($names -join ', ')."
    }
    if ($Denominator -eq 0) { return 'Formüldeki bölen (payda) sıfır.' }
    return $null
}
$oranDurumlari = @{
    CariOran = Get-RatioUnavailableReason @('1','3') $tutarlar['3']
    AsitTest = Get-RatioUnavailableReason @('1','15','3') $tutarlar['3']
    NakitOrani = Get-RatioUnavailableReason @('10','3') $tutarlar['3']
    StokBagimlilik = Get-RatioUnavailableReason @('1','15') $tutarlar['1']
    Kaldirac = Get-RatioUnavailableReason @('1','3') $tutarlar['ToplamAktifler']
    Finansman = Get-RatioUnavailableReason @('5','3') $tutarlar['ToplamBorclar']
    BorcOzsermaye = Get-RatioUnavailableReason @('5','3') $tutarlar['5']
    Otofinansman = Get-RatioUnavailableReason @('5','1') $tutarlar['ToplamAktifler']
    BrutKar = Get-RatioUnavailableReason @('60','62') $brutSatis
}

# --------------------------------------------
# 🚀 PUANLAMA ve YORUMLAR Fonksiyonu ---
# --------------------------------------------
function Get-FinancialRatiosScore {
    param(
        [double]$cariOran,
        [double]$asitTest,
        [double]$nakitOrani,
        [double]$stokBagimlilik,
        [double]$kaldirac,
        [double]$finansman,
        [double]$borcOzsermaye,
        [double]$otofinansman,
        [double]$brutKarMarji
    )

    $results = @{}

# --------------------------------------------
# 🚀 Puanlama ve yorum için yardımcı fonksiyon
# --------------------------------------------
    function Score-Comment {
        param(
            [double]$value,
            [double]$goodThreshold,
            [double]$acceptableThreshold,
            [bool]$inverse, # Küçük değer daha iyiyse $true
            [string]$goodComment,
            [string]$acceptableComment,
            [string]$poorComment
        )
        $score = 0
        $comment = ""

        if ($inverse) { # Küçük değer daha iyidir (örn: Stok Bağımlılık, Kaldıraç)
            if ($value -le $goodThreshold) {
                $score = 5 # Çok İyi
                $comment = $goodComment
            } elseif ($value -le $acceptableThreshold) {
                $score = 3 # Orta
                $comment = $acceptableComment
            } else {
                $score = 1 # Zayıf
                $comment = $poorComment
            }
        } else { # Büyük değer daha iyidir (örn: Cari Oran, Brüt Kar Marjı)
            if ($value -ge $goodThreshold) {
                $score = 5 # Çok İyi
                $comment = $goodComment
            } elseif ($value -ge $acceptableThreshold) {
                $score = 3 # Orta
                $comment = $acceptableComment
            } else {
                $score = 1 # Zayıf
                $comment = $poorComment
            }
        }
        return @{ Puan = $score; Yorum = $comment }
    }

    # Cari Oran (Büyük değer iyi)
    $results.CariOran = Score-Comment -value $cariOran -goodThreshold 2 -acceptableThreshold 1.5 -inverse $false `
        -goodComment "Likidite çok iyi, kısa vadeli borçlar rahat karşılanabilir." `
        -acceptableComment "Likidite yeterli düzeyde, tedbirli olunmalı." `
        -poorComment "Likidite düşük, borç ödeme riski var."

    # Asit Test Oranı (Büyük değer iyi)
    $results.AsitTest = Score-Comment -value $asitTest -goodThreshold 1 -acceptableThreshold 0.75 -inverse $false `
        -goodComment "Likidite çok iyi, stoklara bağlılık az." `
        -acceptableComment "Orta düzeyde likidite, stok hariç dikkatli olunmalı." `
        -poorComment "Likidite zayıf, stoklara aşırı bağımlılık var."

    # Nakit Oranı (Büyük değer iyi)
    $results.NakitOrani = Score-Comment -value $nakitOrani -goodThreshold 0.5 -acceptableThreshold 0.25 -inverse $false `
        -goodComment "Nakit pozisyonu çok güçlü." `
        -acceptableComment "Nakit durumu orta seviyede." `
        -poorComment "Nakit pozisyonu zayıf, riskli."

    # Stok Bağımlılık Oranı (Küçük değer iyi)
    $results.StokBagimlilik = Score-Comment -value $stokBagimlilik -goodThreshold 0.2 -acceptableThreshold 0.4 -inverse $true `
        -goodComment "Stok bağımlılığı çok düşük, likidite iyi." `
        -acceptableComment "Orta düzey stok bağımlılığı, dikkat edilmeli." `
        -poorComment "Stok bağımlılığı yüksek, likidite riski var."

    # Kaldıraç Oranı (Küçük değer iyi)
    $results.Kaldirac = Score-Comment -value $kaldirac -goodThreshold 0.5 -acceptableThreshold 0.7 -inverse $true `
        -goodComment "Borç seviyesi düşük, finansal risk az." `
        -acceptableComment "Borç seviyesi orta, dikkatli olunmalı." `
        -poorComment "Borç seviyesi yüksek, risk fazla."

    # Finansman Oranı (Öz Kaynaklar / Yabancı Kaynaklar; büyük değer iyi)
    $results.Finansman = Score-Comment -value $finansman -goodThreshold 1 -acceptableThreshold 0.5 -inverse $false `
        -goodComment "Sağlıklı özsermaye ile finansman sağlanıyor." `
        -acceptableComment "Orta seviyede borç kullanımı var." `
        -poorComment "Yüksek borç kullanımı, finansal risk var."

    # Borç / Özsermaye Oranı (Küçük değer iyi)
    $results.BorcOzsermaye = Score-Comment -value $borcOzsermaye -goodThreshold 1 -acceptableThreshold 2 -inverse $true `
        -goodComment "Özsermaye güçlü, borç seviyesi makul." `
        -acceptableComment "Orta seviyede borç-özsermaye dengesi." `
        -poorComment "Borç yüksek, özsermaye zayıf."

    # Otofinansman Oranı (Büyük değer iyi)
    $results.Otofinansman = Score-Comment -value $otofinansman -goodThreshold 0.5 -acceptableThreshold 0.3 -inverse $false `
        -goodComment "Varlıkların büyük bölümü öz kaynak ile finanse edilmiş." `
        -acceptableComment "Orta düzeyde öz kaynak finansmanı var." `
        -poorComment "Öz kaynak kullanımı düşük, dış finansman yüksek."

    # Brüt Kar Marjı (Büyük değer iyi)
    $results.BrutKar = Score-Comment -value $brutKarMarji -goodThreshold 25 -acceptableThreshold 15 -inverse $false `
        -goodComment "Brüt kar marjı yüksek, karlılık iyi." `
        -acceptableComment "Orta düzeyde kar marjı var." `
        -poorComment "Brüt kar marjı düşük, karlılık sorunlu."

    return $results
}

# --- Puanlama sonuçlarını al ---
$scoreResults = Get-FinancialRatiosScore `
    -cariOran $cariOran `
    -asitTest $asitTest `
    -nakitOrani $nakitOrani `
    -stokBagimlilik $stokBagimlilik `
    -kaldirac $kaldirac `
    -finansman $finansman `
    -borcOzsermaye $borcOzsermaye `
    -otofinansman $otofinansman `
    -brutKarMarji $brutKarMarji

# Hesaplanamayan oranların puanını genel ortalamaya dahil etme.
foreach ($oranAdi in $oranDurumlari.Keys) {
    if (-not [string]::IsNullOrWhiteSpace($oranDurumlari[$oranAdi])) {
        $scoreResults[$oranAdi].Puan = $null
        $scoreResults[$oranAdi].Yorum = $oranDurumlari[$oranAdi]
    }
}

# --- RAPOR ÇIKTISI ---

# Ana rapor başlıkları (artık arka plan rengi yok)
Write-ColoredLine "Finansal Oran Analizi Raporu" -Color Blue -Width 100
Write-ColoredLine ("=" * 100) -Color DarkBlue -Width 100
Write-ColoredLine "Not: Mizanda bulunmayan 2, 4, 11 ve 61 kodlu isteğe bağlı hesap grupları sıfır kabul edilir." -Color DarkGray
Write-ColoredLine ("Hazırlayan: SMMM Mustafa Orakcı") -Color Blue -Width 100
Write-ColoredLine ("Tarih: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')") -Color Blue -Width 100
Write-ColoredLine ("=" * 100) -Color DarkBlue -Width 100


Write-SectionHeader "LİKİDİTE ORANLARI"

Write-Ratio -Name "1- Cari Oran" `
    -Value $cariOran `
    -Formula "Dönen Varlıklar / Kısa Vadeli Yabancı Kaynaklar" `
    -Description "İşletmenin kısa vadeli borçlarını ödeme gücünü gösterir." `
    -UnavailableReason $oranDurumlari.CariOran `
    -Score $scoreResults.CariOran.Puan `
    -Comment $scoreResults.CariOran.Yorum

Write-Ratio -Name "2- Asit Test Oranı" `
    -Value $asitTest `
    -Formula "(Dönen Varlıklar - Stoklar) / Kısa Vadeli Yabancı Kaynaklar" `
    -Description "Stoklar hariç, likit varlıkların kısa vadeli borçlara oranı." `
    -UnavailableReason $oranDurumlari.AsitTest `
    -Score $scoreResults.AsitTest.Puan `
    -Comment $scoreResults.AsitTest.Yorum

Write-Ratio -Name "3- Nakit Oranı" `
    -Value $nakitOrani `
    -Formula "(Hazır Değerler + Menkul Kıymetler) / Kısa Vadeli Yabancı Kaynaklar" `
    -Description "En likit varlıkların kısa vadeli borçlara oranı." `
    -UnavailableReason $oranDurumlari.NakitOrani `
    -Score $scoreResults.NakitOrani.Puan `
    -Comment $scoreResults.NakitOrani.Yorum

Write-Ratio -Name "4- Stok Bağımlılık Oranı" `
    -Value $stokBagimlilik `
    -Formula "Stoklar / Dönen Varlıklar" `
    -Description "Dönen varlıklar içinde stokların payı. Düşük olması daha iyidir." `
    -UnavailableReason $oranDurumlari.StokBagimlilik `
    -Score $scoreResults.StokBagimlilik.Puan `
    -Comment $scoreResults.StokBagimlilik.Yorum

Write-SectionHeader "MALİ YAPI ORANLARI"

Write-Ratio -Name "1- Kaldıraç Oranı" `
    -Value $kaldirac `
    -Formula "Toplam Borçlar / Toplam Aktifler" `
    -Description "İşletmenin aktiflerini borç ile finanse etme oranı. Düşük olması daha iyidir." `
    -UnavailableReason $oranDurumlari.Kaldirac `
    -Score $scoreResults.Kaldirac.Puan `
    -Comment $scoreResults.Kaldirac.Yorum

Write-Ratio -Name "2- Finansman Oranı" `
    -Value $finansman `
    -Formula "Öz Kaynaklar / Toplam Yabancı Kaynaklar" `
    -Description "Yabancı kaynakların öz kaynaklarla karşılanma gücünü gösterir. Yüksek olması daha iyidir." `
    -UnavailableReason $oranDurumlari.Finansman `
    -Score $scoreResults.Finansman.Puan `
    -Comment $scoreResults.Finansman.Yorum

Write-Ratio -Name "3- Borç / Özsermaye Oranı" `
    -Value $borcOzsermaye `
    -Formula "Toplam Borçlar / Öz Kaynaklar" `
    -Description "Borçların öz sermayeye oranı. Düşük olması daha iyidir." `
    -UnavailableReason $oranDurumlari.BorcOzsermaye `
    -Score $scoreResults.BorcOzsermaye.Puan `
    -Comment $scoreResults.BorcOzsermaye.Yorum

Write-Ratio -Name "4- Otofinansman Oranı" `
    -Value $otofinansman `
    -Formula "Öz Kaynaklar / Toplam Aktifler" `
    -Description "Varlıkların ne kadarının öz kaynakla finanse edildiği. Yüksek olması daha iyidir." `
    -UnavailableReason $oranDurumlari.Otofinansman `
    -Score $scoreResults.Otofinansman.Puan `
    -Comment $scoreResults.Otofinansman.Yorum

Write-SectionHeader "KÂRLILIK ORANLARI"

Write-Ratio -Name "1- Brüt Kar Marjı" `
    -Value $brutKarMarji `
    -Formula "((Brüt Satışlar - Satış İndirimleri - Satış Maliyeti) / Brüt Satışlar) * 100" `
    -Description "Satışların maliyetinden sonra kalan kar marjı yüzdesi. Yüksek olması daha iyidir." `
    -UnavailableReason $oranDurumlari.BrutKar `
    -IsPercentage `
    -Score $scoreResults.BrutKar.Puan `
    -Comment $scoreResults.BrutKar.Yorum

# --- Genel Değerlendirme Özeti ---

$totalScoreSum = 0
$ratioCount = 0
foreach ($ratioName in $scoreResults.Keys) {
    # Bir oranın puanı olmadığında (hesaplama 0/0 gibi) skor null olabilir, kontrol ekleyelim
    if ($scoreResults.$ratioName.Puan -ne $null) {
        $totalScoreSum += $scoreResults.$ratioName.Puan
        $ratioCount++
    }
}

$overallAverageScore = if ($ratioCount -gt 0) { $totalScoreSum / $ratioCount } else { 0 }
$overallComment = ""

if ($ratioCount -eq 0) {
    $overallComment = "Değerlendirme yapılamadı; hesaplanabilir oran bulunamadı."
    $overallColor = 'Yellow'
} elseif ($overallAverageScore -ge 4) {
    $overallComment = "Şirketin finansal durumu genel olarak ÇOK GÜÇLÜ ve SAĞLAM görünüyor. Mevcut politikalarınıza devam etmeniz önerilir."
    $overallColor = 'Green'
} elseif ($overallAverageScore -ge 2.5) {
    $overallComment = "Şirketin finansal durumu ORTA DÜZEYDE ve stabil görünüyor. Bazı alanlarda iyileştirme potansiyeli mevcut, dikkatli olunmalı."
    $overallColor = 'Yellow'
} else {
    $overallComment = "Şirketin finansal durumu ZAYIF ve RİSKLİ görünüyor. Acil ve kapsamlı finansal iyileştirme önlemleri alınması gerekebilir."
    $overallColor = 'Red'
}

Write-SectionHeader "GENEL FİNANSAL DEĞERLENDİRME"
Write-ColoredLine "Hesaplanabilen oran: $ratioCount / 9" -Color $overallColor
if ($ratioCount -gt 0) { Write-ColoredLine ("Ortalama Puan : {0:N2}" -f $overallAverageScore) -Color $overallColor }
Write-ColoredLine ("Genel Yorum   : $overallComment") -Color $overallColor
Write-Host "" # Boş satır

# Rapor kapanış bilgisi (artık arka plan rengi yok)
Write-ColoredLine "RAPOR SONU SMMM MUSTAFA ORAKCI." -Color Blue -Width 100
Write-ColoredLine ("=" * 100) -Color DarkBlue -Width 100
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  ⚖️ Luca Mizan Rasyo Analizi Aracı' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencereden Luca Mizan dosyanızı (.xls, .xlsx veya .csv) seçin." -ForegroundColor Yellow
    Write-Host "📌 Finansal oranları (likidite, kârlılık, kaldıraç vb.) hesaplar ve puanlar." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    $script:selectionCancelled = $false
    StartProgram
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
