# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

try { Clear-Host -ErrorAction Stop } catch { }

# ------------------------------------------------------------------
# 📋 E-DEFTER KONTROL PROGRAMI
# Kaynak e-defter klasöründeki  VKN \ Dönem \ Ay  yapısını tarar ve her ay için
# Yevmiye / Kebir defterlerinin parça, berat ve GİB onaylı berat sayılarını kontrol eder.
# Hiçbir dosyayı değiştirmez, silmez, taşımaz; yalnızca klasör listelerini okur.
# ------------------------------------------------------------------

# Filtreler
$script:EkYearFilter  = ''
$script:EkMonthFilter = @()          # Boş dizi = tüm aylar

# Yapılandırma
$script:EkRoot           = ''
$script:EkRootText       = ''
$script:EkRootStatus     = 'Yüklenmedi'   # OK, NotFound, InvalidPath, MissingProperty, Error
$script:EkRootConfigPath = Join-Path (Get-MFTModuleDataPath -ModuleId 'edefterkontrol' -Category Config) 'edefterkontrol_path_config.json'
$script:EkFirmaPath      = Join-Path (Get-MFTModuleDataPath -ModuleId 'edefterkontrol' -Category Config) 'edefterkontrol_firmalar.json'
$script:EkReportDir      = Get-MFTModuleDataPath -ModuleId 'edefterkontrol' -Category Reports
$script:EkFirmaMap       = @{}
$script:EkFirmaPasif     = @{}   # VKN -> $true: pasif (kapanmış / muhasebeci değişmiş) firma
$script:EkWidthOverride  = 0     # yalnızca test için: konsol genişliğini zorla
$script:EkNoName        = '(adı tanımsız)'   # Firma listesinde adı olmayan müşteri için gösterilir

# Son tarama sonucu
$script:EkRows    = @()
$script:EkReadErrors = @()

$script:EkMonthNames = @{
  '01'='Ocak'; '02'='Şubat'; '03'='Mart'; '04'='Nisan'; '05'='Mayıs'; '06'='Haziran';
  '07'='Temmuz'; '08'='Ağustos'; '09'='Eylül'; '10'='Ekim'; '11'='Kasım'; '12'='Aralık'
}

function Wait-EkKey {
  param([string]$Text = 'Devam etmek için bir tuşa basın.')
  Write-Host $Text -ForegroundColor DarkGray
  try { $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') }
  catch { [void](Read-Host 'Devam etmek için Enter''a basın') }
}

function Clear-EkScreen {
  try { [Console]::Clear() }
  catch { try { Clear-Host -ErrorAction Stop } catch { Write-Host '' } }
}

function Read-EkMenuKey {
  try { return [int]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').VirtualKeyCode }
  catch {
    $key = (Read-Host 'Seçim (H/N/Q)').Trim().ToUpperInvariant()
    switch ($key) { 'H' { return 72 } 'N' { return 78 } 'Q' { return 81 } default { return 0 } }
  }
}

function Get-EkChildItems {
  param([string]$Path, [switch]$File, [switch]$Directory)
  try {
    $args = @{ LiteralPath = $Path; ErrorAction = 'Stop' }
    if ($File) { $args.File = $true }
    if ($Directory) { $args.Directory = $true }
    return @(Get-ChildItem @args)
  } catch {
    $script:EkReadErrors += [PSCustomObject]@{ Path = $Path; Message = $_.Exception.Message }
    return @()
  }
}

function Get-EkMonthFilterText {
  if ($script:EkMonthFilter.Count -eq 0) { return 'Yok (Tümü)' }
  return (($script:EkMonthFilter | ForEach-Object { $script:EkMonthNames[$_] }) -join ', ')
}

# --------------------------------------------
# 👥 Firma adları (VKN -> Firma)
# --------------------------------------------
function Load-EkFirmaMap {
  $map = @{}
  $pasif = @{}
  if (Test-Path -LiteralPath $script:EkFirmaPath) {
    try {
      $raw = Get-Content -LiteralPath $script:EkFirmaPath -Raw -Encoding UTF8 -ErrorAction Stop
      if (-not [string]::IsNullOrWhiteSpace($raw)) {
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        foreach ($p in $obj.PSObject.Properties) {
          $key = [string]$p.Name
          $map[$key] = [string]$p.Value.Firma
          # "Aktif" alanı yoksa firma aktif sayılır (eski dosyalarla uyumlu)
          if ($p.Value.PSObject.Properties.Name -contains 'Aktif') {
            if (-not [bool]$p.Value.Aktif) { $pasif[$key] = $true }
          }
        }
      }
    } catch {
      Write-Warning "⚠️ Firma listesi okunamadı: $($_.Exception.Message)"
    }
  }
  $script:EkFirmaMap = $map
  $script:EkFirmaPasif = $pasif
}

function Get-EkFirmaSummaryText {
  $n = $script:EkFirmaMap.Count
  $p = @($script:EkFirmaPasif.Keys | Where-Object { $script:EkFirmaMap.ContainsKey($_) }).Count
  if ($p -gt 0) { return "$n kayıt ($p pasif)" }
  return "$n kayıt"
}

function Save-EkFirmaMap {
  try {
    $out = [ordered]@{}
    foreach ($k in ($script:EkFirmaMap.Keys | Sort-Object)) {
      $out[$k] = [ordered]@{ Firma = $script:EkFirmaMap[$k]; Aktif = (-not $script:EkFirmaPasif.ContainsKey($k)) }
    }
    $json = '{}'
    if ($out.Count -gt 0) { $json = $out | ConvertTo-Json -Depth 3 }
    [System.IO.File]::WriteAllText($script:EkFirmaPath, $json, (New-Object System.Text.UTF8Encoding($true)))
    return $true
  } catch {
    Write-Warning "🚫 Firma listesi kaydedilemedi: $($_.Exception.Message)"
    Load-EkFirmaMap
    return $false
  }
}

# --------------------------------------------
# 👥 Firma adları menüsü (listele / ekle-güncelle / sil)
# --------------------------------------------
function Show-EkFirmaMenu {
  while ($true) {
    Clear-EkScreen
    Write-Host '👥 Firma Adları' -ForegroundColor Yellow
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host "Kayıtlı firma: $(Get-EkFirmaSummaryText)"
    Write-Host ''
    Write-Host '  1. 📋  Listele'
    Write-Host '  2. ➕  Ekle / Güncelle'
    Write-Host '  3. 🗑️  Sil'
    Write-Host '  4. 🔄  Aktif / Pasif Değiştir'
    Write-Host '  Q  ↩️  Geri Dön'
    Write-Host ''
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host ''
    $choice = Read-Host 'Seçiminiz'

    switch ($choice.ToUpper()) {
      '1' {
        Write-Host ''
        if ($script:EkFirmaMap.Count -eq 0) {
          Write-Host 'Kayıtlı firma yok.' -ForegroundColor DarkYellow
        } else {
          foreach ($k in ($script:EkFirmaMap.Keys | Sort-Object)) {
            $line = ('{0,-13}{1,-34}' -f $k, $script:EkFirmaMap[$k])
            Write-Host $line -NoNewline
            if ($script:EkFirmaPasif.ContainsKey($k)) { Write-Host 'Pasif' -ForegroundColor DarkYellow } else { Write-Host 'Aktif' -ForegroundColor Green }
          }
        }
        Write-Host ''
        Wait-EkKey
      }
      '2' {
        $vkn = (Read-Host 'VKN / TC No (10 veya 11 hane)').Trim()
        if ($vkn -notmatch '^\d{10,11}$') {
          Write-Warning '⚠️ Geçersiz VKN/TC. 10 veya 11 haneli sayı girin.'
        } else {
          $mevcutDurum = 'Aktif'
          if ($script:EkFirmaMap.ContainsKey($vkn)) {
            if ($script:EkFirmaPasif.ContainsKey($vkn)) { $mevcutDurum = 'Pasif' }
            Write-Host "Mevcut firma: $($script:EkFirmaMap[$vkn]) ($mevcutDurum)"
          }
          $ad = (Read-Host 'Firma adı').Trim()
          if ([string]::IsNullOrEmpty($ad)) {
            Write-Warning '⚠️ Firma adı boş olamaz, kayıt yapılmadı.'
          } else {
            $d = (Read-Host "Durum (A = Aktif, P = Pasif; boş = $mevcutDurum)").Trim().ToUpper()
            $pasifYap = ($mevcutDurum -eq 'Pasif')
            if ($d -eq 'A') { $pasifYap = $false }
            elseif ($d -eq 'P') { $pasifYap = $true }
            elseif ($d -ne '') { Write-Warning "⚠️ Geçersiz durum, '$mevcutDurum' olarak bırakıldı." }
            $script:EkFirmaMap[$vkn] = $ad
            if ($pasifYap) { $script:EkFirmaPasif[$vkn] = $true } else { $script:EkFirmaPasif.Remove($vkn) }
            if (Save-EkFirmaMap) { Write-Host '✅ Kaydedildi.' -ForegroundColor Green }
          }
        }
        Wait-EkKey
      }
      '3' {
        $vkn = (Read-Host 'Silinecek VKN / TC No').Trim()
        if ($script:EkFirmaMap.ContainsKey($vkn)) {
          $onay = Read-Host "'$($script:EkFirmaMap[$vkn])' kaydı silinsin mi? (E/H)"
          if ($onay.ToUpper() -eq 'E') {
            $script:EkFirmaMap.Remove($vkn)
            $script:EkFirmaPasif.Remove($vkn)
            if (Save-EkFirmaMap) { Write-Host '✅ Silindi.' -ForegroundColor Green }
          } else {
            Write-Host 'İptal edildi.' -ForegroundColor DarkYellow
          }
        } else {
          Write-Warning '⚠️ Bu VKN listede yok.'
        }
        Wait-EkKey
      }
      '4' {
        $vkn = (Read-Host 'Durumu değişecek VKN / TC No').Trim()
        if ($script:EkFirmaMap.ContainsKey($vkn)) {
          if ($script:EkFirmaPasif.ContainsKey($vkn)) {
            $script:EkFirmaPasif.Remove($vkn)
            $yeni = 'AKTİF'
          } else {
            $script:EkFirmaPasif[$vkn] = $true
            $yeni = 'PASİF'
          }
          if (Save-EkFirmaMap) { Write-Host "✅ '$($script:EkFirmaMap[$vkn])' artık $yeni." -ForegroundColor Green }
        } else {
          Write-Warning '⚠️ Bu VKN listede yok. Önce 2) ile ekleyin.'
        }
        Wait-EkKey
      }
      'Q' { return }
      default {
        Write-Host '[!] Geçersiz seçim.' -ForegroundColor Red
        Wait-EkKey
      }
    }
  }
}

# --------------------------------------------
# 📂 E-defter ana klasörü
# --------------------------------------------
function Load-EkRoot {
  $script:EkRoot = ''
  $script:EkRootText = ''
  $script:EkRootStatus = 'NotFound'
  if (-not (Test-Path -LiteralPath $script:EkRootConfigPath)) { return }
  try {
    $raw = Get-Content -LiteralPath $script:EkRootConfigPath -Raw -Encoding UTF8 -ErrorAction Stop
    $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
    $p = ([string]$cfg.EdefterRoot).Trim()
    $script:EkRootText = $p
    if ([string]::IsNullOrEmpty($p)) { $script:EkRootStatus = 'MissingProperty'; return }
    if (Test-Path -LiteralPath $p -PathType Container) {
      $script:EkRoot = $p
      $script:EkRootStatus = 'OK'
    } else {
      $script:EkRootStatus = 'InvalidPath'
    }
  } catch {
    $script:EkRootStatus = 'Error'
  }
}

function Save-EkRoot {
  param([string]$Path)
  try {
    $json = ([PSCustomObject]@{ EdefterRoot = $Path } | ConvertTo-Json -Depth 1)
    [System.IO.File]::WriteAllText($script:EkRootConfigPath, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "✅ E-Defter ana klasörü '$Path' olarak kaydedildi." -ForegroundColor Green
    return $true
  } catch {
    Write-Warning "🚫 Ayar kaydedilemedi: $($_.Exception.Message)"
    return $false
  }
}

function Edit-EkRoot {
  Clear-EkScreen
  Write-Host '✏️ E-Defter Ana Klasörü (VKN klasörlerinin bulunduğu klasör)' -ForegroundColor Yellow
  Write-Host ('-' * 72) -ForegroundColor DarkGray
  $mevcut = $script:EkRootText
  if ([string]::IsNullOrEmpty($mevcut)) { $mevcut = 'Tanımlı Değil' }
  Write-Host "Mevcut Yol: $mevcut"
  $in = Read-Host 'Yeni yol (mevcut bırakmak için boş geçin)'
  if ([string]::IsNullOrWhiteSpace($in)) {
    Write-Warning '⚠️ Değişiklik yapılmadı.'
  } else {
    $new = $in.Trim().Trim('"')
    if (-not [System.IO.Path]::IsPathRooted($new)) {
      Write-Warning '⚠️ Lütfen mutlak bir yol girin (örn: D:\Users\...\e-defter).'
    } elseif (-not (Test-Path -LiteralPath $new -PathType Container)) {
      Write-Warning "⚠️ Klasör bulunamadı: '$new'"
    } else {
      if (Save-EkRoot -Path $new) { Load-EkRoot }
    }
  }
  Wait-EkKey
}

# --------------------------------------------
# 🔎 Bir ay klasörünü tara: dosya adlarından parça / berat / GİB berat numaralarını çıkar
#   Y  = Yevmiye parçası      YB = Yevmiye beratı      GIB-...-YB = GİB onaylı yevmiye beratı
#   K  = Kebir parçası        KB = Kebir beratı        GIB-...-KB = GİB onaylı kebir beratı
# --------------------------------------------
function Get-EkMonthInfo {
  param([string]$MonthPath, [string]$ExpectedVkn, [string]$ParentYear, [string]$ParentMonth)

  $sets = @{ Y = @{}; YB = @{}; GY = @{}; K = @{}; KB = @{}; GK = @{} }
  $nameIssues = New-Object System.Collections.Generic.List[PSObject]
  $xmlBases = @{}
  $zipBases = @{}
  $ym = ''
  $opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
  $allFiles = @(Get-EkChildItems -Path $MonthPath -File)
  foreach ($zip in @($allFiles | Where-Object { $_.Extension -ieq '.zip' })) { $zipBases[$zip.BaseName] = $true }
  $files = @($allFiles | Where-Object { $_.Extension -ieq '.xml' })

  foreach ($f in $files) {
    $xmlBases[$f.BaseName] = $true
    $m = [regex]::Match($f.Name, '^(?<gib>GIB-)?(?<vkn>\d{10,11})-(?<ym>\d{6})-(?<tip>DR|YB|KB|Y|K)-(?<no>\d{6})\.xml$', $opt)
    if (-not $m.Success) {
      $nameIssues.Add([PSCustomObject]@{ FileYear = $ParentYear; FileMonth = $ParentMonth; Text = "Tanınmayan XML dosya adı: $($f.Name)" })
      continue
    }
    $fileVkn = $m.Groups['vkn'].Value
    $fileYm = $m.Groups['ym'].Value
    $fileYear = $fileYm.Substring(0, 4)
    $fileMonth = $fileYm.Substring(4, 2)
    $fileIssues = New-Object System.Collections.Generic.List[string]
    if (-not $zipBases.ContainsKey($f.BaseName)) {
      $nameIssues.Add([PSCustomObject]@{ FileYear = $fileYear; FileMonth = $fileMonth; Text = "Karşılık gelen ZIP yok: $($f.BaseName).zip" })
    }
    if ($fileVkn -ne $ExpectedVkn) { $fileIssues.Add("dosya VKN/TC $fileVkn, klasör $ExpectedVkn") }
    if ($ParentYear -and $fileYear -ne $ParentYear) { $fileIssues.Add("dosya yılı $fileYear, klasör yılı $ParentYear") }
    if ($ParentMonth -and $fileMonth -ne ('{0:D2}' -f [int]$ParentMonth)) { $fileIssues.Add("dosya ayı $fileMonth, klasör ayı $('{0:D2}' -f [int]$ParentMonth)") }
    $tip = $m.Groups['tip'].Value.ToUpperInvariant()
    if ($m.Groups['gib'].Success -and $tip -notin @('YB','KB')) { $fileIssues.Add('GIB öneki yalnızca YB/KB beratlarında geçerli') }
    if ($fileIssues.Count -gt 0) {
      $nameIssues.Add([PSCustomObject]@{ FileYear = $fileYear; FileMonth = $fileMonth; Text = "$($f.Name): $($fileIssues -join ', ')" })
      continue
    }
    if ($ym -eq '') { $ym = $m.Groups['ym'].Value }
    # DR rapor XML'i tanınır ve doğrulanır, ancak defter/berat adetlerine girmez.
    if ($tip -eq 'DR') { continue }
    if ($m.Groups['gib'].Success) {
      if ($tip -eq 'YB') { $key = 'GY' }
      elseif ($tip -eq 'KB') { $key = 'GK' }
      else { continue }
    } else {
      $key = $tip
    }
    $sets[$key][[int]$m.Groups['no'].Value] = $true
  }

  foreach ($zip in @($allFiles | Where-Object { $_.Extension -ieq '.zip' })) {
    if ($zip.Name -in @('Beratlar.zip','Kebirler.zip','Yevmiyeler.zip')) { continue }
    $m = [regex]::Match($zip.Name, '^(?<gib>GIB-)?(?<vkn>\d{10,11})-(?<ym>\d{6})-(?<tip>DR|YB|KB|Y|K)-(?<no>\d{6})\.zip$', $opt)
    if (-not $m.Success) {
      $nameIssues.Add([PSCustomObject]@{ FileYear = $ParentYear; FileMonth = $ParentMonth; Text = "Tanınmayan ZIP dosya adı: $($zip.Name)" })
      continue
    }
    if (-not $xmlBases.ContainsKey($zip.BaseName)) {
      $zipYear = $m.Groups['ym'].Value.Substring(0, 4)
      $zipMonth = $m.Groups['ym'].Value.Substring(4, 2)
      $nameIssues.Add([PSCustomObject]@{ FileYear = $zipYear; FileMonth = $zipMonth; Text = "ZIP var ancak karşılık gelen XML yok: $($zip.Name)" })
    }
  }
  return [PSCustomObject]@{ Sets = $sets; YM = $ym; NameIssues = @($nameIssues.ToArray()) }
}

# Bir defter türü (yevmiye veya kebir) için eksik/fazla açıklamalarını üretir
function Get-EkSetNotes {
  param([string]$Ad, [hashtable]$Parca, [hashtable]$Berat, [hashtable]$Gib)

  $notes = New-Object System.Collections.Generic.List[string]
  if ($Parca.Count -eq 0) { $notes.Add("$Ad parçası yok") }

  $eksikBerat = @($Parca.Keys | Where-Object { -not $Berat.ContainsKey($_) } | Sort-Object)
  $eksikGib   = @($Parca.Keys | Where-Object { -not $Gib.ContainsKey($_) }   | Sort-Object)
  $fazlaBerat = @($Berat.Keys | Where-Object { -not $Parca.ContainsKey($_) } | Sort-Object)
  $fazlaGib   = @($Gib.Keys   | Where-Object { -not $Parca.ContainsKey($_) } | Sort-Object)

  $fmt = { param($list) (($list | ForEach-Object { '{0:D6}' -f [int]$_ }) -join ', ') }

  if ($eksikBerat.Count -gt 0) { $notes.Add("$Ad berat eksik (" + (& $fmt $eksikBerat) + ')') }
  if ($eksikGib.Count   -gt 0) { $notes.Add("$Ad GİB onaylı berat eksik (" + (& $fmt $eksikGib) + ')') }
  if ($fazlaBerat.Count -gt 0) { $notes.Add("$Ad fazladan berat (" + (& $fmt $fazlaBerat) + ')') }
  if ($fazlaGib.Count   -gt 0) { $notes.Add("$Ad fazladan GİB berat (" + (& $fmt $fazlaGib) + ')') }
  return ,$notes.ToArray()
}

# --------------------------------------------
# 📋 Tüm müşteri / ay satırlarını üret (filtreler uygulanır)
# --------------------------------------------
function Get-EkRows {
  param([string]$Root, [string]$YearFilter, [array]$MonthFilter)

  $rows = New-Object System.Collections.Generic.List[PSObject]
  $script:EkReadErrors = @()
  if ([string]::IsNullOrEmpty($Root) -or -not (Test-Path -LiteralPath $Root -PathType Container)) { return @() }

  $vknDirs = @(Get-EkChildItems -Path $Root -Directory |
    Where-Object { $_.Name -match '^\d{10,11}$' } | Sort-Object Name)

  foreach ($v in $vknDirs) {
    $vkn = $v.Name
    $firma = $script:EkNoName
    if ($script:EkFirmaMap.ContainsKey($vkn) -and -not [string]::IsNullOrWhiteSpace($script:EkFirmaMap[$vkn])) { $firma = $script:EkFirmaMap[$vkn] }
    $rowsBefore = $rows.Count
    $aktif = -not $script:EkFirmaPasif.ContainsKey($vkn)
    $monthReadErrorKeys = @{}
    $periodReadErrorYears = @{}
    $rootReadError = $false

    $periodListErrorCount = $script:EkReadErrors.Count
    $periods = @(Get-EkChildItems -Path $v.FullName -Directory | Sort-Object Name)
    if ($script:EkReadErrors.Count -gt $periodListErrorCount) { $rootReadError = $true }
    foreach ($p in $periods) {
      $pm = [regex]::Match($p.Name, '^(?<sd>\d{2})\.(?<sm>\d{2})\.(?<sy>\d{4})-(?<ed>\d{2})\.(?<em>\d{2})\.(?<ey>\d{4})$')
      if (-not $pm.Success) {
        $foundYear = [regex]::Match($p.Name, '\d{4}').Value
        if (-not $YearFilter -or $foundYear -eq $YearFilter) {
          $rows.Add([PSCustomObject]@{
            VKN = $vkn; Firma = $firma; Year = $foundYear; Month = ''; VknPath = $v.FullName
            PeriodName = $p.Name; PeriodPath = $p.FullName; MonthPath = ''
            YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
            Donem = '-'; Aktif = $aktif; Bos = $true; Durum = 'Uyuşmazlık'
            Aciklama = "Dönem klasör adı beklenen GG.AA.YYYY-GG.AA.YYYY biçiminde değil: $($p.Name)"
          })
        }
        continue
      }
      $periodYear = $pm.Groups['sy'].Value
      $periodEndYear = $pm.Groups['ey'].Value
      if ($YearFilter -and ([int]$YearFilter -lt [int]$periodYear -or [int]$YearFilter -gt [int]$periodEndYear)) { continue }
      $monthListErrorCount = $script:EkReadErrors.Count
      $allMonthDirs = @(Get-EkChildItems -Path $p.FullName -Directory | Sort-Object Name)
      $periodListFailed = $script:EkReadErrors.Count -gt $monthListErrorCount
      if ($periodListFailed) {
        $errorYear = $periodYear
        if ($YearFilter -and [int]$YearFilter -ge [int]$periodYear -and [int]$YearFilter -le [int]$periodEndYear) { $errorYear = $YearFilter }
        $periodReadErrorYears[$errorYear] = $true
      }
      foreach ($badMonthDir in @($allMonthDirs | Where-Object { $_.Name -notmatch '^(0[1-9]|1[0-2])$' })) {
        $badMonth = ''
        $trimmedMonth = $badMonthDir.Name.Trim()
        if ($trimmedMonth -match '^\d{1,2}$' -and [int]$trimmedMonth -ge 1 -and [int]$trimmedMonth -le 12) {
          $badMonth = '{0:D2}' -f [int]$trimmedMonth
        }
        if ($MonthFilter.Count -gt 0 -and (-not $badMonth -or $MonthFilter -notcontains $badMonth)) { continue }
        $rows.Add([PSCustomObject]@{
          VKN = $vkn; Firma = $firma; Year = $periodYear; Month = $badMonth; VknPath = $v.FullName
          PeriodName = $p.Name; PeriodPath = $p.FullName; MonthPath = $badMonthDir.FullName
          YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
          Donem = $(if ($badMonth) { "$periodYear-$badMonth" } else { '-' }); Aktif = $aktif; Bos = $true; Durum = 'Uyuşmazlık'
          Aciklama = "Ay klasör adı 01-12 arasında iki haneli olmalı: $($badMonthDir.Name)"
        })
      }
      $monthDirs = @($allMonthDirs | Where-Object {
        $_.Name -match '^(0[1-9]|1[0-2])$' -and ($MonthFilter.Count -eq 0 -or $MonthFilter -contains $_.Name)
      })

      foreach ($md in $monthDirs) {
        $errorCountBefore = $script:EkReadErrors.Count
        $parentMonth = '{0:D2}' -f [int]$md.Name
        $info = Get-EkMonthInfo -MonthPath $md.FullName -ExpectedVkn $vkn -ParentYear $periodYear -ParentMonth $parentMonth
        if ($script:EkReadErrors.Count -gt $errorCountBefore) {
          $errorYear = $periodYear
          if ($YearFilter -and [int]$YearFilter -ge [int]$periodYear -and [int]$YearFilter -le [int]$periodEndYear) { $errorYear = $YearFilter }
          $monthReadErrorKeys[('{0}|{1}' -f $errorYear, $parentMonth)] = $true
        }
        if ($info.NameIssues.Count -gt 0) {
          foreach ($issue in $info.NameIssues) {
            if ($YearFilter -and $YearFilter -ne $periodYear -and $YearFilter -ne $issue.FileYear) { continue }
            if ($MonthFilter.Count -gt 0 -and $MonthFilter -notcontains $parentMonth -and $MonthFilter -notcontains $issue.FileMonth) { continue }
            $rows.Add([PSCustomObject]@{
              VKN = $vkn; Firma = $firma; Year = $periodYear; Month = $parentMonth
              VknPath = $v.FullName; PeriodName = $p.Name; PeriodPath = $p.FullName; MonthPath = $md.FullName
              YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
              Donem = "$periodYear-$parentMonth"; Aktif = $aktif; Bos = $true
              Durum = 'Uyuşmazlık'; Aciklama = $issue.Text
            })
          }
        }
        # Bu ay klasöründe yalnızca yanlış yıl taşıyan dosyalar varsa,
        # Uyuşmazlık satırı yeterlidir; boş bir ayı ayrıca Eksik sayma.
        if (-not $info.YM -and $info.NameIssues.Count -gt 0) { continue }
        $year = ''
        $month = ''
        if ($info.YM -match '^(\d{4})(\d{2})$') {
          $year = $Matches[1]
          $month = $Matches[2]
        } else {
          $month = '{0:D2}' -f [int]$md.Name
          $year = $periodYear
        }
        if ($YearFilter -ne '' -and $year -ne $YearFilter) { continue }
        if ($MonthFilter.Count -gt 0 -and ($MonthFilter -notcontains $month)) { continue }

        $s = $info.Sets
        $notes = New-Object System.Collections.Generic.List[string]
        $notes.AddRange([string[]](Get-EkSetNotes -Ad 'Yevmiye' -Parca $s.Y -Berat $s.YB -Gib $s.GY))
        $notes.AddRange([string[]](Get-EkSetNotes -Ad 'Kebir'   -Parca $s.K -Berat $s.KB -Gib $s.GK))
        $durum = 'Tamam'
        if ($notes.Count -gt 0) { $durum = 'Eksik' }

        if ($script:EkReadErrors.Count -gt $errorCountBefore) { continue }
        $rows.Add([PSCustomObject]@{
          VKN = $vkn; Firma = $firma; Year = $year; Month = $month
          VknPath = $v.FullName; PeriodName = $p.Name; PeriodPath = $p.FullName; MonthPath = $md.FullName
          YParca = $s.Y.Count; YBerat = $s.YB.Count; YGib = $s.GY.Count
          KParca = $s.K.Count; KBerat = $s.KB.Count; KGib = $s.GK.Count
          Donem = ($year + '-' + $month); Aktif = $aktif; Bos = $false
          Durum = $durum; Aciklama = ($notes.ToArray() -join '; ')
        })
      }
    }

    # Açık yıl/ay filtresinde, seçilen ay klasörü bulunmasa bile o müşteri-ay
    # raporda görünmeli. Bir erişim hatası varsa boşluk varsaymak yerine hata
    # satırını kullan.
    $hasRelevantReadError = $rootReadError -or $monthReadErrorKeys.Count -gt 0 -or $periodReadErrorYears.Count -gt 0
    if ($YearFilter -and $MonthFilter.Count -gt 0) {
      foreach ($selectedMonth in $MonthFilter) {
        $monthAlreadyReported = @($rows | Where-Object {
          $_.VKN -eq $vkn -and $_.Year -eq $YearFilter -and $_.Month -eq $selectedMonth
        }).Count -gt 0
        if ($monthAlreadyReported) { continue }
        $monthKey = '{0}|{1}' -f $YearFilter, $selectedMonth
        if ($rootReadError -or $monthReadErrorKeys.ContainsKey($monthKey) -or $periodReadErrorYears.ContainsKey($YearFilter)) { continue }
        if ($aktif) {
          $missingStatus = 'Aktif/Boş'
          $missingDescription = 'Seçilen aya ait klasör veya e-defter bulunamadı'
        } else {
          $missingStatus = 'Pasif/Boş'
          $missingDescription = 'Pasif firma: seçilen aya ait e-defter yok'
        }
        $rows.Add([PSCustomObject]@{
          VKN = $vkn; Firma = $firma; Year = $YearFilter; Month = $selectedMonth
          VknPath = $v.FullName; PeriodName = ''; PeriodPath = ''; MonthPath = ''
          YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
          Donem = "$YearFilter-$selectedMonth"; Aktif = $aktif; Bos = $true
          Durum = $missingStatus; Aciklama = $missingDescription
        })
      }
    }

    # Filtreye uyan hiç e-defter bulunamadıysa (firma kapanmış, muhasebeci değişmiş, boş klasör vb.)
    # müşteri için tek bir bilgi satırı ekle. Bu bir hata sayılmaz, Eksik sayısına girmez.
    if ($rows.Count -eq $rowsBefore -and -not $hasRelevantReadError) {
      if ($aktif) {
        $bosDurum = 'Aktif/Boş'; $bosAcik = 'Aktif firma: filtrelenen döneme ait e-defter bulunamadı'
      } else {
        $bosDurum = 'Pasif/Boş'; $bosAcik = 'Pasif firma: filtrelenen döneme ait e-defter yok'
      }
      $rows.Add([PSCustomObject]@{
        VKN = $vkn; Firma = $firma; Year = $YearFilter; Month = ''
        VknPath = $v.FullName; PeriodName = ''; PeriodPath = ''; MonthPath = ''
        YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
        Donem = '-'; Aktif = $aktif; Bos = $true
        Durum = $bosDurum; Aciklama = $bosAcik
      })
    }
  }

  # Okunamayan klasörler boş/eksik ay gibi sayılmasın; bağlamı bilinen her
  # erişim hatasını ayrı teşhis satırı olarak rapora ekle.
  foreach ($e in $script:EkReadErrors) {
    $rootPrefix = $Root.TrimEnd('\','/')
    $rel = $e.Path.Substring($rootPrefix.Length).TrimStart('\','/')
    $parts = @($rel -split '[\\/]')
    $errVkn = ''; $errYear = ''; $errMonth = ''; $firma = $script:EkNoName; $aktif = $true
    if ($parts.Count -gt 0 -and $parts[0] -match '^\d{10,11}$') {
      $errVkn = $parts[0]
      if ($script:EkFirmaMap.ContainsKey($errVkn)) { $firma = $script:EkFirmaMap[$errVkn] }
      $aktif = -not $script:EkFirmaPasif.ContainsKey($errVkn)
    }
    if ($parts.Count -gt 1) {
      $errorPeriod = [regex]::Match($parts[1], '^(?<sd>\d{2})\.(?<sm>\d{2})\.(?<sy>\d{4})-(?<ed>\d{2})\.(?<em>\d{2})\.(?<ey>\d{4})$')
      if ($errorPeriod.Success) {
        $startYear = $errorPeriod.Groups['sy'].Value
        $endYear = $errorPeriod.Groups['ey'].Value
        $errYear = $startYear
        if ($YearFilter -and [int]$YearFilter -ge [int]$startYear -and [int]$YearFilter -le [int]$endYear) { $errYear = $YearFilter }
      } elseif ($parts[1] -match '\d{4}') { $errYear = [regex]::Match($parts[1], '\d{4}').Value }
    }
    if ($parts.Count -gt 2 -and $parts[2] -match '^\d{1,2}$') { $errMonth = '{0:D2}' -f [int]$parts[2] }
    if ($YearFilter -and $errYear -and $errYear -ne $YearFilter) { continue }
    $errorMonths = @($errMonth)
    if ($MonthFilter.Count -gt 0) {
      if ($errMonth) {
        if ($MonthFilter -notcontains $errMonth) { continue }
      } elseif ($errVkn) {
        # Okunamayan üst klasör seçili ayları listelemeyi engelliyorsa her
        # seçili müşteri-ay için ayrı kontrol sorunu satırı üret.
        $errorMonths = @($MonthFilter)
      }
    }
    $errorYearForRow = $errYear
    if ($YearFilter) { $errorYearForRow = $YearFilter }
    foreach ($errorMonth in $errorMonths) {
      $rows.Add([PSCustomObject]@{
        VKN = $errVkn; Firma = $firma; Year = $errorYearForRow; Month = $errorMonth
        VknPath = ''; PeriodName = ''; PeriodPath = ''; MonthPath = $e.Path
        YParca = 0; YBerat = 0; YGib = 0; KParca = 0; KBerat = 0; KGib = 0
        Donem = $(if ($errorYearForRow -and $errorMonth) { "$errorYearForRow-$errorMonth" } else { '-' })
        Aktif = $aktif; Bos = $true; Durum = 'Okunamadı'
        Aciklama = "Klasör okunamadı: $($e.Path) — $($e.Message)"
      })
    }
  }
  return $rows.ToArray()
}

function Merge-EkRows {
  param([array]$Rows)

  $merged = New-Object System.Collections.Generic.List[PSObject]
  foreach ($group in ($Rows | Group-Object { '{0}|{1}|{2}' -f $_.VKN, $_.Year, $_.Month })) {
    $items = @($group.Group)
    $dataRows = @($items | Where-Object { $_.Durum -eq 'Tamam' -or $_.Durum -eq 'Eksik' })
    $readErrors = @($items | Where-Object { $_.Durum -eq 'Okunamadı' })
    $structureIssues = @($items | Where-Object { $_.Durum -eq 'Uyuşmazlık' })
    $base = $items[0]
    if ($dataRows.Count -gt 0) { $base = $dataRows[0] }

    $details = @($items | ForEach-Object { $_.Aciklama } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $detailText = $details -join '; '
    $hasUnclassifiableXml = @($structureIssues | Where-Object {
      $_.Aciklama -match 'Tanınmayan XML|dosya VKN/TC|dosya yılı|dosya ayı|XML yok|GIB öneki'
    }).Count -gt 0

    if ($readErrors.Count -gt 0) {
      $sayim = 'Kontrol edilemedi'; $dosya = 'Okunamadı'; $sonuc = 'Kontrol sorunu'
    } elseif ($hasUnclassifiableXml) {
      $sayim = 'Hesaplanamadı'; $dosya = 'Uyuşmazlık'; $sonuc = 'Kontrol sorunu'
    } elseif ($dataRows.Count -gt 0 -and $dataRows[0].Durum -eq 'Eksik') {
      $sayim = 'Eksik/fazla'; $dosya = $(if ($structureIssues.Count -gt 0) { 'Uyuşmazlık' } else { 'Uygun' })
      $sonuc = $(if ($structureIssues.Count -gt 0) { 'Kontrol sorunu' } else { 'Eksik' })
    } elseif ($dataRows.Count -gt 0) {
      $sayim = 'Uyumlu'; $dosya = $(if ($structureIssues.Count -gt 0) { 'Uyuşmazlık' } else { 'Uygun' })
      $sonuc = $(if ($structureIssues.Count -gt 0) { 'Kontrol sorunu' } else { 'Tamam' })
    } else {
      $sayim = $(if ($structureIssues.Count -gt 0) { 'Hesaplanamadı' } else { 'Boş' })
      $dosya = $(if ($structureIssues.Count -gt 0) { 'Uyuşmazlık' } else { 'Boş' })
      if ($structureIssues.Count -gt 0) { $sonuc = 'Kontrol sorunu' } else { $sonuc = $base.Durum }
    }

    $musteriDurumu = if ($base.Aktif -eq $false) { 'Pasif' } else { 'Aktif' }
    if ($sonuc -eq 'Aktif/Boş' -or $sonuc -eq 'Pasif/Boş') { $sonuc = "$musteriDurumu/Boş" }
    else { $sonuc = "$musteriDurumu/$sonuc" }

    $merged.Add([PSCustomObject]@{
      VKN = $base.VKN; Firma = $base.Firma; Year = $base.Year; Month = $base.Month
      VknPath = $base.VknPath; PeriodName = $base.PeriodName; PeriodPath = $base.PeriodPath; MonthPath = $base.MonthPath
      YParca = $base.YParca; YBerat = $base.YBerat; YGib = $base.YGib
      KParca = $base.KParca; KBerat = $base.KBerat; KGib = $base.KGib
      Donem = $base.Donem; Aktif = $base.Aktif; Bos = ($dataRows.Count -eq 0)
      SayimDurum = $sayim; DosyaDurum = $dosya; Sonuc = $sonuc; Durum = $sonuc; Aciklama = $detailText
    })
  }
  return $merged.ToArray()
}

function Update-EkRows {
  Write-Host '⏳ Klasörler taranıyor...' -ForegroundColor Cyan
  $rawRows = @(Get-EkRows -Root $script:EkRoot -YearFilter $script:EkYearFilter -MonthFilter @($script:EkMonthFilter))
  $script:EkRows = @(Merge-EkRows -Rows $rawRows)
}

# Satır "boş" mu? (filtreye uyan e-defteri olmayan müşteri satırı)
function Test-EkBosRow {
  param($Row)
  return ($Row.Bos -eq $true -or $Row.Durum -eq 'Aktif/Boş' -or $Row.Durum -eq 'Pasif/Boş' -or $Row.Durum -eq 'Bulunamadı')
}

function Test-EkAktifRow {
  param($Row)
  return ($Row.Aktif -ne $false)
}

function Get-EkTotals {
  param([array]$Rows)
  $Rows = @($Rows)
  $sum = {
    param($prop)
    $r = ($Rows | Measure-Object -Property $prop -Sum).Sum
    if ($null -eq $r) { return 0 }
    return [int]$r
  }

  # Benzersiz müşteri sayıları her sonuç için ayrı sayılır; farklı aylardaki sonuçlar örtüşebilir.
  $aktifTamam = 0; $aktifEksik = 0; $aktifBos = 0; $aktifKontrol = 0; $aktifMusteri = 0
  $pasifTamam = 0; $pasifEksik = 0; $pasifBos = 0; $pasifKontrol = 0; $pasifMusteri = 0
  $customerRows = @($Rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.VKN) })
  foreach ($g in ($customerRows | Group-Object VKN)) {
    $grp = @($g.Group)
    $isAktif = (Test-EkAktifRow $grp[0])
    if ($isAktif) {
      $aktifMusteri++
      if (@($grp | Where-Object { $_.Durum -eq 'Aktif/Tamam' }).Count -gt 0) { $aktifTamam++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Aktif/Eksik' }).Count -gt 0) { $aktifEksik++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Aktif/Kontrol sorunu' }).Count -gt 0) { $aktifKontrol++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Aktif/Boş' }).Count -gt 0) { $aktifBos++ }
    } else {
      $pasifMusteri++
      if (@($grp | Where-Object { $_.Durum -eq 'Pasif/Tamam' }).Count -gt 0) { $pasifTamam++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Pasif/Eksik' }).Count -gt 0) { $pasifEksik++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Pasif/Kontrol sorunu' }).Count -gt 0) { $pasifKontrol++ }
      if (@($grp | Where-Object { $_.Durum -eq 'Pasif/Boş' }).Count -gt 0) { $pasifBos++ }
    }
  }

  $aktifTamamSatir = @($Rows | Where-Object { $_.Durum -eq 'Aktif/Tamam' }).Count
  $aktifEksikSatir = @($Rows | Where-Object { $_.Durum -eq 'Aktif/Eksik' }).Count
  $aktifKontrolSatir = @($Rows | Where-Object { $_.Durum -eq 'Aktif/Kontrol sorunu' }).Count
  $aktifBosSatir = @($Rows | Where-Object { $_.Durum -eq 'Aktif/Boş' }).Count
  $pasifTamamSatir = @($Rows | Where-Object { $_.Durum -eq 'Pasif/Tamam' }).Count
  $pasifEksikSatir = @($Rows | Where-Object { $_.Durum -eq 'Pasif/Eksik' }).Count
  $pasifKontrolSatir = @($Rows | Where-Object { $_.Durum -eq 'Pasif/Kontrol sorunu' }).Count
  $pasifBosSatir = @($Rows | Where-Object { $_.Durum -eq 'Pasif/Boş' }).Count
  $tamam = $aktifTamamSatir + $pasifTamamSatir
  $eksik = $aktifEksikSatir + $pasifEksikSatir
  $kontrol = $aktifKontrolSatir + $pasifKontrolSatir
  $musteri = @($customerRows | Select-Object -ExpandProperty VKN -Unique).Count
  $yP = (& $sum 'YParca'); $yB = (& $sum 'YBerat'); $yG = (& $sum 'YGib')
  $kP = (& $sum 'KParca'); $kB = (& $sum 'KBerat'); $kG = (& $sum 'KGib')
  $activeReportRows = @($Rows | Where-Object { $_.Aktif -ne $false })
  $passiveReportRows = @($Rows | Where-Object { $_.Aktif -eq $false })
  $activeCountRows = @($activeReportRows | Where-Object { -not (Test-EkBosRow $_) })
  $passiveCountRows = @($passiveReportRows | Where-Object { -not (Test-EkBosRow $_) })
  $sumForRows = {
    param([array]$items, [string]$property)
    $value = ($items | Measure-Object -Property $property -Sum).Sum
    if ($null -eq $value) { return 0 }
    return [int]$value
  }
  $aYP = (& $sumForRows $activeCountRows 'YParca'); $aYB = (& $sumForRows $activeCountRows 'YBerat'); $aYG = (& $sumForRows $activeCountRows 'YGib')
  $aKP = (& $sumForRows $activeCountRows 'KParca'); $aKB = (& $sumForRows $activeCountRows 'KBerat'); $aKG = (& $sumForRows $activeCountRows 'KGib')
  $pYP = (& $sumForRows $passiveCountRows 'YParca'); $pYB = (& $sumForRows $passiveCountRows 'YBerat'); $pYG = (& $sumForRows $passiveCountRows 'YGib')
  $pKP = (& $sumForRows $passiveCountRows 'KParca'); $pKB = (& $sumForRows $passiveCountRows 'KBerat'); $pKG = (& $sumForRows $passiveCountRows 'KGib')
  $activeUnknown = @($activeReportRows | Where-Object { $_.SayimDurum -eq 'Hesaplanamadı' -or $_.SayimDurum -eq 'Kontrol edilemedi' }).Count -gt 0
  $passiveUnknown = @($passiveReportRows | Where-Object { $_.SayimDurum -eq 'Hesaplanamadı' -or $_.SayimDurum -eq 'Kontrol edilemedi' }).Count -gt 0
  $aSagY = $null; $aSagK = $null; $pSagY = $null; $pSagK = $null
  if ($activeUnknown) { $aSagY = $false; $aSagK = $false }
  elseif ($activeCountRows.Count -gt 0) {
    $aSagY = (@($activeCountRows | Where-Object { $_.YParca -le 0 -or $_.YParca -ne $_.YBerat -or $_.YParca -ne $_.YGib }).Count -eq 0)
    $aSagK = (@($activeCountRows | Where-Object { $_.KParca -le 0 -or $_.KParca -ne $_.KBerat -or $_.KParca -ne $_.KGib }).Count -eq 0)
  }
  if ($passiveUnknown) { $pSagY = $false; $pSagK = $false }
  elseif ($passiveCountRows.Count -gt 0) {
    $pSagY = (@($passiveCountRows | Where-Object { $_.YParca -le 0 -or $_.YParca -ne $_.YBerat -or $_.YParca -ne $_.YGib }).Count -eq 0)
    $pSagK = (@($passiveCountRows | Where-Object { $_.KParca -le 0 -or $_.KParca -ne $_.KBerat -or $_.KParca -ne $_.KGib }).Count -eq 0)
  }
  $countRows = @($Rows | Where-Object { -not (Test-EkBosRow $_) })
  $countUnknown = @($Rows | Where-Object { $_.SayimDurum -eq 'Hesaplanamadı' -or $_.SayimDurum -eq 'Kontrol edilemedi' }).Count -gt 0
  $sagY = $null; $sagK = $null
  if ($countUnknown) { $sagY = $false; $sagK = $false }
  elseif ($countRows.Count -gt 0) {
    $sagY = (@($countRows | Where-Object { $_.YParca -le 0 -or $_.YParca -ne $_.YBerat -or $_.YParca -ne $_.YGib }).Count -eq 0)
    $sagK = (@($countRows | Where-Object { $_.KParca -le 0 -or $_.KParca -ne $_.KBerat -or $_.KParca -ne $_.KGib }).Count -eq 0)
  }

  return [PSCustomObject]@{
    RaporSatiri = $Rows.Count
    Tamam = $tamam; Eksik = $eksik; Kontrol = $kontrol
    AktifTamamSatir = $aktifTamamSatir; AktifEksikSatir = $aktifEksikSatir; AktifKontrolSatir = $aktifKontrolSatir; AktifBosSatir = $aktifBosSatir
    PasifTamamSatir = $pasifTamamSatir; PasifEksikSatir = $pasifEksikSatir; PasifKontrolSatir = $pasifKontrolSatir; PasifBosSatir = $pasifBosSatir
    Musteri = $musteri
    Aktif = $aktifMusteri; Pasif = $pasifMusteri
    AktifTamam = $aktifTamam; AktifEksik = $aktifEksik; AktifBos = $aktifBos; AktifKontrol = $aktifKontrol
    PasifTamam = $pasifTamam; PasifEksik = $pasifEksik; PasifBos = $pasifBos; PasifKontrol = $pasifKontrol
    AktifYParca = $aYP; AktifYBerat = $aYB; AktifYGib = $aYG; AktifKParca = $aKP; AktifKBerat = $aKB; AktifKGib = $aKG
    PasifYParca = $pYP; PasifYBerat = $pYB; PasifYGib = $pYG; PasifKParca = $pKP; PasifKBerat = $pKB; PasifKGib = $pKG
    AktifSagY = $aSagY; AktifSagK = $aSagK; PasifSagY = $pSagY; PasifSagK = $pSagK
    Bulunamadi = $aktifBos        # aktif olduğu halde e-defteri bulunamayan müşteri
    YParca = $yP; YBerat = $yB; YGib = $yG; KParca = $kP; KBerat = $kB; KGib = $kG
    YBeratFark = ($yB - $yP); YGibFark = ($yG - $yP); KBeratFark = ($kB - $kP); KGibFark = ($kG - $kP)
    SagY = $sagY
    SagK = $sagK
  }
}

# --------------------------------------------
# 📄 Rapor (ekran)
# --------------------------------------------
function Get-EkShort {
  param([string]$Text, [int]$Max)
  if ($null -eq $Text) { return '' }
  if ($Text.Length -le $Max) { return $Text }
  return $Text.Substring(0, $Max - 1) + '…'
}

# --------------------------------------------
# 📐 Konsol genişliğine uyum (ana menü pencereyi ekranın sağ yarısına yerleştirir; genişlik değişkendir)
# --------------------------------------------
function Get-EkConsoleWidth {
  if ($script:EkWidthOverride -gt 0) { return [int]$script:EkWidthOverride }
  try {
    $w = [int]$Host.UI.RawUI.WindowSize.Width
    if ($w -gt 0) { return $w }
  } catch { }
  return 100
}

function Get-EkLayout {
  $avail = (Get-EkConsoleWidth) - 2        # sağ kenarda satır sarmasın diye pay
  $numW = 6; $gibW = 11; $gibHeader = 'GİB Onaylı'; $durumW = 24
  $fixed = 12 + 8 + (2 * ($numW + $numW + $gibW)) + 2 + $durumW
  $firmaW = $avail - $fixed
  if ($firmaW -lt 12) {
    # Dar pencere: GİB başlığını kısalt
    $gibW = 5; $gibHeader = 'GİB'
    $fixed = 12 + 8 + (2 * ($numW + $numW + $gibW)) + 2 + $durumW
    $firmaW = $avail - $fixed
  }
  if ($firmaW -gt 22) { $firmaW = 22 }
  if ($firmaW -lt 10) { $firmaW = 10 }
  $fmt = '{0,-12}{1,-' + $firmaW + '}{2,-8}{3,' + $numW + '}{4,' + $numW + '}{5,' + $gibW + '}{6,' + $numW + '}{7,' + $numW + '}{8,' + $gibW + '}  '
  return [PSCustomObject]@{
    Fmt = $fmt; FirmaW = $firmaW; GibHeader = $gibHeader; Avail = $avail
    Width = (12 + $firmaW + 8 + (2 * ($numW + $numW + $gibW)) + 2 + $durumW)
    Compact = ($gibW -eq 5)
  }
}

function Write-EkDiffCell {
  param([int]$Value)
  $text = ('{0,10}' -f ('{0:+0;-0;0}' -f $Value))
  if ($Value -ne 0) { Write-Host $text -NoNewline -ForegroundColor Red } else { Write-Host $text -NoNewline }
}

function Write-EkResultCell {
  param($Ok)
  if ($null -eq $Ok) { Write-Host ('{0,12}' -f '—') -NoNewline -ForegroundColor DarkGray; return }
  $text = 'TUTUYOR'
  $color = 'Green'
  if (-not $Ok) { $text = 'TUTMUYOR'; $color = 'Red' }
  Write-Host ('{0,12}' -f $text) -NoNewline -ForegroundColor $color
}

function Get-EkAuditStatusColor {
  param([string]$Value)
  switch ($Value) {
    { $_ -in @('Uyumlu','Uygun') } { return 'Green' }
    { $_ -in @('Eksik/fazla','Uyuşmazlık') } { return 'Red' }
    { $_ -in @('Hesaplanamadı','Kontrol edilemedi','Okunamadı') } { return 'DarkYellow' }
    default { return 'DarkGray' }
  }
}

function Get-EkAuditStatusCssClass {
  param([string]$Value)
  switch ($Value) {
    { $_ -in @('Uyumlu','Uygun') } { return 'ok' }
    { $_ -in @('Eksik/fazla','Uyuşmazlık') } { return 'bad' }
    { $_ -in @('Hesaplanamadı','Kontrol edilemedi','Okunamadı') } { return 'yok' }
    default { return 'pas' }
  }
}

function Format-EkHeaderCell {
  param([string]$Text, [int]$Width)
  if ($Text.Length -ge $Width) { return $Text }
  return $Text.PadRight($Width)
}

function Write-EkCountCell {
  param([int]$Value, [int]$Width, [ConsoleColor]$Color)
  $text = ('{0,' + $Width + '}') -f $Value
  Write-Host $text -NoNewline -ForegroundColor $Color
}

function Write-EkOutcomeSummaryRow {
  param([string]$Label, [int]$Customers, [int]$Rows, [ConsoleColor]$Color)
  Write-Host ('  {0,-30} │ ' -f $Label) -NoNewline -ForegroundColor $Color
  Write-EkCountCell $Customers 16 $Color
  Write-Host ' │ ' -NoNewline
  Write-EkCountCell $Rows 18 $Color
  Write-Host ''
}

function Write-EkLedgerSummaryRow {
  param([string]$Label, [int]$ActiveJournal, [int]$ActiveLedger, [int]$PassiveJournal, [int]$PassiveLedger)
  Write-Host ('  {0,-26} │ ' -f $Label) -NoNewline
  Write-Host ('{0,12}' -f $ActiveJournal) -NoNewline
  Write-Host ' │ ' -NoNewline
  Write-Host ('{0,12}' -f $ActiveLedger) -NoNewline
  Write-Host ' │ ' -NoNewline
  Write-Host ('{0,12}' -f $PassiveJournal) -NoNewline
  Write-Host ' │ ' -NoNewline
  Write-Host ('{0,12}' -f $PassiveLedger)
}

# Alt özet: sonuç dağılımı ve defter sağlaması
function Write-EkSummary {
  param($Tot)
  $t = $Tot

  Write-Host 'SONUÇ DAĞILIMI' -ForegroundColor Yellow
  Write-Host '  Benzersiz müşteri ve müşteri-ay satırı sayıları (müşteriler farklı aylarda birden çok duruma girebilir).' -ForegroundColor DarkGray
  Write-Host ('  ' + (Format-EkHeaderCell 'Sonuç' 30) + ' │ ' + (Format-EkHeaderCell 'Benzersiz müşteri' 16) + ' │ ' + (Format-EkHeaderCell 'Müşteri-ay satırı' 18)) -ForegroundColor Cyan
  Write-EkOutcomeSummaryRow 'Aktif/Tamam' $t.AktifTamam $t.AktifTamamSatir Green
  Write-EkOutcomeSummaryRow 'Aktif/Eksik' $t.AktifEksik $t.AktifEksikSatir Red
  Write-EkOutcomeSummaryRow 'Aktif/Kontrol sorunu' $t.AktifKontrol $t.AktifKontrolSatir Red
  Write-EkOutcomeSummaryRow 'Aktif/Boş' $t.AktifBos $t.AktifBosSatir Magenta
  Write-Host ('  {0,-30} │ ' -f 'Toplam aktif müşteri') -NoNewline
  Write-EkCountCell $t.Aktif 16 White
  Write-Host ' │ ' -NoNewline; Write-Host ('{0,18}' -f '—')
  Write-Host ('  ' + ('─' * 70)) -ForegroundColor DarkGray
  Write-EkOutcomeSummaryRow 'Pasif/Tamam' $t.PasifTamam $t.PasifTamamSatir Green
  Write-EkOutcomeSummaryRow 'Pasif/Eksik' $t.PasifEksik $t.PasifEksikSatir Red
  Write-EkOutcomeSummaryRow 'Pasif/Kontrol sorunu' $t.PasifKontrol $t.PasifKontrolSatir Red
  Write-EkOutcomeSummaryRow 'Pasif/Boş' $t.PasifBos $t.PasifBosSatir DarkGray
  Write-Host ('  {0,-30} │ ' -f 'Toplam pasif müşteri') -NoNewline
  Write-EkCountCell $t.Pasif 16 White
  Write-Host ' │ ' -NoNewline; Write-Host ('{0,18}' -f '—')
  Write-Host ('  ' + ('─' * 70)) -ForegroundColor DarkGray
  Write-Host ('  {0,-30} │ ' -f 'Toplam benzersiz müşteri') -NoNewline
  Write-EkCountCell $t.Musteri 16 White
  Write-Host ' │ ' -NoNewline; Write-Host ('{0,18}' -f '—')
  Write-Host ('  {0,-30} │ ' -f 'Toplam müşteri-ay satırı') -NoNewline
  Write-Host ('{0,16}' -f '—') -NoNewline; Write-Host ' │ ' -NoNewline
  Write-EkCountCell $t.RaporSatiri 18 White
  Write-Host ''
  Write-Host ''

  Write-Host 'DEFTER SAĞLAMASI (müşteri-ay satırları)' -ForegroundColor Yellow
  Write-Host ('  ' + (Format-EkHeaderCell '' 26) + ' │ ' + (Format-EkHeaderCell 'Aktif Yev.' 12) + ' │ ' + (Format-EkHeaderCell 'Aktif Keb.' 12) + ' │ ' + (Format-EkHeaderCell 'Pasif Yev.' 12) + ' │ ' + (Format-EkHeaderCell 'Pasif Keb.' 12)) -ForegroundColor Cyan
  Write-EkLedgerSummaryRow 'Parça sayısı' $t.AktifYParca $t.AktifKParca $t.PasifYParca $t.PasifKParca
  Write-EkLedgerSummaryRow 'Berat sayısı' $t.AktifYBerat $t.AktifKBerat $t.PasifYBerat $t.PasifKBerat
  Write-EkLedgerSummaryRow 'GİB onaylı berat sayısı' $t.AktifYGib $t.AktifKGib $t.PasifYGib $t.PasifKGib
  Write-EkLedgerSummaryRow 'Berat - Parça farkı' ($t.AktifYBerat - $t.AktifYParca) ($t.AktifKBerat - $t.AktifKParca) ($t.PasifYBerat - $t.PasifYParca) ($t.PasifKBerat - $t.PasifKParca)
  Write-EkLedgerSummaryRow 'GİB - Parça farkı' ($t.AktifYGib - $t.AktifYParca) ($t.AktifKGib - $t.AktifKParca) ($t.PasifYGib - $t.PasifYParca) ($t.PasifKGib - $t.PasifKParca)
  Write-Host ('  {0,-26} │ ' -f 'Sağlama') -NoNewline
  Write-EkResultCell $t.AktifSagY; Write-Host ' │ ' -NoNewline
  Write-EkResultCell $t.AktifSagK; Write-Host ' │ ' -NoNewline
  Write-EkResultCell $t.PasifSagY; Write-Host ' │ ' -NoNewline
  Write-EkResultCell $t.PasifSagK; Write-Host ''
  Write-Host '  Not: Bir müşterinin birden çok parçası olabilir, bu yüzden parça sayısı' -ForegroundColor DarkGray
  Write-Host '  müşteri sayısına eşit olmak zorunda değildir. Sağlama: Parça = Berat = GİB.' -ForegroundColor DarkGray
  Write-Host ''
}

function Write-EkReportTable {
  param([array]$Rows)

  $Rows = @($Rows)
  $tot = Get-EkTotals -Rows $Rows
  $L = Get-EkLayout
  $sayimW = 'Sayım'.Length
  $dosyaW = 'Dosya/Klasör'.Length
  $sonucW = 'Sonuç'.Length
  foreach ($r in $Rows) {
    if ([string]$r.SayimDurum -and ([string]$r.SayimDurum).Length -gt $sayimW) { $sayimW = ([string]$r.SayimDurum).Length }
    if ([string]$r.DosyaDurum -and ([string]$r.DosyaDurum).Length -gt $dosyaW) { $dosyaW = ([string]$r.DosyaDurum).Length }
    if ([string]$r.Sonuc -and ([string]$r.Sonuc).Length -gt $sonucW) { $sonucW = ([string]$r.Sonuc).Length }
  }
  $sayimW++; $dosyaW++; $sonucW++
  $tailFmt = '{0,-' + $sayimW + '} │ {1,-' + $dosyaW + '} │ {2,-' + $sonucW + '} │ {3}'
  $sayimFmt = '{0,-' + $sayimW + '}'
  $dosyaFmt = '{0,-' + $dosyaW + '}'
  $sonucFmt = '{0,-' + $sonucW + '}'
  $yearText = $script:EkYearFilter
  if ([string]::IsNullOrEmpty($yearText)) { $yearText = 'Tümü' }

  Write-Host '📋 E-DEFTER KONTROL RAPORU' -ForegroundColor Magenta
  Write-Host "Klasör: $($script:EkRoot)"
  Write-Host "Yıl: $yearText | Ay: $(Get-EkMonthFilterText) | Tarih: $((Get-Date).ToString('dd.MM.yyyy HH:mm'))"
  if ($L.Compact) { Write-Host 'GİB: GİB onaylı berat sayısı' -ForegroundColor DarkGray }
  Write-Host ''

  # Başlık: grup etiketleri "Parça" sütununun ilk harfiyle aynı hizada başlar
  $hdr = $L.Fmt -f 'VKN / TC', 'Firma', 'Dönem', 'Parça', 'Berat', $L.GibHeader, 'Parça', 'Berat', $L.GibHeader
  $idxY = $hdr.IndexOf('Parça', [System.StringComparison]::Ordinal)
  $idxK = $hdr.IndexOf('Parça', $idxY + 1, [System.StringComparison]::Ordinal)
  $grpLine = ((' ' * $idxY) + 'YEVMİYE DEFTERİ').PadRight($idxK) + 'KEBİR DEFTERİ'
  Write-Host $grpLine -ForegroundColor Cyan
  $headerLine = $hdr + ($tailFmt -f 'Sayım', 'Dosya/Klasör', 'Sonuç', 'Açıklama')
  Write-Host $headerLine -ForegroundColor Cyan
  $rule = '─' * $headerLine.Length
  Write-Host $rule

  foreach ($r in $Rows) {
    if (Test-EkBosRow $r) { $v = @('-', '-', '-', '-', '-', '-') }
    else { $v = @($r.YParca, $r.YBerat, $r.YGib, $r.KParca, $r.KBerat, $r.KGib) }
    $text = $L.Fmt -f $r.VKN, (Get-EkShort -Text $r.Firma -Max ($L.FirmaW - 1)), $r.Donem, $v[0], $v[1], $v[2], $v[3], $v[4], $v[5]
    Write-Host $text -NoNewline
    Write-Host ($sayimFmt -f $r.SayimDurum) -NoNewline -ForegroundColor (Get-EkAuditStatusColor $r.SayimDurum)
    Write-Host ' │ ' -NoNewline
    Write-Host ($dosyaFmt -f $r.DosyaDurum) -NoNewline -ForegroundColor (Get-EkAuditStatusColor $r.DosyaDurum)
    Write-Host ' │ ' -NoNewline
    if ($r.Sonuc -match '/Tamam$') { Write-Host ($sonucFmt -f $r.Sonuc) -NoNewline -ForegroundColor Green }
    elseif ($r.Sonuc -match '/Eksik$|/Kontrol sorunu$') { Write-Host ($sonucFmt -f $r.Sonuc) -NoNewline -ForegroundColor Red }
    elseif ($r.Sonuc -eq 'Aktif/Boş') { Write-Host ($sonucFmt -f $r.Sonuc) -NoNewline -ForegroundColor Magenta }
    elseif ($r.Sonuc -eq 'Pasif/Boş') { Write-Host ($sonucFmt -f $r.Sonuc) -NoNewline -ForegroundColor DarkGray }
    else { Write-Host ($sonucFmt -f $r.Sonuc) -NoNewline -ForegroundColor DarkYellow }
    Write-Host ' │ ' -NoNewline
    Write-Host $r.Aciklama -ForegroundColor DarkYellow
  }

  Write-Host $rule
  Write-Host ($L.Fmt -f 'TOPLAM', '', "$($tot.RaporSatiri) satır", $tot.YParca, $tot.YBerat, $tot.YGib, $tot.KParca, $tot.KBerat, $tot.KGib) -NoNewline -ForegroundColor White
  Write-Host ($tailFmt -f '', '', '', 'Sonuç dağılımı alt özette') -ForegroundColor White
  Write-Host ''
  Write-EkSummary -Tot $tot
}

# --------------------------------------------
# 🌐 Rapor (HTML görüntüleme)
# --------------------------------------------
function Export-EkReportHtml {
  param([array]$Rows)
  try {
    $Rows = @($Rows)
    $tot = Get-EkTotals -Rows $Rows
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }
    $okTxt = { param($ok) if ($null -eq $ok) { '<span class="pas">—</span>' } elseif ($ok) { '<span class="ok">TUTUYOR</span>' } else { '<span class="bad">TUTMUYOR</span>' } }
    $okCountHtml = { param($n) '<span class="ok">' + $n + '</span>' }
    $badCountHtml = { param($n) '<span class="bad">' + $n + '</span>' }
    $activeEmptyCountHtml = { param($n) '<span class="aktif-bos">' + $n + '</span>' }
    $passiveEmptyCountHtml = { param($n) '<span class="pas">' + $n + '</span>' }
    $fark  = { param($v) $t = ('{0:+0;-0;0}' -f [int]$v); if ([int]$v -ne 0) { '<span class="bad">' + $t + '</span>' } else { $t } }

    $yearText = $script:EkYearFilter
    if ([string]::IsNullOrEmpty($yearText)) { $yearText = 'Tümü' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="tr"><head><meta charset="utf-8"><title>E-Defter Kontrol Raporu</title>')
    [void]$sb.AppendLine(@'
<style>
  body { font-family: "Segoe UI", Arial, sans-serif; margin: 16px; color: #111; }
  h1 { font-size: 18px; margin: 0 0 4px 0; }
  .meta { font-size: 12px; color: #444; margin-bottom: 10px; }
  table { border-collapse: collapse; width: 100%; font-size: 12px; }
  th, td { border: 1px solid #888; padding: 3px 6px; }
  th { background: #e8e8e8; text-align: center; }
  td.n { text-align: right; }
  td.c { text-align: center; white-space: nowrap; }
  .ok  { color: #0a7a2f; font-weight: bold; }
  .bad { color: #c00000; font-weight: bold; }
  .yok { color: #8a6d00; font-weight: bold; }
  .aktif-bos { color: #c00080; font-weight: bold; }
  .pas { color: #555; font-weight: bold; }
  .acik { font-size: 11px; color: #7a4b00; }
  .tanimsiz { color: #888; font-style: italic; }
  tfoot td { font-weight: bold; background: #f3f3f3; }
  thead { display: table-header-group; }
  tfoot { display: table-row-group; }
  tr { page-break-inside: avoid; }
  .sums { display: flex; flex-wrap: wrap; gap: 24px; margin-top: 14px; align-items: flex-start; page-break-inside: avoid; }
  .sums table { width: auto; font-size: 12px; }
  .sums caption { text-align: left; font-weight: bold; margin-bottom: 3px; }
  .sums td, .sums th { padding: 2px 8px; }
  .sums th { text-align: left; }
  .sums tr.group-separator td { border-top: 2px solid #555; padding: 0; height: 3px; }
  .note { font-size: 11px; color: #555; margin-top: 8px; }
  @page { size: A4 landscape; margin: 10mm; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
</style></head><body>
'@)
    [void]$sb.AppendLine('<h1>E-Defter Kontrol Raporu</h1>')
    [void]$sb.AppendLine('<div class="meta">Klasör: ' + (& $enc $script:EkRoot) + ' | Yıl: ' + (& $enc $yearText) +
      ' | Ay: ' + (& $enc (Get-EkMonthFilterText)) + ' | Tarih: ' + (Get-Date).ToString('dd.MM.yyyy HH:mm') + '</div>')

    [void]$sb.AppendLine('<table><thead>')
    [void]$sb.AppendLine('<tr><th rowspan="2">VKN / TC</th><th rowspan="2">Firma</th><th rowspan="2">Dönem</th>' +
      '<th colspan="3">Yevmiye Defteri</th><th colspan="3">Kebir Defteri</th><th rowspan="2">Sayım</th><th rowspan="2">Dosya/Klasör</th><th rowspan="2">Sonuç</th><th rowspan="2">Açıklama</th></tr>')
    [void]$sb.AppendLine('<tr><th>Parça Sayısı</th><th>Berat Sayısı</th><th>GİB Onaylı Berat Sayısı</th>' +
      '<th>Parça Sayısı</th><th>Berat Sayısı</th><th>GİB Onaylı Berat Sayısı</th></tr>')
    [void]$sb.AppendLine('</thead><tbody>')

    foreach ($r in $Rows) {
      if ($r.Sonuc -match '/Tamam$') { $cls = 'ok' } elseif ($r.Sonuc -match '/Eksik$|/Kontrol sorunu$') { $cls = 'bad' } elseif ($r.Sonuc -eq 'Pasif/Boş') { $cls = 'pas' } elseif ($r.Sonuc -eq 'Aktif/Boş') { $cls = 'aktif-bos' } else { $cls = 'yok' }
      if (Test-EkBosRow $r) { $c = @('-','-','-','-','-','-') } else { $c = @($r.YParca, $r.YBerat, $r.YGib, $r.KParca, $r.KBerat, $r.KGib) }
      if ($r.Firma -eq $script:EkNoName) { $firmaTd = '<td class="tanimsiz">' } else { $firmaTd = '<td>' }
      $sayimCls = Get-EkAuditStatusCssClass $r.SayimDurum
      $dosyaCls = Get-EkAuditStatusCssClass $r.DosyaDurum
      [void]$sb.AppendLine('<tr><td>' + (& $enc $r.VKN) + '</td>' + $firmaTd + (& $enc $r.Firma) + '</td><td class="c">' + (& $enc $r.Donem) + '</td>' +
        '<td class="n">' + $c[0] + '</td><td class="n">' + $c[1] + '</td><td class="n">' + $c[2] + '</td>' +
        '<td class="n">' + $c[3] + '</td><td class="n">' + $c[4] + '</td><td class="n">' + $c[5] + '</td>' +
        '<td class="c ' + $sayimCls + '">' + (& $enc $r.SayimDurum) + '</td><td class="c ' + $dosyaCls + '">' + (& $enc $r.DosyaDurum) + '</td>' +
        '<td class="c ' + $cls + '">' + (& $enc $r.Sonuc) + '</td><td class="acik">' + (& $enc $r.Aciklama) + '</td></tr>')
    }
    [void]$sb.AppendLine('</tbody><tfoot>')
    [void]$sb.AppendLine('<tr><td colspan="3">TOPLAM (' + $tot.RaporSatiri + ' müşteri-ay satırı)</td>' +
      '<td class="n">' + $tot.YParca + '</td><td class="n">' + $tot.YBerat + '</td><td class="n">' + $tot.YGib + '</td>' +
      '<td class="n">' + $tot.KParca + '</td><td class="n">' + $tot.KBerat + '</td><td class="n">' + $tot.KGib + '</td>' +
      '<td class="c" colspan="4">Sonuç dağılımı alt özette</td></tr>')
    [void]$sb.AppendLine('</tfoot></table>')

    # Alt özet: müşteri dağılımı, ay dağılımı, defter sağlaması
    [void]$sb.AppendLine('<div class="sums">')
    [void]$sb.AppendLine('<table><caption>Sonuç Dağılımı</caption>' +
      '<tr><th>Sonuç</th><th>Benzersiz müşteri</th><th>Müşteri-ay satırı</th></tr>' +
      '<tr><td class="ok">Aktif/Tamam</td><td class="n">' + (& $okCountHtml $tot.AktifTamam) + '</td><td class="n">' + (& $okCountHtml $tot.AktifTamamSatir) + '</td></tr>' +
      '<tr><td class="bad">Aktif/Eksik</td><td class="n">' + (& $badCountHtml $tot.AktifEksik) + '</td><td class="n">' + (& $badCountHtml $tot.AktifEksikSatir) + '</td></tr>' +
      '<tr><td class="bad">Aktif/Kontrol sorunu</td><td class="n">' + (& $badCountHtml $tot.AktifKontrol) + '</td><td class="n">' + (& $badCountHtml $tot.AktifKontrolSatir) + '</td></tr>' +
      '<tr><td class="aktif-bos">Aktif/Boş</td><td class="n">' + (& $activeEmptyCountHtml $tot.AktifBos) + '</td><td class="n">' + (& $activeEmptyCountHtml $tot.AktifBosSatir) + '</td></tr>' +
      '<tr><td><b>Toplam aktif müşteri</b></td><td class="n"><b>' + $tot.Aktif + '</b></td><td class="n">—</td></tr>' +
      '<tr class="group-separator"><td colspan="3"></td></tr>' +
      '<tr><td class="ok">Pasif/Tamam</td><td class="n">' + (& $okCountHtml $tot.PasifTamam) + '</td><td class="n">' + (& $okCountHtml $tot.PasifTamamSatir) + '</td></tr>' +
      '<tr><td class="bad">Pasif/Eksik</td><td class="n">' + (& $badCountHtml $tot.PasifEksik) + '</td><td class="n">' + (& $badCountHtml $tot.PasifEksikSatir) + '</td></tr>' +
      '<tr><td class="bad">Pasif/Kontrol sorunu</td><td class="n">' + (& $badCountHtml $tot.PasifKontrol) + '</td><td class="n">' + (& $badCountHtml $tot.PasifKontrolSatir) + '</td></tr>' +
      '<tr><td class="pas">Pasif/Boş</td><td class="n">' + (& $passiveEmptyCountHtml $tot.PasifBos) + '</td><td class="n">' + (& $passiveEmptyCountHtml $tot.PasifBosSatir) + '</td></tr>' +
      '<tr><td><b>Toplam pasif müşteri</b></td><td class="n"><b>' + $tot.Pasif + '</b></td><td class="n">—</td></tr>' +
      '<tr class="group-separator"><td colspan="3"></td></tr>' +
      '<tr><td><b>Toplam benzersiz müşteri</b></td><td class="n"><b>' + $tot.Musteri + '</b></td><td class="n">—</td></tr>' +
      '<tr><td><b>Toplam müşteri-ay satırı</b></td><td class="n">—</td><td class="n"><b>' + $tot.RaporSatiri + '</b></td></tr></table>')
    [void]$sb.AppendLine('<table><caption>Defter Sağlaması (müşteri-ay satırları)</caption>' +
      '<tr><th></th><th>Aktif Yevmiye</th><th>Aktif Kebir</th><th>Pasif Yevmiye</th><th>Pasif Kebir</th></tr>' +
      '<tr><td>Parça sayısı</td><td class="n">' + $tot.AktifYParca + '</td><td class="n">' + $tot.AktifKParca + '</td><td class="n">' + $tot.PasifYParca + '</td><td class="n">' + $tot.PasifKParca + '</td></tr>' +
      '<tr><td>Berat sayısı</td><td class="n">' + $tot.AktifYBerat + '</td><td class="n">' + $tot.AktifKBerat + '</td><td class="n">' + $tot.PasifYBerat + '</td><td class="n">' + $tot.PasifKBerat + '</td></tr>' +
      '<tr><td>GİB onaylı berat sayısı</td><td class="n">' + $tot.AktifYGib + '</td><td class="n">' + $tot.AktifKGib + '</td><td class="n">' + $tot.PasifYGib + '</td><td class="n">' + $tot.PasifKGib + '</td></tr>' +
      '<tr><td>Berat - Parça farkı</td><td class="n">' + (& $fark ($tot.AktifYBerat - $tot.AktifYParca)) + '</td><td class="n">' + (& $fark ($tot.AktifKBerat - $tot.AktifKParca)) + '</td><td class="n">' + (& $fark ($tot.PasifYBerat - $tot.PasifYParca)) + '</td><td class="n">' + (& $fark ($tot.PasifKBerat - $tot.PasifKParca)) + '</td></tr>' +
      '<tr><td>GİB - Parça farkı</td><td class="n">' + (& $fark ($tot.AktifYGib - $tot.AktifYParca)) + '</td><td class="n">' + (& $fark ($tot.AktifKGib - $tot.AktifKParca)) + '</td><td class="n">' + (& $fark ($tot.PasifYGib - $tot.PasifYParca)) + '</td><td class="n">' + (& $fark ($tot.PasifKGib - $tot.PasifKParca)) + '</td></tr>' +
      '<tr><td><b>Sağlama</b></td><td class="c">' + (& $okTxt $tot.AktifSagY) + '</td><td class="c">' + (& $okTxt $tot.AktifSagK) + '</td><td class="c">' + (& $okTxt $tot.PasifSagY) + '</td><td class="c">' + (& $okTxt $tot.PasifSagK) + '</td></tr></table>')
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<div class="note">Not: Benzersiz müşteri sayıları farklı aylardaki sonuçlar nedeniyle örtüşebilir. Müşteri-ay satırı sonuçları toplanarak rapor satırı sayısıyla sağlanır. Bir müşterinin birden çok parçası olabilir; defter sağlaması satır bazındadır.</div>')
    [void]$sb.AppendLine('</body></html>')

    if (-not (Test-Path -LiteralPath $script:EkReportDir)) {
      [void](New-Item -ItemType Directory -Path $script:EkReportDir -Force -ErrorAction Stop)
    }
    $file = Join-Path $script:EkReportDir ('EdefterKontrol_' + (Get-Date).ToString('yyyyMMdd_HHmmss_fff') + '.html')
    [System.IO.File]::WriteAllText($file, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    Start-Process -FilePath $file
    return $file
  } catch {
    Write-Warning "🚫 Rapor hazırlanamadı: $($_.Exception.Message)"
    return $null
  }
}

# --------------------------------------------
# 🎛️ Rapor sorgusu: yıl ve ay filtreleri (boş = tümü)
# --------------------------------------------
function Read-EkYearFilter {
  while ($true) {
    $in = (Read-Host 'Yıl (örn: 2025, boş = tümü)').Trim()
    if ($in -eq '') { return '' }
    if ($in -match '^\d{4}$' -and [int]$in -ge 1900 -and [int]$in -le 2100) { return $in }
    Write-Warning '⚠️ Geçersiz yıl. 4 haneli bir yıl girin veya boş bırakın.'
  }
}

function Read-EkMonthFilter {
  while ($true) {
    $in = (Read-Host 'Ay (örn: 12 veya 01,03,12, boş = tümü)').Trim()
    if ($in -eq '') { return @() }
    $aylar = @()
    $gecerli = $true
    foreach ($part in ($in -split ',')) {
      $s = $part.Trim()
      if ($s -match '^\d{1,2}$' -and [int]$s -ge 1 -and [int]$s -le 12) {
        $aylar += ('{0:D2}' -f [int]$s)
      } else {
        $gecerli = $false
      }
    }
    if ($gecerli -and $aylar.Count -gt 0) { return @($aylar | Sort-Object -Unique) }
    Write-Warning '⚠️ Geçersiz ay. 1-12 arası sayı girin (birden fazlası virgülle) veya boş bırakın.'
  }
}

# Tarama öncesi seçili yılın dönem ve ay klasörlerini kesin biçime göre doğrula.
# Erişim hatalarını burada sahiplenme; normal tarama bunları müşteri-ay satırına yazar.
function Get-EkStructureConflicts {
  param([string]$Root, [string]$YearFilter, [array]$MonthFilter)

  $conflicts = New-Object System.Collections.Generic.List[string]
  if ([string]::IsNullOrWhiteSpace($YearFilter)) { return @() }
  $expectedPeriodName = "01.01.$YearFilter-31.12.$YearFilter"

  try { $vknDirs = @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^\d{10,11}$' }) }
  catch { return @() }

  foreach ($vknDir in $vknDirs) {
    try { $periodDirs = @(Get-ChildItem -LiteralPath $vknDir.FullName -Directory -ErrorAction Stop) }
    catch { continue }

    $yearPeriods = @($periodDirs | Where-Object {
      $pm = [regex]::Match($_.Name, '^(?<sd>\d{2})\.(?<sm>\d{2})\.(?<sy>\d{4})-(?<ed>\d{2})\.(?<em>\d{2})\.(?<ey>\d{4})$')
      $coversYear = $false
      if ($pm.Success) {
        $coversYear = ([int]$YearFilter -ge [int]$pm.Groups['sy'].Value -and [int]$YearFilter -le [int]$pm.Groups['ey'].Value)
      }
      $coversYear -or ($_.Name -match ('(?<!\d)' + [regex]::Escape($YearFilter) + '(?!\d)'))
    })

    $invalidYearPeriods = @($yearPeriods | Where-Object { $_.Name -cne $expectedPeriodName })
    if ($invalidYearPeriods.Count -gt 0 -or $yearPeriods.Count -gt 1) {
      $names = ($yearPeriods | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
      $conflicts.Add("VKN $($vknDir.Name): $YearFilter yılı için dönem klasörü tam olarak '$expectedPeriodName' olmalı. Bulunan: $names")
      continue
    }
    if ($yearPeriods.Count -eq 0) { continue } # Klasörün olmaması boş müşteri-ay olarak raporlanabilir.

    try { $yearEntries = @(Get-ChildItem -LiteralPath $yearPeriods[0].FullName -Force -ErrorAction Stop) }
    catch { continue }

    $invalidYearEntries = @($yearEntries | Where-Object { -not $_.PSIsContainer -or $_.Name -cnotmatch '^(0[1-9]|1[0-2])$' })
    if ($invalidYearEntries.Count -gt 0) {
      $names = ($invalidYearEntries | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
      $conflicts.Add("VKN $($vknDir.Name), ${expectedPeriodName}: yalnızca 01-12 adlı ay klasörleri bulunabilir. Uymayan içerik: $names")
    }
  }
  return $conflicts.ToArray()
}

function Start-EkReport {
  if ([string]::IsNullOrEmpty($script:EkRoot) -or -not (Test-Path -LiteralPath $script:EkRoot -PathType Container)) {
    Write-Warning '🚫 E-Defter ana klasörü tanımlı değil veya bulunamıyor. Ayarlar menüsünden tanımlayın.'
    Wait-EkKey
    return
  }

  while ($true) {
    Clear-EkScreen
    Write-Host '📊 Kontrol Raporu' -ForegroundColor Yellow
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host "📂 $($script:EkRoot)"
    Write-Host ''
    $script:EkYearFilter  = Read-EkYearFilter
    $script:EkMonthFilter = @(Read-EkMonthFilter)
    Write-Host ''

    $conflicts = @(Get-EkStructureConflicts -Root $script:EkRoot -YearFilter $script:EkYearFilter -MonthFilter @($script:EkMonthFilter))
    if ($conflicts.Count -gt 0) {
      Write-Host '⚠️ Klasör yapısı belirsiz. Tarama başlatılmadı:' -ForegroundColor Red
      foreach ($conflict in $conflicts) { Write-Host "  $conflict" -ForegroundColor Red }
      Write-Host 'Düzeltmeden sonra yeniden sorgulayın. Ana menüye dönülüyor.' -ForegroundColor Yellow
      Wait-EkKey
      return
    }

    Update-EkRows
    $next = Show-EkReport
    if ($next -ne 'N') { return }
  }
}

# --------------------------------------------
# 📄 Rapor ekranı: H = HTML görüntüle, N = yeni sorgu, Q = ana menü
# --------------------------------------------
function Show-EkReport {
  $rows = @($script:EkRows | Sort-Object VKN, Year, Month)

  # Ekranı yalnızca ilk girişte çiz; desteklenmeyen tuşları beklerken aynı
  # raporun gereksiz yere temizlenip yeniden yazılmasını önle.
  $renderReport = $true
  while ($true) {
    if ($renderReport) {
    Clear-EkScreen
    if ($rows.Count -eq 0) {
      Write-Host 'Raporlanacak kayıt yok (filtreyi veya klasörü kontrol edin).' -ForegroundColor DarkYellow
      Write-Host ''
    } else {
      Write-EkReportTable -Rows $rows
    }
    Write-Host '📌 H: HTML Görüntüle | N: Yeni Sorgu | Q / Backspace: Ana Menü' -ForegroundColor DarkGray
      $renderReport = $false
    }

    $key = Read-EkMenuKey
    switch ($key) {
      72 {
        if ($rows.Count -gt 0) {
          $path = Export-EkReportHtml -Rows $rows
          if ($path) {
            Write-Host "✅ Rapor hazırlandı: $path" -ForegroundColor Green
            Write-Host 'Rapor tarayıcıda açıldı. İsterseniz tarayıcı menüsünden yazdırabilirsiniz.' -ForegroundColor DarkGray
          }
          Wait-EkKey
          $renderReport = $true
        }
      }
      78 { return 'N' }
      81 { return 'Q' }
      8  { return 'Q' }
      default { }
    }
  }
}

# --------------------------------------------
# ⚙️ Ayarlar (e-defter kaynak klasörü ve firma adları)
# --------------------------------------------
function Show-EkSettingsMenu {
  while ($true) {
    Clear-EkScreen
    Write-Host ('-' * 72) -ForegroundColor Magenta
    Write-Host '⚙️ --- E-DEFTER KONTROL AYARLARI ---' -ForegroundColor Magenta
    Write-Host ('-' * 72) -ForegroundColor Magenta
    Write-Host ''
    $mevcut = $script:EkRootText
    if ([string]::IsNullOrEmpty($mevcut)) { $mevcut = 'Tanımlı Değil' }
    Write-Host "📂 E-Defter klasörü : $mevcut"
    Write-Host "👥 Firma adları     : $(Get-EkFirmaSummaryText)"
    Write-Host ''
    Write-Host '  1. 📂  E-Defter Ana Klasörünü Değiştir'
    Write-Host '  2. 👥  Firma Adları (listele / ekle / güncelle / sil)'
    Write-Host '  Q  ↩️  Ana Menü'
    Write-Host ''
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host ''
    $choice = Read-Host 'Seçiminiz'
    switch ($choice.ToUpper()) {
      '1' { Edit-EkRoot }
      '2' { Show-EkFirmaMenu }
      'Q' { return }
      default {
        Write-Host '[!] Geçersiz seçim.' -ForegroundColor Red
        Wait-EkKey
      }
    }
  }
}

# --------------------------------------------
# 🏠 Ana menü
# --------------------------------------------
function Show-EkMainMenu {
  while ($true) {
    Clear-EkScreen
    Write-Host ('-' * 72) -ForegroundColor Magenta
    Write-Host '📋 E-DEFTER KONTROL PROGRAMI' -ForegroundColor Magenta
    Write-Host ('-' * 72) -ForegroundColor Magenta
    Write-Host ''
    Write-Host '📌 YAPILANDIRMA DURUMU' -ForegroundColor Yellow
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $mevcut = $script:EkRootText
    if ([string]::IsNullOrEmpty($mevcut)) { $mevcut = 'Tanımlı Değil' }
    Write-Host "📂 E-Defter : $mevcut"
    Write-Host "👥 Firma    : $(Get-EkFirmaSummaryText)"

    switch ($script:EkRootStatus) {
      'OK'              { Write-Host '✅ Klasör erişilebilir.' -ForegroundColor Green }
      'InvalidPath'     { Write-Host '⚠️ Ayarlardaki klasör bulunamadı. Ayarlar menüsünden düzeltin.' -ForegroundColor Red }
      'MissingProperty' { Write-Host '⚠️ Ayar dosyasında klasör yolu boş. Ayarlar menüsünden tanımlayın.' -ForegroundColor Red }
      'Error'           { Write-Host '⚠️ Ayar dosyası okunamadı. Ayarlar menüsünden yeniden tanımlayın.' -ForegroundColor Red }
      default           { Write-Host 'ℹ️ E-Defter klasörü henüz tanımlanmadı. Ayarlar menüsünden tanımlayın.' -ForegroundColor Yellow }
    }

    Write-Host ''
    Write-Host ('-' * 72) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  1. 📊  Kontrol Raporu'
    Write-Host '  2. ⚙️  Ayarlar'
    Write-Host '  Q  🚪  Çıkış'

    Write-Host `
        ('-' * 72) -ForegroundColor DarkGray

    Write-Host ''
    $choice = Read-Host 'Seçiminiz'

    switch ($choice.ToUpper()) {
      '1' { Start-EkReport }
      '2' { Show-EkSettingsMenu; Load-EkRoot }
      'Q' { return }
      default {
        Write-Host '[!] Geçersiz seçim.' -ForegroundColor Red
        Wait-EkKey
      }
    }
  }
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
Load-EkFirmaMap
Load-EkRoot
Show-EkMainMenu


