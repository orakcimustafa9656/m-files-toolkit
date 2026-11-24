#requires -Version 5.1
# ============================================================
# yerelyedek.ps1
# Yerel Yedekleme Sistemi
#
# Kullanım:
#   .\yerelyedek.ps1              → normal menülü çalışma
#   .\yerelyedek.ps1 -SelfTest    → Pester tabanlı öz-test
#
# Menü seçenekleri:
#   1. Kontrol Raporu    - Planı gösterir, dosya sistemi değişikliği yapmaz
#   2. Gerçek Yedekleme  - Planı onaylayıp uygular
#   3. Bütünlük Kontrolü - Kaynak ile yedeği SHA-256 ile doğrular
#   4. Ayarlar           - Kaynak ve yedek kökünü yapılandırır
# ============================================================

param(
    [switch]$SelfTest
)

# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}



try { chcp 65001 | Out-Null } catch {}

$Utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding  = $Utf8
[Console]::OutputEncoding = $Utf8
$OutputEncoding           = $Utf8

# Ana menü (M-Files Toolkit) modülleri 'Stop' hata tercihiyle çalıştırır.
$ErrorActionPreference = 'Continue'

# ============================================================
# SABİTLER
# ============================================================

$ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
$SettingsFile = Join-Path (Get-MFTModuleDataPath -ModuleId 'yerelyedek' -Category Config) 'yerelyedek_path_config.json'

$ChangedFolderName = '.Changed'
$DeletedFolderName = '.Deleted'
$TempFolderName    = '.Temp'

$script:ReportLimits = @{
    PlanListPerCategory = 50
    PlanListCopy        = 10
    EventPrintLimit     = 30
    WarningListMax      = 20
    SkipListMax         = 20
}

# IO_REPARSE_TAG_CLOUD ailesi: 0x9000001A, 0x9000101A ... 0x9000F01A
# 'L' suffix'i gerekli: 0x80000000+ PowerShell'de [int] olur, negatife döner.
$script:CloudFilesReparseMask = [uint32]0xFFFF0FFFL
$script:CloudFilesReparseBase = [uint32]0x9000001AL

$script:TimeToleranceTicks  = [int64]0
$script:StrictMetadata      = $true
$script:BackupFileSystem    = ''
$script:LastMetadataWarning = ''
$script:LastMetadataErrorMessage = ''
$script:LastHashErrorType   = ''
$script:LastHashErrorMessage = ''
$script:LastArchiveErrorType = ''
$script:LastArchiveErrorMessage = ''

$script:SettableAttributeMask = [int](
    [System.IO.FileAttributes]::ReadOnly -bor
    [System.IO.FileAttributes]::Hidden -bor
    [System.IO.FileAttributes]::System -bor
    [System.IO.FileAttributes]::Archive -bor
    [System.IO.FileAttributes]::NotContentIndexed
)

# ============================================================
# OLAY ETİKET/RENK HARİTALARI
# ============================================================

$script:EventLabels = @{
    Yeni='[Yeni]'; Degisen='[Değişen]'; Silinen='[Silinen]'; Eksik='[Eksik]'; Fazla='[Fazla]'
    Atlanan='[Atlanan]'; ErisimHatasi='[Erişim Hatası]'; OkumaHatasi='[Okuma Hatası]'
    MetadataDegisen='[Metadata Değişen]'; MetadataErisimHatasi='[Metadata Erişim Hatası]'
    MetadataOkumaHatasi='[Metadata Okuma Hatası]'; Ayni='[Aynı]'
}
$script:EventColors = @{
    Yeni='Green'; Degisen='Yellow'; Silinen='Red'; Eksik='Red'; Fazla='Magenta'
    Atlanan='DarkYellow'; ErisimHatasi='Red'; OkumaHatasi='Magenta'
    MetadataDegisen='Yellow'; MetadataErisimHatasi='Red'; MetadataOkumaHatasi='Magenta'
    Ayni='DarkGray'
}
$script:CounterLabels = @{
    Yeni='Yeni'; Degisen='Değişen'; Silinen='Silinen'; Eksik='Eksik'; Fazla='Fazla'
    ErisimHatasi='Erişim Hatası'; OkumaHatasi='Okuma Hatası'
    MetadataDegisen='Metadata Değişen'; MetadataErisimHatasi='Metadata Erişim Hatası'
    MetadataOkumaHatasi='Metadata Okuma Hatası'; Ayni='Aynı'
    KontrolEdilen='Kontrol Edilen'; KararHatasi='Karar Hatası'
    UygulamaHatasi='Uygulama Hatası'; ToplamHata='Toplam Hata'
    Atlanan='Atlanan Dosya'; Uyari='Uyarı'
}
$script:ActionColors = @{
    COPY='Green'; ARCHIVE_CHANGED='Yellow'; ARCHIVE_DELETED='Red'
    UPDATE_METADATA='DarkCyan'; REPLACE='Magenta'; NOOP='DarkGray'
    SKIP='Red'; ERROR='Red'
}

# ============================================================
# YOL FONKSİYONLARI
# ============================================================

function Convert-ToLongPath {
    <#
    .SYNOPSIS
        Verilen yolu '\\?\' (veya UNC için '\\?\UNC\') önekiyle uzun yol biçimine çevirir.
    .PARAMETER Path
        Dönüştürülecek yol.
    .RETURNS
        Uzun yol biçiminde string.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.StartsWith('\\?\')) { return $full }
    if ($full.StartsWith('\\'))    { return '\\?\UNC\' + $full.TrimStart('\') }
    return '\\?\' + $full
}

function Get-DisplayPath {
    <#
    .SYNOPSIS
        Uzun yol önekini kaldırıp kullanıcıya gösterilebilir yolu döner.
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path.StartsWith('\\?\UNC\')) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))     { return $Path.Substring(4) }
    return $Path
}

function Get-NormalizedFullPath {
    <#
    .SYNOPSIS
        Yolu tam yol biçimine normalleştirir; sürücü kökünün sonundaki '\' korunur.
    .RETURNS
        Normalleştirilmiş tam yol.
    .NOTES
        'D:' gibi eksik kökleri 'D:\' yapar.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $value = $Path.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { throw 'Yol boş olamaz.' }
    if ($value -match '^[A-Za-z]:$') { $value += '\' }

    $full = [System.IO.Path]::GetFullPath($value)
    $root = [System.IO.Path]::GetPathRoot($full)
    if ([string]::IsNullOrWhiteSpace($root)) { throw "Yol kökü alınamadı: $Path" }

    if ($full.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $root }
    return $full.TrimEnd('\')
}

function Get-RelativePathSafe {
    <#
    .SYNOPSIS
        FullPath'in Root altındaki göreli kısmını döner.
    .RETURNS
        Göreli yol (Root ise boş string).
    .NOTES
        Path kök altında değilse throw eder.
    #>
    param(
        [Parameter(Mandatory)][string]$FullPath,
        [Parameter(Mandatory)][string]$Root
    )
    $rootFull = Get-NormalizedFullPath $Root
    $pathFull = Get-NormalizedFullPath $FullPath

    if ($pathFull.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) { return '' }

    $prefix = if ($rootFull.EndsWith('\')) { $rootFull } else { $rootFull + '\' }

    if (-not $pathFull.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Yol kök altında değil: $FullPath"
    }
    return $pathFull.Substring($prefix.Length)
}

function Join-RelativePath {
    param(
        [Parameter(Mandatory)][string]$Root,
        [AllowEmptyString()][string]$Relative
    )
    if ([string]::IsNullOrWhiteSpace($Relative)) { return $Root }
    return Join-Path $Root $Relative
}

function Get-RelativeFromDrive {
    <#
    .SYNOPSIS
        'D:\Users' → 'D\Users' biçiminde sürücü temelli göreli yol üretir.
    .RETURNS
        Göreli yol (sürücü harfi ile başlar).
    #>
    param([Parameter(Mandatory)][string]$Source)
    $full = Get-NormalizedFullPath $Source
    $root = [System.IO.Path]::GetPathRoot($full)
    if ([string]::IsNullOrWhiteSpace($root)) { throw 'Kaynak yolu geçersiz.' }
    if ($root -notmatch '^[A-Za-z]:\\$') {
        throw "Kaynak yolu sürücü tabanlı olmalıdır: $Source"
    }
    $driveLetter = $root.Substring(0, 1)
    $relative    = $full.Substring($root.Length).TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($relative)) { return $driveLetter }
    return Join-Path $driveLetter $relative
}

function Get-DestinationRoot {
    <#
    .SYNOPSIS
        Kaynağa karşılık gelen hedef kökünü hesaplar: <BackupRoot>\<Sürücü>\<Yol>.
    .RETURNS
        Hedef kök yolu.
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$BackupRoot
    )
    $relative = Get-RelativeFromDrive $Source
    return Join-Path $BackupRoot $relative
}

function Get-AuxRoot {
    param(
        [Parameter(Mandatory)][string]$BackupRoot,
        [Parameter(Mandatory)][ValidateSet('Changed','Deleted','Temp')][string]$Kind
    )
    $name = switch ($Kind) {
        'Changed' { $ChangedFolderName }
        'Deleted' { $DeletedFolderName }
        'Temp'    { $TempFolderName }
    }
    return Join-Path $BackupRoot $name
}

function Get-ChangedRoot { param([Parameter(Mandatory)][string]$BackupRoot) Get-AuxRoot -BackupRoot $BackupRoot -Kind 'Changed' }
function Get-DeletedRoot { param([Parameter(Mandatory)][string]$BackupRoot) Get-AuxRoot -BackupRoot $BackupRoot -Kind 'Deleted' }
function Get-TempRoot    { param([Parameter(Mandatory)][string]$BackupRoot) Get-AuxRoot -BackupRoot $BackupRoot -Kind 'Temp'    }

function Test-PathEqualOrUnder {
    <#
    .SYNOPSIS
        Child yolunun Parent ile aynı veya altında olup olmadığını test eder.
    .RETURNS
        $true / $false
    #>
    param(
        [Parameter(Mandatory)][string]$Child,
        [Parameter(Mandatory)][string]$Parent
    )
    $childFull  = Get-NormalizedFullPath $Child
    $parentFull = Get-NormalizedFullPath $Parent

    if ($childFull.Equals($parentFull, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }

    $prefix = if ($parentFull.EndsWith('\')) { $parentFull } else { $parentFull + '\' }
    return $childFull.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-DriveRootSafe {
    param([Parameter(Mandatory)][string]$Path)
    $full = Get-NormalizedFullPath $Path
    $root = [System.IO.Path]::GetPathRoot($full)
    if ([string]::IsNullOrWhiteSpace($root)) { throw "Sürücü kökü alınamadı: $Path" }
    return $root
}

function Test-DifferentDrive {
    <#
    .SYNOPSIS
        Kaynak ve yedek kökünün farklı sürücüde olup olmadığını test eder.
    .RETURNS
        $true → farklı sürücü (yedekleme güvenli); $false → aynı sürücü
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$BackupRoot
    )
    $sourceDrive = Get-DriveRootSafe $Source
    $backupDrive = Get-DriveRootSafe $BackupRoot
    return (-not $sourceDrive.Equals($backupDrive, [System.StringComparison]::OrdinalIgnoreCase))
}

# ============================================================
# REPARSE POINT / JUNCTION KONTROLÜ
# ============================================================

if (-not ([System.Management.Automation.PSTypeName]'YerelYedekNative').Type) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class YerelYedekNative
{
    [StructLayout(LayoutKind.Sequential)]
    public struct FILETIME { public uint dwLowDateTime; public uint dwHighDateTime; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct WIN32_FIND_DATA
    {
        public uint dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternateFileName;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindFirstFileW(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool FindClose(IntPtr hFindFile);
}
'@
}

function Get-ReparsePointTag {
    <#
    .SYNOPSIS
        Yolun reparse tag değerini (dwReserved0) okur.
    .RETURNS
        [uint32] tag ya da $null (reparse değilse veya okunamazsa).
    #>
    param([Parameter(Mandatory)][string]$Path)

    $handle = [IntPtr]::Zero
    try {
        $longPath = Convert-ToLongPath $Path
        $data = New-Object YerelYedekNative+WIN32_FIND_DATA
        $handle = [YerelYedekNative]::FindFirstFileW($longPath, [ref]$data)

        if ($handle -eq [IntPtr]::Zero -or $handle -eq [IntPtr](-1)) { return $null }

        $attributes = [uint32]$data.dwFileAttributes
        if (($attributes -band [uint32][System.IO.FileAttributes]::ReparsePoint) -eq 0) { return $null }

        return [uint32]$data.dwReserved0
    }
    catch {
        Write-RunLog 'UYARI' "ReparsePoint tag okunamadı: $Path -- $($_.Exception.Message)"
        return $null
    }
    finally {
        if ($handle -ne [IntPtr]::Zero -and $handle -ne [IntPtr](-1)) {
            [void][YerelYedekNative]::FindClose($handle)
        }
    }
}

function Test-IsCloudFilesReparseTag {
    <#
    .SYNOPSIS
        Verilen tag IO_REPARSE_TAG_CLOUD ailesine mi ait?
    .RETURNS
        $true / $false
    #>
    param([Parameter(Mandatory)][uint32]$Tag)
    return (($Tag -band $script:CloudFilesReparseMask) -eq $script:CloudFilesReparseBase)
}

function Test-IsBlockedDirectoryReparsePoint {
    <#
    .SYNOPSIS
        Klasörün "içine girilmemesi gereken" bir reparse point olup olmadığını test eder.
    .RETURNS
        $true → içine girme; $false → güvenli (Cloud Files yer tutucu veya reparse değil).
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $attributes = [System.IO.File]::GetAttributes((Convert-ToLongPath $Path))
        if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) { return $false }

        $tag = Get-ReparsePointTag $Path
        if ($null -eq $tag) { return $true }
        if (Test-IsCloudFilesReparseTag $tag) { return $false }
        return $true
    }
    catch {
        Write-RunLog 'UYARI' "ReparsePoint tespiti yapılamadı (güvenli taraf): $Path -- $($_.Exception.Message)"
        return $true
    }
}

function Test-IsUnderRelativePath {
    param(
        [Parameter(Mandatory)][string]$Child,
        [Parameter(Mandatory)][string]$Parent
    )
    if ($Child.Equals($Parent, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $Child.StartsWith($Parent + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-IsUnderAnyRelativePath {
    param(
        [Parameter(Mandatory)][string]$Child,
        [Parameter(Mandatory)]$Parents
    )
    foreach ($parent in $Parents) {
        if ([string]::IsNullOrWhiteSpace([string]$parent)) { continue }
        if (Test-IsUnderRelativePath -Child $Child -Parent ([string]$parent)) { return $true }
    }
    return $false
}

# ============================================================
# DOSYA / KLASÖR KONTROLLERİ
# ============================================================

function Test-DirectoryExists {
    <#
    .SYNOPSIS
        Uzun yol destekli klasör varlık testi.
    .RETURNS
        $true / $false (hata durumunda $false)
    #>
    param([string]$Path)
    try { return [System.IO.Directory]::Exists((Convert-ToLongPath $Path)) } catch { return $false }
}

function Test-FileExists {
    <#
    .SYNOPSIS
        Uzun yol destekli dosya varlık testi.
    .RETURNS
        $true / $false (hata durumunda $false)
    #>
    param([string]$Path)
    try { return [System.IO.File]::Exists((Convert-ToLongPath $Path)) } catch { return $false }
}

function Ensure-Directory {
    <#
    .SYNOPSIS
        Klasör yoksa oluşturur. Hata durumunda ekrana yazmaz; sonucu döner.
    .RETURNS
        $true → klasör mevcut; $false → oluşturulamadı.
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $long = Convert-ToLongPath $Path
        if (-not [System.IO.Directory]::Exists($long)) {
            [System.IO.Directory]::CreateDirectory($long) | Out-Null
        }
        return [System.IO.Directory]::Exists($long)
    }
    catch {
        return $false
    }
}

function Remove-ReadOnlyAttribute {
    <#
    .SYNOPSIS
        Dosya/klasörden ReadOnly niteliğini kaldırır (varsa).
    .RETURNS
        $true / $false
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $long = Convert-ToLongPath $Path
        $attr = [System.IO.File]::GetAttributes($long)
        if (($attr -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
            [System.IO.File]::SetAttributes($long, ($attr -bxor [System.IO.FileAttributes]::ReadOnly))
        }
        return $true
    }
    catch { return $false }
}

# ============================================================
# AYARLAR
# ============================================================

function Load-Settings {
    <#
    .SYNOPSIS
        Ayar JSON'unu okur, doğrular ve döner.
    .RETURNS
        [PSCustomObject]@{Source; BackupRoot} veya $null (yok/geçersizse).
    #>
    if (-not (Test-Path -LiteralPath $SettingsFile)) { return $null }
    try {
        $settings = Get-Content -LiteralPath $SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $settings) { return $null }
        if ([string]::IsNullOrWhiteSpace([string]$settings.Source)) { return $null }
        if ([string]::IsNullOrWhiteSpace([string]$settings.BackupRoot)) { return $null }

        $source = ([string]$settings.Source).Trim()
        $backup = ([string]$settings.BackupRoot).Trim()

        if ($source -match '^[A-Za-z]:$') { $source += '\' }
        if ($backup -match '^[A-Za-z]:$') { $backup += '\' }

        return [PSCustomObject]@{ Source = $source; BackupRoot = $backup }
    }
    catch {
        Write-RunLog 'UYARI' "Ayar dosyası okunamadı: $($_.Exception.Message)"
        return $null
    }
}

function Save-Settings {
    <#
    .SYNOPSIS
        Ayar JSON'unu diske yazar.
    .RETURNS
        $true / $false
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$BackupRoot
    )
    try {
        $obj  = [PSCustomObject]@{ Source = $Source; BackupRoot = $BackupRoot }
        $json = $obj | ConvertTo-Json -Depth 3
        Set-Content -LiteralPath $SettingsFile -Value $json -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch { return $false }
}

function Test-SettingsValid {
    <#
    .SYNOPSIS
        Kaynak/yedek ayarlarının tüm ön koşullarını doğrular.
    .PARAMETER ShowError
        Verildiyse başarısızlık nedenini ekrana yazar.
    .RETURNS
        $true / $false
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$BackupRoot,
        [switch]$ShowError
    )
    try {
        $sourceFull = Get-NormalizedFullPath $Source
        $backupFull = Get-NormalizedFullPath $BackupRoot
        $sourceDrive = Get-DriveRootSafe $sourceFull
        $backupDrive = Get-DriveRootSafe $backupFull

        if (-not (Test-DirectoryExists $sourceFull)) { throw "Kaynak klasörü bulunamadı: $sourceFull" }
        if (Test-IsBlockedDirectoryReparsePoint $sourceFull) { throw "Kaynak klasörü engellenen bir Reparse Point: $sourceFull" }
        if (-not (Test-DirectoryExists $backupDrive)) { throw "Yedek sürücüsü erişilemiyor: $backupDrive" }
        if ($sourceDrive.Equals($backupDrive, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'Kaynak ve yedek farklı sürücülerde olmalıdır.'
        }

        if (Test-DirectoryExists $backupFull) {
            if (Test-IsBlockedDirectoryReparsePoint $backupFull) {
                throw "Yedek konumu engellenen bir Reparse Point: $backupFull"
            }
        }

        foreach ($reserved in @($ChangedFolderName, $DeletedFolderName)) {
            $p = Join-Path $sourceFull $reserved
            if ((Test-DirectoryExists $p) -or (Test-FileExists $p)) {
                throw "Kaynakta ayrılmış isim kullanılıyor: $reserved"
            }
            $archiveRoot = Join-Path $backupFull $reserved
            if (Test-FileExists $archiveRoot) {
                throw "Yedek kökünde ayrılmış arşiv yolu dosya olarak kullanılıyor: $archiveRoot"
            }
            if ((Test-DirectoryExists $archiveRoot) -and (Test-IsBlockedDirectoryReparsePoint $archiveRoot)) {
                throw "Yedek kökündeki arşiv yolu engellenen bir Reparse Point: $archiveRoot"
            }
        }

        $tempRoot = Get-TempRoot $backupFull
        if (Test-FileExists $tempRoot) {
            throw "Yedek kökünde ayrılmış geçici yol dosya olarak kullanılıyor: $tempRoot"
        }
        if ((Test-DirectoryExists $tempRoot) -and (Test-IsBlockedDirectoryReparsePoint $tempRoot)) {
            throw "Yedek kökündeki geçici yol engellenen bir Reparse Point: $tempRoot"
        }

        $destination = Get-DestinationRoot -Source $sourceFull -BackupRoot $backupFull
        if (Test-PathEqualOrUnder -Child $destination -Parent $sourceFull) {
            throw 'Yedek hedefi kaynak klasörünün içine düşüyor.'
        }
        return $true
    }
    catch {
        if ($ShowError) { Write-Host "[HATA] $($_.Exception.Message)" -ForegroundColor Red }
        return $false
    }
}

# ============================================================
# YAPILANDIRMA DURUMU
# ============================================================

function Get-ConfigurationStatus {
    <#
    .SYNOPSIS
        Ana menüde gösterilen yapılandırma durumunu derler.
    .RETURNS
        [PSCustomObject] (Status, Message, Settings, Destination, *Status alanları)
    #>
    if (-not (Test-Path -LiteralPath $SettingsFile)) {
        return [PSCustomObject]@{
            Status='MISSING'; Message='AYARLARDA KAYNAK VE YEDEK KONUMU BELİRLEYİN'
            Settings=$null; Destination=$null
            SourceStatus='YOK'; BackupStatus='YOK'
            DestinationStatus='YOK'; SettingsFileStatus='BULUNAMADI'
        }
    }

    $settings = Load-Settings
    if ($null -eq $settings) {
        return [PSCustomObject]@{
            Status='INVALID'; Message='AYAR DOSYASI GEÇERSİZ'
            Settings=$null; Destination=$null
            SourceStatus='HATA'; BackupStatus='HATA'
            DestinationStatus='HESAPLANAMADI'; SettingsFileStatus='GEÇERSİZ'
        }
    }

    $sourceStatus = if (Test-DirectoryExists $settings.Source) { 'OK' } else { 'YOK' }

    try {
        $backupDrive = Get-DriveRootSafe $settings.BackupRoot
        $backupStatus = if (Test-DirectoryExists $backupDrive) { 'OK' } else { 'YOK' }
    } catch { $backupStatus = 'HATA' }

    $destination = $null
    $destinationStatus = 'YOK'
    try {
        $destination = Get-DestinationRoot -Source $settings.Source -BackupRoot $settings.BackupRoot
        if (Test-DirectoryExists $destination) { $destinationStatus = 'OK' }
        elseif ($sourceStatus -eq 'OK' -and $backupStatus -eq 'OK') { $destinationStatus = 'HAZIRLANACAK' }
        else { $destinationStatus = 'YOK' }
    } catch { $destination = $null; $destinationStatus = 'HESAPLANAMADI' }

    $settingsValid = Test-SettingsValid -Source $settings.Source -BackupRoot $settings.BackupRoot
    if (-not $settingsValid) {
        return [PSCustomObject]@{
            Status='INVALID'; Message='AYARLAR GEÇERSİZ VEYA ERİŞİLEMİYOR'
            Settings=$settings; Destination=$destination
            SourceStatus=$sourceStatus; BackupStatus=$backupStatus
            DestinationStatus=$destinationStatus; SettingsFileStatus='OK'
        }
    }
    return [PSCustomObject]@{
        Status='READY'; Message='AYARLAR HAZIR'
        Settings=$settings; Destination=$destination
        SourceStatus=$sourceStatus; BackupStatus=$backupStatus
        DestinationStatus=$destinationStatus; SettingsFileStatus='OK'
    }
}

# ============================================================
# LOG SİSTEMİ
# ============================================================
$script:RunLogWriter = $null
$script:RunLogPath   = ''

function Get-LogRoot {
    return (Get-MFTModuleDataPath -ModuleId 'yerelyedek' -Category Logs)
}

function Write-RunLog {
    <#
    .SYNOPSIS
        Çalışma log dosyasına tek satır yazar (writer yoksa sessizce döner).
    #>
    param([string]$Level = 'BİLGİ', [string]$Message = '')
    if ($null -eq $script:RunLogWriter) { return }
    try {
        $script:RunLogWriter.WriteLine(
            ('{0} | {1,-5} | {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
        )
    } catch { }
}

function Initialize-RunLog {
    <#
    .SYNOPSIS
        Data\Modules\yerelyedek\Logs altına zaman damgalı log dosyası açar.
    .NOTES
        Log klasörü merkezi Data alanında oluşturulur; yedek hedefinden bağımsızdır.
    #>
    param(
        [Parameter(Mandatory)][string]$Slug,
        [Parameter(Mandatory)][string]$BackupRoot,
        [string[]]$HeaderLines = @()
    )
    if ($null -ne $script:RunLogWriter) { try { $script:RunLogWriter.Dispose() } catch { } }
    $script:RunLogWriter = $null
    $script:RunLogPath   = ''

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $candidates = @((Get-LogRoot))

    $lastError = ''

    for ($i = 0; $i -lt $candidates.Count; $i++) {
        $dir = $candidates[$i]
        try {
            if (-not (Test-DirectoryExists $dir)) {
                if (-not (Ensure-Directory $dir)) { $lastError = "Klasör oluşturulamadı: $dir"; continue }
            }
            $path = Join-Path $dir ('{0}_{1}.log' -f $stamp, $Slug)
            $fs = New-Object System.IO.FileStream(
                (Convert-ToLongPath $path),
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::Read
            )
            $w = New-Object System.IO.StreamWriter($fs, (New-Object System.Text.UTF8Encoding($true)))
            $w.AutoFlush = $true

            $script:RunLogWriter = $w
            $script:RunLogPath   = $path
            break
        } catch { $lastError = $_.Exception.Message; continue }
    }

    if ($null -eq $script:RunLogWriter) {
        Write-Host 'UYARI: Log dosyası oluşturulamadı; işlem logsuz devam ediyor.' -ForegroundColor Yellow
        if ($lastError) { Write-Host "       Sebep: $lastError" -ForegroundColor DarkYellow }
        return
    }


    Write-RunLog 'BİLGİ' ('Başlangıç | Bilgisayar: {0} | Kullanıcı: {1} | PowerShell: {2}' -f $env:COMPUTERNAME, $env:USERNAME, $PSVersionTable.PSVersion)
    foreach ($h in $HeaderLines) { Write-RunLog 'BİLGİ' $h }
}

function Write-RunLogRow {
    param($Row)
    $level = 'İŞLEM'
    if ($Row.Action -eq 'ERROR' -or $Row.Result -eq 'ERROR') { $level = 'HATA' }
    elseif ($Row.Action -eq 'SKIP') { $level = 'ATLA' }

    $msg = '{0,-16} | {1,-8} | {2}' -f $Row.Action, $Row.Result, $Row.RelativePath
    if (-not [string]::IsNullOrWhiteSpace([string]$Row.Error)) { $msg += ' | ' + [string]$Row.Error }
    Write-RunLog $level $msg
}

function Write-RunLogSummary {
    param($Manifest)
    Write-RunLog 'BİLGİ' ('Süre: {0:N1} sn (onay bekleme süresi hariç)' -f (Get-RunElapsed).TotalSeconds)
    if ($null -ne $script:RunPausedTime -and $script:RunPausedTime -gt [TimeSpan]::Zero) {
        Write-RunLog 'BİLGİ' ('Onay bekleme süresi: {0:N1} sn' -f $script:RunPausedTime.TotalSeconds)
    }
    foreach ($g in @($Manifest | Group-Object { '{0} / {1}' -f $_.Action, $_.Result } | Sort-Object Name)) {
        Write-RunLog 'BİLGİ' ('ÖZET | {0,-30} | {1}' -f $g.Name, $g.Count)
    }
}

# ============================================================
# CANLI İLERLEME SATIRI
# ============================================================
$script:ProgressActive    = $false
$script:ProgressStopwatch = $null
$script:ProgressLastMs    = -1000
$script:ProgressLastLength = 0
# Neden 100: bazı kurulumlarda (cmd/mode con ile açılan pencereler) bildirilen genişlik
# görünen genişlikten büyük çıkıyor; satır taşınca satır başı (`r) yalnızca son
# satırı geri aldığı için eski içerik ekranda kalıyor. Kısa tutmak taşmayı önler.
$script:ProgressMaxWidth  = 100

function Get-ConsoleWidth {
    <#
    .SYNOPSIS
        Konsol genişliğini döner; birden fazla kaynağın en küçüğünü alır.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param()

    $candidates = New-Object System.Collections.Generic.List[int]
    try { $candidates.Add([int]$Host.UI.RawUI.WindowSize.Width) } catch { Write-Verbose 'RawUI.WindowSize okunamadı.' }
    try { $candidates.Add([int][Console]::WindowWidth) } catch { Write-Verbose '[Console]::WindowWidth okunamadı.' }

    $valid = @($candidates | Where-Object { $_ -ge 40 })
    if ($valid.Count -eq 0) { return 100 }
    return [int]($valid | Measure-Object -Minimum).Minimum
}

function Get-ProgressLineWidth {
    <#
    .SYNOPSIS
        İlerleme satırının en çok kaplayabileceği genişliği döner (satır taşmasın diye kısıtlı).
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param()

    return [Math]::Min((Get-ConsoleWidth), $script:ProgressMaxWidth)
}

function Format-ShortPath {
    param([string]$Path, [int]$Max)
    if ($Max -lt 12) { $Max = 12 }
    if ($Path.Length -le $Max) { return $Path }

    $parts = $Path -split '\\'
    if ($parts.Count -ge 3) {
        $candidate = $parts[0] + '\...\' + $parts[$parts.Count - 1]
        if ($candidate.Length -le $Max) { return $candidate }
        $tail = $parts[$parts.Count - 1]
        if ($tail.Length -gt ($Max - 4)) { $tail = $tail.Substring($tail.Length - ($Max - 4)) }
        return '...\' + $tail
    }
    return '...' + $Path.Substring($Path.Length - ($Max - 3))
}

function Format-Eta {
    param([double]$Seconds)
    if ($Seconds -lt 0 -or [double]::IsNaN($Seconds) -or [double]::IsInfinity($Seconds)) { return '--:--' }
    $ts = [TimeSpan]::FromSeconds([Math]::Min($Seconds, 359999))
    if ($ts.TotalHours -ge 1) {
        return ('{0}:{1:00}:{2:00}' -f [int][Math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds)
    }
    return ('{0:00}:{1:00}' -f $ts.Minutes, $ts.Seconds)
}

function Start-ProgressLine {
    $script:ProgressStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:ProgressLastMs    = -1000
    $script:ProgressLastLength = 0
    $script:ProgressActive    = $false
}

function Update-ProgressLine {
    param(
        [int]$Index, [int]$Total, [string]$Verb, [string]$Path,
        [switch]$Force
    )
    if ($null -eq $script:ProgressStopwatch) { return }
    $ms = $script:ProgressStopwatch.ElapsedMilliseconds
    if (-not $Force -and (($ms - $script:ProgressLastMs) -lt 120)) { return }
    $script:ProgressLastMs = $ms

    $pct = 100
    if ($Total -gt 0) { $pct = [int][Math]::Floor(100.0 * $Index / $Total) }

    $eta = '--:--'
    if ($Index -ge 3 -and $ms -ge 2000 -and $Total -gt $Index) {
        $eta = Format-Eta ((($ms / 1000.0) / $Index) * ($Total - $Index))
    } elseif ($Total -gt 0 -and $Index -ge $Total) { $eta = '00:00' }

    $sep = ''
    if (-not [string]::IsNullOrEmpty($Path)) { $sep = ': ' }

    if ($Total -gt 0) {
        $barWidth = 10
        $filled = [int][Math]::Floor($barWidth * $pct / 100.0)
        $bar = ('#' * $filled) + ('-' * ($barWidth - $filled))
        $prefix = '[{0:N0} / {1:N0}] [{2}] %{3} | {4}{5}' -f $Index, $Total, $bar, $pct, $Verb, $sep
    } else {
        $prefix = '[{0:N0} dosya] | {1}{2}' -f $Index, $Verb, $sep
    }
    # Neden: toplam bilinmiyorsa (envanter taraması) kalan süre hesaplanamaz; anlamsız '--:--' gösterilmez.
    $suffix = if ($Total -gt 0) { ' | kalan ~{0}' -f $eta } else { '' }

    $width = Get-ProgressLineWidth
    $avail = $width - 1 - $prefix.Length - $suffix.Length
    $text  = $prefix + (Format-ShortPath $Path $avail) + $suffix
    if ($text.Length -gt ($width - 1)) { $text = $text.Substring(0, $width - 1) }

    # Neden: yalnızca önceki satırın uzunluğu kadar boşlukla silinir; tam genişliğe
    # dolgu yapılmaz, böylece satır hiçbir zaman pencere kenarına dayanıp taşmaz.
    $padLength = [Math]::Max($text.Length, $script:ProgressLastLength)
    Write-Host ("`r" + $text.PadRight($padLength)) -NoNewline
    $script:ProgressLastLength = $text.Length
    $script:ProgressActive = $true
}

function Clear-ProgressLine {
    if (-not $script:ProgressActive) { return }
    $clearLength = [Math]::Max($script:ProgressLastLength, 0)
    Write-Host ("`r" + (' ' * $clearLength) + "`r") -NoNewline
    $script:ProgressLastLength = 0
    $script:ProgressActive = $false
    $script:ProgressLastMs = -1000
}

function Stop-ProgressLine {
    Clear-ProgressLine
    $script:ProgressStopwatch = $null
}

function Write-PermanentLine {
    param([string]$Text, [string]$Color = 'Gray')
    Clear-ProgressLine
    Write-Host $Text -ForegroundColor $Color
}

# ============================================================
# PLAN ÖZETİ
# ============================================================

function Format-ByteSize {
    param([double]$Bytes)
    if ($Bytes -lt 1024) { return ('{0:N0} B' -f $Bytes) }
    $units = @('KB', 'MB', 'GB', 'TB')
    $i = -1
    while ($Bytes -ge 1024 -and $i -lt 3) { $Bytes = $Bytes / 1024; $i++ }
    return ('{0:N1} {1}' -f $Bytes, $units[$i])
}

function Write-ProblemEntry {
    <#
    .SYNOPSIS
        Atlanan/hatalı kaydı iki satırda yazar: göreli dosya yolu, altında hata mesajı.
    .NOTES
        Neden: tek satırda yol + hata uzun hatalarda kesiliyordu. Yol, diğer plan
        kategorileri ve logla tutarlı olsun diye göreli verilir; tam yol zaten hata
        mesajının içinde yer alır. Hata metni kısaltılmadan ayrı satırda girintili yazılır.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Row,
        [System.ConsoleColor]$PathColor = [System.ConsoleColor]::DarkGray,
        [System.ConsoleColor]$ErrorColor = [System.ConsoleColor]::DarkYellow
    )

    Write-Host ('  {0}' -f $Row.RelativePath) -ForegroundColor $PathColor

    if (-not [string]::IsNullOrWhiteSpace([string]$Row.Error)) {
        Write-Host ('    {0}' -f [string]$Row.Error) -ForegroundColor $ErrorColor
    }
}

function Write-PlanDigest {
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)]$SourceInventory,
        [int]$MaxPerCategory = $script:ReportLimits.PlanListPerCategory
    )
    $cats = @(
        @{ A='ERROR';           T='HATA';                                             C='Red';     Max=$script:ReportLimits.PlanListPerCategory },
        @{ A='SKIP';            T='ATLANAN (okunamadı, yedeklenmeyecek)';             C='Red';     Max=$script:ReportLimits.PlanListPerCategory },
        @{ A='REPLACE';         T='TİP ÇATIŞMASI';                                    C='Magenta'; Max=$MaxPerCategory },
        @{ A='ARCHIVE_CHANGED'; T='DEĞİŞEN (eski sürüm .Changed klasörüne alınacak)'; C='Yellow';  Max=$MaxPerCategory },
        @{ A='ARCHIVE_DELETED'; T='SİLİNEN (.Deleted klasörüne alınacak)';            C='Red';     Max=$MaxPerCategory },
        @{ A='COPY';            T='YENİ';                                             C='Green';   Max=$script:ReportLimits.PlanListCopy }
    )

    $groups = @{}
    foreach ($g in ($Manifest | Group-Object Action)) { $groups[$g.Name] = $g.Group }

    foreach ($c in $cats) {
        # Neden @(if ...): if ifadesi tek öğeli diziyi açar; Windows PowerShell 5.1'de tek PSCustomObject'in .Count değeri $null döner.
        $rows = @(if ($groups.ContainsKey($c.A)) { $groups[$c.A] | Sort-Object RelativePath })
        if ($rows.Count -eq 0) { continue }

        $dirCount = @($rows | Where-Object {
            ($_.SourceType -eq 'Directory') -or
            (($_.SourceType -eq '') -and ($_.DestinationType -eq 'Directory'))
        }).Count
        $fileCount = $rows.Count - $dirCount

        [double]$bytes = 0
        if ($c.A -in 'COPY', 'ARCHIVE_CHANGED', 'REPLACE') {
            foreach ($r in $rows) {
                if ($r.SourceType -eq 'File' -and $SourceInventory.Files.ContainsKey($r.RelativePath)) {
                    $m = $SourceInventory.Files[$r.RelativePath].Metadata
                    if ($null -ne $m -and $null -ne $m.Length) { $bytes += [double]$m.Length }
                }
            }
        }

        $head = '{0}: {1:N0} dosya' -f $c.T, $fileCount
        if ($dirCount -gt 0) { $head += ', {0:N0} klasör' -f $dirCount }
        if ($bytes -gt 0)    { $head += ' ({0})' -f (Format-ByteSize $bytes) }

        Write-Host $head -ForegroundColor $c.C

        $shown = 0
        foreach ($r in $rows) {
            if ($shown -ge $c.Max) { break }
            if ($c.A -eq 'ERROR' -or $c.A -eq 'SKIP') {
                Write-ProblemEntry -Row $r
            }
            else {
                Write-Host ('  ' + $r.RelativePath) -ForegroundColor DarkGray
            }
            $shown++
        }
        if ($rows.Count -gt $shown) {
            Write-Host ('  ... ve {0:N0} kayıt daha (tamamı log dosyasında)' -f ($rows.Count - $shown)) -ForegroundColor DarkYellow
        }
        Write-Host ''
    }

    $maintenanceCount = if ($groups.ContainsKey('UPDATE_METADATA')) { $groups['UPDATE_METADATA'].Count } else { 0 }
    $noopCount        = if ($groups.ContainsKey('NOOP'))            { $groups['NOOP'].Count }            else { 0 }

    if ($maintenanceCount -gt 0) {
        Write-Host ('DİĞER DÜZENLEMELER: {0:N0} kayıt' -f $maintenanceCount) -ForegroundColor DarkCyan
    }
    Write-Host ('DEĞİŞMEYEN: {0:N0} kayıt' -f $noopCount) -ForegroundColor DarkGray
}

# ============================================================
# ÇALIŞMA OLAYLARI / SAYAÇ / RAPOR
# ============================================================

function Get-RunElapsed {
    <#
    .SYNOPSIS
        Çalışma süresini kullanıcı onayı beklerken geçen süre düşülmüş olarak döner.
    .NOTES
        Neden: "Uygulansın mı?" sorusuna cevap verilene kadar geçen süre işin süresi
        değildir; düşülmezse süre performans göstergesi olarak yanıltıcı olur.
    #>
    [CmdletBinding()]
    [OutputType([TimeSpan])]
    param()

    $paused = if ($null -ne $script:RunPausedTime) { $script:RunPausedTime } else { [TimeSpan]::Zero }
    $elapsed = (Get-Date) - $script:RunStarted - $paused
    if ($elapsed -lt [TimeSpan]::Zero) { return [TimeSpan]::Zero }
    return $elapsed
}

function Initialize-RunReport {
    param([string]$Mode,[string]$Source,[string]$Destination,[string]$BackupRoot)
    $script:RunMode        = $Mode
    $script:RunStarted     = Get-Date
    $script:RunPausedTime  = [TimeSpan]::Zero
    $script:RunSource      = $Source
    $script:RunDestination = $Destination
    $script:RunBackupRoot  = $BackupRoot

    if ($null -ne $script:RunLogWriter) { try { $script:RunLogWriter.Dispose() } catch { } }
    $script:RunLogWriter = $null
    $script:RunLogPath   = ''

    $script:RunSuppressEventOutput = $false
    $script:RunCounters = [ordered]@{
        Yeni=0; Degisen=0; Silinen=0; Eksik=0; Fazla=0; ErisimHatasi=0; OkumaHatasi=0
        MetadataDegisen=0; MetadataErisimHatasi=0; MetadataOkumaHatasi=0; Ayni=0
        KontrolEdilen=0; KararHatasi=0; UygulamaHatasi=0; ToplamHata=0
        Atlanan=0; Uyari=0
    }
    $script:RunWarnings = New-Object System.Collections.Generic.List[string]
    $script:LastArchiveErrorType    = ''
    $script:LastArchiveErrorMessage = ''
    $script:EventPrintCount = @{}
}

function Add-RunEvent {
    <#
    .SYNOPSIS
        Manifest karar olayını sayaca/log'a/konsola işler.
    #>
    param(
        [ValidateSet('Yeni','Degisen','Silinen','Eksik','Fazla','Atlanan','ErisimHatasi','OkumaHatasi',
                     'MetadataDegisen','MetadataErisimHatasi','MetadataOkumaHatasi','Ayni')]
        [string]$Type,
        [string]$Path,
        [string]$Detail = ''
    )
    if ($null -eq $script:RunCounters) { return }

    $label = if ($script:EventLabels.ContainsKey($Type)) { $script:EventLabels[$Type] } else { "[$Type]" }
    $color = if ($script:EventColors.ContainsKey($Type)) { $script:EventColors[$Type] } else { 'Gray' }

    if ($script:RunCounters.Contains($Type)) { $script:RunCounters[$Type]++ }

    $line = if ([string]::IsNullOrWhiteSpace($Detail)) { "$label $Path" } else { "$label $Path -- $Detail" }

    $isErrorType = $Type -in 'ErisimHatasi','OkumaHatasi','MetadataErisimHatasi','MetadataOkumaHatasi'
    $logLevel = if ($isErrorType) { 'HATA' } elseif ($Type -eq 'Atlanan') { 'UYARI' } else { 'FARK' }
    Write-RunLog $logLevel $line

    if (-not $script:RunSuppressEventOutput) {
        $key = [string]$Type
        if (-not $script:EventPrintCount.ContainsKey($key)) { $script:EventPrintCount[$key] = 0 }
        $script:EventPrintCount[$key]++
        if ($isErrorType -or $script:EventPrintCount[$key] -le $script:ReportLimits.EventPrintLimit) {
            Write-PermanentLine $line $color
        }
        elseif ($script:EventPrintCount[$key] -eq ($script:ReportLimits.EventPrintLimit + 1)) {
            Write-PermanentLine ("  ... {0} türündeki diğer kayıtlar yalnızca log dosyasında" -f $label) 'DarkYellow'
        }
    }
}

function Add-RunWarning {
    <#
    .SYNOPSIS
        İkincil metadata uyarısını (klasör değiştirilme zamanı ve nitelikler) kaydeder.
    #>
    param([string]$Path, [string]$Detail = '')
    if ($null -eq $script:RunWarnings) { return }

    $entry = if ([string]::IsNullOrWhiteSpace($Detail)) { $Path } else { "$Path -- $Detail" }
    [void]$script:RunWarnings.Add($entry)
    Write-RunLog 'UYARI' $entry

    if ($null -ne $script:RunCounters -and $script:RunCounters.Contains('Uyari')) {
        $script:RunCounters['Uyari']++
    }
}

function Write-RunWarnings {
    param([int]$Max = $script:ReportLimits.WarningListMax)
    if ($null -eq $script:RunWarnings -or $script:RunWarnings.Count -eq 0) { return }

    Write-Host ''
    Write-Host (
        'UYARI: {0} ikincil metadata uyarısı (klasör değiştirilme zamanı / nitelik; yedeğin geçerliliğini etkilemez)' -f
        $script:RunWarnings.Count
    ) -ForegroundColor Yellow

    $i = 0
    foreach ($w in $script:RunWarnings) {
        if ($i -ge $Max) {
            Write-Host ('  ... ve {0} uyarı daha' -f ($script:RunWarnings.Count - $Max)) -ForegroundColor DarkYellow
            break
        }
        Write-Host "  $w" -ForegroundColor DarkYellow
        $i++
    }
}

function Write-ManifestRow {
    <#
    .SYNOPSIS
        Manifest tablosunda bir satırı hücre hücre yazar; 0 değerleri soluk, HATA > 0 kırmızı olur.
    .NOTES
        Neden hücre hücre: satır rengi eylemi gösterir, ama sıfırlar soluk olunca
        tabloda gerçekten ne olduğu bir bakışta görünür. Sütun genişlikleri başlık
        satırıyla aynıdır (7, 9, 8, 7, 8, 6).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][int[]]$Values,
        [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray,
        [switch]$DimZeros
    )

    $widths = @(7, 9, 8, 7, 8, 6)
    Write-Host ('  {0,-20}' -f $Label) -ForegroundColor $Color -NoNewline
    for ($index = 0; $index -lt $widths.Count; $index++) {
        $cellColor = $Color
        if ($DimZeros -and $Values[$index] -eq 0) { $cellColor = [System.ConsoleColor]::DarkGray }
        elseif ($DimZeros -and $index -eq 5 -and $Values[$index] -gt 0) { $cellColor = [System.ConsoleColor]::Red }
        Write-Host (' ' + $Values[$index].ToString('N0').PadLeft($widths[$index])) -ForegroundColor $cellColor -NoNewline
    }
    Write-Host ''
}

function Get-AppliedByteTotal {
    <#
    .SYNOPSIS
        Başarıyla uygulanan kopyalama işlemlerinin (yeni/değişen/REPLACE dosyalar) toplam boyutunu döner.
    .NOTES
        Boyut kaynak envanterindeki gerçek dosya uzunluğundan gelir; tahmin değildir.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param(
        [Parameter(Mandatory = $true)]$Manifest,
        [Parameter(Mandatory = $true)]$SourceInventory
    )

    [double]$total = 0
    foreach ($row in $Manifest) {
        $isCopyAction = ($row.Action -in 'COPY', 'ARCHIVE_CHANGED', 'REPLACE')
        if ($row.Result -ne 'SUCCESS' -or -not $isCopyAction -or $row.SourceType -ne 'File') { continue }
        if (-not $SourceInventory.Files.ContainsKey($row.RelativePath)) { continue }
        $metadata = $SourceInventory.Files[$row.RelativePath].Metadata
        if ($null -ne $metadata -and $null -ne $metadata.Length) { $total += [double]$metadata.Length }
    }
    return $total
}

function Write-RunSummary {
    param([object]$Manifest = $null)

    if ($null -ne $Manifest) {
        $actionOrder = @('COPY','ARCHIVE_CHANGED','ARCHIVE_DELETED','UPDATE_METADATA','REPLACE','NOOP','SKIP','ERROR')

        Write-Host ''
        Write-Host 'MANİFEST KARARI / OLAY SONUCU' -ForegroundColor Cyan
        Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
        Write-Host ("  {0,-20} {1,7} {2,9} {3,8} {4,7} {5,8} {6,6}" -f 'MANİFEST KARARI','TOPLAM','BAŞARILI','PLANLI','AYNI','ATLANDI','HATA') -ForegroundColor Gray
        Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray

        $manifestTotal = $Manifest.Count
        $totalSuccess = 0; $totalPlanned = 0; $totalNoop = 0; $totalSkipped = 0; $totalErrors = 0

        $byAction = @{}
        foreach ($g in ($Manifest | Group-Object Action)) { $byAction[$g.Name] = $g.Group }

        foreach ($action in $actionOrder) {
            if (-not $byAction.ContainsKey($action)) { continue }
            $rows = @($byAction[$action])

            $success = @($rows | Where-Object { $_.Result -eq 'SUCCESS' }).Count
            $planned = @($rows | Where-Object { $_.Result -eq 'PLANNED' }).Count
            $noop    = @($rows | Where-Object { $_.Result -eq 'NOOP'    }).Count
            $skipped = @($rows | Where-Object { $_.Result -eq 'SKIPPED' }).Count
            $errors  = @($rows | Where-Object { $_.Result -eq 'ERROR'   }).Count

            $totalSuccess += $success; $totalPlanned += $planned; $totalNoop += $noop
            $totalSkipped += $skipped; $totalErrors  += $errors

            $rowColor = if ($script:ActionColors.ContainsKey($action)) { $script:ActionColors[$action] } else { 'Gray' }
            Write-ManifestRow -Label $action -Values @($rows.Count, $success, $planned, $noop, $skipped, $errors) -Color $rowColor -DimZeros
        }

        Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
        Write-ManifestRow -Label 'TOPLAM' -Values @($manifestTotal, $totalSuccess, $totalPlanned, $totalNoop, $totalSkipped, $totalErrors) -Color Cyan

        $script:RunCounters['KontrolEdilen'] = $manifestTotal
        $applicationErrors = [int]$script:RunCounters['UygulamaHatasi']
        $decisionErrors = $totalErrors - $applicationErrors
        if ($decisionErrors -lt 0) { $decisionErrors = 0 }
        $script:RunCounters['KararHatasi'] = $decisionErrors
        $script:RunCounters['ToplamHata']  = $totalErrors

        Write-Host ''
        if ($null -ne $script:RunSource) {
            $sourceDirs  = if ($null -ne $script:RunSourceInventory) { '{0:N0}' -f @($script:RunSourceInventory.Directories.Keys).Count } else { '-' }
            $sourceFiles = if ($null -ne $script:RunSourceInventory) { '{0:N0}' -f @($script:RunSourceInventory.Files.Keys).Count }       else { '-' }
            Write-InfoLine -Label 'Kaynak' -Value ('{0} | {1} klasör / {2} dosya' -f $script:RunSource, $sourceDirs, $sourceFiles) -Icon '📁' -Color Yellow
        }
        if ($null -ne $script:RunDestination) {
            $destDirs  = if ($null -ne $script:RunDestinationInventory) { '{0:N0}' -f @($script:RunDestinationInventory.Directories.Keys).Count } else { '-' }
            $destFiles = if ($null -ne $script:RunDestinationInventory) { '{0:N0}' -f @($script:RunDestinationInventory.Files.Keys).Count }       else { '-' }
            Write-InfoLine -Label 'Hedef' -Value ('{0} | {1} klasör / {2} dosya' -f $script:RunDestination, $destDirs, $destFiles) -Icon '🎯' -Color Green
        }
        $elapsed = Get-RunElapsed
        Write-InfoLine -Label 'Süre' -Value ('{0:hh\:mm\:ss} ({1} sn)' -f $elapsed, $elapsed.TotalSeconds.ToString('0.0')) -Icon '⌛' -Color Gray
        Write-InfoLine -Label 'Manifest' -Value ('{0:N0} satır' -f $manifestTotal) -Icon '🧾' -Color Gray
        if ($null -ne $script:RunSourceInventory) {
            $appliedBytes = Get-AppliedByteTotal -Manifest $Manifest -SourceInventory $script:RunSourceInventory
            if ($appliedBytes -gt 0) {
                Write-InfoLine -Label 'Kopyalanan veri' -Value (Format-ByteSize $appliedBytes) -Icon '📊' -Color Gray
            }
        }
        $errorColor = if ($totalErrors -gt 0) { 'Red' } else { 'Green' }
        $errorIcon  = if ($totalErrors -gt 0) { '❌' } else { '✅' }
        Write-InfoLine -Label 'Hata' -Value ('{0:N0} | Karar hatası: {1:N0} | Uygulama hatası: {2:N0}' -f $totalErrors, $script:RunCounters['KararHatasi'], $script:RunCounters['UygulamaHatasi']) -Icon $errorIcon -Color $errorColor

        # Neden @(if ...): tek öğeli sonuçta .Count $null dönmesin (Windows PowerShell 5.1).
        $skipRows = @(if ($byAction.ContainsKey('SKIP')) { $byAction['SKIP'] })
        $script:RunCounters['Atlanan'] = $skipRows.Count
        if ($skipRows.Count -gt 0) {
            Write-Host ''
            Write-Host ("ATLANAN: {0} dosya okunamadığı için YEDEKLENMEDİ" -f $skipRows.Count) -ForegroundColor Red
            foreach ($s in @($skipRows | Select-Object -First $script:ReportLimits.SkipListMax)) {
                Write-ProblemEntry -Row $s
            }
            if ($skipRows.Count -gt $script:ReportLimits.SkipListMax) {
                Write-Host ("  ... ve {0} dosya daha" -f ($skipRows.Count - $script:ReportLimits.SkipListMax)) -ForegroundColor DarkYellow
            }
        }
        Write-RunWarnings
        return
    }

    if ($script:RunCounters.Contains('ToplamHata')) {
        $script:RunCounters['ToplamHata'] = [int]$script:RunCounters['KararHatasi'] + [int]$script:RunCounters['UygulamaHatasi']
    }
    Write-Host ''; Write-Host 'GENEL ÖZET' -ForegroundColor Cyan; Write-Host '----------' -ForegroundColor DarkGray
    foreach ($k in $script:RunCounters.Keys) {
        $label = if ($script:CounterLabels.ContainsKey($k)) { $script:CounterLabels[$k] } else { $k }
        Write-Host ("  {0,-24}: {1}" -f $label,$script:RunCounters[$k])
    }
    Write-RunWarnings
}

function Save-RunReport {
    param([string]$Result = 'Tamamlandı')
    if ($script:RunCounters.Contains('ToplamHata')) {
        $script:RunCounters['ToplamHata'] = [int]$script:RunCounters['KararHatasi'] + [int]$script:RunCounters['UygulamaHatasi']
    }
    if ($script:RunLogPath) {
        Write-RunLog 'BİLGİ' ('SONUÇ: ' + ($Result -replace '^SONUÇ:\s*',''))
        Write-Host ("Log: {0}" -f $script:RunLogPath) -ForegroundColor DarkCyan
    }
    Close-RunLog
    return $null
}

function Close-RunLog {
    if ($null -eq $script:RunLogWriter) { return }
    try { $script:RunLogWriter.Dispose() } catch { }
    $script:RunLogWriter = $null
}

function Get-ErrorEventType {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $ex = $ErrorRecord.Exception
    if ($ex -is [System.UnauthorizedAccessException] -or $ex -is [System.Security.SecurityException]) { return 'ErisimHatasi' }
    return 'OkumaHatasi'
}

function Add-SkipReadError {
    param([string]$Path, [string]$Detail)
    Add-RunEvent -Type OkumaHatasi -Path $Path -Detail $Detail
}

function Add-ExecutionError {
    param(
        [string]$Path,
        [string]$Detail,
        [string]$EventType = 'ErisimHatasi'
    )
    if ($script:RunCounters.Contains('UygulamaHatasi')) { $script:RunCounters['UygulamaHatasi']++ }
    Add-RunEvent -Type $EventType -Path $Path -Detail "Uygulama hatası: $Detail"
}

# ============================================================
# MANIFEST SATIRI OLUŞTURMA
# ============================================================

function New-MergeManifestRow {
    param(
        [string]$RelativePath, [string]$SourcePath, [string]$DestinationPath,
        [string]$SourceType,   [string]$DestinationType,
        [bool]$SourceExists,   [bool]$DestinationExists,
        [bool]$TypeSame,       [bool]$LengthSame,
        [object]$CreationTimeSame,[bool]$LastWriteTimeSame, [bool]$AttributesSame,
        [string]$SourceSha256, [string]$DestinationSha256, [bool]$Sha256Same,
        [bool]$SourceReparsePoint, [string]$SourceReparseType,
        [bool]$DestinationReparsePoint, [string]$DestinationReparseType,
        [string]$Status, [string]$Action,
        [string]$Result = 'PENDING', [string]$Error = ''
    )
    [PSCustomObject]@{
        RelativePath=$RelativePath; SourcePath=$SourcePath; DestinationPath=$DestinationPath
        SourceExists=$SourceExists; DestinationExists=$DestinationExists
        SourceType=$SourceType; DestinationType=$DestinationType
        TypeSame=$TypeSame; LengthSame=$LengthSame
        CreationTimeSame=$CreationTimeSame; LastWriteTimeSame=$LastWriteTimeSame
        AttributesSame=$AttributesSame
        SourceSha256=$SourceSha256; DestinationSha256=$DestinationSha256; Sha256Same=$Sha256Same
        SourceReparsePoint=$SourceReparsePoint; SourceReparseType=$SourceReparseType
        DestinationReparsePoint=$DestinationReparsePoint; DestinationReparseType=$DestinationReparseType
        Status=$Status; Action=$Action; Result=$Result; Error=$Error
    }
}

function Add-DecisionRow {
    param(
        [Parameter(Mandatory)]$Manifest,
        [Parameter(Mandatory)]$ByPath,
        [Parameter(Mandatory)]$Row
    )
    [void]$Manifest.Add($Row)
    $ByPath[$Row.RelativePath] = $Row
}

function Get-MetaFlags {
    param($SourceMetadata, $DestinationMetadata)
    if ($null -eq $SourceMetadata -or $null -eq $DestinationMetadata) {
        return @{ Length=$false; Creation=$null; Write=$false; Attributes=$false }
    }
    return @{
        Length     = ($SourceMetadata.Length -eq $DestinationMetadata.Length)
        # CreationTime bilerek karşılaştırılmıyor: hedefteki oluşturulma zamanı yerel kalır.
        Creation   = $null
        Write      = (Test-WriteTimeSame $SourceMetadata $DestinationMetadata)
        Attributes = (Test-AttributesSame $SourceMetadata $DestinationMetadata)
    }
}

function Get-FileSystemMetadataSafe {
    <#
    .SYNOPSIS
        Dosya/klasör metadata'sını tek bir PSCustomObject olarak döner.
    .RETURNS
        [PSCustomObject] ya da $null (erişim hatası / bulunamadı).
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $long = Convert-ToLongPath $Path
        if ([System.IO.Directory]::Exists($long)) {
            $info   = New-Object -TypeName System.IO.DirectoryInfo -ArgumentList $long
            $length = $null
        }
        elseif ([System.IO.File]::Exists($long)) {
            $info   = New-Object -TypeName System.IO.FileInfo -ArgumentList $long
            $length = $info.Length
        }
        else { return $null }

        return [PSCustomObject]@{
            Length           = $length
            CreationTimeUtc  = $info.CreationTimeUtc
            LastWriteTimeUtc = $info.LastWriteTimeUtc
            Attributes       = $info.Attributes
        }
    } catch { return $null }
}

# ============================================================
# METADATA KATMANLARI
# ============================================================

function Get-BackupFileSystemName {
    param([Parameter(Mandatory)][string]$BackupRoot)
    try {
        $root = Get-DriveRootSafe $BackupRoot
        if ($root -match '^[A-Za-z]:\\$') {
            $di = New-Object System.IO.DriveInfo -ArgumentList $root.Substring(0, 1)
            if ($di.IsReady) { return [string]$di.DriveFormat }
        }
    } catch { Write-RunLog 'UYARI' "Dosya sistemi adı okunamadı: $($_.Exception.Message)" }
    return ''
}

function Get-TimeToleranceTicks {
    <#
    .SYNOPSIS
        Hedef dosya sistemine göre zaman toleransını tick cinsinden döner.
    .NOTES
        NTFS/ReFS → 0; exFAT → 20 ms; FAT/UNC/bilinmeyen → 2 sn.
    #>
    param([string]$FileSystem)
    switch -Regex ($FileSystem) {
        '^(NTFS|ReFS)$' { return [int64]0 }
        '^exFAT$'       { return [int64]200000 }
        default         { return [int64]20000000 }
    }
}

function Initialize-TimeToleranceContext {
    param([Parameter(Mandatory)][string]$BackupRoot)
    $fs    = Get-BackupFileSystemName $BackupRoot
    $ticks = Get-TimeToleranceTicks -FileSystem $fs

    $script:BackupFileSystem   = $fs
    $script:TimeToleranceTicks = [int64]$ticks
    $script:StrictMetadata     = ($ticks -eq 0)
}

function Get-TimeToleranceInfo {
    <#
    .SYNOPSIS
        Hedef dosya sistemi, zaman toleransı ve metadata modunu nesne olarak döner.
    .NOTES
        Metin yerine nesne döner; ekran ve log aynı kaynaktan beslenir.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    $fileSystem = $script:BackupFileSystem
    if ([string]::IsNullOrWhiteSpace($fileSystem)) { $fileSystem = 'bilinmiyor' }

    $seconds = ([double]$script:TimeToleranceTicks / 10000000).ToString(
        '0.###', [System.Globalization.CultureInfo]::InvariantCulture)

    $mode = if ($script:StrictMetadata) {
        'LastWriteTime + nitelikler'
    } else {
        'LastWriteTime (nitelikler uygulanır, karşılaştırılmaz)'
    }

    [PSCustomObject]@{
        FileSystem       = $fileSystem
        ToleranceSeconds = $seconds
        MetadataMode     = $mode
    }
}

function Write-InfoLine {
    <#
    .SYNOPSIS
        Başlık bloğunda "ikon etiket : değer" satırını sabit sütunlarla yazar.
    .NOTES
        Neden: emoji konsolda 2 hücre kaplar; ikon sütunu 2 hücre + 1 boşluk olarak
        sabitlenir, ikonsuz satırlar 3 boşlukla aynı sütuna hizalanır. Böylece
        etiketler ve ':' işaretleri tüm satırlarda alt alta gelir.
        Write-Host bilinçli kullanılıyor: renkli etkileşimli konsol çıktısı, boru hattı verisi değil.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
        [string]$Icon = '',
        [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray
    )

    $prefix = if ([string]::IsNullOrEmpty($Icon)) { '   ' } else { '{0} ' -f $Icon }
    Write-Host ('{0}{1,-16}: {2}' -f $prefix, $Label, $Value) -ForegroundColor $Color
}

function Write-RunEnvironmentInfo {
    <#
    .SYNOPSIS
        Dosya sistemi, ön karşılaştırma, SHA-256 kapsamı ve metadata politikasını başlıkta gösterir.
    .PARAMETER FullHash
        Tüm ortak dosyaların SHA-256 ile taranacağını belirtir.
    .PARAMETER ShowCopyBehavior
        Kopyalama sırasında uygulanacak metadata davranışını ayrıca gösterir.
    #>
    [CmdletBinding()]
    param(
        [bool]$FullHash = $false,
        [switch]$ShowCopyBehavior,
        [switch]$IntegrityCheck
    )

    $info = Get-TimeToleranceInfo
    if ($IntegrityCheck) {
        $hashMode = 'Tüm ortak dosyalar'
        $preCheck = 'Ortak dosyalar doğrudan SHA-256 ile kontrol edilir'
    } elseif ($FullHash) {
        $hashMode = 'Tüm ortak dosyalar (yeni dosyalar da hashlenir)'
        $preCheck = 'Dosya boyutu + LastWriteTime'
    } else {
        $hashMode = 'Yeni dosyalar + farklı görünen ortaklar'
        $preCheck = 'Dosya boyutu + LastWriteTime'
    }

    Write-InfoLine -Label 'Dosya sistemi' -Value $info.FileSystem -Icon '🔧' -Color DarkCyan
    Write-InfoLine -Label 'Zaman toleransı' -Value ('{0} sn' -f $info.ToleranceSeconds) -Icon '🕒' -Color DarkCyan
    Write-InfoLine -Label 'Ön kontrol' -Value $preCheck -Icon '🔍' -Color DarkCyan
    Write-InfoLine -Label 'SHA-256 kapsamı' -Value $hashMode -Icon '🔐' -Color DarkCyan
    Write-InfoLine -Label 'Metadata' -Value $info.MetadataMode -Icon '🧾' -Color DarkCyan

    if ($ShowCopyBehavior) {
        Write-Host ''
        Write-Host '📋 KOPYALAMA SIRASINDA UYGULANANLAR' -ForegroundColor Cyan
        Write-Host '  • LastWriteTime : Kaynaktan hedefe uygulanır ve tekrar okunarak doğrulanır.' -ForegroundColor DarkGray
        if ($script:StrictMetadata) {
            Write-Host '  • Nitelikler      : Hedefe uygulanır; dosyalarda ayrıca doğrulanır.' -ForegroundColor DarkGray
        } else {
            Write-Host '  • Nitelikler      : Hedefe uygulanır; bu dosya sisteminde ayrıca doğrulanmaz.' -ForegroundColor DarkGray
        }
        Write-Host '  • CreationTime    : Kaynaktan alınmaz; hedef kopyanın oluşturulma zamanı korunur.' -ForegroundColor DarkGray
        Write-Host '  • LastAccessTime  : Kaynakla eşitlenmez; hedefte işletim sistemi yönetir.' -ForegroundColor DarkGray
    }

    if ($script:RunLogPath) {
        Write-InfoLine -Label 'Log' -Value $script:RunLogPath -Icon '📝' -Color DarkCyan
    }
}

function Get-TimeToleranceDescription {
    $info = Get-TimeToleranceInfo
    return ('Hedef dosya sistemi: {0} | Zaman toleransı: {1} sn | Metadata karşılaştırma: {2}' -f $info.FileSystem, $info.ToleranceSeconds, $info.MetadataMode)
}

function Test-TimeClose {
    param([datetime]$A, [datetime]$B)
    $diff = [math]::Abs(($A - $B).Ticks)
    return ($diff -le [int64]$script:TimeToleranceTicks)
}

function Test-WriteTimeSame {
    param($SourceMetadata, $DestinationMetadata)
    if ($null -eq $SourceMetadata -or $null -eq $DestinationMetadata) { return $false }
    return (Test-TimeClose $SourceMetadata.LastWriteTimeUtc $DestinationMetadata.LastWriteTimeUtc)
}

function Test-AttributesSame {
    param($SourceMetadata, $DestinationMetadata)
    if ($null -eq $SourceMetadata -or $null -eq $DestinationMetadata) { return $false }
    $mask = [int]$script:SettableAttributeMask
    return (([int]$SourceMetadata.Attributes -band $mask) -eq ([int]$DestinationMetadata.Attributes -band $mask))
}

function Set-FileSystemMetadata {
    <#
    .SYNOPSIS
        Hedef nesneye kaynak metadata'sını uygular.
    .NOTES
        LastWriteTime kaynaktan uygulanır; CreationTime hedefte doğal kalır.
        Attributes en iyi çabayla uygulanır. LastAccessTime kaynaktan eşitlenmez.
        Hata mesajı $script:LastMetadataErrorMessage içine yazılır; ekrana basılmaz.
    .RETURNS
        $true → LastWriteTime uygulandı; $false → kritik kısım başarısız.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Metadata
    )
    $script:LastMetadataWarning = ''
    $script:LastMetadataErrorMessage = ''

    try {
        $long = Convert-ToLongPath $Path
        $isFile = [System.IO.File]::Exists($long)
        $isDirectory = [System.IO.Directory]::Exists($long)

        if (-not ($isFile -or $isDirectory)) { throw "Hedef nesne bulunamadı: $Path" }

        $info = if ($isFile) {
            New-Object -TypeName System.IO.FileInfo -ArgumentList $long
        } else {
            New-Object -TypeName System.IO.DirectoryInfo -ArgumentList $long
        }

        $originalAttributes = [System.IO.File]::GetAttributes($long)
        $hadReadOnly = (([int]$originalAttributes -band 1) -ne 0)

        if ($hadReadOnly) {
            [System.IO.File]::SetAttributes(
                $long,
                [System.Enum]::ToObject([System.IO.FileAttributes], ([int]$originalAttributes -band (-bnot 1)))
            )
        }

        $warnings = New-Object System.Collections.Generic.List[string]

        try { $info.LastWriteTimeUtc = $Metadata.LastWriteTimeUtc }
        catch {
            if ($hadReadOnly) { try { [System.IO.File]::SetAttributes($long, $originalAttributes) } catch { } }
            throw
        }

        try {
            $current = [int][System.IO.File]::GetAttributes($long)
            $mask    = [int]$script:SettableAttributeMask
            $new     = ($current -band (-bnot $mask)) -bor ([int]$Metadata.Attributes -band $mask)

            if (($new -band (-bnot 128)) -ne 0) { $new = $new -band (-bnot 128) }
            elseif ($new -eq 0) { $new = 128 }

            if ($new -ne $current) {
                [System.IO.File]::SetAttributes(
                    $long,
                    [System.Enum]::ToObject([System.IO.FileAttributes], $new))
            }
        } catch { [void]$warnings.Add("Nitelikler uygulanamadı: $($_.Exception.Message)") }

        $script:LastMetadataWarning = ($warnings -join ' | ')
        return $true
    }
    catch {
        $script:LastMetadataErrorMessage = $_.Exception.Message
        return $false
    }
}

function Test-FileSystemMetadataSame {
    <#
    .SYNOPSIS
        Kaynak/hedef metadata eşleşmesini test eder.
    .PARAMETER CriticalOnly
        Sadece LastWriteTime karşılaştırılır.
    .PARAMETER Full
        StrictMetadata kapalı olsa bile nitelikler de karşılaştırılır. CreationTime karşılaştırılmaz.
    #>
    param(
        $SourceMetadata,
        $DestinationMetadata,
        [switch]$CriticalOnly,
        [switch]$Full
    )
    if ($null -eq $SourceMetadata -or $null -eq $DestinationMetadata) { return $false }

    if (-not (Test-WriteTimeSame $SourceMetadata $DestinationMetadata)) { return $false }
    if ($CriticalOnly) { return $true }

    if ($Full -or $script:StrictMetadata) {
        if (-not (Test-AttributesSame $SourceMetadata $DestinationMetadata)) { return $false }
    }
    return $true
}

function Get-FileHashSafe {
    <#
    .SYNOPSIS
        Dosyanın SHA-256 özetini döner; hata durumunda 1 kez yeniden dener.
    .RETURNS
        Küçük harfli hex string ya da $null.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $script:LastHashErrorType = ''
    $script:LastHashErrorMessage = ''

    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        $sha = $null; $stream = $null
        try {
            $sha = [System.Security.Cryptography.SHA256]::Create()
            $stream = [System.IO.File]::Open(
                (Convert-ToLongPath $Path),
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
            return ([System.BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '')
        }
        catch {
            $script:LastHashErrorType = Get-ErrorEventType $_
            $script:LastHashErrorMessage = $_.Exception.Message
            if ($attempt -eq 0) { Start-Sleep -Milliseconds 300 }
        }
        finally {
            if ($stream) { $stream.Dispose() }
            if ($sha)    { $sha.Dispose() }
        }
    }
    return $null
}

function Move-PathToArchive {
    <#
    .SYNOPSIS
        Hedefi .Changed/.Deleted arşivine taşır, metadata'sını geri yükler ve doğrular.
    .RETURNS
        $true / $false. Hata detayı $script:LastArchiveError* değişkenlerinde.
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    $script:LastArchiveErrorType = ''
    $script:LastArchiveErrorMessage = ''

    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$script:RunBackupRoot) -and
            -not [string]::IsNullOrWhiteSpace([string]$script:RunDestination)) {
            $blocked = Test-IntermediateReparsePoints `
                -BackupRoot $script:RunBackupRoot `
                -DestinationRoot $script:RunDestination `
                -AdditionalPaths @($Source, $Destination)
            if ($null -ne $blocked) { throw "Arşiv yolunda engellenen Reparse Point: $blocked" }
        }

        $isFile = Test-FileExists $Source
        $isDirectory = Test-DirectoryExists $Source
        if (-not ($isFile -or $isDirectory)) { throw "Arşivlenecek kaynak bulunamadı: $Source" }
        if ($isDirectory -and (Test-IsBlockedDirectoryReparsePoint $Source)) {
            throw "Engellenen Reparse Point arşivlenemez: $Source"
        }

        $metadata = Get-FileSystemMetadataSafe $Source
        if ($null -eq $metadata) { throw "Kaynak metadata okunamadı: $Source" }

        $parent = [System.IO.Path]::GetDirectoryName($Destination)
        if (-not (Ensure-Directory $parent)) { throw "Arşiv klasörü oluşturulamadı: $parent" }

        if ((Test-FileExists $Destination) -or (Test-DirectoryExists $Destination)) {
            throw "Arşiv hedefi zaten mevcut: $Destination"
        }

        if ($isFile -and -not (Remove-ReadOnlyAttribute $Source)) {
            $script:LastArchiveErrorType = 'ErisimHatasi'
            $script:LastArchiveErrorMessage = 'ReadOnly niteliği kaldırılamadı.'
            return $false
        }

        if ($isFile) {
            [System.IO.File]::Move((Convert-ToLongPath $Source), (Convert-ToLongPath $Destination))
        } else {
            [System.IO.Directory]::Move((Convert-ToLongPath $Source), (Convert-ToLongPath $Destination))
        }

        if (-not (Set-FileSystemMetadata -Path $Destination -Metadata $metadata)) {
            Add-RunWarning -Path $Destination -Detail 'Arşivlenen nesnenin metadata geri yüklemesi başarısız'
        }

        $moveVerified = (-not (Test-FileExists $Source)) -and
                        (-not (Test-DirectoryExists $Source)) -and
                        ((Test-FileExists $Destination) -or (Test-DirectoryExists $Destination))
        if (-not $moveVerified) { throw "Taşıma sonrası doğrulama başarısız: $Source -> $Destination" }
        return $true
    }
    catch {
        $script:LastArchiveErrorType = Get-ErrorEventType $_
        $script:LastArchiveErrorMessage = $_.Exception.Message
        return $false
    }
}

function Build-Inventory {
    <#
    .SYNOPSIS
        Bir kökün altındaki dosya/klasör envanterini çıkarır.
    .RETURNS
        [PSCustomObject]@{ Root; RootMetadata; Directories; Files; Errors; BlockedReparsePoints; HasErrors }
    #>
    param([Parameter(Mandatory)][string]$Root)

    $dirs  = @{}
    $files = @{}
    $errors = New-Object System.Collections.Generic.List[object]
    $blockedReparsePoints = New-Object System.Collections.Generic.List[string]

    $rootMetadata = Get-FileSystemMetadataSafe $Root
    if ($null -eq $rootMetadata) {
        [void]$errors.Add([PSCustomObject]@{
            EventType='OkumaHatasi'; Path=$Root
            Detail='Kök metadata okunamadı.'
            Message="Kök metadata okunamadı: $Root"
        })
    }

    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    $scanCount = 0
    Start-ProgressLine
    Update-ProgressLine -Index 0 -Total 0 -Verb 'Envanter taranıyor' -Path $Root -Force

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()

        try {
            foreach ($fileLong in [System.IO.Directory]::GetFiles((Convert-ToLongPath $current))) {
                $file = Get-DisplayPath $fileLong
                try {
                    $relative = Get-RelativePathSafe -FullPath $file -Root $Root
                    if ([string]::IsNullOrWhiteSpace($relative)) { continue }

                    $metadata = Get-FileSystemMetadataSafe $file
                    if ($null -eq $metadata) { throw 'Metadata okunamadı.' }

                    $files[$relative] = [PSCustomObject]@{ Length = $metadata.Length; Metadata = $metadata }
                    $scanCount++
                    Update-ProgressLine -Index $scanCount -Total 0 -Verb ("Envanter: {0:N0} dosya" -f $scanCount) -Path $relative
                }
                catch {
                    $eventType = Get-ErrorEventType $_
                    [void]$errors.Add([PSCustomObject]@{
                        EventType=$eventType; Path=$file
                        Detail=$_.Exception.Message
                        Message="Dosya yolu işlenemedi: $file -- $($_.Exception.Message)"
                    })
                }
            }
        }
        catch {
            $eventType = Get-ErrorEventType $_
            [void]$errors.Add([PSCustomObject]@{
                EventType=$eventType; Path=$current
                Detail=$_.Exception.Message
                Message="Dosyalar okunamadı: $current -- $($_.Exception.Message)"
            })
        }

        try {
            foreach ($dirLong in [System.IO.Directory]::GetDirectories((Convert-ToLongPath $current))) {
                $dir = Get-DisplayPath $dirLong
                try {
                    $relative = Get-RelativePathSafe -FullPath $dir -Root $Root
                    if ([string]::IsNullOrWhiteSpace($relative)) { continue }

                    if (Test-IsBlockedDirectoryReparsePoint $dir) {
                        [void]$blockedReparsePoints.Add($relative)
                        continue
                    }

                    $metadata = Get-FileSystemMetadataSafe $dir
                    if ($null -eq $metadata) { throw 'Metadata okunamadı.' }

                    $dirs[$relative] = [PSCustomObject]@{ Metadata = $metadata }
                    Update-ProgressLine -Index $scanCount -Total 0 -Verb ("Envanter: {0:N0} dosya" -f $scanCount) -Path $relative
                    $stack.Push($dir)
                }
                catch {
                    $eventType = Get-ErrorEventType $_
                    [void]$errors.Add([PSCustomObject]@{
                        EventType=$eventType; Path=$dir
                        Detail=$_.Exception.Message
                        Message="Klasör yolu işlenemedi: $dir -- $($_.Exception.Message)"
                    })
                }
            }
        }
        catch {
            $eventType = Get-ErrorEventType $_
            [void]$errors.Add([PSCustomObject]@{
                EventType=$eventType; Path=$current
                Detail=$_.Exception.Message
                Message="Alt klasörler okunamadı: $current -- $($_.Exception.Message)"
            })
        }
    }

    Stop-ProgressLine

    return [PSCustomObject]@{
        Root=$Root; RootMetadata=$rootMetadata
        Directories=$dirs; Files=$files
        Errors=$errors; BlockedReparsePoints=$blockedReparsePoints
        HasErrors=($errors.Count -gt 0)
    }
}

# ============================================================
# ARŞİV YOLLARI
# ============================================================

function Get-ArchiveFilePath {
    param(
        [Parameter(Mandatory)][string]$ArchiveRoot,
        [Parameter(Mandatory)][string]$RelativeFile,
        [Parameter(Mandatory)][string]$Timestamp
    )
    $parent = [System.IO.Path]::GetDirectoryName($RelativeFile)
    $name   = [System.IO.Path]::GetFileName($RelativeFile)

    $folder = if ($parent) {
        Join-Path (Join-Path $ArchiveRoot $parent) $Timestamp
    } else {
        Join-Path $ArchiveRoot $Timestamp
    }

    $path = Join-Path $folder $name
    $n = 1
    while ((Test-FileExists $path) -or (Test-DirectoryExists $path)) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        $ext  = [System.IO.Path]::GetExtension($name)
        $path = Join-Path $folder ("{0}_{1}{2}" -f $base, $n, $ext)
        $n++
    }
    return $path
}

function Get-ArchiveDirectoryPath {
    param(
        [Parameter(Mandatory)][string]$ArchiveRoot,
        [Parameter(Mandatory)][string]$RelativeDirectory,
        [Parameter(Mandatory)][string]$Timestamp
    )
    $base = Join-Path $ArchiveRoot $RelativeDirectory
    $path = Join-Path $base $Timestamp
    $n = 1
    while ((Test-DirectoryExists $path) -or (Test-FileExists $path)) {
        $path = Join-Path $base ("{0}_{1}" -f $Timestamp, $n)
        $n++
    }
    return $path
}

function Get-Depth {
    param([string]$RelativePath)
    if (-not $RelativePath) { return 0 }
    return (($RelativePath -split '\\').Count)
}

function Show-InventoryErrors {
    param(
        [Parameter(Mandatory)]$Inventory,
        [string]$Title = 'Tarama Hataları'
    )
    if (-not $Inventory.HasErrors) { return }

    Write-Host "[$Title]" -ForegroundColor Red
    foreach ($e in $Inventory.Errors) {
        if ($null -ne $e.PSObject.Properties['EventType']) {
            $eventType = [string]$e.EventType
            if ($eventType -notin @('ErisimHatasi','OkumaHatasi','MetadataErisimHatasi','MetadataOkumaHatasi')) {
                $eventType = 'OkumaHatasi'
            }
            Add-RunEvent -Type $eventType -Path ([string]$e.Path) -Detail ([string]$e.Detail)
        } else {
            Add-RunEvent -Type 'OkumaHatasi' -Path $Title -Detail ([string]$e)
        }
    }
}

function Find-ReparseConflict {
    param(
        [Parameter(Mandatory)]$SourceInventory,
        [Parameter(Mandatory)]$TargetInventory
    )
    foreach ($blocked in @($TargetInventory.BlockedReparsePoints)) {
        foreach ($rel in $SourceInventory.Directories.Keys) {
            if (Test-IsUnderRelativePath -Child $rel -Parent $blocked) {
                return [PSCustomObject]@{ Path=$blocked; Type='Engellenen hedef klasör Reparse Point' }
            }
        }
        foreach ($rel in $SourceInventory.Files.Keys) {
            if (Test-IsUnderRelativePath -Child $rel -Parent $blocked) {
                return [PSCustomObject]@{ Path=$blocked; Type='Engellenen hedef klasör Reparse Point' }
            }
        }
    }

    foreach ($rel in $TargetInventory.Files.Keys) {
        $targetItem = $TargetInventory.Files[$rel]
        if ($null -eq $targetItem -or $null -eq $targetItem.Metadata) { continue }
        $attrs = [int]$targetItem.Metadata.Attributes
        if (($attrs -band [int][System.IO.FileAttributes]::ReparsePoint) -eq 0) { continue }

        if ($SourceInventory.Files.ContainsKey($rel) -or $SourceInventory.Directories.ContainsKey($rel)) {
            return [PSCustomObject]@{ Path=$rel; Type='Hedef dosya Reparse Point' }
        }
    }
    return $null
}

function Find-BlockedReparseComponent {
    param([Parameter(Mandatory)][string]$Path)

    $full = Get-NormalizedFullPath $Path
    $root = [System.IO.Path]::GetPathRoot($full)
    if ([string]::IsNullOrWhiteSpace($root)) { throw "Yol kökü alınamadı: $Path" }

    $current = $root
    if ((Test-DirectoryExists $current) -and (Test-IsBlockedDirectoryReparsePoint $current)) { return $current }

    $tail = $full.Substring($root.Length)
    if ([string]::IsNullOrWhiteSpace($tail)) { return $null }

    foreach ($segment in ($tail -split '[\\/]+')) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }
        $current = Join-Path $current $segment
        $isDirectory = Test-DirectoryExists $current
        $attrs = $null
        try { $attrs = [System.IO.File]::GetAttributes((Convert-ToLongPath $current)) } catch { }

        if ($null -ne $attrs -and (($attrs -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
            if ($isDirectory) {
                if (Test-IsBlockedDirectoryReparsePoint $current) { return $current }
            } else { return $current }
        }
        if (-not $isDirectory) { break }
    }

    if (Test-FileExists $full) {
        try {
            $attrs = [System.IO.File]::GetAttributes((Convert-ToLongPath $full))
            if (($attrs -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return $full }
        }
        catch { throw "Dosya Reparse Point denetimi yapılamadı: $full -- $($_.Exception.Message)" }
    }
    return $null
}

function Test-IntermediateReparsePoints {
    param(
        [Parameter(Mandatory)][string]$BackupRoot,
        [Parameter(Mandatory)][string]$DestinationRoot,
        [string[]]$AdditionalPaths = @()
    )
    $backupFull = Get-NormalizedFullPath $BackupRoot
    $destFull   = Get-NormalizedFullPath $DestinationRoot
    if (-not (Test-PathEqualOrUnder -Child $destFull -Parent $backupFull)) {
        throw "Hedef, yedek kökünün altında değil: $destFull"
    }

    $paths = @($backupFull, $destFull) + @($AdditionalPaths)
    foreach ($path in $paths) {
        if ([string]::IsNullOrWhiteSpace([string]$path)) { continue }
        $full = Get-NormalizedFullPath ([string]$path)
        if (-not (Test-PathEqualOrUnder -Child $full -Parent $backupFull)) {
            throw "Denetlenecek yol yedek kökünün dışında: $full"
        }
        $blocked = Find-BlockedReparseComponent -Path $full
        if ($null -ne $blocked) { return [string]$blocked }
    }
    return $null
}

function Assert-NoBlockedReparse {
    param(
        [Parameter(Mandatory)][string]$BackupRoot,
        [Parameter(Mandatory)][string]$DestinationRoot,
        [string[]]$AdditionalPaths = @(),
        [string]$Reason = 'Reparse denetimi'
    )
    $blocked = Test-IntermediateReparsePoints `
        -BackupRoot $BackupRoot `
        -DestinationRoot $DestinationRoot `
        -AdditionalPaths $AdditionalPaths
    if ($null -ne $blocked) { throw "[$Reason] Engellenen Reparse Point: $blocked" }
}

# ============================================================
# AYARLAR MENÜSÜ
# ============================================================

function Wait-ForSettingsMenuEnter {
    while ($true) {
        $response = Read-Host 'Ayarlar Menüsüne dönmek için Enter'
        if ([string]::IsNullOrEmpty($response)) { return }
        Write-Host 'Devam etmek için yalnızca Enter tuşuna basın.' -ForegroundColor DarkGray
    }
}

function Settings-Menu {
    while ($true) {
        $settings = Load-Settings
        if ($null -ne $settings) {
            $Source = $settings.Source
            $BackupRoot = $settings.BackupRoot
        } else { $Source = ''; $BackupRoot = '' }

        $destination = $null
        if (-not [string]::IsNullOrWhiteSpace($Source) -and
            -not [string]::IsNullOrWhiteSpace($BackupRoot)) {
            try { $destination = Get-DestinationRoot -Source $Source -BackupRoot $BackupRoot }
            catch { $destination = $null }
        }

        Clear-Host
        Write-Host ('-' * 72) -ForegroundColor Magenta
        Write-Host '⚙️ YEDEKLEME AYARLARI' -ForegroundColor Magenta
        Write-Host ('-' * 72) -ForegroundColor Magenta
        Write-Host ''
        Write-Host '⚙️ MEVCUT AYARLAR' -ForegroundColor Cyan
        Write-Host ('-' * 72) -ForegroundColor DarkGray
        Write-Host '📁 Mevcut kaynak:'
        if ([string]::IsNullOrWhiteSpace($Source)) { Write-Host '   YOK' -ForegroundColor Red }
        else { Write-Host ('   {0}' -f $Source) -ForegroundColor DarkGray }

        Write-Host '💾 Mevcut yedek:'
        if ([string]::IsNullOrWhiteSpace($BackupRoot)) { Write-Host '   YOK' -ForegroundColor Red }
        else { Write-Host ('   {0}' -f $BackupRoot) -ForegroundColor DarkGray }

        Write-Host '🎯 Hesaplanan hedef:'
        if ($null -ne $destination) { Write-Host ('   {0}' -f $destination) -ForegroundColor DarkGray }
        else { Write-Host '   HESAPLANAMADI' -ForegroundColor Red }

        Write-Host ('-' * 72) -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  1. ⚙️  Ayarları Düzenle'
        Write-Host '  Q  ↩️  Ana Menü'
        Write-Host ''
        Write-Host ('-' * 72) -ForegroundColor DarkGray
        Write-Host ''

        $choice = Read-Host 'Seçiminiz'
        if ($choice -match '^[Qq]$') { return 'MENU' }

        if ($choice -eq '1') {
            $OriginalSource = $Source
            $OriginalBackupRoot = $BackupRoot

            Clear-Host
            Write-Host ('-' * 72) -ForegroundColor Magenta
            Write-Host '⚙️ YEDEKLEME AYARLARI' -ForegroundColor Magenta
            Write-Host ('-' * 72) -ForegroundColor Magenta
            Write-Host ''
            Write-Host '⚙️ AYARLARI DÜZENLE' -ForegroundColor Cyan
            Write-Host ('-' * 72) -ForegroundColor DarkGray
            Write-Host ''
            Write-Host '📁 Mevcut Kaynak    : ' -NoNewline
            if ([string]::IsNullOrWhiteSpace($Source)) { Write-Host 'YOK' -ForegroundColor Red }
            else { Write-Host $Source -ForegroundColor Yellow }

            Write-Host '💾 Mevcut Yedek     : ' -NoNewline
            if ([string]::IsNullOrWhiteSpace($BackupRoot)) { Write-Host 'YOK' -ForegroundColor Red }
            else { Write-Host $BackupRoot -ForegroundColor Yellow }

            Write-Host '🎯 Hesaplanan Hedef : ' -NoNewline
            if ($null -ne $destination) { Write-Host $destination -ForegroundColor Green }
            else { Write-Host 'HESAPLANAMADI' -ForegroundColor Red }
            Write-Host ''

            while ($true) {
                $newSource = Read-Host 'Yeni KAYNAK klasör yolu (boş = değiştirme)'
                if ($newSource -match '^[Qq]$') { break }
                if ([string]::IsNullOrWhiteSpace($newSource)) { break }

                try {
                    $fullSource = Get-NormalizedFullPath $newSource.Trim()
                    if (-not (Test-DirectoryExists $fullSource)) { throw 'Kaynak klasörü bulunamadı.' }

                    foreach ($reserved in @($ChangedFolderName, $DeletedFolderName)) {
                        $reservedPath = Join-Path $fullSource $reserved
                        if ((Test-DirectoryExists $reservedPath) -or (Test-FileExists $reservedPath)) {
                            throw "Kaynakta ayrılmış isim kullanılıyor: $reserved"
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($BackupRoot)) {
                        if (-not (Test-DifferentDrive -Source $fullSource -BackupRoot $BackupRoot)) {
                            throw 'Kaynak ve yedek farklı sürücülerde olmalıdır.'
                        }
                        $testDestination = Get-DestinationRoot -Source $fullSource -BackupRoot $BackupRoot
                        if (Test-PathEqualOrUnder -Child $testDestination -Parent $fullSource) {
                            throw 'Yedek hedefi kaynak klasörünün içine düşüyor.'
                        }
                    }

                    $Source = $fullSource
                    Write-Host ''
                    Write-Host 'Kaynak konumu güncellendi.' -ForegroundColor Green
                    Write-Host ''

                    if (-not [string]::IsNullOrWhiteSpace($BackupRoot)) {
                        try {
                            $previewDestination = Get-DestinationRoot -Source $Source -BackupRoot $BackupRoot
                            Write-Host 'Hesaplanan Hedef : ' -NoNewline
                            Write-Host $previewDestination -ForegroundColor Green
                        } catch {}
                    }
                    break
                }
                catch {
                    Write-Host ''
                    Write-Host "[HATA] $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host ''
                }
            }
            if ($newSource -match '^[Qq]$') { continue }

            while ($true) {
                Write-Host ''
                Write-Host '💾 Mevcut hedef:'
                if ([string]::IsNullOrWhiteSpace($BackupRoot)) { Write-Host '   YOK' -ForegroundColor Red }
                else { Write-Host ('   {0}' -f $BackupRoot) -ForegroundColor DarkGray }
                Write-Host ''

                $newBackup = Read-Host 'Yeni HEDEF klasör yolu (boş = değiştirme)'
                if ($newBackup -match '^[Qq]$') { break }
                if ([string]::IsNullOrWhiteSpace($newBackup)) { break }

                try {
                    $fullBackup = Get-NormalizedFullPath $newBackup.Trim()
                    $backupDrive = Get-DriveRootSafe $fullBackup
                    if (-not (Test-DirectoryExists $backupDrive)) {
                        throw "Yedek sürücüsü erişilemiyor: $backupDrive"
                    }

                    if (-not [string]::IsNullOrWhiteSpace($Source)) {
                        if (-not (Test-DifferentDrive -Source $Source -BackupRoot $fullBackup)) {
                            throw 'Kaynak ve yedek farklı sürücülerde olmalıdır.'
                        }
                        $testDestination = Get-DestinationRoot -Source $Source -BackupRoot $fullBackup
                        if (Test-PathEqualOrUnder -Child $testDestination -Parent $Source) {
                            throw 'Yedek hedefi kaynak klasörünün içine düşüyor.'
                        }
                    }

                    $BackupRoot = $fullBackup
                    Write-Host ''
                    Write-Host 'Yedek konumu güncellendi.' -ForegroundColor Green
                    Write-Host ''

                    if (-not [string]::IsNullOrWhiteSpace($Source)) {
                        try {
                            $previewDestination = Get-DestinationRoot -Source $Source -BackupRoot $BackupRoot
                            Write-Host 'Hesaplanan Hedef : ' -NoNewline
                            Write-Host $previewDestination -ForegroundColor Green
                        } catch {}
                    }
                    break
                }
                catch {
                    Write-Host ''
                    Write-Host "[HATA] $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host ''
                }
            }
            if ($newBackup -match '^[Qq]$') { continue }

            $dirty = ($Source -ne $OriginalSource -or $BackupRoot -ne $OriginalBackupRoot)
            if (-not $dirty) {
                Write-Host ''
                Write-Host 'Ayar değişikliği yapılmadı.' -ForegroundColor Yellow
                Write-Host ''
                Wait-ForSettingsMenuEnter
                continue
            }

            if ([string]::IsNullOrWhiteSpace($Source)) {
                Write-Host ''
                Write-Host '[HATA] Kaynak konumu belirlenmemiş.' -ForegroundColor Red
                Wait-ForSettingsMenuEnter
                continue
            }
            if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
                Write-Host ''
                Write-Host '[HATA] Yedek konumu belirlenmemiş.' -ForegroundColor Red
                Wait-ForSettingsMenuEnter
                continue
            }

            Write-Host ''
            Write-Host 'Ayarlar doğrulanıyor...' -ForegroundColor Cyan

            if (-not (Test-SettingsValid -Source $Source -BackupRoot $BackupRoot -ShowError)) {
                Write-Host ''
                Write-Host 'Ayarlar kaydedilmedi.' -ForegroundColor Red
                Wait-ForSettingsMenuEnter
                continue
            }

            if (Save-Settings -Source $Source -BackupRoot $BackupRoot) {
                $saved = Load-Settings
                if ($null -ne $saved -and $saved.Source -eq $Source -and $saved.BackupRoot -eq $BackupRoot) {
                    Write-Host ''
                    Write-Host '============================================================' -ForegroundColor Green
                    Write-Host '                    AYARLAR KAYDEDİLDİ' -ForegroundColor Green
                    Write-Host '============================================================' -ForegroundColor Green
                    Write-Host ''
                    Write-Host "Kaynak : $Source"
                    Write-Host "Yedek  : $BackupRoot"
                    try {
                        $savedDestination = Get-DestinationRoot -Source $Source -BackupRoot $BackupRoot
                        Write-Host "Hedef  : $savedDestination" -ForegroundColor Green
                    } catch {}
                } else {
                    Write-Host ''
                    Write-Host 'Ayarlar yazıldı ancak doğrulanamadı.' -ForegroundColor Red
                }
            }

            Write-Host ''
            Wait-ForSettingsMenuEnter
            continue
        }

        Write-Host ''
        Write-Host 'Geçersiz seçim.' -ForegroundColor Red
        Start-Sleep -Milliseconds 700
    }
}

# ============================================================
# YEDEKLEME ÇEKİRDEĞİ
# ============================================================

function Confirm-FileMetadataApplied {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$SourceMetadata,
        [string]$Label = ''
    )
    $dstMeta = Get-FileSystemMetadataSafe $Path

    if (-not (Test-WriteTimeSame $SourceMetadata $dstMeta)) {
        throw "LastWriteTime doğrulaması başarısız ($Label)"
    }
    if ($script:StrictMetadata -and
        -not (Test-FileSystemMetadataSame -SourceMetadata $SourceMetadata -DestinationMetadata $dstMeta -Full)) {
        Add-RunWarning -Path $Path -Detail "Dosya nitelikleri tam eşleşmedi ($Label)"
    }
}

function Copy-FileAtomic {
    <#
    .SYNOPSIS
        Kaynak dosyayı .Temp'te hazırlar, doğrular, (gerekirse) eski hedefi arşive alır ve yerine koyar.
    .NOTES
        Hata olursa hedef bozulmaz; arşivlenmiş eski sürüm geri alınır.
    #>
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$TempRoot,
        [string]$ExpectedSha256 = '',
        [string]$ArchivePath = ''
    )
    $parent = [System.IO.Path]::GetDirectoryName($DestinationPath)

    if (-not [string]::IsNullOrWhiteSpace([string]$script:RunBackupRoot) -and
        -not [string]::IsNullOrWhiteSpace([string]$script:RunDestination)) {
        $blocked = Test-IntermediateReparsePoints `
            -BackupRoot $script:RunBackupRoot `
            -DestinationRoot $script:RunDestination `
            -AdditionalPaths @($DestinationPath, $TempRoot)
        if ($null -ne $blocked) { throw "Kopyalama yolunda engellenen Reparse Point: $blocked" }
    }

    if (-not (Test-DirectoryExists $parent)) {
        if (-not (Ensure-Directory $parent)) { throw "Hedef klasör oluşturulamadı: $parent" }
    }
    if (-not (Ensure-Directory $TempRoot)) { throw "Geçici klasör oluşturulamadı: $TempRoot" }

    $tmp = Join-Path $TempRoot ([guid]::NewGuid().ToString('N') + '.tmp')
    $archived = $false
    $newFilePlaced = $false

    try {
        $meta = Get-FileSystemMetadataSafe $SourcePath
        if ($null -eq $meta) { throw 'Kaynak metadata okunamadı' }
        if (($meta.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            $sourceTag = Get-ReparsePointTag $SourcePath
            if (($null -eq $sourceTag) -or (-not (Test-IsCloudFilesReparseTag ([uint32]$sourceTag)))) {
                throw 'Kaynak dosya Reparse Point; bağlantı hedefi izlenerek kopyalanmayacak.'
            }
        }

        [System.IO.File]::Copy((Convert-ToLongPath $SourcePath), (Convert-ToLongPath $tmp), $false)

        if (-not (Set-FileSystemMetadata -Path $tmp -Metadata $meta)) {
            throw $(if ($script:LastMetadataErrorMessage) { $script:LastMetadataErrorMessage } else { 'Geçici dosyaya metadata uygulanamadı' })
        }
        if ($script:LastMetadataWarning) {
            Add-RunWarning -Path $DestinationPath -Detail $script:LastMetadataWarning
        }

        Confirm-FileMetadataApplied -Path $tmp -SourceMetadata $meta -Label 'geçici dosya'

        if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
            $tmpHash = Get-FileHashSafe $tmp
            if ([string]::IsNullOrWhiteSpace($tmpHash)) { throw 'Geçici kopya SHA-256 okunamadı' }
            if ($tmpHash -ne $ExpectedSha256) {
                throw 'Geçici kopya SHA-256 eşleşmiyor (kaynak kopya sırasında değişmiş olabilir)'
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($ArchivePath)) {
            if (-not (Move-PathToArchive -Source $DestinationPath -Destination $ArchivePath)) {
                $archived = (Test-FileExists $ArchivePath) -or (Test-DirectoryExists $ArchivePath)
                $why = if ($script:LastArchiveErrorMessage) { $script:LastArchiveErrorMessage } else { 'bilinmeyen hata' }
                throw "Mevcut hedef arşive taşınamadı: $why"
            }
            $archived = $true
        }
        elseif ((Test-FileExists $DestinationPath) -or (Test-DirectoryExists $DestinationPath)) {
            throw 'Hedef zaten mevcut'
        }

        [System.IO.File]::Move((Convert-ToLongPath $tmp), (Convert-ToLongPath $DestinationPath))
        $newFilePlaced = $true

        Confirm-FileMetadataApplied -Path $DestinationPath -SourceMetadata $meta -Label 'taşıma sonrası'
        return $true
    }
    catch {
        $err = $_.Exception.Message

        if ($newFilePlaced -and (Test-FileExists $DestinationPath)) {
            try {
                [void](Remove-ReadOnlyAttribute $DestinationPath)
                [System.IO.File]::Delete((Convert-ToLongPath $DestinationPath))
                $newFilePlaced = $false
            }
            catch { $err += " | Hatalı yeni dosya hedeften kaldırılamadı: $($_.Exception.Message)" }
        }

        if ($archived -and -not $newFilePlaced -and
            -not (Test-FileExists $DestinationPath) -and
            -not (Test-DirectoryExists $DestinationPath)) {
            try {
                if (Test-DirectoryExists $ArchivePath) {
                    [System.IO.Directory]::Move((Convert-ToLongPath $ArchivePath), (Convert-ToLongPath $DestinationPath))
                } else {
                    [System.IO.File]::Move((Convert-ToLongPath $ArchivePath), (Convert-ToLongPath $DestinationPath))
                }
                $err += ' (eski sürüm yerine geri alındı)'
            }
            catch { $err += " | GERİ ALMA BAŞARISIZ; eski sürüm burada: $ArchivePath" }
        }
        throw $err
    }
    finally {
        if (Test-FileExists $tmp) {
            try {
                [void](Remove-ReadOnlyAttribute $tmp)
                [System.IO.File]::Delete((Convert-ToLongPath $tmp))
            } catch { }
        }
    }
}

function Clear-StaleTemp {
    param([Parameter(Mandatory)][string]$TempRoot)
    try {
        if (-not (Test-DirectoryExists $TempRoot)) { return }
        if (Test-IsBlockedDirectoryReparsePoint $TempRoot) {
            Add-RunWarning -Path $TempRoot -Detail 'Geçici klasör Reparse Point; temizlenmedi'
            return
        }
        foreach ($f in [System.IO.Directory]::GetFiles((Convert-ToLongPath $TempRoot), '*.tmp')) {
            try {
                [void](Remove-ReadOnlyAttribute $f)
                [System.IO.File]::Delete($f)
            }
            catch {
                Add-RunWarning -Path (Get-DisplayPath $f) -Detail "Eski geçici dosya silinemedi: $($_.Exception.Message)"
            }
        }
    } catch {
        Add-RunWarning -Path $TempRoot -Detail "Geçici klasör temizlenemedi: $($_.Exception.Message)"
    }
}

function Remove-TempRootIfEmpty {
    param([Parameter(Mandatory)][string]$TempRoot)
    try {
        $long = Convert-ToLongPath $TempRoot
        if ([System.IO.Directory]::Exists($long) -and
            ([System.IO.Directory]::GetFileSystemEntries($long).Length -eq 0)) {
            [System.IO.Directory]::Delete($long, $false)
        }
    } catch { }
}

function New-BackupMutexName {
    param([Parameter(Mandatory)][string]$BackupRoot)
    $normalized = (Get-NormalizedFullPath $BackupRoot).ToUpperInvariant()
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
        $digest = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
        return "Global\YerelYedek_$digest"
    }
    finally { $sha.Dispose() }
}

function Set-DirectoryMetadataFinal {
    param(
        [Parameter(Mandatory)]$SourceInventory,
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$DestinationRoot
    )
    $pairs = New-Object System.Collections.Generic.List[object]

    foreach ($rel in @($SourceInventory.Directories.Keys | Sort-Object { Get-Depth $_ } -Descending)) {
        [void]$pairs.Add([PSCustomObject]@{
            Rel = $rel
            Src = (Join-RelativePath $Source $rel)
            Dst = (Join-RelativePath $DestinationRoot $rel)
        })
    }
    [void]$pairs.Add([PSCustomObject]@{ Rel='(kök)'; Src=$Source; Dst=$DestinationRoot })

    foreach ($p in $pairs) {
        try {
            if (-not (Test-DirectoryExists $p.Dst)) {
                Add-RunWarning -Path $p.Rel -Detail 'Hedef klasör bulunamadı; klasör metadata atlandı'
                continue
            }
            $sm = Get-FileSystemMetadataSafe $p.Src
            if ($null -eq $sm) { Add-RunWarning -Path $p.Rel -Detail 'Kaynak klasör metadata okunamadı'; continue }

            $dm = Get-FileSystemMetadataSafe $p.Dst
            if (Test-FileSystemMetadataSame -SourceMetadata $sm -DestinationMetadata $dm) { continue }

            if (-not (Set-FileSystemMetadata -Path $p.Dst -Metadata $sm)) {
                Add-RunWarning -Path $p.Rel -Detail "Klasör zamanı uygulanamadı: $($script:LastMetadataErrorMessage)"
                continue
            }
            if ($script:LastMetadataWarning) {
                Add-RunWarning -Path $p.Rel -Detail $script:LastMetadataWarning
            }

            $dm2 = Get-FileSystemMetadataSafe $p.Dst
            if (-not (Test-WriteTimeSame $sm $dm2)) {
                Add-RunWarning -Path $p.Rel -Detail 'Klasör LastWriteTime doğrulanamadı'
            }
        } catch {
            Add-RunWarning -Path $p.Rel -Detail $_.Exception.Message
        }
    }
}

function Set-ManifestResult {
    param([object]$Row,[string]$Result,[string]$Error = '')
    $Row.Result = $Result
    $Row.Error  = $Error
}

function Invoke-ManifestAction {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][string]$ChangedRoot,
        [Parameter(Mandatory)][string]$DeletedRoot,
        [Parameter(Mandatory)][string]$ArchivePrefix,
        [Parameter(Mandatory)][string]$Timestamp,
        [Parameter(Mandatory)][string]$TempRoot,
        [int]$RetryAttempt = 0
    )
    $script:LastArchiveErrorType = ''
    $script:LastArchiveErrorMessage = ''

    try {
        switch ($Row.Action) {
            'NOOP'  { Set-ManifestResult $Row 'NOOP'; return $true }
            'SKIP'  { Set-ManifestResult $Row 'SKIPPED' $Row.Error; return $true }
            'ERROR' { Set-ManifestResult $Row 'ERROR' $Row.Error; return $false }

            'COPY' {
                if ($Row.SourceType -eq 'Directory') {
                    if (-not (Test-DirectoryExists $Row.DestinationPath)) {
                        if (-not (Ensure-Directory $Row.DestinationPath)) { throw 'Hedef klasör oluşturulamadı' }
                    }
                    Set-ManifestResult $Row 'SUCCESS'
                    return $true
                }
                [void](Copy-FileAtomic `
                    -SourcePath $Row.SourcePath `
                    -DestinationPath $Row.DestinationPath `
                    -TempRoot $TempRoot `
                    -ExpectedSha256 $Row.SourceSha256)
                Set-ManifestResult $Row 'SUCCESS'
                return $true
            }

            'ARCHIVE_CHANGED' {
                if (-not (Test-FileExists $Row.DestinationPath)) { throw 'Değişen dosyanın mevcut hedefi bulunamadı' }
                $archive = Get-ArchiveFilePath `
                    -ArchiveRoot $ChangedRoot `
                    -RelativeFile (Join-Path $ArchivePrefix $Row.RelativePath) `
                    -Timestamp $Timestamp
                [void](Copy-FileAtomic `
                    -SourcePath $Row.SourcePath `
                    -DestinationPath $Row.DestinationPath `
                    -TempRoot $TempRoot `
                    -ExpectedSha256 $Row.SourceSha256 `
                    -ArchivePath $archive)
                if (-not (Test-FileExists $archive)) { throw 'Mevcut hedef .Changed arşivinde doğrulanamadı' }
                Set-ManifestResult $Row 'SUCCESS'
                return $true
            }

            'ARCHIVE_DELETED' {
                $archive = if ($Row.DestinationType -eq 'Directory') {
                    Get-ArchiveDirectoryPath -ArchiveRoot $DeletedRoot `
                        -RelativeDirectory (Join-Path $ArchivePrefix $Row.RelativePath) -Timestamp $Timestamp
                } else {
                    Get-ArchiveFilePath -ArchiveRoot $DeletedRoot `
                        -RelativeFile (Join-Path $ArchivePrefix $Row.RelativePath) -Timestamp $Timestamp
                }
                if (-not (Move-PathToArchive -Source $Row.DestinationPath -Destination $archive)) {
                    $why = if ($script:LastArchiveErrorMessage) { $script:LastArchiveErrorMessage } else { 'bilinmeyen hata' }
                    throw "Hedef .Deleted arşivine taşınamadı: $why"
                }
                if (-not (Test-FileExists $archive) -and -not (Test-DirectoryExists $archive)) {
                    throw 'Hedef .Deleted arşivinde doğrulanamadı'
                }
                Set-ManifestResult $Row 'SUCCESS'
                return $true
            }

            'UPDATE_METADATA' {
                if ($Row.SourceType -eq 'File') {
                    if (-not (Test-FileExists $Row.DestinationPath)) { throw 'Metadata hedef dosyası bulunamadı' }
                    $meta = Get-FileSystemMetadataSafe $Row.SourcePath
                    if ($null -eq $meta) { throw 'Kaynak metadata okunamadı' }

                    if (-not (Set-FileSystemMetadata -Path $Row.DestinationPath -Metadata $meta)) {
                        throw $(if ($script:LastMetadataErrorMessage) { $script:LastMetadataErrorMessage } else { 'Metadata uygulanamadı' })
                    }
                    if ($script:LastMetadataWarning) {
                        Add-RunWarning -Path $Row.RelativePath -Detail $script:LastMetadataWarning
                    }
                    Confirm-FileMetadataApplied -Path $Row.DestinationPath -SourceMetadata $meta -Label 'metadata güncelleme'
                } else {
                    if (-not (Test-DirectoryExists $Row.DestinationPath)) { throw 'Metadata hedef klasörü bulunamadı' }
                }
                Set-ManifestResult $Row 'SUCCESS'
                return $true
            }

            'REPLACE' {
                $archive = if ($Row.DestinationType -eq 'Directory') {
                    Get-ArchiveDirectoryPath -ArchiveRoot $DeletedRoot `
                        -RelativeDirectory (Join-Path $ArchivePrefix $Row.RelativePath) -Timestamp $Timestamp
                } else {
                    Get-ArchiveFilePath -ArchiveRoot $DeletedRoot `
                        -RelativeFile (Join-Path $ArchivePrefix $Row.RelativePath) -Timestamp $Timestamp
                }

                if ($Row.SourceType -eq 'Directory') {
                    if (-not (Move-PathToArchive -Source $Row.DestinationPath -Destination $archive)) {
                        $why = if ($script:LastArchiveErrorMessage) { $script:LastArchiveErrorMessage } else { 'bilinmeyen hata' }
                        throw "Tip çatışması hedefi .Deleted arşivine taşınamadı: $why"
                    }
                    if (-not (Test-FileExists $archive) -and -not (Test-DirectoryExists $archive)) {
                        throw 'Tip çatışması arşivi doğrulanamadı'
                    }
                    if (-not (Ensure-Directory $Row.DestinationPath)) { throw 'Yeni hedef klasör oluşturulamadı' }
                } else {
                    [void](Copy-FileAtomic `
                        -SourcePath $Row.SourcePath `
                        -DestinationPath $Row.DestinationPath `
                        -TempRoot $TempRoot `
                        -ExpectedSha256 $Row.SourceSha256 `
                        -ArchivePath $archive)
                    if (-not (Test-FileExists $archive) -and -not (Test-DirectoryExists $archive)) {
                        throw 'Tip çatışması arşivi doğrulanamadı'
                    }
                }
                Set-ManifestResult $Row 'SUCCESS'
                return $true
            }

            default { throw "Bilinmeyen manifest Action: $($Row.Action)" }
        }
    }
    catch {
        $detail = $_.Exception.Message
        $retryableFileAction = $Row.SourceType -eq 'File' -and $Row.Action -in 'COPY','ARCHIVE_CHANGED','REPLACE'
        $transientFileError = $_.Exception -is [System.IO.IOException] -or
            $_.Exception -is [System.UnauthorizedAccessException] -or
            $detail -match 'SHA-256 eşleşmiyor|kaynak kopya sırasında değişmiş olabilir|başka bir işlem tarafından kullanılıyor|sharing violation'

        if ($RetryAttempt -eq 0 -and $retryableFileAction -and $transientFileError) {
            Start-Sleep -Milliseconds 500
            return Invoke-ManifestAction `
                -Row $Row `
                -ChangedRoot $ChangedRoot `
                -DeletedRoot $DeletedRoot `
                -ArchivePrefix $ArchivePrefix `
                -Timestamp $Timestamp `
                -TempRoot $TempRoot `
                -RetryAttempt 1
        }
        if ($RetryAttempt -gt 0 -and $retryableFileAction -and $transientFileError) {
            $detail = "Dosya kilitli veya değişiyor; bir yeniden deneme sonrası da başarısız: $detail"
        }

        Set-ManifestResult $Row 'ERROR' $detail
        $eventType = if ($script:LastArchiveErrorType) { $script:LastArchiveErrorType } else { Get-ErrorEventType $_ }
        Add-ExecutionError -Path $Row.RelativePath -Detail $detail -EventType $eventType
        return $false
    }
}

# ============================================================
# YEDEKLEME — 5 PARÇA + ORKESTRATÖR
# ============================================================

function New-BackupContext {
    <#
    .SYNOPSIS
        Yedekleme çalışması için gerekli tüm türetilmiş yolları ve durumu hazırlar.
    .RETURNS
        [PSCustomObject] $ctx
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$BackupRoot,
        [bool]$TestMode,
        [bool]$FullHash
    )
    $destination = Get-DestinationRoot -Source $Source -BackupRoot $BackupRoot
    $changed     = Get-ChangedRoot $BackupRoot
    $deleted     = Get-DeletedRoot $BackupRoot
    $temp        = Get-TempRoot $BackupRoot
    $archivePrefix = Get-RelativeFromDrive $Source
    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reparsePaths = @($changed, $deleted, $temp)

    return [PSCustomObject]@{
        Source        = $Source
        BackupRoot    = $BackupRoot
        Destination   = $destination
        ChangedRoot   = $changed
        DeletedRoot   = $deleted
        TempRoot      = $temp
        ArchivePrefix = $archivePrefix
        Timestamp     = $timestamp
        TestMode      = $TestMode
        FullHash      = $FullHash
        ReparsePaths  = $reparsePaths
        Mutex         = $null
        MutexHeld     = $false
    }
}

function Get-BackupInventories {
    <#
    .SYNOPSIS
        Kaynak ve hedef envanterlerini çıkarır, hata varsa ekrana basıp $null döner.
    .RETURNS
        [PSCustomObject]@{ Source; Dest } ya da $null.
    #>
    param([Parameter(Mandatory)]$Context)

    Write-Host '[1/3] Kaynak ağacı okunuyor...'
    $srcInv = Build-Inventory -Root $Context.Source
    $script:RunSourceInventory = $srcInv
    if ($srcInv.HasErrors) {
        Show-InventoryErrors -Inventory $srcInv -Title 'Kaynak Tarama Hataları'
        return $null
    }

    Write-Host '[2/3] Hedef ağacı okunuyor...'
    $dstInv = if (Test-DirectoryExists $Context.Destination) {
        Build-Inventory -Root $Context.Destination
    } else {
        [PSCustomObject]@{Directories=@{};Files=@{};RootMetadata=$null;BlockedReparsePoints=@();HasErrors=$false}
    }
    $script:RunDestinationInventory = $dstInv
    if ($dstInv.HasErrors) {
        Show-InventoryErrors -Inventory $dstInv -Title 'Hedef Tarama Hataları'
        return $null
    }

    return [PSCustomObject]@{ Source=$srcInv; Dest=$dstInv }
}

function New-BackupManifest {
    <#
    .SYNOPSIS
        Kaynak ve hedef envanterinden karar (manifest) üretir.
    .RETURNS
        List[object] — her satır bir manifest kaydı.
    #>
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)]$SourceInventory,
        [Parameter(Mandatory)]$DestInventory
    )

    $Source = $Context.Source
    $DestinationRoot = $Context.Destination
    $srcInv = $SourceInventory
    $dstInv = $DestInventory

    $script:RunSuppressEventOutput = $true
    $manifest = New-Object System.Collections.Generic.List[object]
    $byPath   = @{}
    $hashes   = @{}
    $hashErrors = @{}
    $reparseSkips = @{}
    $conflictFile = @{}
    $conflictDir  = @{}

    foreach ($rel in @($srcInv.Files.Keys | Sort-Object)) {
        if ($dstInv.Directories.ContainsKey($rel)) { $conflictDir[$rel] = $true }
    }
    foreach ($rel in @($srcInv.Directories.Keys | Sort-Object)) {
        if ($dstInv.Files.ContainsKey($rel)) { $conflictFile[$rel] = $true }
    }

    $deletedDirs = New-Object System.Collections.Generic.List[string]
    foreach ($rel in @($dstInv.Directories.Keys | Sort-Object { Get-Depth $_ })) {
        if ($srcInv.Directories.ContainsKey($rel) -or $conflictDir.ContainsKey($rel)) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $conflictDir.Keys) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $srcInv.BlockedReparsePoints) { continue }
        $covered = $false
        foreach ($p in $deletedDirs) { if (Test-IsUnderRelativePath -Child $rel -Parent $p) { $covered = $true; break } }
        if (-not $covered) { [void]$deletedDirs.Add($rel) }
    }
    $deletedFiles = New-Object System.Collections.Generic.List[string]
    foreach ($rel in @($dstInv.Files.Keys | Sort-Object)) {
        if ($srcInv.Files.ContainsKey($rel) -or $conflictFile.ContainsKey($rel)) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $conflictDir.Keys) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $srcInv.BlockedReparsePoints) { continue }
        $covered = $false
        foreach ($p in $deletedDirs) { if (Test-IsUnderRelativePath -Child $rel -Parent $p) { $covered = $true; break } }
        if (-not $covered) { [void]$deletedFiles.Add($rel) }
    }

    $fullHash = $Context.FullHash
    $hashCandidateCount = 0
    $hashProgressIndex = 0
    $hashProgressTotal = $srcInv.Files.Count
    if ($fullHash) { Write-Host '[3/3] Tüm ortak dosyalar SHA-256 ile karşılaştırılıyor; yeni dosyalar için kaynak hash değeri alınıyor...' }
    else { Write-Host '[3/3] Yeni dosyalar ve farklı görünen ortak adaylar SHA-256 ile karşılaştırılıyor...' }
    Start-ProgressLine
    Update-ProgressLine -Index 0 -Total $hashProgressTotal -Verb 'SHA-256 karşılaştırılıyor' -Path '' -Force

    foreach ($rel in @($srcInv.Files.Keys | Sort-Object)) {
        $hashProgressIndex++
        $srcMeta = $srcInv.Files[$rel].Metadata
        $srcPathForHash = Join-RelativePath $Source $rel
        $srcIsReparse = (($srcMeta.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)

        if ($srcIsReparse) {
            $srcTag = Get-ReparsePointTag $srcPathForHash
            if (($null -eq $srcTag) -or (-not (Test-IsCloudFilesReparseTag ([uint32]$srcTag)))) {
                $hashErrors[$rel] = 'Kaynak dosya Reparse Point; bağlantı hedefini izlememek için güvenli biçimde atlandı.'
                $reparseSkips[$rel] = $true
                Add-RunEvent -Type Atlanan -Path $rel -Detail $hashErrors[$rel]
                continue
            }
        }

        $hasTargetFile = $dstInv.Files.ContainsKey($rel)
        $mustHash = $fullHash -or (-not $hasTargetFile)
        if ($hasTargetFile) {
            $dstMeta = $dstInv.Files[$rel].Metadata
            if (($srcMeta.Length -ne $dstMeta.Length) -or (-not (Test-WriteTimeSame $srcMeta $dstMeta))) {
                $mustHash = $true
            }
        }
        if (-not $mustHash) {
            Update-ProgressLine -Index $hashProgressIndex -Total $hashProgressTotal -Verb 'SHA-256 karşılaştırılıyor' -Path $rel
            continue
        }

        $hashCandidateCount++
        Update-ProgressLine -Index $hashProgressIndex -Total $hashProgressTotal -Verb 'SHA-256 okunuyor' -Path $rel -Force
        $h1 = Get-FileHashSafe (Join-RelativePath $Source $rel)
        $h2 = ''
        if ($hasTargetFile) { $h2 = Get-FileHashSafe (Join-RelativePath $DestinationRoot $rel) }

        if ([string]::IsNullOrWhiteSpace($h1) -or ($hasTargetFile -and [string]::IsNullOrWhiteSpace($h2))) {
            $msg = if ($script:LastHashErrorMessage) { $script:LastHashErrorMessage } else { 'SHA-256 alınamadı' }
            $hashErrors[$rel] = $msg
            Add-SkipReadError -Path $rel -Detail $msg
        } else {
            $hashes[$rel] = [PSCustomObject]@{Source=$h1;Destination=$h2}
        }
        Update-ProgressLine -Index $hashProgressIndex -Total $hashProgressTotal -Verb 'SHA-256 karşılaştırılıyor' -Path $rel
    }
    Update-ProgressLine -Index $hashProgressTotal -Total $hashProgressTotal -Verb 'SHA-256 tamamlandı' -Path '' -Force
    Stop-ProgressLine
    Write-RunLog 'BİLGİ' ("SHA-256 taranan kaynak dosya: {0:N0} / {1:N0} | Tam tarama: {2}" -f $hashCandidateCount, $srcInv.Files.Count, $fullHash)

    foreach ($rel in @($conflictDir.Keys | Sort-Object)) {
        $src = Join-RelativePath $Source $rel
        $dst = Join-RelativePath $DestinationRoot $rel
        $sm = $srcInv.Files[$rel].Metadata
        $srcIsReparse = (($sm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $srcTag = if ($srcIsReparse) { [string](Get-ReparsePointTag $src) } else { '' }
        $dm = $dstInv.Directories[$rel].Metadata
        $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $dstTag = if ($dstIsReparse) { [string](Get-ReparsePointTag $dst) } else { '' }

        if ($hashErrors.ContainsKey($rel)) {
            $statusText = if ($reparseSkips.ContainsKey($rel)) { 'Reparse Point atlandı' } else { 'SHA-256 okunamadı' }
            $r = New-MergeManifestRow $rel $src $dst 'File' 'Directory' $true $true $false $false $false $false $false '' '' $false $srcIsReparse $srcTag $dstIsReparse $dstTag $statusText 'SKIP' 'SKIPPED' $hashErrors[$rel]
        } else {
            $r = New-MergeManifestRow $rel $src $dst 'File' 'Directory' $true $true $false $false $false $false $false $hashes[$rel].Source '' $false $srcIsReparse $srcTag $dstIsReparse $dstTag 'Tip çatışması' 'REPLACE'
        }
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    foreach ($rel in @($conflictFile.Keys | Sort-Object)) {
        $src = Join-RelativePath $Source $rel
        $dst = Join-RelativePath $DestinationRoot $rel
        $dm = $dstInv.Files[$rel].Metadata
        $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $dstTag = if ($dstIsReparse) { [string](Get-ReparsePointTag $dst) } else { '' }
        $r = New-MergeManifestRow $rel $src $dst 'Directory' 'File' $true $true $false $false $false $false $false '' '' $false $false '' $dstIsReparse $dstTag 'Tip çatışması' 'REPLACE'
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    foreach ($rel in @($deletedDirs | Sort-Object)) {
        $dst = Join-RelativePath $DestinationRoot $rel
        $dm = $dstInv.Directories[$rel].Metadata
        $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $dstTag = if ($dstIsReparse) { [string](Get-ReparsePointTag (Join-RelativePath $DestinationRoot $rel)) } else { '' }
        $r = New-MergeManifestRow $rel '' $dst '' 'Directory' $false $true $false $false $false $false $false '' '' $false $false '' $dstIsReparse $dstTag 'Kaynakta yok' 'ARCHIVE_DELETED'
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    foreach ($rel in @($deletedFiles | Sort-Object)) {
        $dst = Join-RelativePath $DestinationRoot $rel
        $dm = $dstInv.Files[$rel].Metadata
        $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $dstTag = if ($dstIsReparse) { [string](Get-ReparsePointTag $dst) } else { '' }
        $r = New-MergeManifestRow $rel '' $dst '' 'File' $false $true $false $false $false $false $false '' '' $false $false '' $dstIsReparse $dstTag 'Kaynakta yok' 'ARCHIVE_DELETED'
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    foreach ($rel in @($srcInv.Directories.Keys | Sort-Object)) {
        if ($byPath.ContainsKey($rel)) { continue }
        $src = Join-RelativePath $Source $rel
        $dst = Join-RelativePath $DestinationRoot $rel

        if (-not $dstInv.Directories.ContainsKey($rel)) {
            $r = New-MergeManifestRow $rel $src $dst 'Directory' '' $true $false $false $false $false $false $false '' '' $false $false '' $false '' 'Yeni' 'COPY'
        } else {
            $sm = $srcInv.Directories[$rel].Metadata
            $dm = $dstInv.Directories[$rel].Metadata
            $same = Test-FileSystemMetadataSame -SourceMetadata $sm -DestinationMetadata $dm -CriticalOnly
            $f = Get-MetaFlags $sm $dm
            $srcIsReparse = (($sm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
            $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
            $srcTag = if ($srcIsReparse) { [string](Get-ReparsePointTag $src) } else { '' }
            $dstTag = if ($dstIsReparse) { [string](Get-ReparsePointTag $dst) } else { '' }
            $r = New-MergeManifestRow $rel $src $dst 'Directory' 'Directory' $true $true $true $f.Length $f.Creation $f.Write $f.Attributes '' '' $false $srcIsReparse $srcTag $dstIsReparse $dstTag $(if ($same) { 'Aynı' } else { 'Metadata farklı' }) $(if ($same) { 'NOOP' } else { 'UPDATE_METADATA' })
        }
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    foreach ($rel in @($srcInv.Files.Keys | Sort-Object)) {
        if ($byPath.ContainsKey($rel)) { continue }
        $src = Join-RelativePath $Source $rel
        $dst = Join-RelativePath $DestinationRoot $rel
        $sm = $srcInv.Files[$rel].Metadata
        $srcIsReparse = (($sm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        $srcTag = if ($srcIsReparse) { [string](Get-ReparsePointTag $src) } else { '' }

        $destinationIsFile = $dstInv.Files.ContainsKey($rel)
        $destinationIsDirectory = $dstInv.Directories.ContainsKey($rel)
        $dstIsReparse = $false; $dstTag = ''
        if ($destinationIsFile) {
            $dm = $dstInv.Files[$rel].Metadata
            $dstIsReparse = (($dm.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
            if ($dstIsReparse) { $dstTag = [string](Get-ReparsePointTag $dst) }
        }

        if ($hashErrors.ContainsKey($rel)) {
            $destType = if ($destinationIsDirectory) { 'Directory' } elseif ($destinationIsFile) { 'File' } else { '' }
            $destExists = $destinationIsDirectory -or $destinationIsFile
            $statusText = if ($reparseSkips.ContainsKey($rel)) { 'Reparse Point atlandı' } else { 'SHA-256 okunamadı' }
            $r = New-MergeManifestRow $rel $src $dst 'File' $destType $true $destExists $false $false $false $false $false '' '' $false $srcIsReparse $srcTag $dstIsReparse $dstTag $statusText 'SKIP' 'SKIPPED' $hashErrors[$rel]
        }
        elseif (-not $destinationIsFile) {
            $h = $hashes[$rel]
            $r = New-MergeManifestRow $rel $src $dst 'File' '' $true $false $false $false $false $false $false $h.Source '' $false $srcIsReparse $srcTag $false '' 'Yeni' 'COPY'
        }
        else {
            $dm = $dstInv.Files[$rel].Metadata
            $sameMeta = Test-FileSystemMetadataSame -SourceMetadata $sm -DestinationMetadata $dm
            $h = $hashes[$rel]
            if ($null -eq $h) { $sameHash = $true; $srcHash = ''; $dstHash = '' }
            else { $sameHash = ($h.Source -eq $h.Destination); $srcHash = $h.Source; $dstHash = $h.Destination }

            $action = if (-not $sameHash) { 'ARCHIVE_CHANGED' }
                      elseif (-not $sameMeta) { 'UPDATE_METADATA' }
                      else { 'NOOP' }
            $status = if ($sameHash -and $sameMeta) { 'Aynı' }
                      elseif (-not $sameHash) { 'İçerik farklı' }
                      else { 'Metadata farklı' }
            $f = Get-MetaFlags $sm $dm
            $r = New-MergeManifestRow $rel $src $dst 'File' 'File' $true $true $true $f.Length $f.Creation $f.Write $f.Attributes $srcHash $dstHash $sameHash $srcIsReparse $srcTag $dstIsReparse $dstTag $status $action
        }
        Add-DecisionRow -Manifest $manifest -ByPath $byPath -Row $r
    }

    # Manifest + yardımcı sözlükleri birlikte döneriz; execution katmanı bunları kullanır.
    return [PSCustomObject]@{
        Manifest = $manifest
        ByPath   = $byPath
        DeletedDirs  = $deletedDirs
        DeletedFiles = $deletedFiles
    }
}

function Invoke-BackupExecution {
    <#
    .SYNOPSIS
        Manifest satırlarını sırayla uygular (onay zaten alınmış olmalıdır).
    .RETURNS
        [PSCustomObject]@{ Failed; ProgressDone; ProgressTotal }
    #>
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)]$ManifestData,
        [Parameter(Mandatory)]$SourceInventory,
        [Parameter(Mandatory)]$DestInventory
    )

    $manifest = $ManifestData.Manifest
    $SourceInventory = $SourceInventory

    Clear-StaleTemp -TempRoot $Context.TempRoot

    $requiredFreeBytes = [double]0
    foreach ($row in $manifest) {
        if ($row.Action -notin 'COPY','ARCHIVE_CHANGED','REPLACE') { continue }
        if ($row.SourceType -ne 'File' -or -not $SourceInventory.Files.ContainsKey($row.RelativePath)) { continue }
        $requiredFreeBytes += [double]$SourceInventory.Files[$row.RelativePath].Metadata.Length
    }
    if ($requiredFreeBytes -gt 0) {
        try {
            $backupDrive = Get-DriveRootSafe $Context.BackupRoot
            $driveInfo = [System.IO.DriveInfo]::new($backupDrive)
            if (-not $driveInfo.IsReady) { throw "Yedek sürücüsü hazır değil: $backupDrive" }
            $availableFreeBytes = [double]$driveInfo.AvailableFreeSpace
            if ($availableFreeBytes -lt $requiredFreeBytes) {
                $needText = Format-ByteSize $requiredFreeBytes
                $freeText = Format-ByteSize $availableFreeBytes
                throw "Plan için yetersiz boş alan. Gereken yaklaşık: $needText; kullanılabilir: $freeText."
            }
        }
        catch {
            throw "Yedek sürücüsünün boş alanı doğrulanamadı: $($_.Exception.Message)"
        }
    }

    $failed = 0
    $script:RunSuppressEventOutput = $true

    Write-Host ''
    Write-Host 'YEDEKLEME UYGULANIYOR' -ForegroundColor Cyan
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray

    $ordered = @($manifest | Sort-Object `
        @{Expression={switch ($_.Action) {'REPLACE'{10};'ARCHIVE_DELETED'{20};'ARCHIVE_CHANGED'{30};'COPY'{40};'UPDATE_METADATA'{50};'NOOP'{60};'SKIP'{70};'ERROR'{80};default{90}}}},
        @{Expression={$_.RelativePath}})

    $progressTotal = @($ordered | Where-Object { $_.Action -notin 'NOOP','SKIP','ERROR' }).Count
    $progressDone = 0
    Start-ProgressLine

    foreach ($row in $ordered) {
        $isWork = ($row.Action -notin 'NOOP','SKIP','ERROR')

        if ($isWork) {
            $verb = switch ($row.Action) {
                'COPY'            { 'Kopyalanıyor' }
                'ARCHIVE_CHANGED' { 'Değişen arşivleniyor' }
                'ARCHIVE_DELETED' { 'Silinen arşivleniyor' }
                'UPDATE_METADATA' { 'Güncelleniyor' }
                'REPLACE'         { 'Tip çatışması' }
                default           { [string]$row.Action }
            }
            Update-ProgressLine -Index $progressDone -Total $progressTotal -Verb $verb -Path $row.RelativePath
        }

        $ok = Invoke-ManifestAction -Row $row `
            -ChangedRoot $Context.ChangedRoot -DeletedRoot $Context.DeletedRoot `
            -ArchivePrefix $Context.ArchivePrefix -Timestamp $Context.Timestamp -TempRoot $Context.TempRoot
        if (-not $ok) { $failed++ }
        if ($isWork) { $progressDone++ }

        if ($row.Action -ne 'NOOP') { Write-RunLogRow $row }

        if ($row.Result -eq 'ERROR') {
            Write-PermanentLine ("  [HATA] {0}" -f $row.RelativePath) 'Red'
            if (-not [string]::IsNullOrWhiteSpace([string]$row.Error)) {
                Write-PermanentLine ("         {0}" -f $row.Error) 'Red'
            }
        }
    }

    Update-ProgressLine -Index $progressTotal -Total $progressTotal -Verb 'Tamamlandı' -Path '' -Force
    Stop-ProgressLine
    Write-Host ("Uygulanan işlem: {0:N0} / {1:N0}" -f $progressDone,$progressTotal) -ForegroundColor DarkGray

    Write-Host ''
    Write-Host 'Son klasör düzenlemeleri yapılıyor...' -ForegroundColor DarkCyan
    Set-DirectoryMetadataFinal -SourceInventory $SourceInventory -Source $Context.Source -DestinationRoot $Context.Destination
    Remove-TempRootIfEmpty -TempRoot $Context.TempRoot

    return [PSCustomObject]@{
        Failed        = $failed
        ProgressDone  = $progressDone
        ProgressTotal = $progressTotal
    }
}

function Complete-BackupRun {
    <#
    .SYNOPSIS
        Yedekleme sonrası özet/sonuç/log kapanışını yapar.
    .RETURNS
        Sonuç metni (string).
    #>
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)]$ManifestData
    )

    $manifest = $ManifestData.Manifest
    $script:RunCounters['KontrolEdilen'] = $manifest.Count

    # Neden: Hedef envanteri işlem ÖNCESİ alınır; rapor gerçek sonucu göstermeli.
    # Hedefte bir şey değiştiyse yeniden sayılır; sayım başarısızsa eski sayı güncelmiş
    # gibi gösterilmez ('-' basılır).
    $appliedChanges = @($manifest | Where-Object {
        $_.Action -notin 'NOOP', 'SKIP', 'ERROR' -and $_.Result -eq 'SUCCESS'
    }).Count
    if ($appliedChanges -gt 0) {
        Write-Host 'Hedef ağacı işlem sonrası yeniden sayılıyor...'
        try {
            $postInventory = Build-Inventory -Root $Context.Destination
            if ($postInventory.HasErrors) {
                $script:RunDestinationInventory = $null
                Write-RunLog 'UYARI' 'İşlem sonrası hedef sayımı tarama hatası verdi; rapordaki hedef sayısı gösterilmedi.'
            }
            else {
                $script:RunDestinationInventory = $postInventory
            }
        }
        catch {
            $script:RunDestinationInventory = $null
            Write-RunLog 'UYARI' ('İşlem sonrası hedef sayımı başarısız: {0}' -f $_.Exception.Message)
        }
    }

    Write-RunSummary -Manifest $manifest

    Write-Host ''
    $skipCount = @($manifest | Where-Object { $_.Action -eq 'SKIP' }).Count
    if (($script:RunCounters.ToplamHata -eq 0) -and ($skipCount -eq 0)) {
        $finalText = 'SONUÇ: BAŞARILI'; $finalColor = 'Green'
    }
    elseif ($script:RunCounters.ToplamHata -eq 0) {
        $finalText = ("SONUÇ: TAMAMLANDI, ANCAK {0} DOSYA ATLANDI (yedeklenmedi)" -f $skipCount)
        $finalColor = 'Yellow'
    }
    else {
        $finalText = 'SONUÇ: HATA VAR'; $finalColor = 'Red'
    }

    # Ekranda ikonlu, log/raporda sade metin: ikon yalnızca görsel vurgudur.
    $finalIcon = switch ($finalColor) { 'Green' { '✅' } 'Yellow' { '🟡' } default { '❌' } }
    Write-Host ('{0} {1}' -f $finalIcon, $finalText) -ForegroundColor $finalColor
    Write-RunLogSummary -Manifest $manifest
    Save-RunReport -Result $finalText
    return $finalText
}

function Start-Backup {
    <#
    .SYNOPSIS
        Yedekleme orkestratörü. Kontrol veya gerçek yedekleme modunu yürütür.
    .PARAMETER TestMode
        $true → Kontrol Raporu (dosya sistemi değişmez).
    .PARAMETER FullHash
        $true → tüm ortak dosyalar SHA-256 ile taranır.
    .RETURNS
        'MENU' | 'REPORT' | 'CANCEL'
    #>
    param([bool]$TestMode, [bool]$FullHash = $false)

    $settings = Load-Settings
    if ($null -eq $settings) {
        Write-Host '[HATA] Geçerli ayar bulunamadı.' -ForegroundColor Red
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }
    $Source = $settings.Source
    $BackupRoot = $settings.BackupRoot
    if (-not (Test-SettingsValid -Source $Source -BackupRoot $BackupRoot -ShowError)) {
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    $ctx = New-BackupContext -Source $Source -BackupRoot $BackupRoot -TestMode $TestMode -FullHash $FullHash

    # Mutex
    try {
        $backupMutexName = New-BackupMutexName -BackupRoot $ctx.BackupRoot
        $ctx.Mutex = [System.Threading.Mutex]::new($false, $backupMutexName)
        try { $ctx.MutexHeld = $ctx.Mutex.WaitOne(0) }
        catch [System.Threading.AbandonedMutexException] {
            $ctx.MutexHeld = $true
            Write-Host '[UYARI] Önceki yedekleme beklenmedik biçimde sonlandı; kilit devralındı.' -ForegroundColor Yellow
        }
    }
    catch {
        if ($ctx.Mutex) { $ctx.Mutex.Dispose() }
        Write-Host "[HATA] Yedekleme kilidi alınamadı: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }
    if (-not $ctx.MutexHeld) {
        $ctx.Mutex.Dispose()
        Write-Host '[UYARI] Bu yedek kökü başka bir pencerede kullanılıyor. İkinci çalışma başlatılmadı.' -ForegroundColor Yellow
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    try {
        Initialize-RunReport -Mode $(if ($TestMode) { 'Sadece Kontrol' } else { 'Gerçek Yedekleme' }) `
            -Source $ctx.Source -Destination $ctx.Destination -BackupRoot $ctx.BackupRoot

        # Başlangıç reparse denetimi
        try {
            Assert-NoBlockedReparse -BackupRoot $ctx.BackupRoot -DestinationRoot $ctx.Destination `
                -AdditionalPaths $ctx.ReparsePaths -Reason 'Başlangıç'
        }
        catch {
            Write-Host '[HATA] Yedek yollarında engellenen Reparse Point bulundu.' -ForegroundColor Red
            Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
            Save-RunReport -Result 'Reparse denetimi başarısız'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        Initialize-TimeToleranceContext -BackupRoot $ctx.BackupRoot
        Initialize-RunLog -Slug $(if ($TestMode) { 'kontrol' } else { 'yedek' }) -BackupRoot $ctx.BackupRoot `
            -HeaderLines @("Mod: $script:RunMode", "Kaynak: $($ctx.Source)", "Yedek: $($ctx.BackupRoot)",
                           "Hedef: $($ctx.Destination)", (Get-TimeToleranceDescription))

        Clear-Host
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host $(if ($TestMode) { '                   Yedekleme Kontrolü' } else { '                   Gerçek Yedekleme' }) -ForegroundColor Cyan
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-InfoLine -Label 'Kaynak' -Value $ctx.Source -Icon '📁' -Color Yellow
        Write-InfoLine -Label 'Yedek kökü' -Value $ctx.BackupRoot -Icon '💾' -Color Yellow
        Write-InfoLine -Label 'Hedef' -Value $ctx.Destination -Icon '🎯' -Color Green
        Write-InfoLine -Label '.Changed' -Value $ctx.ChangedRoot -Icon '📦' -Color DarkCyan
        Write-InfoLine -Label '.Deleted' -Value $ctx.DeletedRoot -Icon '🚮' -Color DarkCyan
        Write-InfoLine -Label '.Temp' -Value $ctx.TempRoot -Icon '⏳' -Color DarkCyan
        if ($TestMode) {
            Write-RunEnvironmentInfo -FullHash $ctx.FullHash
        } else {
            Write-RunEnvironmentInfo -FullHash $ctx.FullHash -ShowCopyBehavior
        }
        Write-Host ''

        if (-not (Test-DirectoryExists $ctx.Destination)) {
            if ($TestMode) { Write-Host '[KLASÖR] Hedef oluşturulacak.' -ForegroundColor DarkCyan }
            elseif (-not (Ensure-Directory $ctx.Destination)) {
                Save-RunReport -Result 'Hedef oluşturulamadı'
                Read-Host 'Ana Menüye dönmek için Enter'
                return 'MENU'
            }
        }

        try {
            Assert-NoBlockedReparse -BackupRoot $ctx.BackupRoot -DestinationRoot $ctx.Destination `
                -AdditionalPaths $ctx.ReparsePaths -Reason 'Hedef oluşturma sonrası'
        }
        catch {
            Write-Host "[HATA] Hedef oluşturma sonrası reparse denetimi başarısız: $($_.Exception.Message)" -ForegroundColor Red
            Save-RunReport -Result 'Hedef reparse denetimi başarısız'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        $inv = Get-BackupInventories -Context $ctx
        if ($null -eq $inv) {
            Save-RunReport -Result 'Envanter hatası'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        $reparseConflict = Find-ReparseConflict -SourceInventory $inv.Source -TargetInventory $inv.Dest
        if ($null -ne $reparseConflict) {
            Write-Host '[HATA] Hedefte kaynakla çakışan engellenen Reparse Point bulundu.' -ForegroundColor Red
            Write-Host "       $($reparseConflict.Path)" -ForegroundColor Red
            Write-Host "       Tür: $($reparseConflict.Type)" -ForegroundColor Red
            Save-RunReport -Result 'Hedefte reparse çakışması'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        try {
            Assert-NoBlockedReparse -BackupRoot $ctx.BackupRoot -DestinationRoot $ctx.Destination `
                -AdditionalPaths $ctx.ReparsePaths -Reason 'Envanter sonrası'
        }
        catch {
            Write-Host "[HATA] Envanter sonrası reparse denetimi başarısız: $($_.Exception.Message)" -ForegroundColor Red
            Save-RunReport -Result 'Envanter sonrası reparse denetimi başarısız'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        # Karar katmanı
        $manifestData = New-BackupManifest -Context $ctx -SourceInventory $inv.Source -DestInventory $inv.Dest
        $manifest = $manifestData.Manifest

        $actionCounts = @{COPY=0;ARCHIVE_CHANGED=0;ARCHIVE_DELETED=0;UPDATE_METADATA=0;REPLACE=0;NOOP=0;SKIP=0;ERROR=0}
        foreach ($m in $manifest) {
            if ($actionCounts.ContainsKey($m.Action)) { $actionCounts[$m.Action]++ }
        }
        foreach ($k in @($actionCounts.Keys | Sort-Object)) {
            Write-RunLog 'BİLGİ' ('PLAN | {0,-16} | {1}' -f $k, $actionCounts[$k])
        }

        if ($TestMode) {
            foreach ($m in $manifest) {
                if ($m.Action -eq 'NOOP') { $m.Result = 'NOOP' }
                elseif ($m.Action -eq 'SKIP') { $m.Result = 'SKIPPED' }
                elseif ($m.Action -eq 'ERROR') { $m.Result = 'ERROR' }
                else { $m.Result = 'PLANNED' }
            }
            foreach ($m in $manifest) { if ($m.Action -ne 'NOOP') { Write-RunLogRow $m } }

            Write-Host ''
            Write-PlanDigest -Manifest $manifest -SourceInventory $inv.Source
            Write-RunSummary -Manifest $manifest
            Write-RunLogSummary -Manifest $manifest
            Save-RunReport -Result 'KONTROL TAMAMLANDI'
            return 'REPORT'
        }

        Write-Host ''
        Write-Host 'YAPILACAK İŞLEMLER (PLAN)' -ForegroundColor Cyan
        Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
        Write-PlanDigest -Manifest $manifest -SourceInventory $inv.Source

        $planWork = @($manifest | Where-Object { $_.Action -notin 'NOOP','SKIP','ERROR' }).Count

        if ($planWork -gt 0) {
            Write-Host ''
            $deleteCount = $actionCounts.ARCHIVE_DELETED
            $changeCount = $actionCounts.ARCHIVE_CHANGED
            $replaceCount = $actionCounts.REPLACE
            $risk = $deleteCount + $replaceCount

            $destinationEntryCount = $inv.Dest.Files.Count + $inv.Dest.Directories.Count
            $effectiveDeletedEntries = $manifestData.DeletedFiles.Count
            foreach ($deletedDir in $manifestData.DeletedDirs) {
                $effectiveDeletedEntries++
                foreach ($dirRel in $inv.Dest.Directories.Keys) {
                    if (Test-IsUnderRelativePath -Child $dirRel -Parent $deletedDir) { $effectiveDeletedEntries++ }
                }
                foreach ($fileRel in $inv.Dest.Files.Keys) {
                    if (Test-IsUnderRelativePath -Child $fileRel -Parent $deletedDir) { $effectiveDeletedEntries++ }
                }
            }
            $deleteRatio = 0.0
            if ($destinationEntryCount -gt 0) {
                $deleteRatio = [double]$effectiveDeletedEntries / $destinationEntryCount
            }
            if ($deleteCount -gt 0 -and ($destinationEntryCount -gt 0) -and
                ($deleteRatio -ge 0.5 -or ($inv.Source.Files.Count -eq 0 -and $inv.Source.Directories.Count -eq 0))) {
                $ratioText = '{0:P0}' -f $deleteRatio
                Write-Host ("UYARI: Plan hedef girdilerinin yaklaşık {0}'ini .Deleted arşivine taşıyacak. Kaynak boş veya büyük ölçüde farklı görünüyor; kaynağı doğrulayın." -f $ratioText) -ForegroundColor Red
            }

            $newCount      = $actionCounts.COPY
            $metadataCount = $actionCounts.UPDATE_METADATA
            $prompt = 'Plan: yeni {0}, değişen {1}, silinen {2}, REPLACE {3}, metadata {4}. ' -f $newCount, $changeCount, $deleteCount, $replaceCount, $metadataCount
            if ($risk -gt 0) { $prompt += 'Riskli plan. EVET yazarak onaylayın' }
            else { $prompt += 'Uygulansın mı? (E/H)' }

            # Neden: onay bekleme süresi çalışma süresinden düşülür (Get-RunElapsed).
            $confirmStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $go = Read-Host $prompt
            $confirmStopwatch.Stop()
            $script:RunPausedTime += $confirmStopwatch.Elapsed
            $approved = if ($risk -gt 0) { $go -ceq 'EVET' } else { $go -match '^[Ee]$' }

            if (-not $approved) {
                Write-Host 'İşlem iptal edildi. Hiçbir değişiklik yapılmadı.' -ForegroundColor Yellow
                Save-RunReport -Result 'KULLANICI İPTAL ETTİ'
                return 'CANCEL'
            }
        }
        else {
            Write-Host ''
            Write-Host 'Uygulanacak işlem yok; yedek güncel.' -ForegroundColor Green
        }

        try {
            Assert-NoBlockedReparse -BackupRoot $ctx.BackupRoot -DestinationRoot $ctx.Destination `
                -AdditionalPaths $ctx.ReparsePaths -Reason 'Uygulama öncesi'
        }
        catch {
            Write-Host "[HATA] Uygulama öncesi reparse denetimi başarısız; hiçbir plan işlemi uygulanmadı: $($_.Exception.Message)" -ForegroundColor Red
            Save-RunReport -Result 'Uygulama öncesi reparse denetimi başarısız'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        try {
            [void](Invoke-BackupExecution -Context $ctx -ManifestData $manifestData `
                -SourceInventory $inv.Source -DestInventory $inv.Dest)
        }
        catch {
            Write-Host "[HATA] Yedekleme uygulaması başlatılamadı: $($_.Exception.Message)" -ForegroundColor Red
            Save-RunReport -Result 'Uygulama hatası'
            Read-Host 'Ana Menüye dönmek için Enter'
            return 'MENU'
        }

        [void](Complete-BackupRun -Context $ctx -ManifestData $manifestData)
        return 'REPORT'
    }
    finally {
        Stop-ProgressLine
        Close-RunLog
        if ($ctx.MutexHeld -and $ctx.Mutex) {
            try { $ctx.Mutex.ReleaseMutex() } catch { }
        }
        if ($ctx.Mutex) { $ctx.Mutex.Dispose() }
    }
}

# ============================================================
# BÜTÜNLÜK KONTROLÜ
# ============================================================

function Start-IntegrityCheck {
    <#
    .SYNOPSIS
        Kaynak ile yedeği SHA-256, LastWriteTime ve metadata açısından karşılaştırır.
    .RETURNS
        'MENU'
    #>
    $settings = Load-Settings
    if ($null -eq $settings) {
        Write-Host '[HATA] Geçerli ayar bulunamadı.' -ForegroundColor Red
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }
    if (-not (Test-SettingsValid -Source $settings.Source -BackupRoot $settings.BackupRoot -ShowError)) {
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    $source = $settings.Source
    $backup = $settings.BackupRoot
    $destination = Get-DestinationRoot -Source $source -BackupRoot $backup

    Clear-Host
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '                    Bütünlük Kontrolü' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-InfoLine -Label 'Kaynak'     -Value $source      -Icon '📁' -Color Yellow
    Write-InfoLine -Label 'Yedek kökü'  -Value $backup      -Icon '💾' -Color Yellow
    Write-InfoLine -Label 'Hedef'       -Value $destination -Icon '🎯' -Color Green
    Write-Host ''

    Initialize-RunReport -Mode 'Bütünlük Kontrolü' -Source $source -Destination $destination -BackupRoot $backup
    Initialize-TimeToleranceContext -BackupRoot $backup
    Initialize-RunLog -Slug 'butunluk' -BackupRoot $backup `
        -HeaderLines @("Mod: Bütünlük Kontrolü","Kaynak: $source","Yedek: $backup",
                       "Hedef: $destination", (Get-TimeToleranceDescription))

    Write-RunEnvironmentInfo -FullHash $true -IntegrityCheck
    Write-Host ''

    if (-not (Test-DirectoryExists $destination)) {
        Write-Host '[HATA] Hedef bulunamadı.' -ForegroundColor Red
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    $srcInv = Build-Inventory -Root $source
    $dstInv = Build-Inventory -Root $destination

    if ($srcInv.HasErrors) {
        Show-InventoryErrors -Inventory $srcInv -Title 'Kaynak Tarama Hataları'
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }
    if ($dstInv.HasErrors) {
        Show-InventoryErrors -Inventory $dstInv -Title 'Hedef Tarama Hataları'
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    $reparseConflict = Find-ReparseConflict -SourceInventory $srcInv -TargetInventory $dstInv
    if ($null -ne $reparseConflict) {
        Write-Host ''
        Write-Host '[HATA] Hedefte engellenen Reparse Point ile kaynak çakışıyor.' -ForegroundColor Red
        Write-Host "       $($reparseConflict.Path)" -ForegroundColor Red
        Write-Host "       Tür: $($reparseConflict.Type)" -ForegroundColor Red
        Read-Host 'Ana Menüye dönmek için Enter'
        return 'MENU'
    }

    $same = 0; $different = 0; $missing = 0; $extra = 0; $errors = 0
    $metadataDifferent = 0; $softDifferent = 0; $checked = 0

    $progressTotal = $srcInv.Files.Count
    $progressIndex = 0
    Start-ProgressLine

    foreach ($rel in @($srcInv.Files.Keys | Sort-Object)) {
        $progressIndex++
        Update-ProgressLine -Index $progressIndex -Total $progressTotal -Verb 'Doğrulanıyor' -Path $rel

        $src = Join-RelativePath $source $rel
        $dst = Join-RelativePath $destination $rel

        if (-not $dstInv.Files.ContainsKey($rel)) {
            Add-RunEvent -Type Eksik -Path $rel -Detail 'Kaynakta var, yedekte yok'
            $missing++
            continue
        }

        $checked++
        if ($script:RunCounters.Contains('KontrolEdilen')) { $script:RunCounters['KontrolEdilen']++ }

        $h1 = Get-FileHashSafe $src
        $h2 = Get-FileHashSafe $dst

        if ([string]::IsNullOrWhiteSpace($h1) -or [string]::IsNullOrWhiteSpace($h2)) {
            Add-RunEvent -Type $(if ($script:LastHashErrorType) { $script:LastHashErrorType } else { 'OkumaHatasi' }) `
                -Path $rel -Detail $(if ($script:LastHashErrorMessage) { $script:LastHashErrorMessage } else { 'SHA-256 alınamadı' })
            $errors++
            continue
        }

        if ($h1 -eq $h2) {
            $same++
            if (-not (Test-FileSystemMetadataSame -SourceMetadata $srcInv.Files[$rel].Metadata -DestinationMetadata $dstInv.Files[$rel].Metadata -CriticalOnly)) {
                Add-RunEvent -Type MetadataDegisen -Path $rel -Detail 'Dosya LastWriteTime farklı'
                $metadataDifferent++
            }
            elseif (-not (Test-FileSystemMetadataSame -SourceMetadata $srcInv.Files[$rel].Metadata -DestinationMetadata $dstInv.Files[$rel].Metadata)) {
                Add-RunWarning -Path $rel -Detail 'Dosya nitelikleri farklı'
                $softDifferent++
            }
        }
        else {
            Add-RunEvent -Type Degisen -Path $rel -Detail 'SHA-256 farklı'
            $different++
        }
    }

    Update-ProgressLine -Index $progressTotal -Total $progressTotal -Verb 'Tamamlandı' -Path '' -Force
    Stop-ProgressLine

    foreach ($rel in @($dstInv.Files.Keys | Sort-Object)) {
        if ($srcInv.Files.ContainsKey($rel)) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $srcInv.BlockedReparsePoints) { continue }
        Add-RunEvent -Type Fazla -Path $rel -Detail 'Yedekte var, kaynakta yok'
        $extra++
    }

    foreach ($rel in @($srcInv.Directories.Keys | Sort-Object)) {
        if (-not $dstInv.Directories.ContainsKey($rel)) {
            Add-RunEvent -Type Eksik -Path $rel -Detail 'Kaynakta var, yedekte yok olan klasör'
            $missing++
            continue
        }
        if (-not (Test-FileSystemMetadataSame -SourceMetadata $srcInv.Directories[$rel].Metadata -DestinationMetadata $dstInv.Directories[$rel].Metadata)) {
            Add-RunWarning -Path $rel -Detail 'Klasör zamanı/metadata farklı (en iyi çaba)'
            $softDifferent++
        }
    }

    foreach ($rel in @($dstInv.Directories.Keys | Sort-Object)) {
        if ($srcInv.Directories.ContainsKey($rel)) { continue }
        if (Test-IsUnderAnyRelativePath -Child $rel -Parents $srcInv.BlockedReparsePoints) { continue }
        Add-RunEvent -Type Fazla -Path $rel -Detail 'Yedekte var, kaynakta yok olan klasör'
        $extra++
    }

    if (($null -ne $srcInv.RootMetadata) -and ($null -ne $dstInv.RootMetadata) -and
        -not (Test-FileSystemMetadataSame -SourceMetadata $srcInv.RootMetadata -DestinationMetadata $dstInv.RootMetadata)) {
        Add-RunWarning -Path $destination -Detail 'Kök klasör zamanı/metadata farklı (en iyi çaba)'
        $softDifferent++
    }

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '                 Bütünlük Kontrolü Bitti' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host "Kontrol Edilen : $checked"
    Write-Host "İçerik Aynı    : $same"
    Write-Host "İçerik Farklı  : $different"
    Write-Host "Eksik          : $missing"
    Write-Host "Fazla          : $extra"
    Write-Host "Hata           : $errors"
    Write-Host "Dosya Zamanı Farkı : $metadataDifferent"
    Write-Host "İkincil Uyarı      : $softDifferent"
    Write-RunSummary
    Write-Host ''

    if (($different -eq 0) -and ($missing -eq 0) -and ($extra -eq 0) -and ($errors -eq 0) -and ($metadataDifferent -eq 0)) {
        Write-Host 'Sonuç: Yedek ile kaynak uyumlu.' -ForegroundColor Green
        if ($softDifferent -gt 0) {
            Write-Host "Not: $softDifferent ikincil metadata farkı var (klasör değiştirilme zamanı / nitelik); yedeğin geçerliliğini etkilemez." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host 'Sonuç: İncelenmesi gereken farklar var.' -ForegroundColor Red
    }

    Write-RunLog 'BİLGİ' ("ÖZET | kontrol={0} aynı={1} farklı={2} eksik={3} fazla={4} hata={5} dosyaZamanıFarkı={6} ikincilUyarı={7}" -f $checked, $same, $different, $missing, $extra, $errors, $metadataDifferent, $softDifferent)
    Write-RunLog 'BİLGİ' ("Süre: {0:N1} sn" -f (Get-RunElapsed).TotalSeconds)
    Save-RunReport -Result $(if (($different -eq 0) -and ($missing -eq 0) -and ($extra -eq 0) -and ($errors -eq 0) -and ($metadataDifferent -eq 0)) { 'UYUMLU' } else { 'FARKLAR VAR' })

    Write-Host ''
    Write-Host 'Q - Ana Menü'
    while ($true) {
        if ((Read-Host 'Seçiminiz') -match '^[Qq]$') { return 'MENU' }
        Write-Host 'Q kullanın.' -ForegroundColor Yellow
    }
}

# ============================================================
# ÖZ-TESTLER  (Pester 5.x)
# .\yerelyedek.ps1 -SelfTest  ile çalışır
# ============================================================

function Invoke-SelfTests {
    if (-not (Get-Module -ListAvailable -Name Pester)) {
        Write-Host 'Pester modülü bulunamadı. Öz-testler atlanıyor.' -ForegroundColor Yellow
        return
    }
    Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop | Out-Null

    # Test edilecek script'i geçici olarak yükleyip fonksiyonları kapsama al.
    $here = $PSCommandPath

    $container = New-PesterContainer -ScriptBlock {
        BeforeAll {
            # Script'in kendisini dot-source etmeden, zaten bellekte olan fonksiyonları kullanıyoruz.
        }

        Describe 'Path helper' {
            It 'Get-RelativeFromDrive: D:\\Users → D\\Users' {
                (Get-RelativeFromDrive 'D:\Users') | Should -Be 'D\Users'
            }
            It 'Get-RelativeFromDrive: D:\\ (kök) → D' {
                (Get-RelativeFromDrive 'D:\') | Should -Be 'D'
            }
            It 'Get-DestinationRoot: kaynak D:\\Veri, yedek E:\\Y' {
                (Get-DestinationRoot -Source 'D:\Veri' -BackupRoot 'E:\Y') | Should -Be 'E:\Y\D\Veri'
            }
            It 'Test-PathEqualOrUnder: çocuk ebeveynin altında' {
                (Test-PathEqualOrUnder -Child 'C:\a\b\c' -Parent 'C:\a\b') | Should -BeTrue
            }
            It 'Test-PathEqualOrUnder: aynı yol' {
                (Test-PathEqualOrUnder -Child 'C:\a\b' -Parent 'C:\a\b') | Should -BeTrue
            }
            It 'Test-PathEqualOrUnder: kardeş yol değil' {
                (Test-PathEqualOrUnder -Child 'C:\a\bc' -Parent 'C:\a\b') | Should -BeFalse
            }
            It 'Test-IsUnderRelativePath: basit alt' {
                (Test-IsUnderRelativePath -Child 'a\b\c' -Parent 'a\b') | Should -BeTrue
            }
            It 'Test-IsUnderRelativePath: aynı' {
                (Test-IsUnderRelativePath -Child 'a\b' -Parent 'a\b') | Should -BeTrue
            }
            It 'Test-IsUnderRelativePath: kardeş değil' {
                (Test-IsUnderRelativePath -Child 'a\bc' -Parent 'a\b') | Should -BeFalse
            }
            It 'Get-Depth: boş → 0' {
                (Get-Depth '') | Should -Be 0
            }
            It 'Get-Depth: a\\b\\c → 3' {
                (Get-Depth 'a\b\c') | Should -Be 3
            }
            It 'Get-NormalizedFullPath: D: → D:\\' {
                (Get-NormalizedFullPath 'D:') | Should -Be 'D:\'
            }
        }

        Describe 'Format helper' {
            It 'Format-ByteSize: 500 B' {
                (Format-ByteSize 500) | Should -Match '500 B'
            }
            It 'Format-ByteSize: 2048 → KB' {
                (Format-ByteSize 2048) | Should -Match 'KB'
            }
            It 'Format-ByteSize: 1 MB' {
                (Format-ByteSize 1048576) | Should -Match 'MB'
            }
            It 'Format-ByteSize: 2 GB' {
                (Format-ByteSize 2147483648) | Should -Match 'GB'
            }
            It 'Format-Eta: negatif → --:--' {
                (Format-Eta -1) | Should -Be '--:--'
            }
            It 'Format-Eta: 65 saniye → 01:05' {
                (Format-Eta 65) | Should -Be '01:05'
            }
            It 'Format-Eta: 3700 saniye → 1:01:40' {
                (Format-Eta 3700) | Should -Be '1:01:40'
            }
        }

        Describe 'Time tolerance' {
            It 'NTFS → 0 tick' {
                (Get-TimeToleranceTicks 'NTFS') | Should -Be 0
            }
            It 'ReFS → 0 tick' {
                (Get-TimeToleranceTicks 'ReFS') | Should -Be 0
            }
            It 'exFAT → 200000 tick (20 ms)' {
                (Get-TimeToleranceTicks 'exFAT') | Should -Be 200000
            }
            It 'FAT32 → 20000000 tick (2 sn)' {
                (Get-TimeToleranceTicks 'FAT32') | Should -Be 20000000
            }
            It 'bilinmeyen → 20000000 tick (2 sn)' {
                (Get-TimeToleranceTicks 'BilinmeyenFS') | Should -Be 20000000
            }
        }

        Describe 'Cloud Files reparse tag' {
            It 'Aile üyesi: 0x9000001A → True' {
                (Test-IsCloudFilesReparseTag ([uint32]0x9000001AL)) | Should -BeTrue
            }
            It 'Aile üyesi: 0x9000101A → True' {
                (Test-IsCloudFilesReparseTag ([uint32]0x9000101AL)) | Should -BeTrue
            }
            It 'Aile üyesi: 0x9000F01A → True' {
                (Test-IsCloudFilesReparseTag ([uint32]0x9000F01AL)) | Should -BeTrue
            }
            It 'Farklı aile: 0x8000001A → False' {
                (Test-IsCloudFilesReparseTag ([uint32]0x8000001AL)) | Should -BeFalse
            }
            It 'Ayrı tag: 0xA0000003 → False' {
                (Test-IsCloudFilesReparseTag ([uint32]0xA0000003)) | Should -BeFalse
            }
            It 'Ayrı tag: 0x00000000 → False' {
                (Test-IsCloudFilesReparseTag ([uint32]0)) | Should -BeFalse
            }
        }

        Describe 'Aux root' {
            It 'Get-AuxRoot Changed' {
                (Get-AuxRoot -BackupRoot 'E:\Y' -Kind 'Changed') | Should -Be 'E:\Y\.Changed'
            }
            It 'Get-AuxRoot Deleted' {
                (Get-AuxRoot -BackupRoot 'E:\Y' -Kind 'Deleted') | Should -Be 'E:\Y\.Deleted'
            }
            It 'Get-AuxRoot Temp' {
                (Get-AuxRoot -BackupRoot 'E:\Y' -Kind 'Temp') | Should -Be 'E:\Y\.Temp'
            }
            It 'Get-ChangedRoot sarmalayıcı' {
                (Get-ChangedRoot 'E:\Y') | Should -Be 'E:\Y\.Changed'
            }
            It 'Get-DeletedRoot sarmalayıcı' {
                (Get-DeletedRoot 'E:\Y') | Should -Be 'E:\Y\.Deleted'
            }
            It 'Get-TempRoot sarmalayıcı' {
                (Get-TempRoot 'E:\Y') | Should -Be 'E:\Y\.Temp'
            }
        }

        Describe 'Metadata attributes mask' {
            It 'SettableAttributeMask ReadOnly biti içerir' {
                ($script:SettableAttributeMask -band [int][System.IO.FileAttributes]::ReadOnly) | Should -Not -Be 0
            }
            It 'SettableAttributeMask Compressed biti içermez' {
                ($script:SettableAttributeMask -band [int][System.IO.FileAttributes]::Compressed) | Should -Be 0
            }
            It 'SettableAttributeMask ReparsePoint biti içermez' {
                ($script:SettableAttributeMask -band [int][System.IO.FileAttributes]::ReparsePoint) | Should -Be 0
            }
        }

        Describe 'Hashtable başlangıç değerleri' {
            It 'EventLabels tüm ValidateSet üyelerini içerir' {
                foreach ($k in 'Yeni','Degisen','Silinen','Eksik','Fazla','Atlanan','ErisimHatasi','OkumaHatasi','MetadataDegisen','MetadataErisimHatasi','MetadataOkumaHatasi','Ayni') {
                    $script:EventLabels.ContainsKey($k) | Should -BeTrue -Because "$k eksik"
                }
            }
            It 'EventColors tüm ValidateSet üyelerini içerir' {
                foreach ($k in 'Yeni','Degisen','Silinen','Eksik','Fazla','Atlanan','ErisimHatasi','OkumaHatasi','MetadataDegisen','MetadataErisimHatasi','MetadataOkumaHatasi','Ayni') {
                    $script:EventColors.ContainsKey($k) | Should -BeTrue -Because "$k eksik"
                }
            }
            It 'ActionColors tüm Action türlerini içerir' {
                foreach ($k in 'COPY','ARCHIVE_CHANGED','ARCHIVE_DELETED','UPDATE_METADATA','REPLACE','NOOP','SKIP','ERROR') {
                    $script:ActionColors.ContainsKey($k) | Should -BeTrue -Because "$k eksik"
                }
            }
            It 'ReportLimits beklenen anahtarları içerir' {
                foreach ($k in 'PlanListPerCategory','PlanListCopy','EventPrintLimit','WarningListMax','SkipListMax') {
                    $script:ReportLimits.ContainsKey($k) | Should -BeTrue -Because "$k eksik"
                }
            }
        }

        Describe 'Manifest satırı oluşturma' {
            It 'New-MergeManifestRow tüm alanları doldurur' {
                $r = New-MergeManifestRow `
                    'x.txt' 'C:\s\x.txt' 'E:\d\x.txt' 'File' 'File' `
                    $true $true $true $true $true $true $true `
                    'AABB' 'AABB' $true `
                    $false '' $false '' 'Aynı' 'NOOP'
                $r.RelativePath | Should -Be 'x.txt'
                $r.SourceType   | Should -Be 'File'
                $r.Action       | Should -Be 'NOOP'
                $r.Result       | Should -Be 'PENDING'
                $r.Sha256Same   | Should -BeTrue
            }
            It 'New-MergeManifestRow varsayılan Result PENDING' {
                $r = New-MergeManifestRow 'y' 'a' 'b' 'File' 'File' $true $false $false $false $false $false $false '' '' $false $false '' $false '' 'Yeni' 'COPY'
                $r.Result | Should -Be 'PENDING'
                $r.Error  | Should -Be ''
            }
        }

        Describe 'Reparse Point tag' {
            It 'Var olmayan yol → $null' {
                (Get-ReparsePointTag 'C:\Bu\Yol\Kesinlikle\Yok_XYZ_987') | Should -BeNullOrEmpty
            }
        }
    }

    $config = New-PesterConfiguration
    $config.Run.Container = $container
    $config.Run.PassThru  = $true
    $config.Output.Verbosity = 'Detailed'

    $result = Invoke-Pester -Configuration $config
    return $result
}

# ============================================================
# GİRİŞ NOKTASI
# ============================================================

if ($SelfTest) {
    [void](Invoke-SelfTests)
    return
}

# ============================================================
# ANA MENÜ
# ============================================================

while ($true) {
    $config = Get-ConfigurationStatus
    Clear-Host

    Write-Host '------------------------------------------------------------------------' -ForegroundColor Magenta
    Write-Host '💾 --- YEREL YEDEKLEME SİSTEMİ ---' -ForegroundColor Magenta
    Write-Host '------------------------------------------------------------------------' -ForegroundColor Magenta
    Write-Host ''
    Write-Host '📌 YAPILANDIRMA DURUMU'
    Write-Host '------------------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    if ($null -ne $config.Settings) {
        $sourcePath = $config.Settings.Source
        $backupPath = $config.Settings.BackupRoot
        $destinationPath = $config.Destination
        $changedPath = Get-ChangedRoot $backupPath
        $deletedPath = Get-DeletedRoot $backupPath
        $tempPath    = Get-TempRoot $backupPath
    } else {
        $sourcePath = 'AYARLANMAMIŞ'; $backupPath = 'AYARLANMAMIŞ'; $destinationPath = 'AYARLANMAMIŞ'
        $changedPath = 'AYARLANMAMIŞ'; $deletedPath = 'AYARLANMAMIŞ'; $tempPath = 'AYARLANMAMIŞ'
    }

    Write-InfoLine -Label 'Kaynak' -Value $sourcePath -Icon '📁' -Color Yellow
    Write-InfoLine -Label 'Yedek kökü' -Value $backupPath -Icon '💾' -Color Yellow
    Write-InfoLine -Label 'Hedef' -Value $destinationPath -Icon '🎯' -Color Green
    Write-InfoLine -Label '.Changed' -Value $changedPath -Icon '📦' -Color DarkCyan
    Write-InfoLine -Label '.Deleted' -Value $deletedPath -Icon '🚮' -Color DarkCyan
    Write-InfoLine -Label '.Temp' -Value $tempPath -Icon '⏳' -Color DarkCyan
    Write-Host ''

    $sourceStatusLabel = if ($config.SourceStatus -eq 'OK') { '✅ OK' } else { '❌ YOK' }
    $backupStatusLabel = if ($config.BackupStatus -eq 'OK') { '✅ OK' } else { '❌ YOK' }
    $destinationStatusLabel = switch ($config.DestinationStatus) {
        'OK'          { '✅ OK' }
        'HAZIRLANACAK'{ '🟡 HAZIRLANACAK' }
        default       { '❌ HATA' }
    }
    $settingsFileStatusLabel = if ($config.SettingsFileStatus -eq 'OK') { '✅ OK' }
                               elseif ($config.SettingsFileStatus -eq 'GEÇERSİZ') { '❌ GEÇERSİZ' }
                               else { '❌ BULUNAMADI' }
    $overallStatusLabel = if ($config.Status -eq 'READY') { '✅ Ayarlar Hazır' } else { "❌ $($config.Message)" }

    Write-Host ("📁 {0,-15}: {1}" -f 'Kaynak durumu',$sourceStatusLabel)
    Write-Host ("💾 {0,-15}: {1}" -f 'Yedek durumu',$backupStatusLabel)
    Write-Host ("🎯 {0,-15}: {1}" -f 'Hedef durumu',$destinationStatusLabel)
    Write-Host ("⚙️ {0,-15}: {1}" -f 'Ayar dosyası',$settingsFileStatusLabel)
    Write-Host ''
    Write-Host ("📌 {0,-15}: {1}" -f 'Durum',$overallStatusLabel)
    Write-Host ''

    Write-Host '📜 PROGRAM POLİTİKASI' -ForegroundColor Cyan
    Write-Host '  • Kaynak dosyaları okunur; kaynakta değişiklik veya silme yapılmaz.' -ForegroundColor DarkGray
    Write-Host '  • Değişen ve silinen hedefler .Changed ve .Deleted klasörlerinde arşivlenir.' -ForegroundColor DarkGray
    Write-Host '  • Dosyalar .Temp altında hazırlanır, SHA-256 ile doğrulanır ve sonra hedefe alınır.' -ForegroundColor DarkGray
    Write-Host '  • Ön kontrol dosya boyutu ve LastWriteTime kullanır; gerçek yedeklemede SHA-256 kapsamı H/E seçimine bağlıdır.' -ForegroundColor DarkGray
    Write-Host '  • Kopyalamada LastWriteTime ve nitelikler uygulanır; CreationTime kaynakla eşitlenmez.' -ForegroundColor DarkGray
    Write-Host '  • Bütünlük kontrolü ortak dosyaları SHA-256 ile karşılaştırır; hatalar kaydedilir.' -ForegroundColor DarkGray
    Write-Host '  • İşlem kayıtları Data\Modules\yerelyedek\Logs klasöründe tutulur.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '------------------------------------------------------------------------' -ForegroundColor DarkGray

    Write-Host '  1. 🔍  Kontrol Raporu - Planı gösterir, dosyaları değiştirmez'
    Write-Host '  2. 💾  Gerçek Yedekleme - Planı onayla ve uygula'
    Write-Host '  3. 🛡️  Bütünlük Kontrolü - Kaynak ile yedeği doğrula'
    Write-Host '  4. ⚙️  Ayarlar - Kaynak ve yedek kökünü yapılandır'
    Write-Host '  Q  🚪  Çıkış'
    Write-Host '------------------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    $choice = Read-Host 'Seçiminiz'

    try {
        switch -Regex ($choice) {
            '^1$' {
                if ($config.Status -ne 'READY') {
                    Write-Host ("Ayarlar hazır değil: {0}" -f $config.Message) -ForegroundColor Red
                    Read-Host 'Ayarları düzenlemek için 4, menüye dönmek için Enter'
                    break
                }
                $backupResult = Start-Backup -TestMode $true
                if ($backupResult -eq 'REPORT') { Read-Host 'Raporu kapatıp ana menüye dönmek için Enter' }
            }

            '^2$' {
                if ($config.Status -ne 'READY') {
                    Write-Host ("Ayarlar hazır değil: {0}" -f $config.Message) -ForegroundColor Red
                    Read-Host 'Ayarları düzenlemek için 4, menüye dönmek için Enter'
                    break
                }
                $fullHashChoice = (Read-Host 'Tüm ortak dosyalar SHA-256 ile taransın mı? (E/H, varsayılan H)') -match '^[Ee]$'
                $backupResult = Start-Backup -TestMode $false -FullHash $fullHashChoice
                if ($backupResult -eq 'REPORT') { Read-Host 'Raporu kapatıp ana menüye dönmek için Enter' }
            }

            '^3$' {
                if ($config.Status -ne 'READY') {
                    Write-Host ("Ayarlar hazır değil: {0}" -f $config.Message) -ForegroundColor Red
                    Read-Host 'Ayarları düzenlemek için 4, menüye dönmek için Enter'
                    break
                }
                [void](Start-IntegrityCheck)
            }

            '^4$' { [void](Settings-Menu) }

            '^[Qq]$' {
                if (-not $env:MFT_MODULE_ID) {
                    Write-Host ''
                    Write-Host 'Programdan çıkılıyor. İyi günler!' -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 500
                }
                return
            }

            default {
                Write-Host 'Geçersiz seçim.' -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
    catch {
        Write-Host ''
        Write-Host "[HATA] $($_.Exception.Message)" -ForegroundColor Red
        if ($_.InvocationInfo) {
            Write-Host "       Satır $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkYellow
        }
        Write-Host ''
        Read-Host 'Devam etmek için Enter'
    }
}
