#requires -Version 5.1


# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}

# ============================================================
# E-DEFTER ZIPLEME PROGRAMI - v2.1.0
# PowerShell 5.1
# ============================================================
#
# V2 REFACTOR / GÜVENLİK HEDEFLERİ
#   - Tek dosyalı yapı korunur.
#   - Ana menü korunur: 1 / 2 / 3 / 4.
#   - İşlem başladıktan sonra yarıda kesme yoktur.
#   - Hatalı kayıt sonraki kayıtları durdurmaz.
#   - Mevcut hedef ZIP'lerine dokunulmaz.
#   - Yeni ZIP önce benzersiz TMP olarak oluşturulur.
#   - TMP ZIP: okunabilirlik + dosya sayısı + yollar + SHA-256 ile doğrulanır.
#   - .silinenler özel klasördür ve normal müşteri ağacından ayrı işlenir.
#   - Bütünlük modülü salt okunurdur.
#   - RunRecords + konsol + zaman damgalı TXT log korunur.
#
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# KONSOL / UTF-8
# ============================================================

try {
    chcp 65001 | Out-Null
}
catch {}

$utf8 = New-Object System.Text.UTF8Encoding($true)

try { [Console]::InputEncoding = $utf8 } catch {}
try { [Console]::OutputEncoding = $utf8 } catch {}
$OutputEncoding = $utf8

# ============================================================
# SABİTLER
# ============================================================

$ProgramFolder = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($ProgramFolder)) {
    $ProgramFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$SettingsFile      = Join-Path (Get-MFTModuleDataPath -ModuleId 'edefterzip' -Category Config) 'edefterzip_path_config.json'
$LogFolder         = Get-MFTModuleDataPath -ModuleId 'edefterzip' -Category Logs
$ReportFolder      = Get-MFTModuleDataPath -ModuleId 'edefterzip' -Category Reports
$script:LogFile           = $null
$script:LogInitialized    = $false
$SpecialFolderName = '.silinenler'
$ProgramVersion = '2.1.0'
$script:LastRecoveryTarget = ''
$Header            = '🗜️ --- E-DEFTER ZİPLEME PROGRAMI ---'
$Line              = '-' * 72

Add-Type -AssemblyName System.IO.Compression.FileSystem

# ============================================================
# GLOBAL KAYITLAR
# ============================================================

$global:RunRecords = New-Object 'System.Collections.Generic.List[object]'

# ============================================================
# EKRAN / YARDIMCI
# ============================================================

function Show-Header {
    Clear-Host
    Write-Host $Line -ForegroundColor Magenta
    Write-Host $Header -ForegroundColor Magenta
    Write-Host $Line -ForegroundColor Magenta
}

function Pause-Program {
    Write-Host ''
    Write-Host 'Devam etmek için bir tuşa basın...' -ForegroundColor DarkGray

    try {
        $null = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        Read-Host 'Devam etmek için ENTER''a basın'
    }
}

function Test-FolderExists {
    param([Parameter(Mandatory=$true)][string]$Path)
    try {
        return (Test-Path -LiteralPath $Path -PathType Container)
    }
    catch {
        return $false
    }
}

function Test-FileExists {
    param([Parameter(Mandatory=$true)][string]$Path)
    try {
        return (Test-Path -LiteralPath $Path -PathType Leaf)
    }
    catch {
        return $false
    }
}

function Test-PathExists {
    param([Parameter(Mandatory=$true)][string]$Path)
    try {
        return (Test-Path -LiteralPath $Path)
    }
    catch {
        return $false
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        [void][System.IO.Directory]::CreateDirectory($Path)
        return [System.IO.Directory]::Exists($Path)
    }
    catch {
        return $false
    }
}

# ============================================================
# KAYNAK / HEDEF GÜVENLİK
# ============================================================

function Test-SourceTargetSeparation {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot
    )

    try {
        $sourceFull = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')
        $targetFull = [System.IO.Path]::GetFullPath($TargetRoot).TrimEnd('\')

        if ([string]::Equals($sourceFull, $targetFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{ Success = $false; Message = 'Kaynak ve hedef aynı klasör olamaz.' }
        }

        if ($targetFull.StartsWith(($sourceFull + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{ Success = $false; Message = 'Hedef klasör kaynak klasörünün altında olamaz.' }
        }

        if ($sourceFull.StartsWith(($targetFull + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{ Success = $false; Message = 'Kaynak klasör hedef klasörünün altında olamaz.' }
        }

        return [PSCustomObject]@{ Success = $true; Message = '' }
    }
    catch {
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

# ============================================================
# AYARLAR
# ============================================================

function Load-Settings {
    if (Test-FileExists -Path $SettingsFile) {
        try {
            $json = Get-Content -LiteralPath $SettingsFile -Raw -Encoding UTF8
            $settings = $json | ConvertFrom-Json

            $source = ''
            $target = ''

            if ($null -ne $settings.SourceFolder) { $source = [string]$settings.SourceFolder }
            if ($null -ne $settings.TargetFolder) { $target = [string]$settings.TargetFolder }

            return [PSCustomObject]@{
                SourceFolder = $source
                TargetFolder = $target
            }
        }
        catch {
            Write-Host ''
            Write-Host '⚠️ Ayar dosyası okunamadı.' -ForegroundColor Yellow
        }
    }

    return [PSCustomObject]@{
        SourceFolder = ''
        TargetFolder = ''
    }
}

function Save-Settings {
    param([Parameter(Mandatory=$true)][PSCustomObject]$Settings)

    try {
        $json = $Settings | ConvertTo-Json -Depth 5
        Set-Content -LiteralPath $SettingsFile -Value $json -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-Host ''
        Write-Host ('❌ Ayarlar kaydedilemedi: {0}' -f $_.Exception.Message) -ForegroundColor Red
        return $false
    }
}

function Wait-ForSettingsMenuEnter {
    while ($true) {
        $response = Read-Host 'Ayarlar Menüsüne dönmek için Enter'
        if ([string]::IsNullOrEmpty($response)) { return }
        Write-Host 'Devam etmek için yalnızca Enter tuşuna basın.' -ForegroundColor DarkGray
    }
}

function Configure-Settings {
    param([Parameter(Mandatory=$true)][ref]$CurrentSettingsRef)

    while ($true) {
        $originalSourceFolder = $CurrentSettingsRef.Value.SourceFolder
        $originalTargetFolder = $CurrentSettingsRef.Value.TargetFolder

        Clear-Host
        Write-Host $Line -ForegroundColor Magenta
        Write-Host '⚙️ E-DEFTER ZIP AYARLARI' -ForegroundColor Magenta
        Write-Host $Line -ForegroundColor Magenta
        Write-Host ''
        Write-Host '⚙️ MEVCUT AYARLAR' -ForegroundColor Cyan
        Write-Host $Line -ForegroundColor DarkGray
        Write-Host '📁 Mevcut kaynak:'
        Write-Host ('   {0}' -f $CurrentSettingsRef.Value.SourceFolder) -ForegroundColor DarkGray
        Write-Host '💾 Mevcut hedef:'
        Write-Host ('   {0}' -f $CurrentSettingsRef.Value.TargetFolder) -ForegroundColor DarkGray
        Write-Host $Line -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  1. ⚙️  Ayarları Düzenle'
        Write-Host '  Q  ↩️  Ana Menü'
        Write-Host ''
        Write-Host $Line -ForegroundColor DarkGray

        $choice = Read-Host 'Seçiminiz'

        if ($choice.ToLower() -eq 'q') { return }

        if ($choice -ne '1') {
            Write-Host ''
            Write-Host '❌ Geçersiz seçim.' -ForegroundColor Red
            Wait-ForSettingsMenuEnter
            continue
        }

        Clear-Host
        Write-Host $Line -ForegroundColor Magenta
        Write-Host '⚙️ E-DEFTER ZIP AYARLARI' -ForegroundColor Magenta
        Write-Host $Line -ForegroundColor Magenta
        Write-Host ''
        Write-Host '⚙️ AYARLARI DÜZENLE' -ForegroundColor Cyan
        Write-Host $Line -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '📁 Mevcut kaynak:'
        Write-Host ('   {0}' -f $CurrentSettingsRef.Value.SourceFolder) -ForegroundColor DarkGray
        Write-Host ''

        $newSource = Read-Host 'Yeni KAYNAK klasör yolu (boş = değiştirme)'
        if (-not [string]::IsNullOrWhiteSpace($newSource)) {
            $newSource = $newSource.Trim()
            if (-not [System.IO.Path]::IsPathRooted($newSource)) {
                Write-Host ''
                Write-Host '❌ Mutlak bir klasör yolu girilmelidir.' -ForegroundColor Red
            }
            else {
                $CurrentSettingsRef.Value.SourceFolder = $newSource
                Write-Host ''
                Write-Host '✅ Kaynak ayarlandı.' -ForegroundColor Green
            }
        }

        Write-Host ''
        Write-Host '💾 Mevcut hedef:'
        Write-Host ('   {0}' -f $CurrentSettingsRef.Value.TargetFolder) -ForegroundColor DarkGray
        Write-Host ''

        $newTarget = Read-Host 'Yeni HEDEF klasör yolu (boş = değiştirme)'
        if (-not [string]::IsNullOrWhiteSpace($newTarget)) {
            $newTarget = $newTarget.Trim()
            if (-not [System.IO.Path]::IsPathRooted($newTarget)) {
                Write-Host ''
                Write-Host '❌ Mutlak bir klasör yolu girilmelidir.' -ForegroundColor Red
            }
            else {
                $CurrentSettingsRef.Value.TargetFolder = $newTarget
                Write-Host ''
                Write-Host '✅ Hedef ayarlandı.' -ForegroundColor Green
            }
        }

        [void](Save-Settings -Settings $CurrentSettingsRef.Value)
        if (
            $CurrentSettingsRef.Value.SourceFolder -eq $originalSourceFolder -and
            $CurrentSettingsRef.Value.TargetFolder -eq $originalTargetFolder
        ) {
            Write-Host ''
            Write-Host 'Ayar değişikliği yapılmadı.' -ForegroundColor Yellow
	    Write-Host ''
        }
        Wait-ForSettingsMenuEnter
    }
}

function Show-ConfigurationStatus {
    param([Parameter(Mandatory=$true)][PSCustomObject]$Settings)

    $sourceConfigured = -not [string]::IsNullOrWhiteSpace($Settings.SourceFolder)
    $targetConfigured = -not [string]::IsNullOrWhiteSpace($Settings.TargetFolder)
    $sourceExists = $sourceConfigured -and (Test-FolderExists -Path $Settings.SourceFolder)
    $targetExists = $targetConfigured -and (Test-FolderExists -Path $Settings.TargetFolder)

    Write-Host ''
    Write-Host '📌 YAPILANDIRMA DURUMU'
    Write-Host $Line -ForegroundColor DarkGray
    Write-Host ''

    Write-Host "📁 Kaynak        : $($Settings.SourceFolder)" -ForegroundColor Yellow

    Write-Host "💾 Hedef         : $($Settings.TargetFolder)" -ForegroundColor Green

    Write-Host ''

    if ($sourceExists) {
        Write-Host '📁 Kaynak durumu : ✅ OK'
    }
    elseif ($sourceConfigured) {
        Write-Host '📁 Kaynak durumu : ❌ KLASÖR YOK' -ForegroundColor Red
    }
    else {
        Write-Host '📁 Kaynak durumu : ⚠️ AYARLANMADI' -ForegroundColor Yellow
    }

    if ($targetExists) {
        Write-Host '💾 Hedef durumu  : ✅ OK'
    }
    elseif ($targetConfigured) {
        Write-Host '💾 Hedef durumu  : ⚠️ YOK (işlemde oluşturulur)' -ForegroundColor Yellow
    }
    else {
        Write-Host '💾 Hedef durumu  : ⚠️ AYARLANMADI' -ForegroundColor Yellow
    }

    $settingsReady = $sourceExists -and $targetConfigured

    Write-Host ''
    Write-Host '📌 Durum         : ' -NoNewline

    if ($settingsReady) {
        Write-Host '✅ Ayarlar Hazır' -ForegroundColor Green
    }
    else {
        Write-Host '⚠️ Ayarlar Kontrol Edilmeli' -ForegroundColor Yellow
    }
}
# ============================================================
# HASH / ZIP OKUMA
# ============================================================

function Get-Sha256FromStream {
    param([Parameter(Mandatory=$true)][System.IO.Stream]$Stream)

    $sha256 = $null
    try {
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hash = $sha256.ComputeHash($Stream)
        return ([BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant())
    }
    finally {
        if ($null -ne $sha256) { $sha256.Dispose() }
    }
}

function Get-Sha256FromFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        return (Get-Sha256FromStream -Stream $stream)
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Test-ZipReadable {
    param([Parameter(Mandatory=$true)][string]$ZipPath)

    $archive = $null
    try {
        if (-not (Test-FileExists -Path $ZipPath)) { return $false }
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)

        foreach ($entry in $archive.Entries) {
            $null = $entry.FullName
            $null = $entry.Length
            $null = $entry.CompressedLength

            $stream = $null
            try {
                $stream = $entry.Open()
                $buffer = New-Object byte[] 8192
                while ($stream.Read($buffer, 0, $buffer.Length) -gt 0) { }
            }
            finally {
                if ($null -ne $stream) { $stream.Dispose() }
            }
        }

        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

# ============================================================
# KAYNAK DOSYALARI
# Hidden/System dosyalar da dahil edilir.
# ============================================================

function Get-SourceFiles {
    # -Force: gizli/system dosyalar da envantere dahil edilir.
    param([Parameter(Mandatory=$true)][string]$Folder)

    try {
        $files = @(Get-ChildItem -LiteralPath $Folder -Recurse -File -Force -ErrorAction Stop | Sort-Object FullName)
        return [PSCustomObject]@{
            Success      = $true
            Files        = $files
            Count        = $files.Count
            ErrorMessage = ''
        }
    }
    catch {
        return [PSCustomObject]@{
            Success      = $false
            Files        = @()
            Count        = 0
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Eski çağrı adını korumak için küçük uyumluluk sarmalayıcısı.
function Get-SourceFileCount {
    param([Parameter(Mandatory=$true)][string]$Folder)
    return (Get-SourceFiles -Folder $Folder)
}

# ============================================================
# KAYNAK / ZIP BÜTÜNLÜK KARŞILAŞTIRMASI
# ============================================================

function Test-ZipMatchesSource {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [Parameter(Mandatory=$true)][string]$ZipPath
    )

    $archive = $null

    try {
        # Kaynak klasörü gizli/system olsa bile -Force ile bul.
        $sourceItem = $null
        try {
            $sourceItem = Get-Item -LiteralPath $SourceFolder -Force -ErrorAction Stop
        }
        catch {
            return [PSCustomObject]@{ Success=$false; ResultType='SourceError'; Message='Kaynak ay klasörü bulunamadı veya okunamadı.'; Detail=$_.Exception.Message }
        }

        if (-not $sourceItem.PSIsContainer) {
            return [PSCustomObject]@{ Success=$false; ResultType='SourceError'; Message='Kaynak yolu klasör değil.'; Detail=$SourceFolder }
        }

        if (-not (Test-FileExists -Path $ZipPath)) {
            return [PSCustomObject]@{ Success=$false; ResultType='MissingZip'; Message='ZIP dosyası bulunamadı.'; Detail=$ZipPath }
        }

        $sourceRoot = $sourceItem.FullName.TrimEnd([char]'\',[char]'/')

        $sourceResult = Get-SourceFiles -Folder $sourceRoot
        if (-not $sourceResult.Success) {
            return [PSCustomObject]@{ Success=$false; ResultType='SourceError'; Message='Kaynak ay okunamadı.'; Detail=$sourceResult.ErrorMessage }
        }

        $sourceMap = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($file in $sourceResult.Files) {
            $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart([char]'\',[char]'/').Replace('\','/')
            if ($sourceMap.ContainsKey($relativePath)) {
                return [PSCustomObject]@{ Success=$false; ResultType='SourceError'; Message='Kaynakta yinelenen dosya yolu bulundu.'; Detail=$relativePath }
            }
            $sourceMap.Add($relativePath, $file.FullName)
        }

        try {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        }
        catch {
            return [PSCustomObject]@{ Success=$false; ResultType='ZipUnreadable'; Message='ZIP okunamıyor.'; Detail=$_.Exception.Message }
        }

        $archiveRootPrefix = $sourceItem.Name.Replace('\','/') + '/'
        $zipMap = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }
            if ($entry.FullName.EndsWith('/')) { continue }

            $entryPath = $entry.FullName.Replace('\','/')

            if ($entryPath.StartsWith($archiveRootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $entryPath = $entryPath.Substring($archiveRootPrefix.Length)
            }

            if ([string]::IsNullOrWhiteSpace($entryPath)) { continue }

            if ($zipMap.ContainsKey($entryPath)) {
                return [PSCustomObject]@{ Success=$false; ResultType='DuplicateZipPath'; Message='ZIP içinde yinelenen dosya yolu bulundu.'; Detail=$entryPath }
            }

            $zipMap.Add($entryPath, $entry)
        }

        if ($sourceMap.Count -ne $zipMap.Count) {
            return [PSCustomObject]@{ Success=$false; ResultType='FileCountMismatch'; Message='Dosya sayısı uyuşmuyor.'; Detail=('Kaynak: {0} / ZIP: {1}' -f $sourceMap.Count,$zipMap.Count) }
        }

        foreach ($path in $sourceMap.Keys) {
            if (-not $zipMap.ContainsKey($path)) {
                return [PSCustomObject]@{ Success=$false; ResultType='PathMismatch'; Message='ZIP içinde kaynak dosyası bulunamadı.'; Detail=$path }
            }
        }

        foreach ($path in $zipMap.Keys) {
            if (-not $sourceMap.ContainsKey($path)) {
                return [PSCustomObject]@{ Success=$false; ResultType='PathMismatch'; Message='ZIP içinde kaynakta olmayan dosya var.'; Detail=$path }
            }
        }

        foreach ($path in $sourceMap.Keys) {
            $sourceHash = Get-Sha256FromFile -Path $sourceMap[$path]
            $stream = $null
            try {
                $stream = $zipMap[$path].Open()
                $zipHash = Get-Sha256FromStream -Stream $stream
            }
            catch {
                return [PSCustomObject]@{ Success=$false; ResultType='ZipUnreadable'; Message='ZIP içindeki dosya okunamadı.'; Detail=$path }
            }
            finally {
                if ($null -ne $stream) { $stream.Dispose() }
            }

            if ($sourceHash -ne $zipHash) {
                return [PSCustomObject]@{ Success=$false; ResultType='ShaMismatch'; Message='SHA-256 uyuşmuyor.'; Detail=$path }
            }
        }

        return [PSCustomObject]@{ Success=$true; ResultType='IntegrityOK'; Message='Bütünlük OK'; Detail='' }
    }
    catch {
        return [PSCustomObject]@{ Success=$false; ResultType='CheckError'; Message='Bütünlük kontrolü sırasında hata oluştu.'; Detail=$_.Exception.Message }
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

# ============================================================
# YENİ ZIP OLUŞTURMA
# ============================================================

function Create-NewZip {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFolder,
        [Parameter(Mandatory=$true)][string]$DestinationZip
    )

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $uniqueId = [guid]::NewGuid().ToString('N')
    $temporaryZip = '{0}.{1}_{2}.tmp' -f $DestinationZip,$timestamp,$uniqueId
    $temporaryCreatedByThisRun = $false

    try {
        if (Test-PathExists -Path $DestinationZip) {
            return [PSCustomObject]@{ Success=$false; RaceDetected=$true; Message='Hedefte aynı isimde bir nesne mevcut.' }
        }

        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $SourceFolder,
            $temporaryZip,
            [System.IO.Compression.CompressionLevel]::Optimal,
            $true
        )

        $temporaryCreatedByThisRun = $true

        if (-not (Test-FileExists -Path $temporaryZip)) {
            throw 'Geçici ZIP dosyası oluşturulamadı.'
        }

        if (-not (Test-ZipReadable -ZipPath $temporaryZip)) {
            throw 'Oluşturulan TMP ZIP okunamadı.'
        }

        $matchResult = Test-ZipMatchesSource -SourceFolder $SourceFolder -ZipPath $temporaryZip
        if (-not $matchResult.Success) {
            throw ('ZIP kaynakla doğrulanamadı: {0} {1}' -f $matchResult.Message,$matchResult.Detail)
        }

        if (Test-PathExists -Path $DestinationZip) {
            return [PSCustomObject]@{ Success=$false; RaceDetected=$true; Message='ZIP doğrulanırken hedefte aynı isimde bir nesne oluştu.' }
        }

        [System.IO.File]::Move($temporaryZip,$DestinationZip)
        $temporaryCreatedByThisRun = $false

        if (-not (Test-FileExists -Path $DestinationZip)) {
            throw 'ZIP hedef konumda oluşmadı.'
        }

        return [PSCustomObject]@{ Success=$true; RaceDetected=$false; Message='' }
    }
    catch {
        if ($temporaryCreatedByThisRun -and (Test-PathExists -Path $temporaryZip)) {
            Remove-Item -LiteralPath $temporaryZip -Force -ErrorAction SilentlyContinue
        }

        if ($_.Exception.Message -like '*aynı isimde bir nesne*') {
            return [PSCustomObject]@{ Success=$false; RaceDetected=$true; Message=$_.Exception.Message }
        }

        return [PSCustomObject]@{ Success=$false; RaceDetected=$false; Message=$_.Exception.Message }
    }
}

# ============================================================
# KAYIT / KONSOL
# ============================================================

function Initialize-Log {
    if ($script:LogInitialized) {
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $script:LogFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $script:LogFolder -Force | Out-Null
        }

        $script:LogFile = Join-Path $script:LogFolder ('EdefterZip_{0}.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        Add-Content -LiteralPath $script:LogFile -Value ('[{0}] Program başlatıldı. Sürüm: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $script:ProgramVersion) -Encoding UTF8
        $script:LogInitialized = $true
    }
    catch {
        # Log yazımı programın asıl işleyişini durdurmamalıdır.
    }
}

function Add-RunRecord {
    param(
        [string]$Customer,
        [string]$Period,
        [string]$Month,
        [string]$Type,
        [string]$Status,
        [string]$StatusText,
        [string]$ZipPath,
        [int]$FileCount = 0,
        [long]$ZipSize = 0,
        [string]$ErrorType = '',
        [string]$ErrorMessage = ''
    )

    $record = [PSCustomObject]@{
        Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Customer       = $Customer
        Period         = $Period
        Month          = $Month
        Type           = $Type
        Status         = $Status
        StatusText     = $StatusText
        ZipPath        = $ZipPath
        FileCount      = $FileCount
        ZipSizeBytes   = $ZipSize
        ErrorType      = $ErrorType
        ErrorMessage   = $ErrorMessage
    }

    [void]$global:RunRecords.Add($record)

    try {
        Initialize-Log
        $logLine = '[{0}] Tür={1} | Durum={2} | Müşteri={3} | Dönem={4} | Ay={5} | Açıklama={6} | ZIP={7} | Dosya={8} | Boyut={9} | Hata={10} {11}' -f `
            $record.Timestamp, $record.Type, $record.Status, $record.Customer, $record.Period, $record.Month, `
            $record.StatusText, $record.ZipPath, $record.FileCount, $record.ZipSizeBytes, $record.ErrorType, $record.ErrorMessage
        Add-Content -LiteralPath $LogFile -Value $logLine -Encoding UTF8
    }
    catch {
        # Log yazımı programın asıl işleyişini durdurmamalıdır.
    }

    return $record
}

function Show-RunRecord {
    param([Parameter(Mandatory=$true)][PSCustomObject]$Record)

    $color = 'White'

    switch ($Record.Status) {
        'New'                    { $color='Green' }
        'Existing'               { $color='DarkGray' }
        'Empty'                  { $color='Cyan' }
        'Error'                  { $color='Red' }
        'Race'                   { $color='Yellow' }
        'Info'                   { $color='Yellow' }
        'IntegrityOK'            { $color='Green' }
        'IntegrityError'         { $color='Red' }
        'MissingZip'             { $color='Yellow' }
        'IntegritySourceError'   { $color='Red' }
        'IntegrityCustomerError' { $color='Red' }
        'IntegrityPeriodError'   { $color='Red' }
        'IntegrityTargetError'   { $color='Red' }
        'TargetExtra'            { $color='Yellow' }
        'Special'                { $color='Green' }
    }

    if ($Record.Type -eq 'Special') {
        Write-Host ('🗑️ .silinenler             → {0}' -f $Record.StatusText) -ForegroundColor $color
    }
    elseif ($Record.Type -eq 'Info' -or $Record.Type -eq 'IntegrityInfo') {
        Write-Host ('ℹ️ {0}' -f $Record.StatusText) -ForegroundColor $color
    }
    else {
        Write-Host ('👤 {0,-18} 📅 {1,-21} 📁 {2,-4} → {3}' -f $Record.Customer,$Record.Period,$Record.Month,$Record.StatusText) -ForegroundColor $color
    }

    if (-not [string]::IsNullOrWhiteSpace($Record.ErrorMessage)) {
        Write-Host ('   ↳ {0}' -f $Record.ErrorMessage) -ForegroundColor DarkGray
    }
}

# ============================================================
# ÖZETLER
# ============================================================

function Get-ZipRunSummary {
    param([Parameter(Mandatory=$true)][System.Collections.Generic.List[object]]$Records)

    $monthRecords = @($Records | Where-Object { $_.Type -eq 'Month' })
    $newRecords = @($monthRecords | Where-Object { $_.Status -eq 'New' })
    $existingRecords = @($monthRecords | Where-Object { $_.Status -eq 'Existing' })
    $emptyRecords = @($Records | Where-Object { $_.Status -eq 'Empty' })
    $errorRecords = @($Records | Where-Object { $_.Status -eq 'Error' })
    $raceRecords = @($Records | Where-Object { $_.Status -eq 'Race' })
    $infoRecords = @($Records | Where-Object { $_.Status -eq 'Info' })

    return [PSCustomObject]@{
        TotalRecordCount = $Records.Count
        NewCount = $newRecords.Count
        NewCustomerCount = @($newRecords | Select-Object -ExpandProperty Customer -Unique).Count
        ExistingCount = $existingRecords.Count
        ExistingCustomerCount = @($existingRecords | Select-Object -ExpandProperty Customer -Unique).Count
        EmptyCount = $emptyRecords.Count
        SpecialNew = @($Records | Where-Object { $_.Type -eq 'Special' -and $_.Status -eq 'Special' }).Count
        SpecialExisting = @($Records | Where-Object { $_.Type -eq 'Special' -and $_.Status -eq 'Existing' }).Count
        RaceCount = $raceRecords.Count
        InfoCount = $infoRecords.Count
        ErrorCount = $errorRecords.Count
        ErrorTargetFolderCount = @($errorRecords | Where-Object { $_.ErrorType -eq 'Hedef klasörü oluşturulamadı' }).Count
        ErrorCustomerFolderCount = @($errorRecords | Where-Object { $_.ErrorType -eq 'Müşteri klasörü okunamadı' }).Count
        ErrorPeriodFolderCount = @($errorRecords | Where-Object { $_.ErrorType -eq 'Dönem klasörü okunamadı' }).Count
        ErrorSourceMonthCount = @($errorRecords | Where-Object { $_.ErrorType -eq 'Kaynak ay okunamadı' }).Count
        ErrorZipProcessCount = @($errorRecords | Where-Object { $_.ErrorType -eq 'ZIP oluşturulamadı / doğrulanamadı' }).Count
    }
}

function Get-IntegritySummary {
    param([Parameter(Mandatory=$true)][System.Collections.Generic.List[object]]$Records)

    return [PSCustomObject]@{
        TotalRecordCount = $Records.Count
        IntegrityOkCount = @($Records | Where-Object { $_.Status -eq 'IntegrityOK' }).Count
        IntegrityErrorCount = @($Records | Where-Object { $_.Status -eq 'IntegrityError' }).Count
        MissingZipCount = @($Records | Where-Object { $_.Status -eq 'MissingZip' }).Count
        TargetExtraCount = @($Records | Where-Object { $_.Status -eq 'TargetExtra' }).Count
        IntegritySourceErrorCount = @($Records | Where-Object { $_.Status -eq 'IntegritySourceError' }).Count
        IntegrityCustomerErrorCount = @($Records | Where-Object { $_.Status -eq 'IntegrityCustomerError' }).Count
        IntegrityPeriodErrorCount = @($Records | Where-Object { $_.Status -eq 'IntegrityPeriodError' }).Count
        IntegrityTargetErrorCount = @($Records | Where-Object { $_.Status -eq 'IntegrityTargetError' }).Count
        IntegrityInfoCount = @($Records | Where-Object { $_.Status -eq 'Info' }).Count
    }
}

# ============================================================
# BİRLİKTE KULLANILAN BAŞLANGIÇ KONTROLLERİ
# ============================================================

function Test-ModulePreconditions {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot,
        [Parameter(Mandatory=$true)][bool]$CreateTargetIfMissing
    )

    if (-not (Test-FolderExists -Path $SourceRoot)) {
        return [PSCustomObject]@{ Success=$false; Message='❌ Kaynak klasör bulunamadı.'; Pause=$true }
    }

    $separation = Test-SourceTargetSeparation -SourceRoot $SourceRoot -TargetRoot $TargetRoot
    if (-not $separation.Success) {
        return [PSCustomObject]@{ Success=$false; Message=('❌ {0}' -f $separation.Message); Pause=$true }
    }

    if ($CreateTargetIfMissing -and -not (Test-FolderExists -Path $TargetRoot)) {
        Write-Host ''
        Write-Host '📁 Hedef klasör oluşturuluyor...' -ForegroundColor Yellow
        if (-not (Ensure-Folder -Path $TargetRoot)) {
            return [PSCustomObject]@{ Success=$false; Message='❌ Hedef klasörü oluşturulamadı.'; Pause=$true }
        }
    }

    return [PSCustomObject]@{ Success=$true; Message=''; Pause=$false }
}

# ============================================================
# .SİLİNENLER - ZIPLEME
# ============================================================

function Process-SpecialFolderForZip {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot
    )

    $sourceFolder = Join-Path $SourceRoot $SpecialFolderName
    $targetZip = Join-Path $TargetRoot ($SpecialFolderName + '.zip')

    # .silinenler özel ve gizli bir klasördür.
    # Varlığı mutlaka -Force ile kontrol edilir.
    $folderItem = $null
    try {
        $folderItem = Get-Item -LiteralPath $sourceFolder -Force -ErrorAction Stop
    }
    catch {
        # Klasör gerçekten yoksa bu hata değildir.
        if ($_.Exception -is [System.Management.Automation.ItemNotFoundException] -or
            $_.Exception.Message -like '*Could not find item*') {
            return
        }

        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Error' `
            -StatusText '❌ .silinenler klasörü kontrol edilemedi' -ZipPath $targetZip `
            -ErrorType 'Kaynak ay okunamadı' -ErrorMessage $_.Exception.Message
        Show-RunRecord -Record $record
        return
    }

    if (-not $folderItem.PSIsContainer) {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Error' `
            -StatusText '❌ .silinenler klasör değil' -ZipPath $targetZip `
            -ErrorType 'Kaynak ay okunamadı' -ErrorMessage '.silinenler yolu klasör değil.'
        Show-RunRecord -Record $record
        return
    }

    Write-Host ''
    Write-Host '🗑️ .silinenler kontrol ediliyor...' -ForegroundColor DarkYellow

    try {

        if (Test-FileExists -Path $targetZip) {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Existing' `
                -StatusText '⏭️ .silinenler ZIP mevcut — dokunulmadı' -ZipPath $targetZip
            Show-RunRecord -Record $record
            return
        }

        $filesResult = Get-SourceFiles -Folder $sourceFolder
        if (-not $filesResult.Success) {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Error' `
                -StatusText '.silinenler kaynak okunamadı — ZIP oluşturulmadı' -ZipPath $targetZip `
                -ErrorType 'Kaynak ay okunamadı' -ErrorMessage $filesResult.ErrorMessage
            Show-RunRecord -Record $record
            return
        }

        if ($filesResult.Count -eq 0) {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Empty' `
                -StatusText '⚪ .silinenler boş — ZIP oluşturulmadı' -ZipPath $targetZip -FileCount 0
            Show-RunRecord -Record $record
            return
        }

        $zipResult = Create-NewZip -SourceFolder $sourceFolder -DestinationZip $targetZip

        if ($zipResult.RaceDetected) {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Race' `
                -StatusText '⚠️ .silinenler hedefinde nesne oluştu — dokunulmadı' -ZipPath $targetZip `
                -FileCount $filesResult.Count -ErrorMessage $zipResult.Message
            Show-RunRecord -Record $record
            return
        }

        if (-not $zipResult.Success) {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Error' `
                -StatusText '❌ .silinenler ZIP oluşturulamadı / doğrulanamadı' -ZipPath $targetZip `
                -FileCount $filesResult.Count -ErrorType 'ZIP oluşturulamadı / doğrulanamadı' -ErrorMessage $zipResult.Message
            Show-RunRecord -Record $record
            return
        }

        $zipSize = 0
        try { $zipSize = (Get-Item -LiteralPath $targetZip -ErrorAction Stop).Length } catch {}

        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Special' `
            -StatusText '🆕 .silinenler ZIP oluşturuldu ve doğrulandı' -ZipPath $targetZip `
            -FileCount $filesResult.Count -ZipSize $zipSize
        Show-RunRecord -Record $record
    }
    catch {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'Error' `
            -StatusText '❌ .silinenler işlenemedi' -ZipPath $targetZip `
            -ErrorType 'Kaynak ay okunamadı' -ErrorMessage $_.Exception.Message
        Show-RunRecord -Record $record
    }
}

# ============================================================
# .SİLİNENLER - BÜTÜNLÜK
# ============================================================

function Process-SpecialFolderIntegrity {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot
    )

    $sourceFolder = Join-Path $SourceRoot $SpecialFolderName
    $targetZip = Join-Path $TargetRoot ($SpecialFolderName + '.zip')

    # Gizli klasörü mutlaka -Force ile bul. Gerçekten yoksa bu bir hata değildir.
    $folderItem = $null
    try {
        $folderItem = Get-Item -LiteralPath $sourceFolder -Force -ErrorAction Stop
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.ItemNotFoundException] -or
            $_.Exception.Message -like '*Could not find item*') {
            return
        }

        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegritySourceError' `
            -StatusText '❌ .silinenler klasörü kontrol edilemedi' -ZipPath $targetZip `
            -ErrorMessage $_.Exception.Message
        Show-RunRecord -Record $record
        return
    }

    if (-not $folderItem.PSIsContainer) {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegritySourceError' `
            -StatusText '❌ .silinenler klasör değil' -ZipPath $targetZip `
            -ErrorMessage '.silinenler yolu klasör değil.'
        Show-RunRecord -Record $record
        return
    }

    Write-Host ''
    Write-Host '🗑️ .silinenler kontrol ediliyor...' -ForegroundColor DarkYellow

    if (-not (Test-FileExists -Path $targetZip)) {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'MissingZip' `
            -StatusText '⚠️ .silinenler ZIP bulunamadı' -ZipPath $targetZip
        Show-RunRecord -Record $record
        return
    }

    $result = Test-ZipMatchesSource -SourceFolder $sourceFolder -ZipPath $targetZip

    switch ($result.ResultType) {
        'IntegrityOK' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityOK' `
                -StatusText '✅ .silinenler bütünlük OK' -ZipPath $targetZip
        }
        'SourceError' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegritySourceError' `
                -StatusText '❌ .silinenler kaynak okunamadı' -ZipPath $targetZip -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
        }
        'FileCountMismatch' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityError' `
                -StatusText '❌ .silinenler dosya sayısı uyuşmuyor' -ZipPath $targetZip -ErrorMessage $result.Detail
        }
        'PathMismatch' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityError' `
                -StatusText '❌ .silinenler dosya yolu uyuşmuyor' -ZipPath $targetZip -ErrorMessage $result.Detail
        }
        'ShaMismatch' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityError' `
                -StatusText '❌ .silinenler SHA-256 uyuşmuyor' -ZipPath $targetZip -ErrorMessage $result.Detail
        }
        'ZipUnreadable' {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityError' `
                -StatusText '❌ .silinenler ZIP okunamıyor' -ZipPath $targetZip -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
        }
        default {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Special' -Status 'IntegrityError' `
                -StatusText '❌ .silinenler bütünlük kontrolü başarısız' -ZipPath $targetZip -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
        }
    }

    Show-RunRecord -Record $record
}

# ============================================================
# 1. MODÜL - ZIPLEME
# ============================================================

function Start-EdefterZipProcess {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot
    )

    Show-Header

    Initialize-Log
    $pre = Test-ModulePreconditions -SourceRoot $SourceRoot -TargetRoot $TargetRoot -CreateTargetIfMissing $true
    if (-not $pre.Success) {
        Write-Host ''
        Write-Host $pre.Message -ForegroundColor Red
        Pause-Program
        return
    }

    $global:RunRecords.Clear()

    Write-Host ''
    Write-Host ('📁 Kaynak : {0}' -f $SourceRoot)
    Write-Host ('💾 Hedef  : {0}' -f $TargetRoot)
    Write-Host ''
    Write-Host '🔎 Kaynak klasörler taranıyor...' -ForegroundColor DarkYellow
    Write-Host ''

    try {
        $customers = @(
            Get-ChildItem -LiteralPath $SourceRoot -Directory -Force -ErrorAction Stop |
            Where-Object { $_.Name -ne $SpecialFolderName } |
            Sort-Object Name
        )
    }
    catch {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Access' -Status 'Error' `
            -StatusText '❌ Müşteri klasörleri okunamadı' -ZipPath '' -ErrorType 'Müşteri klasörü okunamadı' -ErrorMessage $_.Exception.Message
        Show-RunRecord -Record $record
        $summary = Get-ZipRunSummary -Records $global:RunRecords
        return (Show-ZipCompletionScreen -Source $SourceRoot -Target $TargetRoot -Summary $summary)
    }

    if ($customers.Count -eq 0) {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'Info' -Status 'Info' `
            -StatusText 'ℹ️ Kaynakta müşteri klasörü bulunamadı' -ZipPath ''
        Show-RunRecord -Record $record
    }

    foreach ($customer in $customers) {
        $customerName = $customer.Name
        $customerPath = $customer.FullName

        try {
            $periods = @(Get-ChildItem -LiteralPath $customerPath -Directory -Force -ErrorAction Stop | Sort-Object Name)
        }
        catch {
            $record = Add-RunRecord -Customer $customerName -Period '-' -Month '-' -Type 'Access' -Status 'Error' `
                -StatusText '❌ Müşteri klasörü okunamadı' -ZipPath '' -ErrorType 'Müşteri klasörü okunamadı' -ErrorMessage $_.Exception.Message
            Show-RunRecord -Record $record
            continue
        }

        if ($periods.Count -eq 0) {
            $record = Add-RunRecord -Customer $customerName -Period '-' -Month '-' -Type 'Info' -Status 'Info' `
                -StatusText 'ℹ️ Dönem klasörü bulunamadı' -ZipPath ''
            Show-RunRecord -Record $record
            continue
        }

        foreach ($period in $periods) {
            $periodName = $period.Name
            $periodPath = $period.FullName

            try {
                $months = @(Get-ChildItem -LiteralPath $periodPath -Directory -Force -ErrorAction Stop | Sort-Object Name)
            }
            catch {
                $record = Add-RunRecord -Customer $customerName -Period $periodName -Month '-' -Type 'Access' -Status 'Error' `
                    -StatusText '❌ Dönem klasörü okunamadı' -ZipPath '' -ErrorType 'Dönem klasörü okunamadı' -ErrorMessage $_.Exception.Message
                Show-RunRecord -Record $record
                continue
            }

            if ($months.Count -eq 0) {
                $record = Add-RunRecord -Customer $customerName -Period $periodName -Month '-' -Type 'Info' -Status 'Info' `
                    -StatusText 'ℹ️ Ay klasörü bulunamadı' -ZipPath ''
                Show-RunRecord -Record $record
                continue
            }

            foreach ($month in $months) {
                $monthName = $month.Name
                $monthPath = $month.FullName
                $targetPeriodFolder = Join-Path (Join-Path $TargetRoot $customerName) $periodName
                $zipPath = Join-Path $targetPeriodFolder ($monthName + '.zip')

                # Hedef dönem klasörü yalnızca gerekliyse oluşturulur.
                if (-not (Test-FolderExists -Path $targetPeriodFolder)) {
                    if (-not (Ensure-Folder -Path $targetPeriodFolder)) {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Error' `
                            -StatusText '❌ Hedef klasörü oluşturulamadı' -ZipPath $zipPath -ErrorType 'Hedef klasörü oluşturulamadı' `
                            -ErrorMessage 'Hedef dönem klasörü oluşturulamadı.'
                        Show-RunRecord -Record $record
                        continue
                    }
                }

                # Mevcut ZIP'e hiçbir şekilde dokunma.
                if (Test-FileExists -Path $zipPath) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Existing' `
                        -StatusText '⏭️ ZIP mevcut — dokunulmadı' -ZipPath $zipPath
                    Show-RunRecord -Record $record
                    continue
                }

                $fileCountResult = Get-SourceFileCount -Folder $monthPath
                if (-not $fileCountResult.Success) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Error' `
                        -StatusText '❌ Kaynak ay okunamadı — ZIP oluşturulmadı' -ZipPath $zipPath -ErrorType 'Kaynak ay okunamadı' `
                        -ErrorMessage $fileCountResult.ErrorMessage
                    Show-RunRecord -Record $record
                    continue
                }

                if ($fileCountResult.Count -eq 0) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Empty' `
                        -StatusText '⚪ Boş ay — ZIP oluşturulmadı' -ZipPath $zipPath -FileCount 0
                    Show-RunRecord -Record $record
                    continue
                }

                $zipResult = Create-NewZip -SourceFolder $monthPath -DestinationZip $zipPath

                if ($zipResult.RaceDetected) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Race' `
                        -StatusText '⚠️ Hedefte nesne oluştu — dokunulmadı' -ZipPath $zipPath -FileCount $fileCountResult.Count `
                        -ErrorMessage $zipResult.Message
                    Show-RunRecord -Record $record
                    continue
                }

                if (-not $zipResult.Success) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'Error' `
                        -StatusText '❌ ZIP oluşturulamadı / doğrulanamadı' -ZipPath $zipPath -FileCount $fileCountResult.Count `
                        -ErrorType 'ZIP oluşturulamadı / doğrulanamadı' -ErrorMessage $zipResult.Message
                    Show-RunRecord -Record $record
                    continue
                }

                $zipSize = 0
                try { $zipSize = (Get-Item -LiteralPath $zipPath -ErrorAction Stop).Length } catch {}

                $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'Month' -Status 'New' `
                    -StatusText '🆕 ZIP oluşturuldu ve doğrulandı' -ZipPath $zipPath -FileCount $fileCountResult.Count -ZipSize $zipSize
                Show-RunRecord -Record $record
            }
        }
    }

    Process-SpecialFolderForZip -SourceRoot $SourceRoot -TargetRoot $TargetRoot

    $summary = Get-ZipRunSummary -Records $global:RunRecords
    return (Show-ZipCompletionScreen -Source $SourceRoot -Target $TargetRoot -Summary $summary)
}

# ============================================================
# ZIPLEME SONUÇ EKRANI
# ============================================================

function Show-ZipCompletionScreen {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Target,
        [Parameter(Mandatory=$true)][PSCustomObject]$Summary
    )

    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host '✅ ZIPLEME TAMAMLANDI' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ''
    Write-Host ('📦 Yeni ZIP''ler       : {0} müşteri / {1} ay' -f $Summary.NewCustomerCount,$Summary.NewCount) -ForegroundColor Green
    Write-Host ('⏭️ Mevcut ZIP''ler     : {0} müşteri / {1} ay' -f $Summary.ExistingCustomerCount,$Summary.ExistingCount) -ForegroundColor DarkGray
    Write-Host ('⚪ Boş ay             : {0} (ZIP oluşturulmadı)' -f $Summary.EmptyCount) -ForegroundColor Cyan

    $specialText = 'Yok'
    if ($Summary.SpecialNew -gt 0) { $specialText = 'oluşturuldu' }
    elseif ($Summary.SpecialExisting -gt 0) { $specialText = 'Mevcut' }
    Write-Host ('🗑️ .silinenler        : {0}' -f $specialText) -ForegroundColor DarkGray

    Write-Host ('⚠️ Hedef çakışması    : {0}' -f $Summary.RaceCount) -ForegroundColor Yellow
    Write-Host ''

    if ($Summary.ErrorCount -gt 0) {
        Write-Host '❌ HATA' -ForegroundColor Red
        Write-Host ('   Toplam                         : {0}' -f $Summary.ErrorCount) -ForegroundColor Red
        Write-Host ('   ├─ Hedef klasörü oluşturulamadı : {0}' -f $Summary.ErrorTargetFolderCount) -ForegroundColor Red
        Write-Host ('   ├─ Müşteri klasörü okunamadı    : {0}' -f $Summary.ErrorCustomerFolderCount) -ForegroundColor Red
        Write-Host ('   ├─ Dönem klasörü okunamadı      : {0}' -f $Summary.ErrorPeriodFolderCount) -ForegroundColor Red
        Write-Host ('   ├─ Kaynak ay okunamadı          : {0}' -f $Summary.ErrorSourceMonthCount) -ForegroundColor Red
        Write-Host ('   └─ ZIP oluşturulamadı/doğrulanamadı : {0}' -f $Summary.ErrorZipProcessCount) -ForegroundColor Red
    }
    else {
        Write-Host '❌ HATA               : 0' -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host ('📊 Toplam işlem kaydı : {0}' -f $Summary.TotalRecordCount) -ForegroundColor DarkCyan
    Write-Host ''

    Write-Host $Line -ForegroundColor DarkGray
    Write-Host '  Q  ↩️  Ana Menü'
    Write-Host $Line -ForegroundColor DarkGray

    while ($true) {
        $choice = Read-Host 'Seçiminiz'
        if ($choice.ToLower() -eq 'q') { return }
        Write-Host ''
        Write-Host '❌ Geçersiz seçim. Ana menüye dönmek için Q''ya basın.' -ForegroundColor Red
    }
}

# ============================================================
# 2. MODÜL - BÜTÜNLÜK
# ============================================================

function Start-IntegrityModule {
    param(
        [Parameter(Mandatory=$true)][string]$SourceRoot,
        [Parameter(Mandatory=$true)][string]$TargetRoot
    )

    Show-Header

    Initialize-Log
    $pre = Test-ModulePreconditions -SourceRoot $SourceRoot -TargetRoot $TargetRoot -CreateTargetIfMissing $false
    if (-not $pre.Success) {
        Write-Host ''
        Write-Host $pre.Message -ForegroundColor Red
        Pause-Program
        return
    }

    $global:RunRecords.Clear()

    Write-Host ''
    Write-Host ('📁 Kaynak : {0}' -f $SourceRoot)
    Write-Host ('💾 Hedef  : {0}' -f $TargetRoot)
    Write-Host ''
    Write-Host '🔎 Mevcut ZIP''ler kontrol ediliyor...' -ForegroundColor DarkYellow
    Write-Host ''

    try {
        $customers = @(
            Get-ChildItem -LiteralPath $SourceRoot -Directory -Force -ErrorAction Stop |
            Where-Object { $_.Name -ne $SpecialFolderName } |
            Sort-Object Name
        )
    }
    catch {
        $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'IntegrityAccess' -Status 'IntegrityCustomerError' `
            -StatusText '❌ Müşteri klasörleri okunamadı' -ZipPath '' -ErrorMessage $_.Exception.Message
        Show-RunRecord -Record $record
        $summary = Get-IntegritySummary -Records $global:RunRecords
        return (Show-IntegrityCompletionScreen -Source $SourceRoot -Target $TargetRoot -Summary $summary)
    }

    foreach ($customer in $customers) {
        $customerName = $customer.Name
        $customerPath = $customer.FullName

        try {
            $periods = @(Get-ChildItem -LiteralPath $customerPath -Directory -Force -ErrorAction Stop | Sort-Object Name)
        }
        catch {
            $record = Add-RunRecord -Customer $customerName -Period '-' -Month '-' -Type 'IntegrityAccess' -Status 'IntegrityCustomerError' `
                -StatusText '❌ Müşteri klasörü okunamadı' -ZipPath '' -ErrorMessage $_.Exception.Message
            Show-RunRecord -Record $record
            continue
        }

        foreach ($period in $periods) {
            $periodName = $period.Name
            $periodPath = $period.FullName

            try {
                $months = @(Get-ChildItem -LiteralPath $periodPath -Directory -Force -ErrorAction Stop | Sort-Object Name)
            }
            catch {
                $record = Add-RunRecord -Customer $customerName -Period $periodName -Month '-' -Type 'IntegrityAccess' -Status 'IntegrityPeriodError' `
                    -StatusText '❌ Dönem klasörü okunamadı' -ZipPath '' -ErrorMessage $_.Exception.Message
                Show-RunRecord -Record $record
                continue
            }

            foreach ($month in $months) {
                $monthName = $month.Name
                $monthPath = $month.FullName
                $targetPeriodFolder = Join-Path (Join-Path $TargetRoot $customerName) $periodName
                $zipPath = Join-Path $targetPeriodFolder ($monthName + '.zip')

                if (-not (Test-FolderExists -Path $targetPeriodFolder)) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityTarget' -Status 'IntegrityTargetError' `
                        -StatusText '❌ Hedef dönem klasörü bulunamadı' -ZipPath $zipPath -ErrorMessage 'Hedef dönem klasörü bulunamadı; kontrol yapılamadı.'
                    Show-RunRecord -Record $record
                    continue
                }

                if (-not (Test-FileExists -Path $zipPath)) {
                    $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'MissingZip' `
                        -StatusText '⚠️ ZIP bulunamadı — oluşturulmadı' -ZipPath $zipPath
                    Show-RunRecord -Record $record
                    continue
                }

                Write-Host ('👤 {0,-18} 📅 {1,-21} 📁 {2,-4} → 🔍 Kontrol ediliyor...' -f $customerName,$periodName,$monthName) -ForegroundColor DarkYellow

                $result = Test-ZipMatchesSource -SourceFolder $monthPath -ZipPath $zipPath

                switch ($result.ResultType) {
                    'IntegrityOK' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityOK' `
                            -StatusText '✅ Bütünlük OK' -ZipPath $zipPath
                    }
                    'SourceError' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegritySourceError' `
                            -StatusText '❌ Kaynak ay okunamadı — ZIP değiştirilmedi' -ZipPath $zipPath -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
                    }
                    'ZipUnreadable' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ ZIP okunamıyor' -ZipPath $zipPath -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
                    }
                    'FileCountMismatch' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ Dosya sayısı uyuşmuyor' -ZipPath $zipPath -ErrorMessage $result.Detail
                    }
                    'PathMismatch' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ Dosya yolu uyuşmuyor' -ZipPath $zipPath -ErrorMessage $result.Detail
                    }
                    'ShaMismatch' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ SHA-256 uyuşmuyor' -ZipPath $zipPath -ErrorMessage $result.Detail
                    }
                    'DuplicateZipPath' {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ ZIP içinde yinelenen dosya yolu' -ZipPath $zipPath -ErrorMessage $result.Detail
                    }
                    default {
                        $record = Add-RunRecord -Customer $customerName -Period $periodName -Month $monthName -Type 'IntegrityMonth' -Status 'IntegrityError' `
                            -StatusText '❌ Bütünlük kontrolü başarısız' -ZipPath $zipPath -ErrorMessage ('{0} {1}' -f $result.Message,$result.Detail)
                    }
                }

                Show-RunRecord -Record $record
            }
        }
    }

    Process-SpecialFolderIntegrity -SourceRoot $SourceRoot -TargetRoot $TargetRoot

    # Hedefte kaynakta karşılığı olmayan ZIP'ler.
    if (Test-FolderExists -Path $TargetRoot) {
        try {
            $targetZipFiles = @(Get-ChildItem -LiteralPath $TargetRoot -Recurse -File -Force -Filter '*.zip' -ErrorAction Stop)

            foreach ($targetZip in $targetZipFiles) {
                if ($targetZip.Name -eq '.silinenler.zip') { continue }

                $relativeTarget = $targetZip.FullName.Substring($TargetRoot.TrimEnd('\').Length).TrimStart('\','/')
                $parts = $relativeTarget.Split('\')

                if ($parts.Count -lt 3) {
                    $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'TargetExtra' -Status 'TargetExtra' `
                        -StatusText '⚠️ Hedefte yapıya uymayan ZIP bulundu' -ZipPath $targetZip.FullName -ErrorMessage $relativeTarget
                    Show-RunRecord -Record $record
                    continue
                }

                $customerNameFromTarget = $parts[0]
                $periodNameFromTarget = $parts[1]
                $monthNameFromTarget = [System.IO.Path]::GetFileNameWithoutExtension($parts[$parts.Count - 1])

                $expectedSourceMonth = Join-Path (Join-Path (Join-Path $SourceRoot $customerNameFromTarget) $periodNameFromTarget) $monthNameFromTarget

                if (-not (Test-FolderExists -Path $expectedSourceMonth)) {
                    $record = Add-RunRecord -Customer $customerNameFromTarget -Period $periodNameFromTarget -Month $monthNameFromTarget -Type 'TargetExtra' -Status 'TargetExtra' `
                        -StatusText '⚠️ Hedefte kaynakta karşılığı olmayan ZIP' -ZipPath $targetZip.FullName -ErrorMessage $relativeTarget
                    Show-RunRecord -Record $record
                }
            }
        }
        catch {
            $record = Add-RunRecord -Customer '-' -Period '-' -Month '-' -Type 'IntegrityTarget' -Status 'IntegrityTargetError' `
                -StatusText '❌ Hedef ZIP listesi okunamadı' -ZipPath '' -ErrorMessage $_.Exception.Message
            Show-RunRecord -Record $record
        }
    }

    $summary = Get-IntegritySummary -Records $global:RunRecords
    return (Show-IntegrityCompletionScreen -Source $SourceRoot -Target $TargetRoot -Summary $summary)
}

# ============================================================
# BÜTÜNLÜK SONUÇ EKRANI
# ============================================================

function Show-IntegrityCompletionScreen {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Target,
        [Parameter(Mandatory=$true)][PSCustomObject]$Summary
    )

    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host '✅ KAYNAK / ZIP KONTROLÜ TAMAMLANDI' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ''
    Write-Host ('✅ Bütünlük OK         : {0}' -f $Summary.IntegrityOkCount) -ForegroundColor Green
    Write-Host ('❌ Bütünlük hatası     : {0}' -f $Summary.IntegrityErrorCount) -ForegroundColor Red
    Write-Host ('⚠️ ZIP bulunamadı      : {0}' -f $Summary.MissingZipCount) -ForegroundColor Yellow
    Write-Host ('⚠️ Hedefte fazlalık ZIP: {0}' -f $Summary.TargetExtraCount) -ForegroundColor Yellow

    $accessErrorCount = $Summary.IntegritySourceErrorCount + $Summary.IntegrityCustomerErrorCount + $Summary.IntegrityPeriodErrorCount + $Summary.IntegrityTargetErrorCount
    if ($accessErrorCount -gt 0) {
        Write-Host ''
        Write-Host '❌ ERİŞİM / OKUMA HATALARI' -ForegroundColor Red
        Write-Host ('   Kaynak ay okunamadı       : {0}' -f $Summary.IntegritySourceErrorCount) -ForegroundColor Red
        Write-Host ('   Müşteri klasörü okunamadı: {0}' -f $Summary.IntegrityCustomerErrorCount) -ForegroundColor Red
        Write-Host ('   Dönem klasörü okunamadı  : {0}' -f $Summary.IntegrityPeriodErrorCount) -ForegroundColor Red
        Write-Host ('   Hedef okunamadı           : {0}' -f $Summary.IntegrityTargetErrorCount) -ForegroundColor Red
    }

    Write-Host ''
    Write-Host ('📊 Toplam işlem kaydı : {0}' -f $Summary.TotalRecordCount) -ForegroundColor DarkCyan
    Write-Host ''

    Write-Host $Line
    Write-Host '  Q  ↩️  Ana Menü'
    Write-Host $Line

    while ($true) {
        $choice = Read-Host 'Seçiminiz'
        if ($choice.ToLower() -eq 'q') { return }
        Write-Host ''
        Write-Host '❌ Geçersiz seçim. Ana menüye dönmek için Q''ya basın.' -ForegroundColor Red
    }
}

# ============================================================
# KURTARMA YARDIMCILARI
# ============================================================

function Read-RecoveryFolderPath {
    param(
        [Parameter(Mandatory=$true)][string]$Label,
        [string]$DefaultPath = '',
        [bool]$MustExist = $false
    )

    while ($true) {
        $prompt = $Label
        if (-not [string]::IsNullOrWhiteSpace($DefaultPath)) {
            $prompt = '{0} [{1}]' -f $Label,$DefaultPath
        }

        $rawPath = Read-Host $prompt
        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            $rawPath = $DefaultPath
        }

        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            Write-Host '❌ Klasör yolu boş bırakılamaz.' -ForegroundColor Red
            continue
        }

        $rawPath = $rawPath.Trim()
        if ($rawPath.Length -ge 2 -and $rawPath.StartsWith('"') -and $rawPath.EndsWith('"')) {
            $rawPath = $rawPath.Substring(1,$rawPath.Length - 2)
        }

        try {
            $fullPath = [System.IO.Path]::GetFullPath($rawPath)
        }
        catch {
            Write-Host ('❌ Geçersiz klasör yolu: {0}' -f $_.Exception.Message) -ForegroundColor Red
            continue
        }

        if (Test-FileExists -Path $fullPath) {
            Write-Host '❌ Girilen yol bir dosya. Klasör yolu girin.' -ForegroundColor Red
            continue
        }

        if ($MustExist -and -not (Test-FolderExists -Path $fullPath)) {
            Write-Host '❌ Klasör bulunamadı. Lütfen mevcut bir klasör girin.' -ForegroundColor Red
            continue
        }

        return $fullPath.TrimEnd('\')
    }
}

function Test-RecoveryPathNoReparse {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [Parameter(Mandatory=$true)][string]$CandidatePath
    )

    try {
        $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\')
        $candidateFull = [System.IO.Path]::GetFullPath($CandidatePath)

        if ($candidateFull -ne $rootFull -and -not $candidateFull.StartsWith(($rootFull + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{ Success=$false; Message='Yol kurtarma hedefinin dışında.' }
        }

        $rootItem = Get-Item -LiteralPath $rootFull -Force -ErrorAction Stop
        if (($rootItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            return [PSCustomObject]@{ Success=$false; Message='Kurtarma hedefi bir bağlantı/junction olamaz.' }
        }

        $relative = $candidateFull.Substring($rootFull.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($relative)) {
            return [PSCustomObject]@{ Success=$true; Message='' }
        }
        $parts = $relative.Split([char]'\')
        $current = $rootFull

        foreach ($part in $parts) {
            if ([string]::IsNullOrWhiteSpace($part)) { continue }
            $current = Join-Path $current $part
            if (Test-PathExists -Path $current) {
                $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    return [PSCustomObject]@{ Success=$false; Message=('Junction/sembolik bağlantı üzerinden yazma engellendi: {0}' -f $current) }
                }
            }
        }

        return [PSCustomObject]@{ Success=$true; Message='' }
    }
    catch {
        return [PSCustomObject]@{ Success=$false; Message=$_.Exception.Message }
    }
}

function Get-RecoveryEntryInfo {
    param(
        [Parameter(Mandatory=$true)][string]$EntryName,
        [Parameter(Mandatory=$true)][string]$ExpectedRoot
    )

    if ([string]::IsNullOrWhiteSpace($EntryName)) {
        throw 'ZIP içinde boş dosya yolu bulundu.'
    }

    $normalized = $EntryName.Replace('/','\')
    $isDirectory = $normalized.EndsWith('\')

    if ($normalized.StartsWith('\') -or $normalized -match '^[A-Za-z]:' -or [System.IO.Path]::IsPathRooted($normalized)) {
        throw ('Mutlak ZIP yolu reddedildi: {0}' -f $EntryName)
    }

    $trimmed = $normalized.TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw ('Geçersiz ZIP yolu: {0}' -f $EntryName)
    }

    $parts = $trimmed.Split([char]'\')
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($part in $parts) {
        if ([string]::IsNullOrWhiteSpace($part) -or $part -eq '.' -or $part -eq '..') {
            throw ('Güvensiz ZIP yol bileşeni: {0}' -f $EntryName)
        }

        if ($part.IndexOf(':') -ge 0 -or $part.EndsWith('.') -or $part.EndsWith(' ') -or $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') {
            throw ('Windows için geçersiz ZIP dosya adı: {0}' -f $EntryName)
        }

        foreach ($badChar in $invalidChars) {
            if ($part.IndexOf($badChar) -ge 0) {
                throw ('Geçersiz karakter içeren ZIP yolu: {0}' -f $EntryName)
            }
        }
    }

    if ($parts[0] -ine $ExpectedRoot) {
        throw ('ZIP içindeki üst klasör beklenen yapıda değil. Beklenen: {0} / Gelen: {1}' -f $ExpectedRoot,$parts[0])
    }
    if (-not $isDirectory -and $parts.Count -lt 2) {
        throw ('ZIP dosyası beklenen üst klasörün altında değil: {0}' -f $EntryName)
    }

    return [PSCustomObject]@{
        RelativePath = ($parts -join '\')
        IsDirectory = $isDirectory
    }
}

function Test-RecoveryArchive {
    param(
        [Parameter(Mandatory=$true)][string]$ZipPath,
        [Parameter(Mandatory=$true)][string]$ExpectedRoot
    )

    $archive = $null
    $fileCount = 0
    $directoryCount = 0
    [long]$totalBytes = 0
    $entryNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)

        foreach ($entry in $archive.Entries) {
            $entryInfo = Get-RecoveryEntryInfo -EntryName $entry.FullName -ExpectedRoot $ExpectedRoot

            if (-not $entryNames.Add($entryInfo.RelativePath)) {
                throw ('ZIP içinde yinelenen dosya/klasör yolu: {0}' -f $entryInfo.RelativePath)
            }

            if ($entryInfo.IsDirectory) {
                $directoryCount++
                continue
            }

            $fileCount++
            $totalBytes += [long]$entry.Length

            $stream = $null
            [long]$readTotal = 0
            try {
                $stream = $entry.Open()
                $buffer = New-Object byte[] 65536
                while (($read = $stream.Read($buffer,0,$buffer.Length)) -gt 0) {
                    $readTotal += $read
                }
            }
            finally {
                if ($null -ne $stream) { $stream.Dispose() }
            }

            if ($readTotal -ne [long]$entry.Length) {
                throw ('ZIP dosyası beklenen boyutta okunamadı: {0}' -f $entry.FullName)
            }
        }

        if ($fileCount -eq 0) {
            throw 'ZIP içinde geri yüklenecek dosya bulunamadı.'
        }

        return [PSCustomObject]@{
            Success = $true
            FileCount = $fileCount
            DirectoryCount = $directoryCount
            TotalBytes = $totalBytes
            Message = ''
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            FileCount = $fileCount
            DirectoryCount = $directoryCount
            TotalBytes = $totalBytes
            Message = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Get-RecoveryArchivePlan {
    param(
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$ZipFile,
        [Parameter(Mandatory=$true)][string]$BackupRoot,
        [Parameter(Mandatory=$true)][string]$RecoveryRoot
    )

    $relative = $ZipFile.FullName.Substring($BackupRoot.TrimEnd('\').Length).TrimStart('\','/')
    $parts = $relative.Split([char]'\')

    if ($relative -ieq '.silinenler.zip') {
        return [PSCustomObject]@{
            Success = $true
            ZipPath = $ZipFile.FullName
            RelativeZipPath = $relative
            VKN = '-'
            Period = '-'
            Month = '-'
            ExpectedRoot = '.silinenler'
            DestinationRoot = $RecoveryRoot
            IsSpecial = $true
            Message = ''
        }
    }

    if ($parts.Count -ne 3) {
        return [PSCustomObject]@{ Success=$false; Message=('ZIP yolu beklenen yapıda değil: {0}' -f $relative) }
    }

    $customerName = $parts[0]
    $periodName = $parts[1]
    $monthName = [System.IO.Path]::GetFileNameWithoutExtension($parts[2])

    if ([string]::IsNullOrWhiteSpace($customerName) -or $customerName -eq '.' -or $customerName -eq '..' -or $customerName -ieq '.silinenler') {
        return [PSCustomObject]@{ Success=$false; Message=('Geçersiz VKN klasör adı: {0}' -f $relative) }
    }

    if ([System.IO.Path]::GetExtension($parts[2]) -ine '.zip' -or $monthName -notmatch '^(0[1-9]|1[0-2])$') {
        return [PSCustomObject]@{ Success=$false; Message=('Ay ZIP adı beklenen yapıda değil: {0}' -f $relative) }
    }

    $periodParts = $periodName.Split('-')
    if ($periodParts.Count -ne 2) {
        return [PSCustomObject]@{ Success=$false; Message=('Dönem klasör adı beklenen yapıda değil: {0}' -f $periodName) }
    }

    try {
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $periodStart = [datetime]::ParseExact($periodParts[0],'dd.MM.yyyy',$culture)
        $periodEnd = [datetime]::ParseExact($periodParts[1],'dd.MM.yyyy',$culture)
        if ($periodEnd -lt $periodStart) { throw 'Dönem bitiş tarihi başlangıç tarihinden önce.' }
    }
    catch {
        return [PSCustomObject]@{ Success=$false; Message=('Geçersiz dönem tarihi: {0}' -f $periodName) }
    }

    $destinationPeriod = Join-Path (Join-Path $RecoveryRoot $customerName) $periodName

    return [PSCustomObject]@{
        Success = $true
        ZipPath = $ZipFile.FullName
        RelativeZipPath = $relative
        VKN = $customerName
        Period = $periodName
        Month = $monthName
        ExpectedRoot = $monthName
        DestinationRoot = $destinationPeriod
        IsSpecial = $false
        Message = ''
    }
}

function Initialize-RecoveryLog {
    param([Parameter(Mandatory=$true)][string]$LogPath)

    try {
        $logDirectory = Split-Path -Parent $LogPath
        [void][System.IO.Directory]::CreateDirectory($logDirectory)
        Set-Content -LiteralPath $LogPath -Value ('[{0}] eDefterZip kurtarma işlemi başladı.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')) -Encoding UTF8
        return $true
    }
    catch {
        return $false
    }
}

function Add-RecoveryLogLine {
    param(
        [string]$LogPath,
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    try {
        Add-Content -LiteralPath $LogPath -Value ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Message) -Encoding UTF8
    }
    catch {}
}

function Test-RecoveryEntryAgainstExistingFile {
    param(
        [Parameter(Mandatory=$true)][System.IO.Compression.ZipArchiveEntry]$Entry,
        [Parameter(Mandatory=$true)][string]$ExistingPath
    )

    $stream = $null
    try {
        if ((Get-Item -LiteralPath $ExistingPath -Force -ErrorAction Stop).Length -ne [long]$Entry.Length) {
            return $false
        }

        $stream = $Entry.Open()
        $entryHash = Get-Sha256FromStream -Stream $stream
        $existingHash = Get-Sha256FromFile -Path $ExistingPath
        return ($entryHash -eq $existingHash)
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Invoke-RecoveryArchive {
    param(
        [Parameter(Mandatory=$true)][PSCustomObject]$Plan,
        [Parameter(Mandatory=$true)][string]$RecoveryRoot,
        [Parameter(Mandatory=$true)][ref]$FileResultsRef,
        [Parameter(Mandatory=$true)][string]$LogPath
    )

    $archive = $null
    $written = 0
    $sameExisting = 0
    $differentExisting = 0
    $fileErrors = 0
    $foldersCreated = 0

    try {
        $baseSafety = Test-RecoveryPathNoReparse -Root $RecoveryRoot -CandidatePath $Plan.DestinationRoot
        if (-not $baseSafety.Success) { throw $baseSafety.Message }
        [void][System.IO.Directory]::CreateDirectory($Plan.DestinationRoot)
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Plan.ZipPath)

        foreach ($entry in $archive.Entries) {
            $tempPath = $null
            $entryPathForLog = $entry.FullName

            try {
                $entryInfo = Get-RecoveryEntryInfo -EntryName $entry.FullName -ExpectedRoot $Plan.ExpectedRoot
                $relativeFsPath = $entryInfo.RelativePath.Replace('\',[System.IO.Path]::DirectorySeparatorChar)
                $destinationPath = [System.IO.Path]::GetFullPath((Join-Path $Plan.DestinationRoot $relativeFsPath))
                $destinationPrefix = [System.IO.Path]::GetFullPath($Plan.DestinationRoot).TrimEnd('\') + '\'

                if (-not $destinationPath.StartsWith($destinationPrefix,[System.StringComparison]::OrdinalIgnoreCase)) {
                    throw 'ZIP dosya yolu kurtarma hedefinin dışına çıkıyor.'
                }

                $pathSafety = Test-RecoveryPathNoReparse -Root $RecoveryRoot -CandidatePath $destinationPath
                if (-not $pathSafety.Success) { throw $pathSafety.Message }

                if ($entryInfo.IsDirectory) {
                    if (Test-FileExists -Path $destinationPath) {
                        $differentExisting++
                        [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                            ZIP = $Plan.RelativeZipPath
                            Dosya = $entry.FullName
                            Durum = 'Klasör yerine dosya mevcut'
                            Aciklama = 'Mevcut nesne değiştirilmedi.'
                        })
                        continue
                    }

                    if (-not (Test-FolderExists -Path $destinationPath)) {
                        [void][System.IO.Directory]::CreateDirectory($destinationPath)
                        $foldersCreated++
                    }
                    continue
                }

                if (Test-PathExists -Path $destinationPath) {
                    if (Test-FolderExists -Path $destinationPath) {
                        $differentExisting++
                        [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                            ZIP = $Plan.RelativeZipPath
                            Dosya = $entry.FullName
                            Durum = 'Dosya yerine klasör mevcut'
                            Aciklama = 'Mevcut klasör değiştirilmedi.'
                        })
                    }
                    else {
                        if (Test-RecoveryEntryAgainstExistingFile -Entry $entry -ExistingPath $destinationPath) {
                            $sameExisting++
                            [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                                ZIP = $Plan.RelativeZipPath
                                Dosya = $entry.FullName
                                Durum = 'Mevcut - içerik aynı'
                                Aciklama = 'Dosya atlandı; üzerine yazılmadı.'
                            })
                        }
                        else {
                            $differentExisting++
                            [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                                ZIP = $Plan.RelativeZipPath
                                Dosya = $entry.FullName
                                Durum = 'Mevcut - içerik farklı'
                                Aciklama = 'Dosya atlandı; mevcut içerik korundu.'
                            })
                        }
                    }
                    continue
                }

                $parentPath = Split-Path -Parent $destinationPath
                if (-not (Test-FolderExists -Path $parentPath)) {
                    [void][System.IO.Directory]::CreateDirectory($parentPath)
                }

                $pathSafety = Test-RecoveryPathNoReparse -Root $RecoveryRoot -CandidatePath $destinationPath
                if (-not $pathSafety.Success) { throw $pathSafety.Message }

                $tempPath = Join-Path $parentPath ('.edfz_restore_' + [guid]::NewGuid().ToString('N') + '.tmp')
                $inputStream = $null
                $outputStream = $null
                [long]$copiedBytes = 0

                try {
                    $inputStream = $entry.Open()
                    $outputStream = [System.IO.File]::Open($tempPath,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None)
                    $buffer = New-Object byte[] 65536

                    while (($read = $inputStream.Read($buffer,0,$buffer.Length)) -gt 0) {
                        $outputStream.Write($buffer,0,$read)
                        $copiedBytes += $read
                    }

                    $outputStream.Flush($true)
                }
                finally {
                    if ($null -ne $outputStream) { $outputStream.Dispose() }
                    if ($null -ne $inputStream) { $inputStream.Dispose() }
                }

                if ($copiedBytes -ne [long]$entry.Length) {
                    throw ('Çıkarılan dosyanın boyutu uyuşmuyor. Beklenen: {0}, Gelen: {1}' -f $entry.Length,$copiedBytes)
                }

                # Arşiv akışının SHA-256 değeri ile diske yazılan geçici dosyayı karşılaştır.
                $verifyStream = $null
                try {
                    $verifyStream = $entry.Open()
                    $entrySha256 = Get-Sha256FromStream -Stream $verifyStream
                }
                finally {
                    if ($null -ne $verifyStream) { $verifyStream.Dispose() }
                }

                $tempSha256 = Get-Sha256FromFile -Path $tempPath
                if ($entrySha256 -ne $tempSha256) {
                    throw 'Çıkarılan dosyanın SHA-256 değeri ZIP içeriğiyle uyuşmuyor.'
                }

                [System.IO.File]::SetLastWriteTime($tempPath,$entry.LastWriteTime.LocalDateTime)

                if (Test-PathExists -Path $destinationPath) {
                    if ((Test-FileExists -Path $destinationPath) -and (Test-RecoveryEntryAgainstExistingFile -Entry $entry -ExistingPath $destinationPath)) {
                        $sameExisting++
                        [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                            ZIP = $Plan.RelativeZipPath
                            Dosya = $entry.FullName
                            Durum = 'Yarışma - mevcut içerik aynı'
                            Aciklama = 'Dosya başka işlem tarafından oluşturuldu; atlandı.'
                        })
                    }
                    else {
                        $differentExisting++
                        [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                            ZIP = $Plan.RelativeZipPath
                            Dosya = $entry.FullName
                            Durum = 'Yarışma - hedefte nesne mevcut'
                            Aciklama = 'Mevcut nesne değiştirilmedi.'
                        })
                    }
                    continue
                }

                [System.IO.File]::Move($tempPath,$destinationPath)
                $tempPath = $null
                $written++

                [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                    ZIP = $Plan.RelativeZipPath
                    Dosya = $entry.FullName
                    Durum = 'Yeni dosya'
                    Aciklama = 'Dosya çıkarıldı; boyut ve SHA-256 doğrulandı; LastWriteTime arşivden uygulandı.'
                })
            }
            catch {
                $fileErrors++
                [void]$FileResultsRef.Value.Add([PSCustomObject]@{
                    ZIP = $Plan.RelativeZipPath
                    Dosya = $entryPathForLog
                    Durum = 'HATA'
                    Aciklama = $_.Exception.Message
                })
                Add-RecoveryLogLine -LogPath $LogPath -Message ('Dosya hatası | ZIP={0} | Dosya={1} | {2}' -f $Plan.RelativeZipPath,$entryPathForLog,$_.Exception.Message)
            }
            finally {
                if ($tempPath -and (Test-FileExists -Path $tempPath)) {
                    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                }
            }
        }

        $status = 'Tamamlandı'
        if ($fileErrors -gt 0) { $status = 'Hatalı kayıt var' }
        elseif ($differentExisting -gt 0) { $status = 'Tamamlandı - mevcut çakışma var' }

        return [PSCustomObject]@{
            ZIP = $Plan.RelativeZipPath
            Durum = $status
            YeniDosya = $written
            MevcutAyni = $sameExisting
            MevcutFarkli = $differentExisting
            Hata = $fileErrors
            OlusturulanKlasor = $foldersCreated
            Aciklama = ''
        }
    }
    catch {
        Add-RecoveryLogLine -LogPath $LogPath -Message ('ZIP hatası | {0} | {1}' -f $Plan.RelativeZipPath,$_.Exception.Message)
        return [PSCustomObject]@{
            ZIP = $Plan.RelativeZipPath
            Durum = 'ZIP HATASI'
            YeniDosya = $written
            MevcutAyni = $sameExisting
            MevcutFarkli = $differentExisting
            Hata = ($fileErrors + 1)
            OlusturulanKlasor = $foldersCreated
            Aciklama = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Start-RecoveryZipRestore {
    param([string]$DefaultBackupSource = '')

    Show-Header
    Write-Host '🛟 ZIP YEDEKLERİNİ GERİ YÜKLE' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Yedek Kaynak: Aylık ZIP dosyalarının bulunduğu ana klasör.'
    Write-Host 'Kurtarma Hedefi: ZIP içeriğinin geri açılacağı ana klasör.'
    Write-Host 'Not: Program ayarlarındaki Kaynak klasör kurtarma hedefi olarak kullanılmaz.' -ForegroundColor Yellow
    Write-Host ''

    $backupRoot = Read-RecoveryFolderPath -Label '📦 Yedek Kaynak' -DefaultPath $DefaultBackupSource -MustExist $true

    $backupItem = Get-Item -LiteralPath $backupRoot -Force
    if (($backupItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        Write-Host '❌ Yedek Kaynak junction/sembolik bağlantı olamaz.' -ForegroundColor Red
        Pause-Program
        return
    }

    $backupLeaf = Split-Path -Leaf $backupRoot
    $backupParent = Split-Path -Parent $backupRoot
    $defaultRecovery = Join-Path $backupParent ($backupLeaf + '_Kurtarilan')
    $recoveryRoot = Read-RecoveryFolderPath -Label '📂 Kurtarma Hedefi' -DefaultPath $defaultRecovery -MustExist $false

    $separation = Test-SourceTargetSeparation -SourceRoot $backupRoot -TargetRoot $recoveryRoot
    if (-not $separation.Success) {
        Write-Host ('❌ {0}' -f $separation.Message) -ForegroundColor Red
        Pause-Program
        return
    }

    Write-Host ''
    Write-Host '📦 YEDEK KAYNAK ZIP DOSYALARI TARANIYOR...' -ForegroundColor Cyan

    try {
        $zipFiles = @(Get-ChildItem -LiteralPath $backupRoot -Recurse -File -Force -Filter '*.zip' -ErrorAction Stop | Sort-Object FullName)
    }
    catch {
        Write-Host ('❌ Yedek Kaynak okunamadı: {0}' -f $_.Exception.Message) -ForegroundColor Red
        Pause-Program
        return
    }

    if ($zipFiles.Count -eq 0) {
        Write-Host '❌ Yedek Kaynak içinde ZIP dosyası bulunamadı.' -ForegroundColor Red
        Pause-Program
        return
    }

    $plans = New-Object 'System.Collections.Generic.List[object]'
    $invalidArchives = New-Object 'System.Collections.Generic.List[object]'
    $logicalZipNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($zipFile in $zipFiles) {
        $plan = Get-RecoveryArchivePlan -ZipFile $zipFile -BackupRoot $backupRoot -RecoveryRoot $recoveryRoot
        if (-not $plan.Success) {
            [void]$invalidArchives.Add([PSCustomObject]@{ ZIP=$zipFile.FullName; Hata=$plan.Message })
            continue
        }

        if (-not $logicalZipNames.Add($plan.RelativeZipPath)) {
            [void]$invalidArchives.Add([PSCustomObject]@{ ZIP=$zipFile.FullName; Hata='Yinelenen ZIP yolu.' })
            continue
        }

        [void]$plans.Add($plan)
    }

    if ($invalidArchives.Count -gt 0) {
        Write-Host ''
        Write-Host '❌ Beklenmeyen veya geçersiz ZIP yolları bulundu. Hiçbir dosya çıkarılmayacak.' -ForegroundColor Red
        foreach ($invalid in $invalidArchives) {
            Write-Host ('  • {0} | {1}' -f $invalid.ZIP,$invalid.Hata) -ForegroundColor Red
        }
        Pause-Program
        return
    }

    # Bütün arşivleri, herhangi bir dosya çıkarmadan önce tam olarak oku ve doğrula.
    $preflight = New-Object 'System.Collections.Generic.List[object]'
    $preflightErrors = 0
    $totalFiles = 0
    [long]$totalBytes = 0
    $zipIndex = 0

    Write-Host ''
    Write-Host ('🔍 {0} ZIP dosyası okunabilirlik ve yol güvenliği kontrolünden geçiriliyor...' -f $plans.Count) -ForegroundColor DarkYellow

    foreach ($plan in $plans) {
        $zipIndex++
        $check = Test-RecoveryArchive -ZipPath $plan.ZipPath -ExpectedRoot $plan.ExpectedRoot

        if ($check.Success) {
            $statusText = 'OK'
            $totalFiles += $check.FileCount
            $totalBytes += [long]$check.TotalBytes
        }
        else {
            $statusText = 'HATA'
            $preflightErrors++
        }

        [void]$preflight.Add([PSCustomObject]@{
            ZIP = $plan.RelativeZipPath
            Durum = $statusText
            DosyaSayisi = $check.FileCount
            AcilmisBoyut = $check.TotalBytes
            Hata = $check.Message
        })

        if (-not $check.Success) {
            Write-Host ('❌ {0} | {1}' -f $plan.RelativeZipPath,$check.Message) -ForegroundColor Red
        }

        if (($zipIndex % 10) -eq 0 -or $zipIndex -eq $plans.Count) {
            Write-Progress -Activity 'ZIP yedekleri ön kontrolden geçiriliyor' -Status ('{0} / {1}' -f $zipIndex,$plans.Count) -PercentComplete (($zipIndex / $plans.Count) * 100)
        }
    }

    Write-Progress -Activity 'ZIP yedekleri ön kontrolden geçiriliyor' -Completed

    $runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportRoot = $ReportFolder
    try { [void][System.IO.Directory]::CreateDirectory($reportRoot) }
    catch { throw ('Kurtarma rapor klasörü oluşturulamadı: {0}. Hata: {1}' -f $reportRoot,$_.Exception.Message) }

    $preflightCsv = Join-Path $reportRoot ('Kurtarma_OnKontrol_' + $runStamp + '.csv')
    $preflight | Export-Csv -LiteralPath $preflightCsv -NoTypeInformation -Encoding UTF8

    if ($preflightErrors -gt 0) {
        Write-Host ''
        Write-Host ('❌ Ön kontrolde {0} ZIP hatası bulundu. Güvenlik gereği hiçbir ZIP çıkarılmadı.' -f $preflightErrors) -ForegroundColor Red
        Write-Host ('Ön kontrol raporu: {0}' -f $preflightCsv)
        Pause-Program
        return
    }

    $totalMB = [Math]::Round(($totalBytes / 1MB),2)
    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host 'KURTARMA ÖN KONTROLÜ TAMAMLANDI' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    $vknCount = @($plans | Where-Object { -not $_.IsSpecial } | Select-Object -ExpandProperty VKN -Unique).Count
    $monthZipCount = @($plans | Where-Object { -not $_.IsSpecial }).Count
    $specialZipCount = @($plans | Where-Object { $_.IsSpecial }).Count
    Write-Host ('📦 Yedek Kaynak ZIP   : {0}' -f $plans.Count)
    Write-Host ('📁 VKN klasörü        : {0}' -f $vknCount)
    Write-Host ('📅 Aylık ZIP          : {0}' -f $monthZipCount)
    Write-Host ('🗑️ .silinenler ZIP    : {0}' -f $specialZipCount)
    Write-Host ('📄 Arşiv dosya sayısı : {0}' -f $totalFiles)
    Write-Host ('💾 Açılacak veri      : {0} MB' -f $totalMB)
    Write-Host ''
    Write-Host ('📦 Yedek Kaynak   : {0}' -f $backupRoot)
    Write-Host ('📂 Kurtarma Hedefi: {0}' -f $recoveryRoot)
    Write-Host ''
    Write-Host 'Güvenlik: Mevcut dosyaların üzerine yazılmaz.' -ForegroundColor Green
    Write-Host 'Aynı içerikteki mevcut dosyalar atlanır; farklı içerikteki dosyalar korunur ve raporlanır.' -ForegroundColor Yellow
    Write-Host 'İşlem kesilirse yeniden çalıştırıldığında mevcut dosyalar üzerine yazılmadan devam edilebilir.' -ForegroundColor Green

    if (Test-FolderExists -Path $recoveryRoot) {
        Write-Host 'Kurtarma Hedefi mevcut; işlem birleştirme şeklinde çalışacak.' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Kurtarma Hedefi işlem onayından sonra oluşturulacak.' -ForegroundColor Yellow
    }

    $confirm = (Read-Host 'Kurtarma başlatılsın mı? (E/H)').Trim().ToUpperInvariant()
    if ($confirm -ne 'E') {
        Write-Host 'Kurtarma işlemi iptal edildi. Dosyalara dokunulmadı.' -ForegroundColor Yellow
        Pause-Program
        return
    }

    try {
        [void][System.IO.Directory]::CreateDirectory($recoveryRoot)
    }
    catch {
        Write-Host ('❌ Kurtarma Hedefi oluşturulamadı: {0}' -f $_.Exception.Message) -ForegroundColor Red
        Pause-Program
        return
    }

    $rootItem = Get-Item -LiteralPath $recoveryRoot -Force
    if (($rootItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        Write-Host '❌ Kurtarma Hedefi junction/sembolik bağlantı olamaz. İşlem durduruldu.' -ForegroundColor Red
        Pause-Program
        return
    }

    $script:LastRecoveryTarget = $recoveryRoot
    $logPath = Join-Path $reportRoot ('EdefterZip_Kurtarma_' + $runStamp + '.txt')
    $logReady = Initialize-RecoveryLog -LogPath $logPath
    if (-not $logReady) {
        Write-Host '⚠️ Kurtarma TXT logu oluşturulamadı; CSV raporları yine kaydedilecek.' -ForegroundColor Yellow
    }

    Add-RecoveryLogLine -LogPath $logPath -Message ('Yedek Kaynak={0} | Kurtarma Hedefi={1} | ZIP={2} | Dosya={3}' -f $backupRoot,$recoveryRoot,$plans.Count,$totalFiles)

    $zipResults = New-Object 'System.Collections.Generic.List[object]'
    $fileResults = New-Object 'System.Collections.Generic.List[object]'
    $zipIndex = 0
    $sumWritten = 0
    $sumSame = 0
    $sumDifferent = 0
    $sumFileErrors = 0
    $sumFolders = 0

    Write-Host ''
    Write-Host '📂 KURTARMA BAŞLADI...' -ForegroundColor Cyan

    foreach ($plan in $plans) {
        $zipIndex++
        $result = Invoke-RecoveryArchive -Plan $plan -RecoveryRoot $recoveryRoot -FileResultsRef ([ref]$fileResults) -LogPath $logPath
        [void]$zipResults.Add($result)

        $sumWritten += $result.YeniDosya
        $sumSame += $result.MevcutAyni
        $sumDifferent += $result.MevcutFarkli
        $sumFileErrors += $result.Hata
        $sumFolders += $result.OlusturulanKlasor

        $color = 'Green'
        if ($result.Hata -gt 0) { $color = 'Red' }
        elseif ($result.MevcutFarkli -gt 0) { $color = 'Yellow' }

        Write-Host ('[{0}/{1}] {2} | Yeni: {3} | Aynı mevcut: {4} | Farklı mevcut: {5} | Hata: {6}' -f `
            $zipIndex,$plans.Count,$plan.RelativeZipPath,$result.YeniDosya,$result.MevcutAyni,$result.MevcutFarkli,$result.Hata) -ForegroundColor $color

        Add-RecoveryLogLine -LogPath $logPath -Message ('ZIP={0} | Durum={1} | Yeni={2} | MevcutAyni={3} | MevcutFarkli={4} | Hata={5}' -f `
            $result.ZIP,$result.Durum,$result.YeniDosya,$result.MevcutAyni,$result.MevcutFarkli,$result.Hata)

        if (($zipIndex % 5) -eq 0 -or $zipIndex -eq $plans.Count) {
            Write-Progress -Activity 'ZIP yedekleri geri yükleniyor' -Status ('{0} / {1}' -f $zipIndex,$plans.Count) -PercentComplete (($zipIndex / $plans.Count) * 100)
        }
    }

    Write-Progress -Activity 'ZIP yedekleri geri yükleniyor' -Completed

    # .silinenler kaynakta gizli özel klasördür; kurtarılan kopyada da gizli niteliğini koru.
    $restoredSpecialFolder = Join-Path $recoveryRoot $SpecialFolderName
    if (Test-FolderExists -Path $restoredSpecialFolder) {
        try {
            $specialItem = Get-Item -LiteralPath $restoredSpecialFolder -Force
            if (($specialItem.Attributes -band [System.IO.FileAttributes]::Hidden) -eq 0) {
                $specialItem.Attributes = ($specialItem.Attributes -bor [System.IO.FileAttributes]::Hidden)
                Add-RecoveryLogLine -LogPath $logPath -Message '.silinenler klasörüne Hidden niteliği uygulandı.'
            }
        }
        catch {
            Add-RecoveryLogLine -LogPath $logPath -Message ('.silinenler Hidden niteliği uygulanamadı: {0}' -f $_.Exception.Message)
        }
    }

    $zipCsv = Join-Path $reportRoot ('Kurtarma_ZIP_Sonuclari_' + $runStamp + '.csv')
    $fileCsv = Join-Path $reportRoot ('Kurtarma_Dosya_Detaylari_' + $runStamp + '.csv')
    $zipResults | Export-Csv -LiteralPath $zipCsv -NoTypeInformation -Encoding UTF8
    $fileResults | Export-Csv -LiteralPath $fileCsv -NoTypeInformation -Encoding UTF8

    Add-RecoveryLogLine -LogPath $logPath -Message ('BİTTİ | Yeni={0} | MevcutAyni={1} | MevcutFarkli={2} | Hata={3} | Klasor={4}' -f `
        $sumWritten,$sumSame,$sumDifferent,$sumFileErrors,$sumFolders)

    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host 'KURTARMA İŞLEMİ TAMAMLANDI' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ('📦 İşlenen ZIP           : {0}' -f $plans.Count)
    Write-Host ('📄 Yeni çıkarılan dosya  : {0}' -f $sumWritten) -ForegroundColor Green
    Write-Host ('✅ Mevcut - içerik aynı  : {0}' -f $sumSame) -ForegroundColor DarkGray
    Write-Host ('⚠️ Mevcut - içerik farklı: {0}' -f $sumDifferent) -ForegroundColor Yellow
    Write-Host ('📁 Oluşturulan klasör    : {0}' -f $sumFolders)
    Write-Host ('❌ Dosya/ZIP hatası      : {0}' -f $sumFileErrors) -ForegroundColor $(if ($sumFileErrors -eq 0) { 'Green' } else { 'Red' })
    Write-Host ''
    Write-Host ('Kurtarma Hedefi: {0}' -f $recoveryRoot)
    Write-Host ('Ön kontrol CSV : {0}' -f $preflightCsv)
    Write-Host ('ZIP sonuç CSV  : {0}' -f $zipCsv)
    Write-Host ('Dosya detay CSV: {0}' -f $fileCsv)
    if ($logReady) { Write-Host ('TXT log        : {0}' -f $logPath) }
    Write-Host ''
    Write-Host 'Not: Dosya LastWriteTime değerleri ZIP içinden uygulanır. Klasör tarihleri için alt menüdeki tarih düzenleme işlemini çalıştırın.' -ForegroundColor Yellow
    Pause-Program
}

# ============================================================
# KURTARMA - KLASÖR TARİHLERİNİ YENİDEN OLUŞTUR
# ============================================================

function ConvertTo-RecoveryPeriodDates {
    param([Parameter(Mandatory=$true)][string]$PeriodName)

    $parts = $PeriodName.Split('-')
    if ($parts.Count -ne 2) { return $null }

    try {
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $startDate = [datetime]::ParseExact($parts[0],'dd.MM.yyyy',$culture)
        $endDate = [datetime]::ParseExact($parts[1],'dd.MM.yyyy',$culture)
        if ($endDate -lt $startDate) { return $null }
        return [PSCustomObject]@{ StartDate=$startDate; EndDate=$endDate }
    }
    catch {
        return $null
    }
}

function Start-RecoveryTimestampModule {
    $defaultPath = ''
    if (-not [string]::IsNullOrWhiteSpace($script:LastRecoveryTarget)) {
        $defaultPath = $script:LastRecoveryTarget
    }

    Show-Header
    Write-Host '🕒 KLASÖR TARİHLERİNİ YENİDEN OLUŞTUR' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Bu işlem yalnızca Kurtarma Hedefi klasörlerinin tarihlerini düzenler.'
    Write-Host 'Ay: En yeni GİB önekli dosya tarihi.'
    Write-Host 'Dönem: Son dolu ay içindeki en yeni GİB dosya tarihi.'
    Write-Host 'VKN: Kronolojik son dönemin son dolu ayındaki en yeni GİB dosya tarihi.'
    Write-Host '.silinenler: İçindeki en yeni dosya tarihi.'
    Write-Host ''

    $recoveryRoot = Read-RecoveryFolderPath -Label '📂 Kurtarma Hedefi' -DefaultPath $defaultPath -MustExist $true
    $rootItem = Get-Item -LiteralPath $recoveryRoot -Force
    if (($rootItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        Write-Host '❌ Kurtarma Hedefi junction/sembolik bağlantı olamaz.' -ForegroundColor Red
        Pause-Program
        return
    }

    $monthPlans = New-Object 'System.Collections.Generic.List[object]'
    $periodPlans = New-Object 'System.Collections.Generic.List[object]'
    $customerPlans = New-Object 'System.Collections.Generic.List[object]'
    $specialPlans = New-Object 'System.Collections.Generic.List[object]'
    $skippedMonths = New-Object 'System.Collections.Generic.List[object]'
    $skippedPeriods = New-Object 'System.Collections.Generic.List[object]'
    $skippedCustomers = New-Object 'System.Collections.Generic.List[object]'
    $emptyMonths = 0

    Write-Host ''
    Write-Host '🔎 Klasör ve GİB dosyaları taranıyor...' -ForegroundColor DarkYellow

    try {
        $customers = @(Get-ChildItem -LiteralPath $recoveryRoot -Directory -Force -ErrorAction Stop | Where-Object { $_.Name -ne $SpecialFolderName } | Sort-Object Name)
    }
    catch {
        Write-Host ('❌ Kurtarma Hedefi okunamadı: {0}' -f $_.Exception.Message) -ForegroundColor Red
        Pause-Program
        return
    }

    foreach ($customer in $customers) {
        $periodInfos = New-Object 'System.Collections.Generic.List[object]'
        $periods = @(Get-ChildItem -LiteralPath $customer.FullName -Directory -Force -ErrorAction Stop | Sort-Object Name)

        foreach ($period in $periods) {
            $periodDates = ConvertTo-RecoveryPeriodDates -PeriodName $period.Name
            if ($null -eq $periodDates) { continue }

            $monthInfos = New-Object 'System.Collections.Generic.List[object]'
            $months = @(Get-ChildItem -LiteralPath $period.FullName -Directory -Force -ErrorAction Stop | Where-Object { $_.Name -match '^(0[1-9]|1[0-2])$' } | Sort-Object Name)

            foreach ($month in $months) {
                $monthFiles = @(Get-ChildItem -LiteralPath $month.FullName -File -Force -Recurse -ErrorAction Stop)
                if ($monthFiles.Count -eq 0) {
                    $emptyMonths++
                    continue
                }

                $gibFiles = @($monthFiles | Where-Object { $_.Name -match '^[gG][iIİı][bB]' })
                if ($gibFiles.Count -eq 0) {
                    [void]$skippedMonths.Add([PSCustomObject]@{ VKN=$customer.Name; Donem=$period.Name; Ay=$month.Name; Path=$month.FullName; Neden='Dosya var ancak GİB önekli dosya yok.' })
                    [void]$monthInfos.Add([PSCustomObject]@{ MonthNumber=[int]$month.Name; FileCount=$monthFiles.Count; GibFiles=@(); TargetUtc=$null; BasisFile='' })
                    continue
                }

                $latestGib = $gibFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
                [void]$monthPlans.Add([PSCustomObject]@{
                    Seviye='Ay'; VKN=$customer.Name; Donem=$period.Name; Ay=$month.Name
                    Path=$month.FullName; TargetUtc=$latestGib.LastWriteTimeUtc; BasisFile=$latestGib.FullName
                    FileCount=$monthFiles.Count; GibFileCount=$gibFiles.Count
                })

                [void]$monthInfos.Add([PSCustomObject]@{
                    MonthNumber=[int]$month.Name; FileCount=$monthFiles.Count; GibFiles=$gibFiles
                    TargetUtc=$latestGib.LastWriteTimeUtc; BasisFile=$latestGib.FullName
                })
            }

            $nonEmptyMonths = @($monthInfos | Where-Object { $_.FileCount -gt 0 } | Sort-Object MonthNumber -Descending)
            $periodTarget = $null
            $periodBasis = ''
            $lastNonEmptyMonth = $null

            if ($nonEmptyMonths.Count -gt 0) {
                $lastNonEmptyMonth = $nonEmptyMonths[0]
                if ($lastNonEmptyMonth.GibFiles.Count -gt 0) {
                    $periodGibLatest = $lastNonEmptyMonth.GibFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
                    $periodTarget = $periodGibLatest.LastWriteTimeUtc
                    $periodBasis = $periodGibLatest.FullName

                    [void]$periodPlans.Add([PSCustomObject]@{
                        Seviye='Dönem'; VKN=$customer.Name; Donem=$period.Name; Ay=$lastNonEmptyMonth.MonthNumber.ToString('00')
                        Path=$period.FullName; TargetUtc=$periodTarget; BasisFile=$periodBasis
                        FileCount=$lastNonEmptyMonth.FileCount; GibFileCount=$lastNonEmptyMonth.GibFiles.Count
                    })
                }
                else {
                    [void]$skippedPeriods.Add([PSCustomObject]@{ VKN=$customer.Name; Donem=$period.Name; Path=$period.FullName; Neden='Son dolu ayda GİB önekli dosya yok.' })
                }
            }

            [void]$periodInfos.Add([PSCustomObject]@{
                PeriodName=$period.Name; PeriodEnd=$periodDates.EndDate; HasFiles=($nonEmptyMonths.Count -gt 0)
                TargetUtc=$periodTarget; BasisFile=$periodBasis; Path=$period.FullName
            })
        }

        $periodsWithFiles = @($periodInfos | Where-Object { $_.HasFiles } | Sort-Object PeriodEnd -Descending)
        if ($periodsWithFiles.Count -gt 0) {
            $latestPeriod = $periodsWithFiles[0]
            if ($null -ne $latestPeriod.TargetUtc) {
                [void]$customerPlans.Add([PSCustomObject]@{
                    Seviye='VKN'; VKN=$customer.Name; Donem=$latestPeriod.PeriodName; Ay='-'
                    Path=$customer.FullName; TargetUtc=$latestPeriod.TargetUtc; BasisFile=$latestPeriod.BasisFile
                    FileCount=0; GibFileCount=0
                })
            }
            else {
                [void]$skippedCustomers.Add([PSCustomObject]@{ VKN=$customer.Name; Path=$customer.FullName; Neden='Kronolojik son dönemin son dolu ayında GİB önekli dosya yok.' })
            }
        }
    }

    $specialFolder = Join-Path $recoveryRoot $SpecialFolderName
    if (Test-FolderExists -Path $specialFolder) {
        $specialFiles = @(Get-ChildItem -LiteralPath $specialFolder -File -Force -Recurse -ErrorAction Stop)
        if ($specialFiles.Count -gt 0) {
            $latestSpecial = $specialFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
            [void]$specialPlans.Add([PSCustomObject]@{
                Seviye='.silinenler'; VKN='-'; Donem='-'; Ay='-'; Path=$specialFolder
                TargetUtc=$latestSpecial.LastWriteTimeUtc; BasisFile=$latestSpecial.FullName
                FileCount=$specialFiles.Count; GibFileCount=0
            })
        }
    }

    $orderedPlans = @()
    $orderedPlans += @($monthPlans.ToArray())
    $orderedPlans += @($periodPlans.ToArray())
    $orderedPlans += @($customerPlans.ToArray())
    $orderedPlans += @($specialPlans.ToArray())

    if ($orderedPlans.Count -eq 0) {
        Write-Host '❌ Uygulanabilecek tarih planı bulunamadı.' -ForegroundColor Red
        Pause-Program
        return
    }

    $uniquePaths = @($orderedPlans | Select-Object -ExpandProperty Path -Unique)
    if ($uniquePaths.Count -ne $orderedPlans.Count) {
        Write-Host '❌ Plan içerisinde yinelenen klasör yolu bulundu. İşlem iptal edildi.' -ForegroundColor Red
        Pause-Program
        return
    }

    foreach ($plan in $orderedPlans) {
        $fullPath = [System.IO.Path]::GetFullPath($plan.Path)
        if (-not $fullPath.StartsWith(($recoveryRoot.TrimEnd('\') + '\'),[System.StringComparison]::OrdinalIgnoreCase)) {
            throw ('Plan Kurtarma Hedefi dışında yol içeriyor: {0}' -f $fullPath)
        }
        if (-not (Test-FolderExists -Path $fullPath) -or -not (Test-FileExists -Path $plan.BasisFile)) {
            throw ('Plan klasörü veya tarih kaynağı dosyası bulunamadı: {0}' -f $fullPath)
        }
        $pathSafety = Test-RecoveryPathNoReparse -Root $recoveryRoot -CandidatePath $fullPath
        if (-not $pathSafety.Success) { throw $pathSafety.Message }
        $basis = Get-Item -LiteralPath $plan.BasisFile -Force
        if ($basis.LastWriteTimeUtc.Ticks -ne $plan.TargetUtc.Ticks) {
            throw ('Tarih kaynağı tarama sırasında değişti: {0}' -f $plan.BasisFile)
        }
    }

    $runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportRoot = $ReportFolder
    try { [void][System.IO.Directory]::CreateDirectory($reportRoot) }
    catch { throw ('Kurtarma rapor klasörü oluşturulamadı: {0}. Hata: {1}' -f $reportRoot,$_.Exception.Message) }

    $beforeCsv = Join-Path $reportRoot ('Klasor_Tarih_Oncesi_' + $runStamp + '.csv')
    $afterCsv = Join-Path $reportRoot ('Klasor_Tarih_Sonrasi_' + $runStamp + '.csv')

    $beforeRecords = foreach ($plan in $orderedPlans) {
        $item = Get-Item -LiteralPath $plan.Path -Force
        [PSCustomObject]@{
            Seviye=$plan.Seviye; VKN=$plan.VKN; Donem=$plan.Donem; Ay=$plan.Ay; Yol=$plan.Path
            CreationTimeOnce=$item.CreationTime.ToString('dd.MM.yyyy HH:mm:ss')
            LastWriteTimeOnce=$item.LastWriteTime.ToString('dd.MM.yyyy HH:mm:ss')
            UygulanacakTarih=$plan.TargetUtc.ToLocalTime().ToString('dd.MM.yyyy HH:mm:ss')
            TarihKaynakDosyasi=$plan.BasisFile
        }
    }
    $beforeRecords | Export-Csv -LiteralPath $beforeCsv -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host 'KLASÖR TARİH PLANI' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ('Ay klasörü       : {0}' -f $monthPlans.Count)
    Write-Host ('Dönem klasörü    : {0}' -f $periodPlans.Count)
    Write-Host ('VKN klasörü      : {0}' -f $customerPlans.Count)
    Write-Host ('.silinenler      : {0}' -f $specialPlans.Count)
    Write-Host ('Toplam plan      : {0}' -f $orderedPlans.Count)
    Write-Host ('GİB dosyası yok  : {0} ay atlandı' -f $skippedMonths.Count) -ForegroundColor Yellow
    Write-Host ('Boş ay           : {0}' -f $emptyMonths)
    Write-Host ('Dönem atlandı    : {0}' -f $skippedPeriods.Count)
    Write-Host ('VKN atlandı      : {0}' -f $skippedCustomers.Count)
    Write-Host ('Kurtarma Hedefi  : {0}' -f $recoveryRoot)
    Write-Host ''
    Write-Host 'Uygulama sırası: Ay → Dönem → VKN → .silinenler' -ForegroundColor Green
    Write-Host 'Dosyaların tarihleri ve içerikleri değiştirilmeyecek.' -ForegroundColor Green
    Write-Host ('Önceki tarihler CSV: {0}' -f $beforeCsv)

    $confirm = (Read-Host 'Klasör tarihleri uygulansın mı? (E/H)').Trim().ToUpperInvariant()
    if ($confirm -ne 'E') {
        Write-Host 'İşlem iptal edildi. Klasör tarihleri değiştirilmedi.' -ForegroundColor Yellow
        Pause-Program
        return
    }

    $results = New-Object 'System.Collections.Generic.List[object]'
    $done = 0
    $success = 0
    $failed = 0

    foreach ($plan in $orderedPlans) {
        $done++
        try {
            [System.IO.Directory]::SetCreationTimeUtc($plan.Path,$plan.TargetUtc)
            [System.IO.Directory]::SetLastWriteTimeUtc($plan.Path,$plan.TargetUtc)

            $check = Get-Item -LiteralPath $plan.Path -Force
            $check.Refresh()
            $creationOk = ($check.CreationTimeUtc.Ticks -eq $plan.TargetUtc.Ticks)
            $writeOk = ($check.LastWriteTimeUtc.Ticks -eq $plan.TargetUtc.Ticks)
            $status = 'OK'

            if ($creationOk -and $writeOk) { $success++ }
            else { $failed++; $status='FARK VAR' }

            [void]$results.Add([PSCustomObject]@{
                Seviye=$plan.Seviye; Yol=$plan.Path
                UygulananTarih=$plan.TargetUtc.ToLocalTime().ToString('dd.MM.yyyy HH:mm:ss')
                CreationTimeOK=$creationOk; LastWriteTimeOK=$writeOk; Sonuc=$status
            })
        }
        catch {
            $failed++
            [void]$results.Add([PSCustomObject]@{
                Seviye=$plan.Seviye; Yol=$plan.Path; UygulananTarih=$plan.TargetUtc.ToLocalTime().ToString('dd.MM.yyyy HH:mm:ss')
                CreationTimeOK=$false; LastWriteTimeOK=$false; Sonuc=('HATA: ' + $_.Exception.Message)
            })
        }

        if (($done % 25) -eq 0 -or $done -eq $orderedPlans.Count) {
            Write-Progress -Activity 'Klasör tarihleri yeniden oluşturuluyor' -Status ('{0} / {1}' -f $done,$orderedPlans.Count) -PercentComplete (($done / $orderedPlans.Count) * 100)
        }
    }

    Write-Progress -Activity 'Klasör tarihleri yeniden oluşturuluyor' -Completed
    $results | Export-Csv -LiteralPath $afterCsv -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host $Line -ForegroundColor Cyan
    Write-Host 'KLASÖR TARİHİ DÜZENLEME SONUCU' -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ('Toplam plan          : {0}' -f $orderedPlans.Count)
    Write-Host ('Başarılı doğrulanan  : {0}' -f $success) -ForegroundColor Green
    Write-Host ('Hatalı / farklı      : {0}' -f $failed) -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Red' })
    Write-Host ('Önceki tarihler CSV  : {0}' -f $beforeCsv)
    Write-Host ('Sonuç CSV            : {0}' -f $afterCsv)

    if ($success -eq $orderedPlans.Count -and $failed -eq 0) {
        Write-Host '✅ Bütün klasör tarihleri uygulandı ve tekrar okunarak doğrulandı.' -ForegroundColor Green
    }
    else {
        Write-Host '⚠️ Bazı klasörlerin tarihleri doğrulanamadı. Sonuç CSV dosyasını inceleyin.' -ForegroundColor Yellow
    }

    $script:LastRecoveryTarget = $recoveryRoot
    Pause-Program
}

function Start-RecoveryMenu {
    param([string]$DefaultBackupSource = '')

    while ($true) {
        Show-Header
        Write-Host '🛟 E-DEFTER KURTARMA MENÜSÜ' -ForegroundColor Cyan
        Write-Host $Line -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  1. 📦 ZIP Yedeklerini Geri Yükle'
        Write-Host '  2. 🕒 Klasör Tarihlerini Yeniden Oluştur'
        Write-Host '  Q. ↩️ Ana Menü'
        Write-Host $Line -ForegroundColor DarkGray
        Write-Host ''

        $choice = (Read-Host 'Seçiminiz').Trim().ToUpperInvariant()

        switch ($choice) {
            '1' {
                try { Start-RecoveryZipRestore -DefaultBackupSource $DefaultBackupSource }
                catch { Write-Host ('❌ Kurtarma modülünde beklenmeyen hata: {0}' -f $_.Exception.Message) -ForegroundColor Red; Pause-Program }
            }
            '2' {
                try { Start-RecoveryTimestampModule }
                catch { Write-Host ('❌ Tarih düzenleme modülünde beklenmeyen hata: {0}' -f $_.Exception.Message) -ForegroundColor Red; Pause-Program }
            }
            'Q' { return }
            default {
                Write-Host '❌ Geçersiz seçim.' -ForegroundColor Red
                Start-Sleep -Milliseconds 700
            }
        }
    }
}


# ============================================================
# ANA MENÜ
# ============================================================

$currentSettings = Load-Settings

while ($true) {
    Show-Header

    Show-ConfigurationStatus -Settings $currentSettings

    Write-Host ''
    Write-Host '📜 PROGRAM POLİTİKASI' -ForegroundColor Cyan
    Write-Host '  • Kaynak e-Defter dosyaları okunur; kaynakta değişiklik veya silme yapılmaz.' -ForegroundColor DarkGray
    Write-Host '  • ZIP yalnızca hedefte yoksa oluşturulur; mevcut ZIP dosyalarının üzerine yazılmaz.' -ForegroundColor DarkGray
    Write-Host '  • Yeni ZIP geçici dosyada hazırlanır, okunabilirlik ve SHA-256 ile doğrulanır; sonra hedefe alınır.' -ForegroundColor DarkGray
    Write-Host '  • Bütünlük kontrolü salt okunurdur. Boş aylar için ZIP oluşturulmaz; .silinenler ayrı işlenir.' -ForegroundColor DarkGray
    Write-Host '  • Hatalar kayda alınır, diğer kayıtların işlenmesi sürer. Ayarlar ve işlem kayıtları yerel dosyalarda tutulur.' -ForegroundColor DarkGray
    Write-Host '  • Kurtarma menüsü ZIP yedeklerini bağımsız yollarla açar; mevcut dosyaların üzerine yazmaz.' -ForegroundColor DarkGray

    Write-Host ''
    Write-Host $Line -ForegroundColor DarkGray
    Write-Host '  1. 🗜️  E-Defter Zipleme İşlemi Başlat'
    Write-Host '  2. 🔍  E-Defter ZIP Bütünlük Kontrolü'
    Write-Host '  3. 🛟  E-Defter Kurtarma Menüsü'
    Write-Host '  4. ⚙️  Ayarlar'
    Write-Host '  Q  🚪  Çıkış'
    Write-Host $Line -ForegroundColor DarkGray
    Write-Host ''

    $choice = (Read-Host 'Seçiminiz').ToUpper()

    switch ($choice) {
        '1' {
            if ([string]::IsNullOrWhiteSpace($currentSettings.SourceFolder) -or [string]::IsNullOrWhiteSpace($currentSettings.TargetFolder)) {
                Show-Header
                Write-Host ''
                Write-Host '❌ ZIP işlemi başlatılamıyor.' -ForegroundColor Red
                Write-Host ''
                if ([string]::IsNullOrWhiteSpace($currentSettings.SourceFolder)) { Write-Host '  📁 Kaynak klasör ayarlanmamış.' -ForegroundColor Red }
                if ([string]::IsNullOrWhiteSpace($currentSettings.TargetFolder)) { Write-Host '  💾 Hedef klasör ayarlanmamış.' -ForegroundColor Red }
                Pause-Program
                continue
            }

            Start-EdefterZipProcess -SourceRoot $currentSettings.SourceFolder -TargetRoot $currentSettings.TargetFolder
        }

        '2' {
            if ([string]::IsNullOrWhiteSpace($currentSettings.SourceFolder) -or [string]::IsNullOrWhiteSpace($currentSettings.TargetFolder)) {
                Show-Header
                Write-Host ''
                Write-Host '❌ Kontrol işlemi başlatılamıyor.' -ForegroundColor Red
                Write-Host ''
                if ([string]::IsNullOrWhiteSpace($currentSettings.SourceFolder)) { Write-Host '  📁 Kaynak klasör ayarlanmamış.' -ForegroundColor Red }
                if ([string]::IsNullOrWhiteSpace($currentSettings.TargetFolder)) { Write-Host '  💾 Hedef klasör ayarlanmamış.' -ForegroundColor Red }
                Pause-Program
                continue
            }

            Start-IntegrityModule -SourceRoot $currentSettings.SourceFolder -TargetRoot $currentSettings.TargetFolder
        }

        '3' {
            Start-RecoveryMenu -DefaultBackupSource $currentSettings.TargetFolder
        }

        '4' {
            Configure-Settings -CurrentSettingsRef ([ref]$currentSettings)
            $currentSettings = Load-Settings
        }

        'Q' {
            if (-not $env:MFT_MODULE_ID) {
                Write-Host ''
                Write-Host '🚪 Programdan çıkılıyor. İyi günler!' -ForegroundColor Cyan
                Start-Sleep -Milliseconds 500
            }
            return
        }

        default {
            Write-Host ''
            Write-Host '❌ Geçersiz seçim.' -ForegroundColor Red
            Start-Sleep -Milliseconds 700
        }
    }
}
exit 0
