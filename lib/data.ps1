#requires -Version 5.1
# Merkezi kalici kullanici verisi ve gecici dosya yolu API'si.
# Neden: Bu dosya dot-source edilir; Set-StrictMode cagirmaz, cagiranin modunu degistirmez.
#
# Duzen (hepsi Git disinda, /Data/ altinda; klasorler ilk ihtiyac aninda olusur):
#   Data\Common\Config\                 ortak ayarlar (terminal.json)
#   Data\Modules\<Id>\<Kategori>\       modulun kalici verisi
#   Data\Temp\Update\                   guncelleyici gecici dosyalari
#   Data\Temp\Sessions\<Oturum>\        menu <-> panel iletisimi
#   Data\Temp\Modules\<Id>\             modulun gecici dosyalari
# <Id>, module.meta.json icindeki kucuk harfli sabit kimliktir (ornek: gibberatoku).

$script:MFTDataToolkitRoot = Split-Path -Parent $PSScriptRoot

function Get-MFTToolkitRoot {
    <#
    .SYNOPSIS
        Program ana dizinini dondurur.
    .NOTES
        Panel, modulu calistirirken MFT_TOOLKIT_ROOT ortam degiskenini ayarlar; yoksa lib klasorunun ustu kullanilir.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $toolkitRoot = $env:MFT_TOOLKIT_ROOT
    if ([string]::IsNullOrWhiteSpace($toolkitRoot)) { $toolkitRoot = $script:MFTDataToolkitRoot }
    if ([string]::IsNullOrWhiteSpace($toolkitRoot)) { throw 'M-Files Toolkit ana dizini belirlenemedi.' }
    return $toolkitRoot
}

function Initialize-MFTDirectory {
    <#
    .SYNOPSIS
        Klasor yoksa olusturur ve yolunu dondurur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        # Neden: -WhatIf yalnızca program dosyalarını etkileyen adımları simüle eder; altyapı klasörleri her zaman oluşturulmalı.
        $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop -WhatIf:$false
    }
    return $Path
}

function Get-MFTModuleDataPath {
    <#
    .SYNOPSIS
        Modulun kalici veri klasorunu dondurur; yoksa olusturur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][ValidatePattern('^[a-z0-9][a-z0-9_-]*$')][string]$ModuleId,
        [Parameter(Mandatory = $true)][ValidateSet('Config', 'Data', 'Logs', 'Reports', 'Projects', 'Cache')][string]$Category
    )

    $modulesRoot = Join-Path -Path (Get-MFTToolkitRoot) -ChildPath 'Data\Modules'
    $categoryPath = Join-Path -Path (Join-Path -Path $modulesRoot -ChildPath $ModuleId) -ChildPath $Category
    return (Initialize-MFTDirectory -Path $categoryPath)
}

function Get-MFTCommonDataPath {
    <#
    .SYNOPSIS
        Moduller arasi ortak kalici veri klasorunu dondurur; yoksa olusturur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Config', 'Data', 'Logs', 'Cache')][string]$Category
    )

    $categoryPath = Join-Path -Path (Join-Path -Path (Get-MFTToolkitRoot) -ChildPath 'Data\Common') -ChildPath $Category
    return (Initialize-MFTDirectory -Path $categoryPath)
}

function Get-MFTTempPath {
    <#
    .SYNOPSIS
        Uygulamanin ortak gecici klasorunu (Update veya Sessions) dondurur; yoksa olusturur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Update', 'Sessions')][string]$Area
    )

    $areaPath = Join-Path -Path (Join-Path -Path (Get-MFTToolkitRoot) -ChildPath 'Data\Temp') -ChildPath $Area
    return (Initialize-MFTDirectory -Path $areaPath)
}

function Get-MFTModuleTempPath {
    <#
    .SYNOPSIS
        Modulun gecici dosya klasorunu dondurur; yoksa olusturur.
    .NOTES
        Hedef klasorun yaninda tutulmasi gereken gecici dosyalar (ZIP, yedek, Excel) ve buyuk ISO dosyalari
        icin Windows %TEMP% bu fonksiyonla degil, ilgili modulde dogrudan kullanilir.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][ValidatePattern('^[a-z0-9][a-z0-9_-]*$')][string]$ModuleId
    )

    $modulePath = Join-Path -Path (Join-Path -Path (Get-MFTToolkitRoot) -ChildPath 'Data\Temp\Modules') -ChildPath $ModuleId
    return (Initialize-MFTDirectory -Path $modulePath)
}
