#Requires -Version 5.1
<#
.SYNOPSIS
    M-Files Toolkit GitHub güncelleyicisi (baslat.bat tarafından çağrılır, sessiz çalışır).
.DESCRIPTION
    En son kararlı GitHub Release sürümünü denetler, yeni sürüm varsa program dosyalarını seçici biçimde günceller.
    Sonucu Data\Temp\Update\status.json dosyasına yazar; menü bu mesajı açılış ekranının en üstünde gösterir.

    Güvenlik kuralları:
    - Data klasörüne ve .git klasörüne hiçbir koşulda dokunulmaz.
    - Çalışma kopyası bir Git deposuysa (geliştirici makinesi) dosyalar değiştirilmez, yalnızca bilgi verilir.
    - Başka bir Toolkit oturumu açıksa program dosyalarının üzerine yazılmaz.
    - Güncelleme önce geçici klasöre indirilir ve yapısı doğrulanır; doğrulanmayan arşiv uygulanmaz.
    - version.txt en son yazılır; yarım kalan güncelleme bir sonraki açılışta yeniden denenir.
    - updater.ps1 ve baslat.bat çalışırken değiştirilemez; yeni sürümleri .new olarak bırakılır, baslat.bat uygular.
.PARAMETER Repository
    GitHub deposu (sahip/ad).
.EXAMPLE
    .\updater.ps1 -WhatIf -Verbose
    Dosyaları değiştirmeden güncelleme akışını ve kararlarını gösterir.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [ValidatePattern('^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')]
    [string]$Repository = 'orakcimustafa9656/m-files-toolkit'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$toolkitRoot = $PSScriptRoot
$env:MFT_TOOLKIT_ROOT = $toolkitRoot
$apiUrl = "https://api.github.com/repos/$Repository/releases/latest"
$requestHeaders = @{ 'User-Agent' = 'M-Files-Toolkit-Updater'; 'Accept' = 'application/vnd.github+json' }

# Güncelleyici yalnızca bu iki ortak dosyaya bağımlıdır; ikisi de program kodudur ve Data'ya dokunmaz.
. (Join-Path -Path $toolkitRoot -ChildPath 'lib\ortak.ps1')

function ConvertTo-NormalizedVersion {
    <#
    .SYNOPSIS
        Sürümü üç parçalı hale getirir; böylece 1.5 ile 1.5.0 eşit sayılır ([version] tipi eksik parçayı -1 sayar).
    #>
    [CmdletBinding()]
    [OutputType([version])]
    param(
        [Parameter(Mandatory = $true)][version]$Version
    )

    return [version]::new($Version.Major, $Version.Minor, [Math]::Max(0, $Version.Build))
}

function Write-UpdateStatus {
    <#
    .SYNOPSIS
        Güncelleme sonucunu menünün açılış ekranında göstereceği mesaj dosyasına yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Success', 'Info', 'Warning', 'Error')][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $status = [PSCustomObject]@{ Level = $Level; Message = $Message; Ticks = [DateTime]::UtcNow.Ticks }
    $statusPath = Join-Path -Path (Get-MFTTempPath -Area Update) -ChildPath 'status.json'
    Write-AtomicText -Path $statusPath -Text ($status | ConvertTo-Json -Compress)
    Write-Verbose "[$Level] $Message"
}

function Write-PanelSplitRatio {
    <#
    .SYNOPSIS
        terminal.json içindeki MenuWidth değerinden sağ panelin oranını hesaplayıp split.txt dosyasına yazar.
    .NOTES
        Neden: Windows Terminal "split-pane -s" parametresi YENİ açılan (sağ) panelin oranını belirler; menü %30 ise -s 0.70 gerekir.
        Ayar yoksa, sayı değilse ya da 0.15-0.60 aralığının dışındaysa varsayılan menü genişliği 0.30 kullanılır.
        terminal.json yoksa varsayılanlarla oluşturulur; kullanıcı yalnızca değeri değiştirir.
    #>
    [CmdletBinding()]
    param()

    $defaultMenuWidth = 0.30
    $configPath = Join-Path -Path (Get-MFTCommonDataPath -Category Config) -ChildPath 'terminal.json'
    $menuWidth = $defaultMenuWidth

    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        $defaultJson = ([ordered]@{ MenuWidth = $defaultMenuWidth } | ConvertTo-Json)
        [System.IO.File]::WriteAllText($configPath, $defaultJson, [System.Text.UTF8Encoding]::new($false))
    }
    else {
        try {
            $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $property = $config.PSObject.Properties['MenuWidth']
            $parsed = 0.0
            if ($property -and [double]::TryParse(([string]$property.Value), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed) -and $parsed -ge 0.15 -and $parsed -le 0.60) {
                $menuWidth = $parsed
            }
            else { Write-Warning "terminal.json içindeki MenuWidth geçersiz; varsayılan $defaultMenuWidth kullanılıyor." }
        }
        catch { Write-Warning "terminal.json okunamadı; varsayılan $defaultMenuWidth kullanılıyor." }
    }

    $rightWidth = (1 - $menuWidth).ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
    $splitPath = Join-Path -Path (Get-MFTTempPath -Area Update) -ChildPath 'split.txt'
    [System.IO.File]::WriteAllText($splitPath, $rightWidth, [System.Text.Encoding]::ASCII)
}

function Get-DevelopmentCopyInfo {
    <#
    .SYNOPSIS
        Çalışma kopyası bir Git deposuysa (geliştirici makinesi) commit ve değişiklik bilgisini döndürür; değilse $null.
    .NOTES
        Neden: Sürüm numarası tek başına geliştirici durumunu anlatmaz (yeni kod henüz Release edilmemiş olabilir).
        .git varlığı, kopyanın GitHub'dan indirilmiş bir paket değil, geliştirme deposu olduğunu güvenle gösterir.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    if (-not (Test-Path -LiteralPath (Join-Path -Path $toolkitRoot -ChildPath '.git'))) { return $null }

    $commit = 'bilinmiyor'
    $isDirty = $false
    $git = Get-Command -Name 'git.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($git) {
        $previousPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $head = (& $git.Source -C $toolkitRoot rev-parse --short HEAD 2>$null)
            if ($LASTEXITCODE -eq 0 -and $head) { $commit = [string]$head }
            $changes = @(& $git.Source -C $toolkitRoot status --porcelain 2>$null)
            if ($LASTEXITCODE -eq 0 -and $changes.Count -gt 0) { $isDirty = $true }
        }
        catch { Write-Verbose "Git bilgisi okunamadı: $($_.Exception.Message)" }
        finally { $ErrorActionPreference = $previousPreference }
    }
    return [PSCustomObject]@{ Commit = $commit; IsDirty = $isDirty }
}

function Get-LatestRelease {
    <#
    .SYNOPSIS
        GitHub'daki son kararlı sürümü (etiket, arşiv bağlantısı) döndürür.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    # Neden: Windows PowerShell 5.1 varsayılanı eski TLS olabilir; GitHub TLS 1.2 ister.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $release = Invoke-RestMethod -Uri $apiUrl -Headers $requestHeaders -Method Get -TimeoutSec 20

    $tag = [string]$release.tag_name
    $parsed = $null
    $hasVersion = [version]::TryParse(($tag -replace '^[vV]', ''), [ref]$parsed)
    $zipProperty = $release.PSObject.Properties['zipball_url']
    return [PSCustomObject]@{
        Tag        = $tag
        Version    = if ($hasVersion) { ConvertTo-NormalizedVersion -Version $parsed } else { $null }
        ArchiveUrl = if ($zipProperty) { [string]$zipProperty.Value } else { '' }
    }
}

function Save-ToolkitRelease {
    <#
    .SYNOPSIS
        Sürüm arşivini Data\Temp\Update altındaki geçici klasöre indirir, açar ve kaynak kök klasörünü döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$ArchiveUrl,
        [Parameter(Mandatory = $true)][string]$WorkDirectory
    )

    $null = New-Item -ItemType Directory -Path $WorkDirectory -Force -WhatIf:$false
    $archivePath = Join-Path -Path $WorkDirectory -ChildPath 'release.zip'
    $extractPath = Join-Path -Path $WorkDirectory -ChildPath 'Extracted'

    $downloadParameters = @{
        Uri             = $ArchiveUrl
        Headers         = @{ 'User-Agent' = 'M-Files-Toolkit-Updater' }
        OutFile         = $archivePath
        UseBasicParsing = $true
        TimeoutSec      = 120
    }
    Invoke-WebRequest @downloadParameters

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    # Neden: Arşivdeki hiçbir kayıt çıkarma klasörünün dışına yazamamalı (../ ya da mutlak yol içeren kayıt).
    $archive = [System.IO.Compression.ZipFile]::OpenRead($archivePath)
    try {
        foreach ($entry in $archive.Entries) {
            $segments = $entry.FullName -split '[\\/]'
            if ($entry.FullName.StartsWith('/') -or $entry.FullName -match '^[A-Za-z]:' -or $segments -contains '..') {
                throw "Arşivde güvenli olmayan yol var: $($entry.FullName)"
            }
        }
    }
    finally { $archive.Dispose() }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($archivePath, $extractPath)

    $sourceRoots = @(Get-ChildItem -LiteralPath $extractPath -Directory)
    if ($sourceRoots.Count -ne 1) { throw 'İndirilen arşivde tek bir kaynak klasörü bulunamadı.' }
    return $sourceRoots[0].FullName
}

function Test-ReleaseLayout {
    <#
    .SYNOPSIS
        İndirilen arşivin beklenen program yapısına sahip olup olmadığını denetler; eksik varsa açıklayıcı hata verir.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot
    )

    foreach ($required in @('menu.ps1', 'panel.ps1', 'baslat.bat', 'updater.ps1', 'version.txt', 'lib\ortak.ps1', 'lib\data.ps1', 'lib\menu.ps1', 'lib\panel.ps1')) {
        if (-not (Test-Path -LiteralPath (Join-Path -Path $SourceRoot -ChildPath $required) -PathType Leaf)) {
            throw "Arşivde beklenen dosya yok: $required"
        }
    }
    $modulesPath = Join-Path -Path $SourceRoot -ChildPath 'modules'
    $metaFiles = @(Get-ChildItem -LiteralPath $modulesPath -Filter 'module.meta.json' -File -Recurse -ErrorAction SilentlyContinue)
    if ($metaFiles.Count -eq 0) { throw 'Arşivde hiç modül bulunamadı (modules klasörü boş ya da yok).' }
    $null = Get-ToolkitVersionInfo -Root $SourceRoot
}

function Get-RelativeProgramPath {
    <#
    .SYNOPSIS
        Dosyanın kök klasöre göre göreli yolunu (ters eğik çizgili) döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$FullName
    )

    return $FullName.Substring($Root.Length).TrimStart([char[]]@('\', '/'))
}

function Test-UserDataPath {
    <#
    .SYNOPSIS
        Yolun kullanıcı verisine (Data) ya da Git klasörüne ait olup olmadığını döndürür; bu yollara dokunulmaz.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    $firstSegment = ($RelativePath -split '[\\/]')[0]
    return ($firstSegment -ieq 'Data' -or $firstSegment -ieq '.git')
}

function Test-FileContentDifferent {
    <#
    .SYNOPSIS
        Hedef dosya yoksa ya da içeriği kaynaktan farklıysa $true döndürür (SHA-256).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) { return $true }
    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    $destinationHash = (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash
    return ($sourceHash -ne $destinationHash)
}

function Update-ProgramFile {
    <#
    .SYNOPSIS
        Kaynak klasördeki program dosyalarını hedefe kopyalar; değişmeyen dosyalara dokunmaz.
    .NOTES
        Atlananlar: Data, .git (kullanıcı verisi), version.txt (en son yazılır).
        updater.ps1 ve baslat.bat çalışırken değiştirilemez; farklıysa yanına .new olarak bırakılır.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot
    )

    $deferredNames = @('updater.ps1', 'baslat.bat')
    $changed = [System.Collections.Generic.List[string]]::new()

    foreach ($file in @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -Force)) {
        $relativePath = Get-RelativeProgramPath -Root $SourceRoot -FullName $file.FullName
        if (Test-UserDataPath -RelativePath $relativePath) { continue }
        if ($relativePath -ieq 'version.txt') { continue }

        $destination = Join-Path -Path $toolkitRoot -ChildPath $relativePath
        if ($deferredNames -contains $relativePath) { $destination = "$destination.new" }
        $targetForCompare = if ($deferredNames -contains $relativePath) { Join-Path -Path $toolkitRoot -ChildPath $relativePath } else { $destination }

        if (-not (Test-FileContentDifferent -SourcePath $file.FullName -DestinationPath $targetForCompare)) { continue }
        if ($PSCmdlet.ShouldProcess($relativePath, 'Güncelle')) {
            $null = New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force
            Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
        }
        $changed.Add($relativePath)
    }
    return $changed.ToArray()
}

function Remove-StaleProgramFile {
    <#
    .SYNOPSIS
        Yeni sürümde artık bulunmayan dosyaları yalnızca modules ve lib klasörlerinden siler.
    .NOTES
        Neden: Kaldırılan ya da yeniden adlandırılan bir modül klasörü eski kurulumda kalırsa menüde hayalet modül olarak görünür.
        Karşılaştırma indirilen arşivin kendisiyle yapılır; ayrı bir dosya listesi (manifest) gerekmez. Data'ya dokunulmaz.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot
    )

    $removed = [System.Collections.Generic.List[string]]::new()
    foreach ($folderName in @('modules', 'lib')) {
        $localFolder = Join-Path -Path $toolkitRoot -ChildPath $folderName
        if (-not (Test-Path -LiteralPath $localFolder -PathType Container)) { continue }

        foreach ($file in @(Get-ChildItem -LiteralPath $localFolder -File -Recurse -Force)) {
            $relativePath = Get-RelativeProgramPath -Root $toolkitRoot -FullName $file.FullName
            if (Test-Path -LiteralPath (Join-Path -Path $SourceRoot -ChildPath $relativePath) -PathType Leaf) { continue }
            if ($PSCmdlet.ShouldProcess($relativePath, 'Eski dosyayı sil')) {
                Remove-Item -LiteralPath $file.FullName -Force
            }
            $removed.Add($relativePath)
        }

        # Boşalan klasörleri en derinden başlayarak temizle.
        $directories = @(Get-ChildItem -LiteralPath $localFolder -Directory -Recurse -Force | Sort-Object -Property { $_.FullName.Length } -Descending)
        foreach ($directory in $directories) {
            if (@(Get-ChildItem -LiteralPath $directory.FullName -Force).Count -eq 0) {
                if ($PSCmdlet.ShouldProcess($directory.FullName, 'Boş klasörü sil')) {
                    Remove-Item -LiteralPath $directory.FullName -Force
                }
            }
        }
    }
    return $removed.ToArray()
}

function Invoke-ToolkitUpdate {
    <#
    .SYNOPSIS
        Güncelleme akışının tamamı: sürüm karşılaştırma, güvenlik kontrolleri, indirme, uygulama ve durum mesajı.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $localInfo = $null
    try { $localInfo = Get-ToolkitVersionInfo -Root $toolkitRoot }
    catch {
        Write-UpdateStatus -Level Warning -Message 'Sürüm karşılaştırması yapılamadı (yerel sürüm bilgisi okunamıyor).'
        return
    }
    $localVersion = ConvertTo-NormalizedVersion -Version $localInfo.Version

    $development = Get-DevelopmentCopyInfo
    $developmentSuffix = ''
    if ($null -ne $development) {
        $dirtyText = if ($development.IsDirty) { ', kaydedilmemiş değişiklik var' } else { '' }
        $developmentSuffix = " [Geliştirici kopyası: git $($development.Commit)$dirtyText]"
    }

    try { $latest = Get-LatestRelease }
    catch {
        Write-Verbose "GitHub denetimi başarısız: $($_.Exception.Message)"
        Write-UpdateStatus -Level Warning -Message "Güncelleme kontrolü yapılamadı, yerel sürümle devam ediliyor.$developmentSuffix"
        return
    }
    if ($null -eq $latest.Version) {
        Write-UpdateStatus -Level Warning -Message "Sürüm karşılaştırması yapılamadı (GitHub etiketi '$($latest.Tag)' sürüm numarası değil).$developmentSuffix"
        return
    }

    if ($latest.Version -eq $localVersion) {
        Write-UpdateStatus -Level Success -Message "Sürümler eşleşiyor (v$localVersion).$developmentSuffix"
        return
    }
    if ($latest.Version -lt $localVersion) {
        Write-UpdateStatus -Level Info -Message "Yerel geliştirici sürümü GitHub sürümünden ileride (yerel v$localVersion, GitHub v$($latest.Version)).$developmentSuffix"
        return
    }

    # Buraya gelindiyse GitHub sürümü daha yeni.
    if ($null -ne $development) {
        Write-UpdateStatus -Level Warning -Message "GitHub'da daha yeni sürüm var (v$($latest.Version)); geliştirici kopyası olduğu için dosyalara dokunulmadı.$developmentSuffix"
        return
    }
    if ([string]::IsNullOrWhiteSpace($latest.ArchiveUrl)) {
        Write-UpdateStatus -Level Warning -Message 'Yeni sürüm bulundu ancak GitHub sürüm arşivi bulunamadı; yerel sürümle devam ediliyor.'
        return
    }

    # Neden: Açık bir oturumun çalışan program dosyalarının üzerine yazılmaz; oturumu kullanıcı kapatmadan güncelleme yapılmaz.
    Remove-StaleToolkitSession
    $openSessions = @(Get-ToolkitSession | Where-Object { $_.IsAlive })
    if ($openSessions.Count -gt 0) {
        Write-UpdateStatus -Level Warning -Message "Yeni sürüm bulundu (v$($latest.Version)) ancak başka bir Toolkit oturumu açık; güncelleme ertelendi. Diğer pencereyi kapatıp yeniden başlatın."
        return
    }

    $workDirectory = Join-Path -Path (Get-MFTTempPath -Area Update) -ChildPath ([guid]::NewGuid().ToString('N'))
    try {
        Write-Information -MessageData "Yeni sürüm bulundu (v$($latest.Version)), güncelleniyor..." -InformationAction Continue
        $sourceRoot = Save-ToolkitRelease -ArchiveUrl $latest.ArchiveUrl -WorkDirectory $workDirectory
        Test-ReleaseLayout -SourceRoot $sourceRoot

        $archiveVersion = ConvertTo-NormalizedVersion -Version (Get-ToolkitVersionInfo -Root $sourceRoot).Version
        if ($archiveVersion -ne $latest.Version) {
            throw "Arşivin sürümü (v$archiveVersion) GitHub etiketiyle (v$($latest.Version)) uyuşmuyor."
        }

        # Neden: Fonksiyon tek öğe ya da boş dizi döndürünce StrictMode'da .Count okunamaz; @() ile her zaman dizi olur.
        $changed = @(Update-ProgramFile -SourceRoot $sourceRoot)
        $removed = @(Remove-StaleProgramFile -SourceRoot $sourceRoot)

        # version.txt en son: önceki adımlar yarım kalırsa sürüm eski görünür ve güncelleme bir sonraki açılışta yeniden denenir.
        if ($PSCmdlet.ShouldProcess('version.txt', 'Güncelle')) {
            Copy-Item -LiteralPath (Join-Path -Path $sourceRoot -ChildPath 'version.txt') -Destination (Join-Path -Path $toolkitRoot -ChildPath 'version.txt') -Force
        }
        Write-Verbose "$($changed.Count) dosya güncellendi, $($removed.Count) eski dosya silindi."
        Write-UpdateStatus -Level Success -Message "Yeni sürüm bulundu, güncellendi: v$($latest.Version) ($($changed.Count) dosya)."
    }
    catch {
        Write-UpdateStatus -Level Error -Message "Güncelleme tamamlanamadı: $($_.Exception.Message) Bir sonraki açılışta yeniden denenecek."
    }
    finally {
        if (Test-Path -LiteralPath $workDirectory) { Remove-Item -LiteralPath $workDirectory -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false }
    }
}

try {
    Write-Information -MessageData 'Güncelleme kontrol ediliyor...' -InformationAction Continue
    Write-PanelSplitRatio
    Invoke-ToolkitUpdate
}
catch {
    # Neden: Güncelleyici hatası programın açılmasını engellememeli; hata mesajı açılış ekranında gösterilir.
    Write-UpdateStatus -Level Error -Message "Güncelleyici hatası: $($_.Exception.Message)"
}
