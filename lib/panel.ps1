#Requires -Version 5.1
Set-StrictMode -Version Latest

# Sağ panel: menüden gelen modülü bu konsolda çalıştırır.

function Write-PanelText {
    <#
    .SYNOPSIS
        Panele renkli (ANSI) metin yazar; modül çıktılarıyla karışmaması için doğrudan konsola gider.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [ValidateSet('Amber', 'Muted', 'Green', 'Red', 'Text')][string]$Color = 'Text'
    )

    $esc = [char]27
    $code = switch ($Color) {
        'Amber' { '38;2;245;167;66' }
        'Muted' { '38;2;123;129;144' }
        'Green' { '38;2;76;208;138' }
        'Red' { '38;2;239;107;107' }
        default { '38;2;215;217;222' }
    }
    [Console]::WriteLine("$esc[${code}m$Text$esc[0m")
}

function Show-StartupPanel {
    <#
    .SYNOPSIS
        Toolkit ilk açıldığında sağ panelde kısa, turuncu temalı açılış ekranını gösterir.
    .NOTES
        Kimlik doğrulaması yapıldığı izlenimi vermemek için kullanıcı profili yüklendi mesajı kullanılır.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Version
    )

    Clear-Host
    $esc = [char]27
    $operator = if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { $env:USERNAME } else { 'Operator' }
    $osName = if ([Environment]::OSVersion.Version.Build -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    $hostName = if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) { $env:COMPUTERNAME } else { 'Bilinmiyor' }

    Write-PanelText -Text '╭──────────────────────────────────────────────────────────────╮' -Color Amber
    Write-PanelText -Text '  M-FILES SYSTEM INITIALIZATION                         READY' -Color Amber
    Write-PanelText -Text '╰──────────────────────────────────────────────────────────────╯' -Color Amber
    Write-PanelText -Text '[i] BOOT SEQUENCE INITIATED...' -Color Amber
    Write-PanelText -Text '[i] Connecting to M-FILES CORE INTERFACE...' -Color Amber
    Write-PanelText -Text "[✔] USER PROFILE LOADED: $operator" -Color Green
    Write-PanelText -Text '[i] Preparing module execution panel...' -Color Amber

    $barWidth = 28
    foreach ($percent in @(0, 20, 40, 60, 80, 100)) {
        $filled = [int][Math]::Round($barWidth * $percent / 100)
        $bar = ('▓' * $filled).PadRight($barWidth, '░')
        [Console]::Write("`r$esc[38;2;245;167;66m $bar $percent%$esc[0m")
        Start-Sleep -Milliseconds 75
    }
    [Console]::WriteLine()
    Write-PanelText -Text '[✔] SESSION INITIALIZED' -Color Green
    Write-PanelText -Text '──────────────────────────────────────────────────────────────' -Color Amber
    Write-PanelText -Text '◆  M-FILES TOOLKIT  ◆' -Color Amber
    Write-PanelText -Text 'Welcome back, Operator.' -Color Text
    Write-PanelText -Text '' -Color Muted
    Write-PanelText -Text ("Session  : {0:yyyy-MM-dd HH:mm}" -f (Get-Date)) -Color Muted
    Write-PanelText -Text ("Host     : {0} · {1}" -f $hostName, $osName) -Color Muted
    Write-PanelText -Text ("Operator : {0}" -f $operator) -Color Muted
    Write-PanelText -Text ("Version  : {0}" -f $Version) -Color Muted
    Write-PanelText -Text 'Program developed by: Mustafa Orakcı' -Color Muted
    Write-PanelText -Text '' -Color Muted
    Write-PanelText -Text 'Verimli çalışmalar dileriz...' -Color Amber
    Start-Sleep -Milliseconds 900
}

function Show-IdlePanel {
    <#
    .SYNOPSIS
        Boştaki paneli çizer: kısa bir yönlendirme dışında ekran boştur.
    #>
    [CmdletBinding()]
    param()

    Clear-Host
    Write-PanelText -Text ' MODÜL PANELİ' -Color Amber
    Write-PanelText -Text ' Soldaki menüden bir modül seçin.' -Color Muted
}

function Wait-AnyKey {
    <#
    .SYNOPSIS
        Kullanıcı bir tuşa basana kadar bekler (modül çıktısı okunabilsin diye).
    #>
    [CmdletBinding()]
    param()

    $null = [Console]::ReadKey($true)
}

function Invoke-ToolkitModule {
    <#
    .SYNOPSIS
        Modül betiğini bu konsolda çalıştırır; hata olursa ayrıntılı hata bilgisini gösterir.
    .NOTES
        Modül StrictMode kapalı bir alt kapsamda çalışır: eski modüller StrictMode açıkken tanımsız
        değişkenlerde hata verebilir; bu nedenle modül kapsamı StrictMode kapalı çalıştırılır.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][string]$SessionDirectory,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$ScriptPath
    )

    Publish-PanelStatus -Directory $SessionDirectory -State busy -Key $Key -Title $Title
    Clear-Host
    Write-PanelText -Text "[i] Modül çalıştırılıyor: $Key $Title" -Color Amber

    $succeeded = $false
    $previousPreference = $ErrorActionPreference
    $previousModuleId = $env:MFT_MODULE_ID
    $previousToolkitRoot = $env:MFT_TOOLKIT_ROOT
    try {
        $moduleDirectory = Split-Path -Parent $ScriptPath
        $metaPath = Join-Path $moduleDirectory 'module.meta.json'
        if (Test-Path -LiteralPath $metaPath -PathType Leaf) {
            $moduleMeta = Get-Content -LiteralPath $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
            $moduleIdProperty = $moduleMeta.PSObject.Properties['Id']
            if (-not $moduleIdProperty -or [string]$moduleIdProperty.Value -notmatch '^[a-z0-9][a-z0-9_-]*$') { throw "Modul kimligi (Id) eksik veya gecersiz: $metaPath" }
            $env:MFT_MODULE_ID = [string]$moduleIdProperty.Value
            $env:MFT_TOOLKIT_ROOT = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $moduleDirectory))
        }
        if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf) -or $ScriptPath -notlike '*.ps1') {
            throw "Script bulunamadı: $ScriptPath"
        }
        $ErrorActionPreference = 'Stop'
        # Neden: Out-Host modülün normal çıktısını (Write-Output, tablolar) ekrana basar ve fonksiyonun dönüş değerini kirletmez.
        & {
            Set-StrictMode -Off
            & $ScriptPath
        } | Out-Host
        $succeeded = $true
        Write-PanelText -Text "[✔] $Title işlemi tamamlandı." -Color Green
    }
    catch {
        Write-PanelText -Text "[✖] Çalıştırma hatası: $($_.Exception.Message)" -Color Red
        if ($_.InvocationInfo) {
            Write-PanelText -Text "[✖] Satır: $($_.InvocationInfo.ScriptLineNumber)" -Color Red
            Write-PanelText -Text "[✖] $($_.InvocationInfo.Line.Trim())" -Color Red
        }
    }
    finally {
        $ErrorActionPreference = $previousPreference
        $env:MFT_MODULE_ID = $previousModuleId
        $env:MFT_TOOLKIT_ROOT = $previousToolkitRoot
    }

    Write-PanelText -Text '' -Color Muted
    Write-PanelText -Text ' Menüye dönmek için bir tuşa basın…' -Color Muted
    Wait-AnyKey

    Show-IdlePanel
    Publish-PanelStatus -Directory $SessionDirectory -State idle
    Move-TerminalFocus -Direction left
    return $succeeded
}
