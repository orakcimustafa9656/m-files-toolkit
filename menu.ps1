#Requires -Version 5.1
<#
.SYNOPSIS
    M-Files Toolkit sol panel menüsü (Windows Terminal bölmesinde çalışır).
.DESCRIPTION
    Modülleri module.meta.json dosyalarından bulur, gruplu listeler ve seçilen modülü
    sağdaki modül paneline (panel.ps1) gönderir. Tek başına çalıştırmak için baslat.bat kullanın.
.PARAMETER Session
    Menü ve modül panelinin ortak oturum adı (baslat.bat üretir).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z0-9_-]+$')]
    [string]$Session
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path -Path $PSScriptRoot -ChildPath 'lib\ortak.ps1')
. (Join-Path -Path $PSScriptRoot -ChildPath 'lib\menu.ps1')

# Sürüm ve tarih tek kaynaktan (version.txt) okunur; menü, güncelleyici ve paket aynı değeri kullanır.
$versionInfo = Get-ToolkitVersionInfo -Root $PSScriptRoot
$toolkitVersion = [string]$versionInfo.Version
$toolkitUpdated = [string]$versionInfo.Date
$esc = [char]27

function Get-HostInfo {
    <#
    .SYNOPSIS
        Sol panelin kalıcı karşılama alanında gösterilecek bilgisayar ve işletim sistemi bilgisini üretir.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $osCaption = $null
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os) { $osCaption = [string]$os.Caption }
    }
    catch { }

    if ([string]::IsNullOrWhiteSpace($osCaption)) {
        $osCaption = if ([Environment]::OSVersion.Version.Build -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    }
    else {
        $osCaption = $osCaption -replace '^Microsoft\s+', ''
    }
    $computer = if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) { $env:COMPUTERNAME } else { 'Bilinmiyor' }
    return '{0} • {1}' -f $computer, $osCaption
}

function Invoke-MenuLoop {
    <#
    .SYNOPSIS
        Menünün etkileşimli döngüsü: boyut ve panel durumunu izler, tuşları işler, ekranı yeniden çizer.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SessionDirectory,
        [Parameter(Mandatory = $true)][string]$ModuleRoot
    )

    $catalog = Get-ToolkitModule -ModuleRoot $ModuleRoot
    $size = $Host.UI.RawUI.WindowSize
    $state = Initialize-MenuState -Modules $catalog.Modules -Width $size.Width -Height $size.Height
    if ($catalog.Warnings.Count -gt 0) { Write-MenuMessage -State $state -Text $catalog.Warnings[0] -Kind Warning -Seconds 8 }

    $updateNotice = Read-UpdateNotice
    $hostInfo = Get-HostInfo
    $operator = if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { $env:USERNAME } else { 'Operator' }
    $sessionInfo = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $panelStatus = Get-PanelStatus -Directory $SessionDirectory
    $panelWasAlive = ($panelStatus.State -ne 'offline')
    $panelOfflineChecks = 0
    $statusTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $needDraw = $true
    $clearNext = $true
    $messageWasActive = $false
    $previousScreenLines = $null

    # Açılış efektini sol panelde göster; tamamlanan metinler ilk menü çiziminde temizlenmez.
    Show-MenuBootSequence -Width $size.Width -HostInfo $hostInfo -Operator $operator `
        -SessionInfo $sessionInfo -Version $toolkitVersion -Updated $toolkitUpdated -UpdateNotice $updateNotice
    $clearNext = $false

    while ($true) {
        $size = $Host.UI.RawUI.WindowSize
        if ($size.Width -ne $state.Width -or $size.Height -ne $state.Height) {
            $state.Width = $size.Width
            $state.Height = $size.Height
            $needDraw = $true
            $clearNext = $true
        }

        if ($statusTimer.ElapsedMilliseconds -ge 300) {
            $newStatus = Get-PanelStatus -Directory $SessionDirectory
            if ($newStatus.State -ne $panelStatus.State -or $newStatus.Key -ne $panelStatus.Key) { $needDraw = $true }
            $panelStatus = $newStatus
            $statusTimer.Restart()

            # Sağ panel bir kez ayağa kalktıktan sonra kapanırsa (X ile ya da çöküşle) menü tek başına açık kalmaz.
            # Neden: iki ölçüm üst üste beklenir; anlık bir okuma hatası yüzünden yanlışlıkla kapanma olmasın.
            if ($panelStatus.State -ne 'offline') { $panelWasAlive = $true; $panelOfflineChecks = 0 }
            elseif ($panelWasAlive) {
                $panelOfflineChecks++
                if ($panelOfflineChecks -ge 2) { return }
            }
        }

        $messageActive = ($state.Message.Length -gt 0 -and [DateTime]::UtcNow -lt $state.MessageUntil)
        if ($messageWasActive -and -not $messageActive) { $needDraw = $true }
        $messageWasActive = $messageActive

        if ($needDraw) {
            $screenParameters = @{
                State       = $state
                PanelStatus = $panelStatus
                HostInfo    = $hostInfo
                Operator    = $operator
                SessionInfo = $sessionInfo
                Version     = $toolkitVersion
                Updated     = $toolkitUpdated
                UpdateNotice = $updateNotice
                AsLines     = $true
            }
            $renderedScreen = Format-MenuScreen @screenParameters
            $currentScreenLines = [string[]]$renderedScreen.Lines
            Write-MenuScreenDiff -Lines $currentScreenLines -PreviousLines $previousScreenLines -Clear:$clearNext
            $previousScreenLines = $currentScreenLines
            $needDraw = $false
            $clearNext = $false
        }

        if (-not [Console]::KeyAvailable) {
            Start-Sleep -Milliseconds 40
            continue
        }

        $keyInfo = [Console]::ReadKey($true)
        $action = Invoke-MenuKey -State $state -KeyInfo $keyInfo
        $needDraw = $true

        switch ($action.Action) {
            'Exit' {
                # Neden: modül çalışırken çıkmak yarım kalan işlem (yazma, yedekleme) demektir; önce modül bitmeli.
                $panelStatus = Get-PanelStatus -Directory $SessionDirectory
                if ($panelStatus.State -eq 'busy') {
                    Write-MenuMessage -State $state -Kind Warning -Text "Sağ panelde çalışan modül var: $($panelStatus.Key) $($panelStatus.Title). Önce onu bitirin."
                }
                else {
                    Send-PanelCommand -Directory $SessionDirectory -Action Exit
                    return
                }
            }
            'Refresh' {
                $catalog = Get-ToolkitModule -ModuleRoot $ModuleRoot
                $state.Modules = $catalog.Modules
                Select-MenuFirstVisible -State $state
                Write-MenuMessage -State $state -Text "Liste yenilendi ($($catalog.Modules.Count) modül)." -Kind Success
                $clearNext = $false
            }
            'Run' {
                $panelStatus = Get-PanelStatus -Directory $SessionDirectory
                if ($panelStatus.State -eq 'busy') {
                    Write-MenuMessage -State $state -Kind Warning -Text "Sağ panel meşgul: $($panelStatus.Key) $($panelStatus.Title). Önce onu bitirin."
                }
                elseif ($panelStatus.State -eq 'offline') {
                    Write-MenuMessage -State $state -Kind Error -Text 'Sağ panel yanıt vermiyor. Pencereyi kapatıp baslat.bat ile yeniden açın.'
                }
                else {
                    $commandParameters = @{
                        Directory  = $SessionDirectory
                        Action     = 'Run'
                        Key        = $action.Module.Key
                        Title      = $action.Module.Title
                        ScriptPath = $action.Module.ScriptPath
                    }
                    Send-PanelCommand @commandParameters
                    Write-MenuMessage -State $state -Kind Success -Seconds 2 -Text "Açılıyor: $($action.Module.Key) $($action.Module.Title)"
                    Move-TerminalFocus -Direction right
                }
            }
            default { }
        }
    }
}

Initialize-ConsoleEncoding
$Host.UI.RawUI.WindowTitle = 'M-Files Toolkit · Menü'
$sessionDirectory = Get-SessionDirectory -Session $Session
# Neden: güncelleyici ve panel, oturumun canlı olup olmadığını bu kayıttan anlar; ölü oturum artıkları açılışta temizlenir.
Publish-ToolkitSession -Directory $sessionDirectory
Remove-StaleToolkitSession
$moduleRoot = Join-Path -Path $PSScriptRoot -ChildPath 'modules'

# Alternatif ekran belleği: menü kapanınca terminal eski içeriğine döner. İmleç gizlenir.
[Console]::Write("$esc[?1049h$esc[?25l")
try {
    Invoke-MenuLoop -SessionDirectory $sessionDirectory -ModuleRoot $moduleRoot
}
finally {
    [Console]::Write("$esc[?25h$esc[?1049l")
}
