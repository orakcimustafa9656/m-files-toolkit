#Requires -Version 5.1
<#
.SYNOPSIS
    M-Files Toolkit sağ panel: menüden gelen modülü çalıştırır.
.DESCRIPTION
    Başta boş bekler; menü bir modül seçince onu bu konsolda çalıştırır. Modüller Excel COM ve WinForms
    kullandığı için bu betik STA modunda başlatılmalıdır (baslat.bat bunu yapar).
.PARAMETER Session
    Menü ve modül panelinin ortak oturum adı.
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
. (Join-Path -Path $PSScriptRoot -ChildPath 'lib\panel.ps1')

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning 'Panel STA modunda değil; Excel/WinForms kullanan modüller çalışmayabilir. baslat.bat ile açın.'
}

Initialize-ConsoleEncoding
$Host.UI.RawUI.WindowTitle = 'M-Files Toolkit · Modül'
$sessionDirectory = Get-SessionDirectory -Session $Session

# Sol menü panelin hazır olduğunu ilk açılıştan itibaren bilsin.
Publish-PanelStatus -Directory $sessionDirectory -State idle
Show-IdlePanel
$heartbeat = [System.Diagnostics.Stopwatch]::StartNew()
$menuWasSeen = $false

try {
while ($true) {
    $command = Receive-PanelCommand -Directory $sessionDirectory
    if ($null -ne $command) {
        if ($command.Action -eq 'Exit') { break }
        if ($command.Action -eq 'Run') {
            $moduleParameters = @{
                SessionDirectory = $sessionDirectory
                Key              = [string]$command.Key
                Title            = [string]$command.Title
                ScriptPath       = [string]$command.ScriptPath
            }
            $null = Invoke-ToolkitModule @moduleParameters
            $heartbeat.Restart()
            continue
        }
    }

    # Neden: boştaki panel her saniye "yaşıyorum" der; menü eski kaydı "yanıt vermiyor" sayar.
    if ($heartbeat.ElapsedMilliseconds -ge 1000) {
        Publish-PanelStatus -Directory $sessionDirectory -State idle
        $heartbeat.Restart()

        # Sol menü kapandıysa (Q ya da X) bu panel de kapanır; tek başına açık kalmaz.
        # Neden: menü ilk kez görülmeden kontrol edilmez, çünkü iki panel aynı anda başlar.
        $menuAlive = Test-ToolkitMenuAlive -Directory $sessionDirectory
        if ($menuAlive -eq $true) { $menuWasSeen = $true }
        elseif ($menuAlive -eq $false -and $menuWasSeen) { break }
    }
    Start-Sleep -Milliseconds 150
}
}
finally {
    # Panel oturumun son kapanan tarafıdır; oturum klasörünü burada silmek Exit komutunun okunmasını engellemez.
    Remove-Item -LiteralPath $sessionDirectory -Recurse -Force -ErrorAction SilentlyContinue
}
