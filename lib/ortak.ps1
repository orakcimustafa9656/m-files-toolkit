#Requires -Version 5.1
Set-StrictMode -Version Latest
. (Join-Path -Path $PSScriptRoot -ChildPath 'data.ps1')

# Menü (sol panel) ile modül paneli (sağ panel) arasındaki ortak yardımcılar.
# İletişim, oturuma özel geçici klasörde iki küçük dosyayla yapılır:
#   cmd.txt     menü -> panel   (çalıştırılacak modül ya da çıkış komutu, JSON)
#   status.json panel -> menü   (boşta / meşgul bilgisi, her saniye tazelenir)

function Initialize-ConsoleEncoding {
    <#
    .SYNOPSIS
        Konsol girdi ve çıktısını UTF-8 yapar (Türkçe karakterler ve çizgi sembolleri için).
    #>
    [CmdletBinding()]
    param()

    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
}

function Get-SessionDirectory {
    <#
    .SYNOPSIS
        Oturuma özel geçici iletişim klasörünü döndürür; yoksa oluşturur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[A-Za-z0-9_-]+$')]
        [string]$Session
    )

    $path = Join-Path -Path (Get-MFTTempPath -Area Sessions) -ChildPath $Session
    $null = New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop
    return $path
}

function Get-ToolkitVersionInfo {
    <#
    .SYNOPSIS
        Program surumunu ve tarihini kok dizindeki version.txt dosyasindan okur.
    .NOTES
        Tek kaynak: version.txt (1. satir surum, 2. satir tarih). Surum [version] tipidir; boylece 1.5.10, 1.5.9'dan buyuk sayilir.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][string]$Root
    )

    $path = Join-Path -Path $Root -ChildPath 'version.txt'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "version.txt bulunamadı: $path" }

    $lines = @(Get-Content -LiteralPath $path -Encoding UTF8 -ErrorAction Stop)
    $parsedVersion = $null
    if ($lines.Count -lt 1 -or -not [version]::TryParse($lines[0].Trim(), [ref]$parsedVersion)) {
        throw "version.txt içindeki sürüm okunamadı: $path"
    }
    $dateText = if ($lines.Count -ge 2) { $lines[1].Trim() } else { '' }
    return [PSCustomObject]@{ Version = $parsedVersion; Date = $dateText }
}

function Test-ToolkitProcessAlive {
    <#
    .SYNOPSIS
        İşlem numarasının hala bu program oturumuna ait çalışan bir işlem olup olmadığını denetler.
    .NOTES
        Neden: Windows işlem numaralarını yeniden kullanır; yalnızca numaraya bakmak çöken bir oturumu "açık" gösterebilir.
        Bu yüzden komut satırında beklenen betik adı da doğrulanır. Komut satırı okunamazsa temkinli davranıp "canlı" sayılır.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][int]$ProcessId,
        [Parameter(Mandatory = $true)][string]$ScriptName
    )

    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($null -eq $process) { return $false }

    $cimProcess = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction SilentlyContinue
    if ($null -eq $cimProcess -or [string]::IsNullOrWhiteSpace([string]$cimProcess.CommandLine)) { return $true }
    return ([string]$cimProcess.CommandLine).IndexOf($ScriptName, [StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Publish-ToolkitSession {
    <#
    .SYNOPSIS
        Menü işleminin numarasını oturum klasörüne yazar; güncelleyici ve panel oturumun canlı olup olmadığını buradan anlar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Directory
    )

    $session = [PSCustomObject]@{ MenuProcessId = $PID; StartedTicks = [DateTime]::UtcNow.Ticks }
    Write-AtomicText -Path (Join-Path -Path $Directory -ChildPath 'session.json') -Text ($session | ConvertTo-Json -Compress)
}

function Test-ToolkitMenuAlive {
    <#
    .SYNOPSIS
        Oturumun menü işleminin canlı olup olmadığını döndürür; session.json henüz yazılmadıysa $null.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory = $true)][string]$Directory
    )

    $path = Join-Path -Path $Directory -ChildPath 'session.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $session = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop
        return (Test-ToolkitProcessAlive -ProcessId ([int]$session.MenuProcessId) -ScriptName 'menu.ps1')
    }
    # Neden: dosya tam o anda yazılıyor olabilir; belirsiz durumda "canlı" sayılır, yanlışlıkla kapanma olmaz.
    catch { return $true }
}

function Get-ToolkitSession {
    <#
    .SYNOPSIS
        Data\Temp\Sessions altındaki oturumları listeler; her biri için IsAlive bilgisini döndürür.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    foreach ($directory in @(Get-ChildItem -LiteralPath (Get-MFTTempPath -Area Sessions) -Directory -ErrorAction SilentlyContinue)) {
        $alive = Test-ToolkitMenuAlive -Directory $directory.FullName
        # session.json henüz yoksa oturum yeni başlıyor olabilir; 2 dakikadan eskiyse artık sayılır.
        if ($null -eq $alive) { $alive = ($directory.LastWriteTimeUtc -gt [DateTime]::UtcNow.AddMinutes(-2)) }
        [PSCustomObject]@{ Name = $directory.Name; Path = $directory.FullName; IsAlive = [bool]$alive }
    }
}

function Remove-StaleToolkitSession {
    <#
    .SYNOPSIS
        Yalnızca işlemi artık yaşamayan oturum klasörlerini siler; açık oturumlara dokunmaz.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    foreach ($session in @(Get-ToolkitSession | Where-Object { -not $_.IsAlive })) {
        if ($PSCmdlet.ShouldProcess($session.Path, 'Eski oturum klasörünü sil')) {
            Remove-Item -LiteralPath $session.Path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Read-UpdateNotice {
    <#
    .SYNOPSIS
        Güncelleyicinin bıraktığı son durum mesajını okur; mesaj yoksa ya da 5 dakikadan eskiyse $null döndürür.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    $path = Join-Path -Path (Get-MFTTempPath -Area Update) -ChildPath 'status.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $notice = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop
        $age = [TimeSpan]::FromTicks([DateTime]::UtcNow.Ticks - [long]$notice.Ticks)
        if ($age.TotalMinutes -gt 5 -or [string]::IsNullOrWhiteSpace([string]$notice.Message)) { return $null }
        return [PSCustomObject]@{ Level = [string]$notice.Level; Message = [string]$notice.Message }
    }
    catch { return $null }
}

function Write-AtomicText {
    <#
    .SYNOPSIS
        Dosyayı önce geçici adla yazıp yeniden adlandırır; okuyan taraf yarım dosya görmez.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $temporaryPath = "$Path.tmp"
    [System.IO.File]::WriteAllText($temporaryPath, $Text, [System.Text.UTF8Encoding]::new($false))
    # Neden: -WhatIf çalıştırmalarında bile durum/iletişim dosyaları yazılmalı; yalnızca program dosyaları simüle edilir.
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -ErrorAction Stop -WhatIf:$false
}

function Send-PanelCommand {
    <#
    .SYNOPSIS
        Menüden modül paneline komut gönderir (Run: modülü çalıştır, Exit: paneli kapat).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][ValidateSet('Run', 'Exit')][string]$Action,
        [string]$Key = '',
        [string]$Title = '',
        [string]$ScriptPath = ''
    )

    $command = [PSCustomObject]@{
        Action     = $Action
        Key        = $Key
        Title      = $Title
        ScriptPath = $ScriptPath
    }
    $writeParameters = @{
        Path = Join-Path -Path $Directory -ChildPath 'cmd.txt'
        Text = ($command | ConvertTo-Json -Compress)
    }
    Write-AtomicText @writeParameters
}

function Receive-PanelCommand {
    <#
    .SYNOPSIS
        Bekleyen komut varsa okuyup siler ve nesne olarak döndürür; yoksa $null.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][string]$Directory
    )

    $path = Join-Path -Path $Directory -ChildPath 'cmd.txt'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }

    try {
        $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $path -Force -ErrorAction Stop
        return ($text | ConvertFrom-Json -ErrorAction Stop)
    }
    # Neden: dosya tam o anda değiştiriliyor olabilir; bir sonraki turda yeniden denenir.
    catch [System.IO.IOException] { return $null }
    catch [System.ArgumentException] {
        Write-Warning "Geçersiz komut dosyası yok sayıldı: $($_.Exception.Message)"
        return $null
    }
}

function Publish-PanelStatus {
    <#
    .SYNOPSIS
        Modül panelinin durumunu (idle ya da busy) menünün okuyacağı dosyaya yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][ValidateSet('idle', 'busy')][string]$State,
        [string]$Key = '',
        [string]$Title = ''
    )

    $status = [PSCustomObject]@{
        State        = $State
        Key          = $Key
        Title        = $Title
        ProcessId    = $PID
        UpdatedTicks = [DateTime]::UtcNow.Ticks
    }
    $writeParameters = @{
        Path = Join-Path -Path $Directory -ChildPath 'status.json'
        Text = ($status | ConvertTo-Json -Compress)
    }
    Write-AtomicText @writeParameters
}

function Get-PanelStatus {
    <#
    .SYNOPSIS
        Panelin durumunu döndürür: idle, busy ya da offline (panel kapalı ya da yanıt vermiyor).
    .NOTES
        Boştaki panel her saniye durum yazar; 5 saniyeden eski kayıt "offline" sayılır.
        Meşgul panel modül çalıştırırken yazamaz; bu durumda işlem numarasının yaşadığına bakılır.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][string]$Directory
    )

    $offline = [PSCustomObject]@{ State = 'offline'; Key = ''; Title = '' }
    $path = Join-Path -Path $Directory -ChildPath 'status.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $offline }

    try {
        $status = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch [System.IO.IOException] { return $offline }
    catch [System.ArgumentException] { return $offline }

    # Neden: sağ panel X ile kapatılırsa işlem hemen ölür; 5 saniyelik zaman aşımını beklemeden menü bunu görmeli.
    $alive = $null -ne (Get-Process -Id ([int]$status.ProcessId) -ErrorAction SilentlyContinue)
    if (-not $alive) { return $offline }

    if ($status.State -ne 'busy') {
        $age = [TimeSpan]::FromTicks([DateTime]::UtcNow.Ticks - [long]$status.UpdatedTicks)
        if ($age.TotalSeconds -gt 5) { return $offline }
    }
    return [PSCustomObject]@{ State = [string]$status.State; Key = [string]$status.Key; Title = [string]$status.Title }
}

function Move-TerminalFocus {
    <#
    .SYNOPSIS
        Windows Terminal'de odağı sol ya da sağ bölmeye taşır (en iyi çaba).
    .NOTES
        Windows Terminal dışında ya da wt.exe bulunamazsa sessizce hiçbir şey yapmaz.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('left', 'right')][string]$Direction
    )

    if (-not $env:WT_SESSION) { return }
    $terminal = Get-Command -Name 'wt.exe' -CommandType Application -ErrorAction SilentlyContinue
    if (-not $terminal) { return }

    $startParameters = @{
        FilePath     = $terminal.Source
        ArgumentList = @('-w', '0', 'move-focus', $Direction)
        WindowStyle  = 'Hidden'
        ErrorAction  = 'Stop'
    }
    try { Start-Process @startParameters }
    # Neden: odak değişmezse bile program çalışır; kullanıcı fareyle ya da Alt+Ok ile geçebilir.
    catch { Write-Verbose "Odak taşınamadı: $($_.Exception.Message)" }
}
