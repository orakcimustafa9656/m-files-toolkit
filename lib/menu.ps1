#Requires -Version 5.1
Set-StrictMode -Version Latest

# Sol panel menüsü: modül keşfi, durum yönetimi ve ekran çizimi.
# Çizim ve tuş işleme saf fonksiyonlardır (konsola dokunmaz), bu yüzden ayrı ayrı test edilebilir.

$script:Esc = [char]27

# Sol bölmenin sağ kenarına çizilen turuncu ayırıcı çizgi (▐). Windows Terminal bölme çizgisinin rengi değiştirilemez,
# bu yüzden çizgiyi menü kendisi çizer. İnce çizgi için [char]0x2502 (│) yazın.
# $script:DividerChar = [char]0x2590
$script:Ansi = @{
    Reset     = "$([char]27)[0m"
    Bold      = "$([char]27)[1m"
    Amber     = "$([char]27)[38;2;245;167;66m"
    AmberDim  = "$([char]27)[38;2;185;118;42m"
    Text      = "$([char]27)[38;2;215;217;222m"
    Muted     = "$([char]27)[38;2;123;129;144m"
    Line      = "$([char]27)[38;2;58;62;70m"
    Green     = "$([char]27)[38;2;76;208;138m"
    Red       = "$([char]27)[38;2;239;107;107m"
    OnAmber   = "$([char]27)[48;2;245;167;66m$([char]27)[38;2;17;17;17m"
    OnSoft    = "$([char]27)[48;2;42;36;24m$([char]27)[38;2;215;217;222m"
}

function Get-MenuCulture {
    <#
    .SYNOPSIS
        Türkçe kültürü döndürür (İ/ı büyük-küçük harf kuralları için); bulunamazsa invariant.
    #>
    [CmdletBinding()]
    [OutputType([System.Globalization.CultureInfo])]
    param()

    try { return [System.Globalization.CultureInfo]::GetCultureInfo('tr-TR') }
    catch [System.Globalization.CultureNotFoundException] { return [System.Globalization.CultureInfo]::InvariantCulture }
}

function Get-MetaValue {
    <#
    .SYNOPSIS
        Meta nesnesinden alan okur; alan yoksa StrictMode hatası vermeden $null döndürür.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Meta,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $property = $Meta.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-ModuleGroup {
    <#
    .SYNOPSIS
        Menü grubunu yalnızca fiziksel modules\<Grup>\<Modül> klasör yapısından alır.
    .NOTES
        Grup klasörü adı olduğu gibi kullanılır. Meta içindeki Group veya modül Key'i grubu belirlemez.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][string]$ModuleRoot,
        [Parameter(Mandatory = $true)][string]$ModuleDirectory
    )

    $root = [System.IO.Path]::GetFullPath($ModuleRoot).TrimEnd([char[]]@('\', '/'))
    $modulePath = [System.IO.Path]::GetFullPath($ModuleDirectory)
    if (-not $modulePath.StartsWith($root + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    $relativeDirectory = $modulePath.Substring($root.Length).TrimStart([char[]]@('\', '/'))
    $parts = @($relativeDirectory -split '[\\/]+' | Where-Object { $_.Length -gt 0 })

    # Beklenen fiziksel düzen: modules\<grup>\<modül>\module.meta.json
    if ($parts.Count -ne 2) { return $null }
    return ([string]$parts[0]).Trim()
}

function Get-ToolkitModule {
    <#
    .SYNOPSIS
        modules klasörünün tüm alt dizinlerindeki module.meta.json dosyalarını okuyup menü öğelerini döndürür.
    .NOTES
        Yalnızca modules\<Grup>\<Modül>\module.meta.json yapısını kabul eder.
        Grup adı ve sırası fiziksel grup klasöründen gelir; modül Key'i grup belirlemede kullanılmaz.
        Id modülün sabit kimliğidir; Key menü sıralaması içindir. Betik adı varsayılan olarak <Id>.ps1'dir. Yinelenen Id/Key ve eksik betik atlanır.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][string]$ModuleRoot
    )

    $modules = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $usedKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $usedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if (-not (Test-Path -LiteralPath $ModuleRoot -PathType Container)) {
        $warnings.Add("Modül klasörü bulunamadı: $ModuleRoot")
        return [PSCustomObject]@{ Modules = @(); Warnings = $warnings.ToArray() }
    }

    # Yalnızca gerçek meta dosyalarını tarar; kategori klasörlerinin kendisini modül sanmaz.
    $metaFiles = @(
        Get-ChildItem -LiteralPath $ModuleRoot -Filter 'module.meta.json' -File -Recurse -ErrorAction SilentlyContinue |
            Sort-Object -Property FullName
    )

    foreach ($metaFile in $metaFiles) {
        $moduleDirectory = $metaFile.DirectoryName
        try {
            $meta = Get-Content -LiteralPath $metaFile.FullName -Raw -Encoding UTF8 -ErrorAction Stop |
                ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            $warnings.Add("Meta dosyası okunamadı ($($metaFile.FullName)): $($_.Exception.Message)")
            continue
        }

        $key = Get-MetaValue -Meta $meta -Name 'Key'
        $moduleId = Get-MetaValue -Meta $meta -Name 'Id'
        $title = Get-MetaValue -Meta $meta -Name 'Title'
        $scriptFileName = Get-MetaValue -Meta $meta -Name 'ScriptFileName'
        if (-not ($key -is [string] -and $key -match '^\d+$')) { $warnings.Add("Key geçersiz: $moduleDirectory"); continue }
        if (-not ($moduleId -is [string] -and $moduleId -match '^[a-z0-9][a-z0-9_-]*$')) { $warnings.Add("Id eksik veya geçersiz: $moduleDirectory"); continue }
        if ($usedIds.Contains($moduleId)) { $warnings.Add("Yinelenen Id '$moduleId' atlandı: $moduleDirectory"); continue }
        if ([string]::IsNullOrWhiteSpace([string]$title)) { $warnings.Add("Title eksik: $moduleDirectory"); continue }
        # Standart: betik adı <Id>.ps1'dir; ScriptFileName yalnızca bu kuralın dışına çıkan modül için yazılır.
        if ([string]::IsNullOrWhiteSpace([string]$scriptFileName)) { $scriptFileName = "$moduleId.ps1" }
        if ($usedKeys.Contains($key)) { $warnings.Add("Yinelenen Key '$key' atlandı: $moduleDirectory"); continue }

        $scriptPath = Join-Path -Path $moduleDirectory -ChildPath ([string]$scriptFileName)
        if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
            $warnings.Add("Script dosyası bulunamadı: $scriptPath")
            continue
        }
        $group = Get-ModuleGroup -ModuleRoot $ModuleRoot -ModuleDirectory $moduleDirectory
        if ([string]::IsNullOrWhiteSpace($group)) {
            $warnings.Add("Fiziksel grup klasörü dışında kalan modül atlandı. Beklenen yapı: modules\<Grup>\<Modül>\module.meta.json ($moduleDirectory)")
            continue
        }

        # Grup sırası fiziksel klasör adındaki sayısal önekten gelir; modül Key'i gruplama için kullanılmaz.
        if ($group -match '^\s*(\d+)') { $groupIndex = [int]$matches[1] }
        else { $groupIndex = 100000 }
        $groupSortName = $group
        $null = $usedKeys.Add($key)
        $null = $usedIds.Add($moduleId)

        $modules.Add([PSCustomObject]@{
            Id            = $moduleId
            Key           = $key
            Title         = [string]$title
            Group         = [string]$group
            GroupIndex    = $groupIndex
            GroupSortName = $groupSortName
            GroupFolder   = $group
            SortKey       = [int]$key
            ScriptPath    = $scriptPath
        })
    }

    if ($metaFiles.Count -eq 0) { $warnings.Add("Modül meta dosyası bulunamadı: $ModuleRoot") }
    $sorted = @($modules | Sort-Object -Property GroupIndex, GroupSortName, SortKey, Title)
    return [PSCustomObject]@{ Modules = $sorted; Warnings = $warnings.ToArray() }
}

function Initialize-MenuState {
    <#
    .SYNOPSIS
        Menünün değişken durumunu (arama kutusu, seçili modül, kaydırma, mesaj) başlatır.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Modules,
        [Parameter(Mandatory = $true)][int]$Width,
        [Parameter(Mandatory = $true)][int]$Height
    )

    $state = [PSCustomObject]@{
        Modules      = $Modules
        Width        = $Width
        Height       = $Height
        Buffer       = ''
        SelectedKey  = ''
        Offset       = 0
        Message      = ''
        MessageKind  = 'Info'
        MessageUntil = [DateTime]::MinValue
    }
    Select-MenuFirstVisible -State $state
    return $state
}

# Türkçe harflerin aksansız karşılıkları (karakter kodlarıyla; dosya kodlamasından bağımsız).
$script:SearchCharMap = @{
    ([char]0x00E7) = 'c'; ([char]0x00C7) = 'c'; ([char]0x011F) = 'g'; ([char]0x011E) = 'g'
    ([char]0x0131) = 'i'; ([char]0x0130) = 'i'; ([char]0x00F6) = 'o'; ([char]0x00D6) = 'o'
    ([char]0x015F) = 's'; ([char]0x015E) = 's'; ([char]0x00FC) = 'u'; ([char]0x00DC) = 'u'
    ([char]0x00E2) = 'a'; ([char]0x00C2) = 'a'; ([char]0x00EE) = 'i'; ([char]0x00CE) = 'i'
    ([char]0x00FB) = 'u'; ([char]0x00DB) = 'u'
}

function ConvertTo-SearchKey {
    <#
    .SYNOPSIS
        Arama için metni sadeleştirir: Türkçe harfler aksansız karşılığına çevrilir, küçük harfe indirilir.
    .NOTES
        Kültür ve ICU ayarına bağlı değildir; "insaat" yazınca "İnşaat" bulunur.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyString()][string]$Text = ''
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($character in $Text.ToCharArray()) {
        if ($script:SearchCharMap.ContainsKey($character)) { $null = $builder.Append($script:SearchCharMap[$character]) }
        else { $null = $builder.Append($character) }
    }
    return $builder.ToString().ToLowerInvariant()
}

function Test-ModuleMatch {
    <#
    .SYNOPSIS
        Modülün arama metniyle eşleşip eşleşmediğini söyler (numara ya da ad/grup; Türkçe ve aksan duyarsız).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][psobject]$Module,
        [AllowEmptyString()][string]$Text = ''
    )

    $needle = $Text.Trim()
    if ($needle.Length -eq 0) { return $true }

    if ($needle -match '^\d+$') {
        return $Module.Key.StartsWith($needle, [StringComparison]::Ordinal) -or $Module.Key -eq $needle.PadLeft(2, '0')
    }

    $needleKey = ConvertTo-SearchKey -Text $needle
    return (ConvertTo-SearchKey -Text $Module.Title).Contains($needleKey) -or (ConvertTo-SearchKey -Text $Module.Group).Contains($needleKey)
}

function Get-VisibleModule {
    <#
    .SYNOPSIS
        Arama metnine uyan modülleri menü sırasıyla döndürür.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory = $true)][psobject]$State
    )

    return @($State.Modules | Where-Object { Test-ModuleMatch -Module $_ -Text $State.Buffer })
}

function Select-MenuFirstVisible {
    <#
    .SYNOPSIS
        Seçimi görünür listenin ilk modülüne alır (liste boşsa seçimi boşaltır).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$State
    )

    $visible = @(Get-VisibleModule -State $State)
    $State.SelectedKey = if ($visible.Count -gt 0) { $visible[0].Key } else { '' }
}

function Write-MenuMessage {
    <#
    .SYNOPSIS
        Alt satırda birkaç saniye gösterilecek mesajı ayarlar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$State,
        [Parameter(Mandatory = $true)][string]$Text,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')][string]$Kind = 'Info',
        [int]$Seconds = 4
    )

    $State.Message = $Text
    $State.MessageKind = $Kind
    $State.MessageUntil = [DateTime]::UtcNow.AddSeconds($Seconds)
}

function Invoke-MenuKey {
    <#
    .SYNOPSIS
        Bir tuş vuruşunu işler ve yapılacak eylemi döndürür: None, Run, Exit ya da Refresh.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)][psobject]$State,
        [Parameter(Mandatory = $true)][System.ConsoleKeyInfo]$KeyInfo
    )

    $result = [PSCustomObject]@{ Action = 'None'; Module = $null }
    $visible = @(Get-VisibleModule -State $State)
    $index = -1
    for ($i = 0; $i -lt $visible.Count; $i++) {
        if ($visible[$i].Key -eq $State.SelectedKey) { $index = $i; break }
    }

    switch ($KeyInfo.Key) {
        ([ConsoleKey]::UpArrow) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[[Math]::Max(0, $index - 1)].Key } }
        ([ConsoleKey]::DownArrow) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[[Math]::Min($visible.Count - 1, $index + 1)].Key } }
        ([ConsoleKey]::PageUp) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[[Math]::Max(0, $index - 10)].Key } }
        ([ConsoleKey]::PageDown) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[[Math]::Min($visible.Count - 1, $index + 10)].Key } }
        ([ConsoleKey]::Home) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[0].Key } }
        ([ConsoleKey]::End) { if ($visible.Count -gt 0) { $State.SelectedKey = $visible[$visible.Count - 1].Key } }
        ([ConsoleKey]::F5) { $result.Action = 'Refresh' }
        ([ConsoleKey]::Escape) {
            $State.Buffer = ''
            Select-MenuFirstVisible -State $State
        }
        ([ConsoleKey]::Backspace) {
            if ($State.Buffer.Length -gt 0) {
                $State.Buffer = $State.Buffer.Substring(0, $State.Buffer.Length - 1)
                Select-MenuFirstVisible -State $State
            }
        }
        ([ConsoleKey]::Enter) {
            $text = $State.Buffer.Trim()
            if ($text -ieq 'q') { $result.Action = 'Exit'; break }

            $target = $null
            if ($text.Length -eq 0) {
                $target = $State.Modules | Where-Object { $_.Key -eq $State.SelectedKey } | Select-Object -First 1
            }
            else {
                if ($text -match '^\d+$') {
                    $target = $State.Modules | Where-Object { $_.Key -eq $text.PadLeft(2, '0') } | Select-Object -First 1
                }
                if (-not $target) { $target = $visible | Select-Object -First 1 }
            }

            if ($target) {
                $State.Buffer = ''
                $State.SelectedKey = $target.Key
                $result.Action = 'Run'
                $result.Module = $target
            }
            else {
                Write-MenuMessage -State $State -Text "Modül bulunamadı: $text" -Kind Warning
            }
        }
        default {
            if ($KeyInfo.KeyChar -ne [char]0 -and -not [char]::IsControl($KeyInfo.KeyChar)) {
                $State.Buffer += $KeyInfo.KeyChar
                Select-MenuFirstVisible -State $State
            }
        }
    }
    return $result
}

function Format-Truncated {
    <#
    .SYNOPSIS
        Metni verilen görünür genişliğe kısaltır (sığmazsa sonuna … koyar).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyString()][string]$Text = '',
        [Parameter(Mandatory = $true)][int]$Width
    )

    if ($Width -le 0) { return '' }
    if ($Text.Length -le $Width) { return $Text }
    if ($Width -eq 1) { return '…' }
    return $Text.Substring(0, $Width - 1) + '…'
}

function Get-UpdateNoticeColor {
    <#
    .SYNOPSIS
        Güncelleyici mesaj düzeyine (Success, Info, Warning, Error) karşılık gelen ANSI rengini döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Level
    )

    switch ($Level) {
        'Success' { return $script:Ansi.Green }
        'Warning' { return $script:Ansi.Amber }
        'Error' { return $script:Ansi.Red }
        default { return $script:Ansi.Muted }
    }
}

function Write-MenuBootRow {
    <#
    .SYNOPSIS
        Açılış efektinin tek satırını sol panelde belirtilen satıra yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$Row,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text
    )

    [Console]::Write("$($script:Esc)[$($Row);1H$($script:Esc)[2K$Text$($script:Ansi.Reset)")
}

function Show-MenuBootSequence {
    <#
    .SYNOPSIS
        Eski M-Files açılış metnini sol panelin üst bölümünde satır satır gösterir.
    .DESCRIPTION
        Animasyon tamamlandıktan sonra aynı satırlar menü çiziminde korunur; sağ panelde
        açılış animasyonu çalıştırılmaz.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$Width,
        [Parameter(Mandatory = $true)][string]$HostInfo,
        [Parameter(Mandatory = $true)][string]$Operator,
        [Parameter(Mandatory = $true)][string]$SessionInfo,
        [Parameter(Mandatory = $true)][string]$Version,
        [Parameter(Mandatory = $true)][string]$Updated,
        [PSCustomObject]$UpdateNotice
    )

    $a = $script:Ansi
    $inner = [Math]::Max(12, $Width - 2)
    $barWidth = [Math]::Max(4, [Math]::Min(22, $inner - 6))
    $rule = $a.Line + (' ' + ('─' * $inner)) + $a.Reset

    # İlk temizleme yalnızca sol panelin sanal terminal ekranını etkiler.
    [Console]::Write("$($script:Esc)[2J$($script:Esc)[H")

    # Güncelleyici mesajı varsa BOOT SEQUENCE'ten önce, en üst satırda gösterilir; diğer satırlar bir aşağı kayar.
    $o = 0
    if ($null -ne $UpdateNotice) {
        $o = 1
        $noticeColor = Get-UpdateNoticeColor -Level $UpdateNotice.Level
        Write-MenuBootRow -Row 1 -Text (" $noticeColor$(Format-Truncated -Text $UpdateNotice.Message -Width $inner)$($a.Reset)")
        Start-Sleep -Milliseconds 250
    }

    Write-MenuBootRow -Row ($o + 1) -Text (" $($a.Amber)$(Format-Truncated -Text '[i] BOOT SEQUENCE INITIATED...' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 150
    Write-MenuBootRow -Row ($o + 2) -Text (" $($a.Amber)$(Format-Truncated -Text '[i] Connecting to M-FILES CORE INTERFACE...' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 150
    $userLine = '[✔] USER: {0} authenticated' -f $Operator
    Write-MenuBootRow -Row ($o + 3) -Text (" $($a.Green)$(Format-Truncated -Text $userLine -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 150
    Write-MenuBootRow -Row ($o + 4) -Text (" $($a.Amber)$(Format-Truncated -Text '[i] Loading cognitive modules...' -Width $inner)$($a.Reset)")

    # Dolum adım adım ilerler; son kare kalıcı menü görünümüyle aynıdır.
    foreach ($percent in @(0, 20, 40, 60, 80, 100)) {
        $filled = [int][Math]::Round($barWidth * $percent / 100)
        $bar = ('▓' * $filled).PadRight($barWidth, '░')
        $progress = "$bar $percent%"
        Write-MenuBootRow -Row ($o + 5) -Text " $($a.Amber)$progress$($a.Reset)"
        Start-Sleep -Milliseconds 85
    }

    Start-Sleep -Milliseconds 100
    Write-MenuBootRow -Row ($o + 6) -Text (" $($a.Green)$(Format-Truncated -Text '[✔] Session authorized under CODEX: [ALPHA-0927]' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 130
    Write-MenuBootRow -Row ($o + 7) -Text $rule
    Start-Sleep -Milliseconds 100
    Write-MenuBootRow -Row ($o + 8) -Text (" $($a.Bold)$($a.Amber)$(Format-Truncated -Text '🔸 💾 M-FILES TOOLKIT 🔸' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 100
    Write-MenuBootRow -Row ($o + 9) -Text (" $($a.Text)$(Format-Truncated -Text 'Welcome back, Operator.' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 90
    Write-MenuBootRow -Row ($o + 10) -Text (" $($a.Muted)$(Format-Truncated -Text ('🗓️ Session: {0}' -f $SessionInfo) -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 90
    Write-MenuBootRow -Row ($o + 11) -Text (" $($a.Muted)$(Format-Truncated -Text ('🖥️ Host: {0}' -f $HostInfo) -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 90
    Write-MenuBootRow -Row ($o + 12) -Text (" $($a.Muted)$(Format-Truncated -Text 'Program developed by: Mustafa Orakcı' -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 90
    Write-MenuBootRow -Row ($o + 13) -Text (" $($a.Muted)$(Format-Truncated -Text ('🔖 Sürüm: {0} • Güncellenme: {1}' -f $Version, $Updated) -Width $inner)$($a.Reset)")
    Start-Sleep -Milliseconds 90
    Write-MenuBootRow -Row ($o + 14) -Text $rule
    Start-Sleep -Milliseconds 300
}


function Write-MenuScreenDiff {
    <#
    .SYNOPSIS
        Yalnızca değişen satırları terminale yazar; gezinme sırasında tüm menünün titremesini önler.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][AllowEmptyString()][string[]]$Lines,
        [AllowNull()][string[]]$PreviousLines,
        [switch]$Clear
    )

    $esc = $script:Esc
    if ($Clear) {
        [Console]::Write("$($script:Esc)[2J")
        $PreviousLines = $null
    }

    $oldCount = if ($null -eq $PreviousLines) { 0 } else { $PreviousLines.Count }
    $maxCount = [Math]::Max($Lines.Count, $oldCount)

    for ($index = 0; $index -lt $maxCount; $index++) {
        $newLine = if ($index -lt $Lines.Count) { [string]$Lines[$index] } else { '' }
        $oldLine = if ($index -lt $oldCount) { [string]$PreviousLines[$index] } else { $null }
        if ($null -ne $oldLine -and $newLine -ceq $oldLine) { continue }
        [Console]::Write("$($script:Esc)[$($index + 1);1H$($script:Esc)[2K$newLine$($script:Ansi.Reset)")
    }
}

function Format-MenuScreen {
    <#
    .SYNOPSIS
        Sol menünün tüm ekranını ANSI renk kodlarıyla tek bir metin olarak üretir.
    .NOTES
        Eski M-Files karşılama alanı sol panelin en üstünde kalıcıdır.
        Sağ panelde açılış animasyonu gösterilmez; sağ panel modül bekleme alanı olarak kalır.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory = $true)][psobject]$State,
        [Parameter(Mandatory = $true)][psobject]$PanelStatus,
        [Parameter(Mandatory = $true)][string]$HostInfo,
        [Parameter(Mandatory = $true)][string]$Operator,
        [Parameter(Mandatory = $true)][string]$SessionInfo,
        [Parameter(Mandatory = $true)][string]$Version,
        [Parameter(Mandatory = $true)][string]$Updated,
        [PSCustomObject]$UpdateNotice,
        [switch]$Clear,
        [switch]$AsLines
    )

    $a = $script:Ansi
    $height = [Math]::Max(20, $State.Height)
    $inner = [Math]::Max(12, $State.Width - 2)
    $culture = Get-MenuCulture
    $lines = New-Object -TypeName 'string[]' -ArgumentList $height
    for ($i = 0; $i -lt $height; $i++) { $lines[$i] = '' }

    # Güncelleyici mesajı varsa en üst satırı kullanır; karşılama alanı ve liste bir satır aşağı kayar.
    $o = 0
    if ($null -ne $UpdateNotice) {
        $o = 1
        $lines[0] = " $(Get-UpdateNoticeColor -Level $UpdateNotice.Level)$(Format-Truncated -Text $UpdateNotice.Message -Width $inner)$($a.Reset)"
    }

    $rule = $a.Line + (' ' + ('─' * $inner)) + $a.Reset
    $barWidth = [Math]::Max(4, [Math]::Min(22, $inner - 6))
    $progress = ('▓' * $barWidth) + ' 100%'

    # Kalıcı karşılama/kimlik alanı: eski ekrandaki içerik sol panelde tutulur.
    $userLineText = '[✔] USER: {0} authenticated' -f $Operator
    $sessionLineText = '🗓️ Session: {0}' -f $SessionInfo
    $hostLineText = '🖥️ Host: {0}' -f $HostInfo
    $versionLineText = '🔖 Sürüm: {0} • Güncellenme: {1}' -f $Version, $Updated
    $lines[$o + 0]  = " $($a.Amber)$(Format-Truncated -Text '[i] BOOT SEQUENCE INITIATED...' -Width $inner)$($a.Reset)"
    $lines[$o + 1]  = " $($a.Amber)$(Format-Truncated -Text '[i] Connecting to M-FILES CORE INTERFACE...' -Width $inner)$($a.Reset)"
    $lines[$o + 2]  = " $($a.Green)$(Format-Truncated -Text $userLineText -Width $inner)$($a.Reset)"
    $lines[$o + 3]  = " $($a.Amber)$(Format-Truncated -Text '[i] Loading cognitive modules...' -Width $inner)$($a.Reset)"
    $lines[$o + 4]  = " $($a.Amber)$progress$($a.Reset)"
    $lines[$o + 5]  = " $($a.Green)$(Format-Truncated -Text '[✔] Session authorized under CODEX: [ALPHA-0927]' -Width $inner)$($a.Reset)"
    $lines[$o + 6]  = $rule
    $lines[$o + 7]  = " $($a.Bold)$($a.Amber)$(Format-Truncated -Text '🔸 💾 M-FILES TOOLKIT 🔸' -Width $inner)$($a.Reset)"
    $lines[$o + 8]  = " $($a.Text)$(Format-Truncated -Text 'Welcome back, Operator.' -Width $inner)$($a.Reset)"
    $lines[$o + 9]  = " $($a.Muted)$(Format-Truncated -Text $sessionLineText -Width $inner)$($a.Reset)"
    $lines[$o + 10] = " $($a.Muted)$(Format-Truncated -Text $hostLineText -Width $inner)$($a.Reset)"
    $lines[$o + 11] = " $($a.Muted)$(Format-Truncated -Text 'Program developed by: Mustafa Orakcı' -Width $inner)$($a.Reset)"
    $lines[$o + 12] = " $($a.Muted)$(Format-Truncated -Text $versionLineText -Width $inner)$($a.Reset)"
    $lines[$o + 13] = $rule

    $commandText = if ($State.Buffer.Length -eq 0) {
        "$($a.Muted)$(Format-Truncated -Text 'numara ya da ad yaz…' -Width ($inner - 3))$($a.Reset)"
    }
    else {
        "$($a.Text)$(Format-Truncated -Text $State.Buffer -Width ($inner - 4))$($a.Amber)█$($a.Reset)"
    }
    $lines[$o + 14] = " $($a.Green)▶$($a.Reset) $commandText"
    $lines[$o + 15] = $rule

    # Satırlar: grup başlıkları ve modüller.
    $visible = @(Get-VisibleModule -State $State)
    $rows = [System.Collections.Generic.List[object]]::new()
    $lastGroup = $null
    foreach ($module in $visible) {
        if ($module.GroupFolder -ne $lastGroup) {
            $rows.Add([PSCustomObject]@{ Type = 'Header'; Module = $module })
            $lastGroup = $module.GroupFolder
        }
        $rows.Add([PSCustomObject]@{ Type = 'Item'; Module = $module })
    }

    $areaTop = 16 + $o
    $areaHeight = [Math]::Max(1, $height - $areaTop - 3)

    $selectedRow = -1
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($rows[$i].Type -eq 'Item' -and $rows[$i].Module.Key -eq $State.SelectedKey) { $selectedRow = $i; break }
    }
    if ($selectedRow -ge 0) {
        $firstWanted = $selectedRow
        if ($selectedRow -gt 0 -and $rows[$selectedRow - 1].Type -eq 'Header') { $firstWanted = $selectedRow - 1 }
        if ($firstWanted -lt $State.Offset) { $State.Offset = $firstWanted }
        if ($selectedRow -ge $State.Offset + $areaHeight) { $State.Offset = $selectedRow - $areaHeight + 1 }
    }
    $State.Offset = [Math]::Max(0, [Math]::Min($State.Offset, [Math]::Max(0, $rows.Count - $areaHeight)))

    if ($rows.Count -eq 0) {
        $lines[$areaTop] = " $($a.Muted)(eşleşen modül yok)$($a.Reset)"
    }

    for ($slot = 0; $slot -lt $areaHeight; $slot++) {
        $rowIndex = $State.Offset + $slot
        if ($rowIndex -ge $rows.Count) { break }
        $row = $rows[$rowIndex]
        $module = $row.Module

        if ($row.Type -eq 'Header') {
            $header = Format-Truncated -Text $module.Group.ToUpper($culture) -Width ($inner - 1)
            $lines[$areaTop + $slot] = " $($a.AmberDim)$header$($a.Reset)"
            continue
        }

        $running = ($PanelStatus.State -eq 'busy' -and $PanelStatus.Key -eq $module.Key)
        $marker = if ($running) { ' ▶' } else { '' }
        $titleWidth = $inner - 4 - $marker.Length
        $title = Format-Truncated -Text $module.Title -Width $titleWidth
        if ($module.Key -eq $State.SelectedKey) {
            $plain = (" $($module.Key) $title$marker").PadRight($inner)
            $lines[$areaTop + $slot] = " $($a.OnAmber)$plain$($a.Reset)"
        }
        else {
            $markerText = if ($running) { "$($a.Green)$marker$($a.Reset)" } else { '' }
            $lines[$areaTop + $slot] = " $($a.Amber)$($module.Key)$($a.Text) $title$($a.Reset)$markerText"
        }
    }

    # Alt bölüm: ayraç, mesaj ya da panel durumu, tuş ipuçları.
    $lines[$height - 3] = $rule
    $messageActive = ($State.Message.Length -gt 0 -and [DateTime]::UtcNow -lt $State.MessageUntil)
    if ($messageActive) {
        $color = switch ($State.MessageKind) { 'Warning' { $a.Amber } 'Error' { $a.Red } 'Success' { $a.Green } default { $a.Text } }
        $lines[$height - 2] = " $color$(Format-Truncated -Text $State.Message -Width $inner)$($a.Reset)"
    }
    else {
        $statusLine = switch ($PanelStatus.State) {
            'busy' { "$($a.Green)▶ Sağ panel: $($PanelStatus.Key) $($PanelStatus.Title) çalışıyor" }
            'offline' { "$($a.Red)✖ Sağ panel yanıt vermiyor" }
            default { "$($a.Muted)● Sağ panel hazır" }
        }
        $plainLength = [Regex]::Replace($statusLine, "$([char]27)\[[0-9;]*m", '').Length
        if ($plainLength -gt $inner) { $statusLine = $statusLine.Substring(0, $statusLine.Length - ($plainLength - $inner)) }
        $lines[$height - 2] = " $statusLine$($a.Reset)"
    }
    $lines[$height - 1] = " $($a.Muted)$(Format-Truncated -Text '↑↓ gez · Enter aç · Esc temizle · Q çıkış' -Width $inner)$($a.Reset)"

    if ($AsLines) {
        return [PSCustomObject]@{ Lines = $lines }
    }

    $builder = [System.Text.StringBuilder]::new()
    if ($Clear) { $null = $builder.Append("$($script:Esc)[2J") }
    for ($i = 0; $i -lt $height; $i++) {
        $null = $builder.Append("$($script:Esc)[$($i + 1);1H$($script:Esc)[2K$($lines[$i])$($a.Reset)")
    }
    return $builder.ToString()
}
