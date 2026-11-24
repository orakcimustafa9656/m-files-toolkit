#requires -Version 5.1
<#
Optik Medya Arşiv Sistemi
v1.8.6
Klavye kontrollü SOURCE / PROJECT arşiv yöneticisi.
- SOURCE ve PROJECT aynı listeleme tasarımını kullanır.
- SOURCE: seç, boyut hesapla, sürücü değiştir, PROJECT'e ekle.
- PROJECT: sanal arşivi düzenle, SOURCE'dan gelen içerik üzerinde silme ve yazma işlemleri yap.
- Reparse point türleri tag ile ayrıştırılır; junction/symlink/uyumsuz reparse atlanır.
- Cloud Files / OneDrive türü reparse pointler erişilebildiği sürece seçilebilir.
- Her yazma ayrı SESSION_yyyyMMdd_HHmmss_fff altında tutulur.
- Eski SESSION'lara dokunulmaz; append modunda önceki UDF filesystem import edilir.
- Geçmiş SESSION'larla NEW / CHANGED / PROTECTED karşılaştırması yapılmaz.
- MANIFEST SHA256 yazılacak içeriğin kimliğidir.
- Image önce geçici IMG dosyasına alınır; fiziksel yazma bu sabit image stream'den yapılır.
- VERIFY güncel kaynakla değil, medya dosya SHA256 değerleri ile MANIFEST SHA256 değerlerini karşılaştırır.
- WRITE / FINALIZE / VERIFY / LOG sonuçları birbirinden ayrı raporlanır. Finalize doğrulanınca AutoVerify engellenmez.
- UDF 2.50 + Redundant UDF Metadata korunur.
- ARŞİV ara klasörü kullanılmaz.
- v1.7.5: Salt-okunur fiziksel medya, genel WRITE_PROTECTED kontrolünden önce ayrı sınıflandırılır; finalized sonrası doğrulama ve AutoVerify zinciri korunur.
- v1.7.8: F9 sonrası ana döngüden çıkış hatası giderildi; MANIFEST-RECHECK uyuşmazlığında yazma iptal edilir; kapasite ön kontrolü; UI mesajları; geçici dosya temizliği; UI hata günlüğü.
# - v1.7.9: F9 yazma akışı onarıldı; TEMP kapasite hatası loglanıyor; preflight kapasite UX; volume bekleme 120 sn; UDF import recursive; log çift kontrolü kaldırıldı.
# - v1.7.10: "medya yok" durumu ayrı raporlanıyor (NO_MEDIA profili); kullanıcıya net "medya takın" mesajı.
# - v1.7.11: Mode (d-a-r-h-s) ve Tarih (dd.MM.yyyy HH:mm) sütunları eklendi; üst/alt border rengi SOURCE için mavi, PROJECT için kırmızı.
# - v1.7.12: UI sütun genişlikleri 'dir' çıktısına yaklaştırıldı; Mode 8, Tarih 20 (5 boşluk), Boyut 14 (sağa hizalı).
# - v1.7.13: Sütun başlıkları İngilizce (dir uyumlu); PROJECT klasörlerinin LastWriteUtc değeri dolduruluyor; çakışma ekranında A = birleştir (recursive merge).
# - v1.7.14: Büyük fonksiyonlar anlamlı alt fonksiyonlara bölündü (Write-UiScreen, Remove-ProjectSelectedNode, Select-WriteSettings, Show-WriteCompletionSummary). Davranış değişmedi.
# - v1.8.0: Manifest ve UDF Image aşamalarına ilerleme çubuğu eklendi; olu kod temizlendi (~100 satır).
# - v1.8.1: UDF klasör oluşturma aşamasına ilerleme çubuğu; boş OptikMedyaArsiv klasörü temizliği; artık $script:JustPaused satırı silindi.
# - v1.8.2: F1 (klasör ekle), F9 (yazma hazırlık), F2 (boyut), F3 (yenile) işlemlerine status satırında ilerleme geri bildirimi. Navigasyon ekranı korunur.
# - v1.8.3: Merge progress bar atlanan dosyaları da sayar (bar %100'e gider); (+N yeni) göstergesi; F4 silme sonrası anlık yenileme; büyük işlem sonrası mesaj 1.5sn sabit.
# - v1.8.4: Konsol UI katmanı eklendi (Write-UiRule/Section/Row, Read-UiConfirm); yazma özeti ve onay ekranı yeni tasarıma taşındı (Enter = iptal, uyarı ve kullanım yüzdesi). Yazma mantığı değişmedi.
# - v1.8.5: Hedef konsol akışı: banner, YAZMA HIZI menüsü, [n/4] adım satırları, canlı ilerleme alanı (çubuk + tek satırda akan yol), kısa volume label (MEDYA_yyyyMMdd_HHmmss), yazıcı adı boşlukları temizlendi. Yazma/IMAPI mantığı değişmedi.
# - v1.8.5-stil: Tekil isim adlandırması, isimli parametreler, tüm fonksiyonlarda [CmdletBinding()] ve [OutputType()]; davranış değişmedi.
# - v1.8.6-kalite: Reparse tag, salt-okunur medya türü ve yeniden denenebilir HRESULT değerleri isimli sabit oldu; iç içe Measure-ProjectDirectoryCount dışarı alındı; Remove-ProjectNodesWithRollback asıl hatayı baştan saklar ve sayım yenileme hatasını mesaja ekler; 16 fonksiyona comment-based help eklendi. Davranış değişmedi (tek istisna: rollback sonrası sayım yenilenemezse mesaj bunu belirtir).
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',
    Justification='Bu betik renkli, imleç konumlu bir konsol arayüzüdür; Write-Host görüntüleme katmanıdır, iş çıktısı değildir. Veri çıktıları nesne olarak döner.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',
    Justification='New/Set/Update/Reset/Start/Stop fonksiyonları bellek içi nesne ve arayüz durumunu değiştirir; kalıcı değişiklikler (yazma, silme) zaten açık kullanıcı onayı ekranlarından geçer.')]
param()

# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}


$script:SavedErrorActionPreference = $ErrorActionPreference
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# UYGULAMA
# ============================================================
$script:AppVersion = '1.8.6'
$script:AppDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {[IO.Path]::GetFullPath($PSScriptRoot)} else {(Get-Location).Path}
$script:LogDir = Get-MFTModuleDataPath -ModuleId 'optikmedya' -Category Logs

# ============================================================
# YAZMA MOTORU SABİTLERİ
# ============================================================
$script:FsiFileSystemUDF = 4
$script:UdfRevision250 = 0x250
$script:BytesPerBlock = 2048
$script:Settings = [pscustomobject]@{AutoVerify=$true; EjectAfter=$true; Logging=$true; Multisession=$false; Finalize=$true}
$script:HResultMap = @{
    '0xC0AA0202'='Medya mevcut değil.'
    '0xC0AA0203'='Medya uyumsuz veya fiziksel format desteklenmiyor.'
    '0xC0AA0204'='Medya ters takılmış.'
    '0xC0AA0205'='Medya/sürücü hazır hale geliyor.'
    '0xC0AA0206'='Medya formatlama işlemi devam ediyor.'
    '0xC0AA0207'='Sürücü veya medya meşgul.'
    '0xC0AA0208'='Geçersiz mode parametreleri.'
    '0xC0AA020B'='Recorder özelliği mevcut ancak etkin/current değil.'
    '0xC0AA020D'='Recorder komutu zaman aşımına uğradı.'
    '0xC0AA0510'='Diske veri yazılırken hata oluştu.'
    '0xC0AAB15A'='IMAPI mevcut medya için yazmaya uyumlu bir UDF revision desteklemiyor. Multisession append yapılamaz.'
}
$script:MediaState = @{
    OverwriteOnly=0x1; Blank=0x2; Appendable=0x4; FinalSession=0x8
    Damaged=0x400; EraseRequired=0x800; NonEmptySession=0x1000
    WriteProtected=0x2000; Finalized=0x4000; UnsupportedMedia=0x8000
}
# Yalnizca bir kez yazilabilen medya: CD-R, DVD+R, DVD+R DL, DVD-R, DVD-R DL, BD-R.
# Yeniden yazilabilir turler (CD-RW, DVD+/-RW, DVD-RAM, BD-RE) bilerek yok: arsiv silinebilir olmamali.
$script:WritableMediaTypes = @(2,6,8,9,11,18)
# Okunur-yalnız fiziksel medya türleri (IMAPI_MEDIA_PHYSICAL_TYPE): 1 = CD-ROM, 4 = DVD-ROM, 17 = BD-ROM.
$script:ReadOnlyPhysicalMediaTypes = @(1,4,17)
# Sürücü/medya hazırlanırken dönen, beklenip yeniden denenebilen IMAPI HRESULT değerleri (hazır hale geliyor / formatlama sürüyor / meşgul).
$script:ImapiRetryableHResults = @('0xC0AA0205','0xC0AA0206','0xC0AA0207')
# Reparse tag değerleri ondalık yazılır: 0xA0000003 gibi hex literal'ler PowerShell'de negatif Int32 okunur.
$script:ReparseTag = [pscustomobject]@{
    MountPoint      = [uint32]2684354563   # 0xA0000003
    SymbolicLink    = [uint32]2684354572   # 0xA000000C
    WslSymlink      = [uint32]2684354589   # 0xA000001D
    ProjFs          = [uint32]2415919132   # 0x9000001C
    ProjFsTombstone = [uint32]2684354594   # 0xA0000022
    CloudFilesMask  = [uint32]4294905855   # 0xFFFF0FFF (IO_REPARSE_TAG_CLOUD_MASK'in tersi)
    CloudFilesBase  = [uint32]2415919130   # 0x9000001A
    NameSurrogate   = [uint32]536870912    # 0x20000000
}

# 1. UYGULAMA BAŞLANGICI

function Initialize-AppDirectory {
    [CmdletBinding()]
    param()
    foreach ($dir in @($script:AppDir,$script:LogDir)) {
        if (-not (Test-Path -LiteralPath $dir)) {[void](New-Item -ItemType Directory -Path $dir -Force)}
    }
}
function Set-ArchiveThemeColor {
    <#
    .SYNOPSIS
        Runtime'da tema rengini değiştirir.
    .DESCRIPTION
        Belirtilen renk token'ını yeni ConsoleColor değeriyle günceller.
        Değişiklik sonraki Write-UiScreen çağrısında etkili olur.
    .PARAMETER Token
        Renk token adı (Primary, Success, Warning, Danger, Muted, Info,
        Text, Subtle, SourceBorder, ProjectBorder, SelectionBg,
        SelectionFg, PartialBg, FullSelectBg, FullSelectFg).
    .PARAMETER Color
        Yeni ConsoleColor değeri.
    .EXAMPLE
        PS> Set-ArchiveThemeColor -Token Primary -Color Magenta
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Primary','Success','Warning','Danger','Muted','Info','Text','Subtle',
                     'SourceBorder','ProjectBorder','SelectionBg','SelectionFg',
                     'PartialBg','FullSelectBg','FullSelectFg')]
        [string]$Token,
        [Parameter(Mandatory)]
        [ConsoleColor]$Color
    )
    $script:Theme.Color[$Token] = $Color
}

function Test-ArchiveThemeContrast {
    <#
    .SYNOPSIS
        Tema kontrastlarını doğrular.
    .DESCRIPTION
        Seçili satır ve kısmi seçim arka plan/ön plan renklerinin
        aynı olup olmadığını kontrol eder. Aynıysa uyarı verir.
    .OUTPUTS
        PSCustomObject { Ok, Warnings }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $warnings = New-Object System.Collections.Generic.List[string]
    if ($script:Theme.Color.SelectionBg -eq $script:Theme.Color.SelectionFg) {
        [void]$warnings.Add('SelectionBg ve SelectionFg aynı — seçili satır okunamaz.')
    }
    if ($script:Theme.Color.PartialBg -eq $script:Theme.Color.PartialFg) {
        [void]$warnings.Add('PartialBg ve PartialFg aynı — kısmi seçim okunamaz.')
    }
    if ($script:Theme.Color.FullSelectBg -eq $script:Theme.Color.FullSelectFg) {
        [void]$warnings.Add('FullSelectBg ve FullSelectFg aynı — seçili satır okunamaz.')
    }
    foreach ($w in $warnings) { Write-Warning $w }
    return [pscustomobject]@{
        Ok       = ($warnings.Count -eq 0)
        Warnings = @($warnings)
    }
}
function Clear-StaleTempArtifact {
    # Çökme/iptal sonrası kalan eski geçici image ve manifest klasörlerini siler.
    # Yalnızca belirtilen süreden eski olanlara dokunur; kullanımda olan dosyalar silinemez ve sessizce atlanır.
    [CmdletBinding()]
    param([int]$OlderThanHours=2)
    try {
        $limit=(Get-Date).AddHours(-$OlderThanHours)
        foreach ($f in @(Get-ChildItem -LiteralPath $env:TEMP -Filter 'OptikMedyaArsiv_SESSION_*.img' -File -Force -ErrorAction SilentlyContinue)) {
            if ($f.LastWriteTime -lt $limit) {Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue}
        }
        $dir=Join-Path $env:TEMP 'OptikMedyaArsiv'
        if (Test-Path -LiteralPath $dir -PathType Container) {
            foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -Force -ErrorAction SilentlyContinue)) {
                if ($d.LastWriteTime -lt $limit) {Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue}
            }
            # Alt klasorler silindikten sonra ana klasor bos kaldiysa onu da sil.
            try {
                $remaining=@(Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue)
                if ($remaining.Count -eq 0) {
                    $dirItem=Get-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue
                    if ($null -ne $dirItem -and $dirItem.LastWriteTime -lt $limit) {
                        Remove-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue
                    }
                }
            } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}
function Write-UiErrorLog {
    [CmdletBinding()]
    param([object]$Err,[string]$Where)
    try {
        $line='{0} [{1}] {2}' -f (Get-Date -Format 's'),$Where,$Err.Exception.ToString()
        Add-Content -LiteralPath (Join-Path $script:LogDir 'ui_errors.txt') -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}
function Wait-ConsoleInput {
    [CmdletBinding()]
    param([string]$Message='Devam etmek için Enter')
    Write-Host ''
    Write-Host $Message -ForegroundColor $script:Theme.Color.Muted
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    try {
        while ($true) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Enter) { break }
        }
    } catch {
        try { [void](Read-Host) } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
    try {
        Start-Sleep -Milliseconds 150
        while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}
function Format-Size {
    [CmdletBinding()]
    param([Int64]$Bytes)
    if ($Bytes -lt 1KB) {return ('{0} B' -f $Bytes)}
    if ($Bytes -lt 1MB) {return ('{0:N1} KB' -f ($Bytes/1KB))}
    if ($Bytes -lt 1GB) {return ('{0:N2} MB' -f ($Bytes/1MB))}
    return ('{0:N2} GB' -f ($Bytes/1GB))
}
# DESTEK KATMANI COM / IMAPI
function Unregister-ComObject {
    [CmdletBinding()]
    param([object]$ComObject)
    if ($null -eq $ComObject) {return}
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            try {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}
function Get-HResultHex {
    <#
    .SYNOPSIS
    Bir istisnadan HRESULT değerini onaltılık metin olarak çıkarır.
    .DESCRIPTION
    Dönen metin (ör. 0xC0AA0205), $script:ImapiRetryableHResults ve hata tablosundaki anahtarlarla karşılaştırılabilir.
    .PARAMETER Exception
    HRESULT bilgisi okunacak istisna.
    .OUTPUTS
    System.String
    .EXAMPLE
    PS> Get-HResultHex -Exception $_.Exception
    #>
    [OutputType([string])]
    [CmdletBinding()]
    param([System.Exception]$Exception)
    try {
        $hr=[Runtime.InteropServices.Marshal]::GetHRForException($Exception)
        return ('0x{0:X8}' -f ($hr -band 0xFFFFFFFF))
    } catch {return ''}
}
function Get-FriendlyError {
    [OutputType([string])]
    [CmdletBinding()]
    param([System.Exception]$Exception)
    if ($null -eq $Exception) { return 'Bilinmeyen hata.' }
    $hr=Get-HResultHex -Exception $Exception
    if ($hr -and $script:HResultMap.ContainsKey($hr)) {return ('{0} {1}' -f $hr,$script:HResultMap[$hr])}
    if ($hr) {return ('{0} {1}' -f $hr,$Exception.Message)}
    if ([string]::IsNullOrWhiteSpace($Exception.Message)) { return 'Bilinmeyen hata.' }
    return $Exception.Message
}
function Get-BooleanText { [OutputType([string])] [CmdletBinding()] param([bool]$Value); if ($Value) {'AÇIK'} else {'KAPALI'} }

# ============================================================
# UI KATMANI (renk ve stil tek yerden)
# Cyan: başlık | DarkGray: etiket/not | White: değer | Green: başarılı | Yellow: dikkat/soru | Red: yalnız hata
# ============================================================
$script:UiWidth                 = 62
$script:UiListWidth             = 120
$script:UiNameMinWidth          = 24
$script:UiColumnReserved        = 85
$script:BufferSizeBytes         = 1MB
$script:VolumeWaitSeconds       = 120
$script:VerifyVolumeWaitSeconds = 300
# ============================================================
# TEMA KATMANI — tüm görsel ayarlar burada
# Renkler: ConsoleColor adı (string). Sembol: string.
# Runtime değişiklik: Set-ArchiveThemeColor
# ============================================================
$script:Theme = @{
    Color = @{
        Primary       = 'Cyan'
        Success       = 'Green'
        Warning       = 'Yellow'
        Danger        = 'Red'
        Muted         = 'DarkGray'
        Info          = 'DarkCyan'
        Text          = 'White'
        Subtle        = 'Gray'
        SourceBorder  = 'Blue'
        ProjectBorder = 'Red'
        SelectionBg   = 'Gray'
        SelectionFg   = 'Black'
        PartialBg     = 'DarkYellow'
        PartialFg     = 'Black'
        FullSelectBg  = 'DarkRed'
        FullSelectFg  = 'White'
    }
    Symbol = @{
        Check      = '✓'
        Cross      = '✗'
        Warn       = '⚠'
        Play       = '▶'
        Arrow      = '→'
        Ellipsis   = '…'
        Bullet     = '·'
        BarFull    = '█'
        BarEmpty   = '░'
        Rule       = '─'
        DoubleRule = '═'
    }
}
function Write-UiRule {
    <#
    .SYNOPSIS
    Yatay ayırıcı çizgi yazdırır.
    .EXAMPLE
    PS> Write-UiRule
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    Write-Host (' ' + ('─' * ($script:UiWidth - 2))) -ForegroundColor $script:Theme.Color.Muted
}
function Write-UiSection {
    <#
    .SYNOPSIS
    Bölüm başlığı yazdırır.
    .PARAMETER Title
    Bölüm başlığı metni.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Title)
    $tail = [Math]::Max(3, $script:UiWidth - $Title.Length - 5)
    Write-Host ''
    Write-Host (' ── {0} {1}' -f $Title,('─' * $tail)) -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''
}
function Write-UiRow {
    <#
    .SYNOPSIS
    Etiket-değer çiftini hizalı satır olarak yazdırır.
    .PARAMETER Label
    Sol taraftaki etiket.
    .PARAMETER Value
    Sağ taraftaki değer.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Label,[string]$Value,[string]$Note='',[ConsoleColor]$ValueColor=$script:Theme.Color.Text,[ConsoleColor]$NoteColor=$script:Theme.Color.Muted)
    Write-Host ('  {0,-14}' -f $Label) -NoNewline -ForegroundColor $script:Theme.Color.Muted
    if ([string]::IsNullOrEmpty($Note)) { Write-Host $Value -ForegroundColor $ValueColor; return }
    Write-Host $Value -NoNewline -ForegroundColor $ValueColor
    Write-Host ('  ' + $Note) -ForegroundColor $NoteColor
}
function Read-UiConfirm {
    <#
    .SYNOPSIS
    Kullanıcıdan Evet/Hayır onayı alır.
    .PARAMETER Question
    Soru metni.
    .OUTPUTS
    System.Boolean
    #>
    [OutputType([bool])]
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Question)
    Write-UiRule
    while ($true) {
        Write-Host ('  {0}   ' -f $Question) -NoNewline -ForegroundColor $script:Theme.Color.Warning
        Write-Host '[E] Evet  [H] Hayır  (varsayılan: H) ' -NoNewline -ForegroundColor $script:Theme.Color.Muted
        switch (([string](Read-Host)).Trim().ToUpperInvariant()) {
            'E'     { return $true }
            'EVET'  { return $true }
            'H'     { return $false }
            'HAYIR' { return $false }
            ''      { return $false }
            default { Write-Host '  E veya H girin.' -ForegroundColor $script:Theme.Color.Warning }
        }
    }
}

$script:UiStepTotal = 4
$script:UiLive = $false
$script:UiTop = 0
$script:UiStepNo = 0
$script:UiStepName = ''
$script:UiClock = [Diagnostics.Stopwatch]::StartNew()
$script:UiLastDraw = 0
$script:UiDirCount = -1

function Get-UiMediaText {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Media)
    $state = if ($Media.IsBlank) {'boş, yazılabilir'} elseif ($Media.IsAppendable) {'yazılabilir (eklemeli)'} else {''}
    if ($state) { return ('{0} · {1}' -f $Media.MediaName,$state) }
    return [string]$Media.MediaName
}
function Write-UiBanner {
    <#
    .SYNOPSIS
    Üst ve alt kenar çizgili banner yazdırır.
    .PARAMETER Title
    Sol ana başlık.
    .PARAMETER Right
    Sağ kısa bilgi.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Title,[string]$Right)
    $inner = $script:UiWidth - 2
    $pad = [Math]::Max(1, $inner - 4 - $Title.Length - $Right.Length)
    Write-Host ('╭' + ('─' * $inner) + '╮') -ForegroundColor $script:Theme.Color.Primary
    Write-Host '│  ' -NoNewline -ForegroundColor $script:Theme.Color.Primary
    Write-Host $Title -NoNewline -ForegroundColor $script:Theme.Color.Text
    Write-Host (' ' * $pad) -NoNewline
    Write-Host $Right -NoNewline -ForegroundColor $script:Theme.Color.Success
    Write-Host '  │' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ('╰' + ('─' * $inner) + '╯') -ForegroundColor $script:Theme.Color.Primary
}
function Write-UiStep {
    <#
    .SYNOPSIS
    Tamamlanmış adımın sonuç satırını yazdırır.
    .PARAMETER No
    Adım numarası.
    .PARAMETER Name
    Adım adı.
    .PARAMETER Result
    Adım sonucu.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([int]$No,[string]$Name,[string]$Result,[string]$Time='',[bool]$Failed=$false)
    $mark = if ($Failed) {'✗ '} else {'✓ '}
    $markColor = if ($Failed) {$script:Theme.Color.Danger} else {$script:Theme.Color.Success}
    $resColor = if ($Failed) {$script:Theme.Color.Danger} else {$script:Theme.Color.Text}
    $gap = [Math]::Max(2, $script:UiWidth - 24 - $Result.Length - $Time.Length)
    Write-Host ('  [{0}/{1}] ' -f $No,$script:UiStepTotal) -NoNewline -ForegroundColor $script:Theme.Color.Muted
    Write-Host ('{0,-14}' -f $Name) -NoNewline -ForegroundColor $script:Theme.Color.Subtle
    Write-Host $mark -NoNewline -ForegroundColor $markColor
    if ($Time) { Write-Host $Result -NoNewline -ForegroundColor $resColor; Write-Host ((' ' * $gap) + $Time) -ForegroundColor $script:Theme.Color.Muted }
    else { Write-Host $Result -ForegroundColor $resColor }
}
function Start-UiStep {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([int]$No,[string]$Name,[Int64]$Count=0,[string]$Unit='')
    $script:UiStepNo = $No
    $script:UiStepName = $Name
    $script:UiClock = [Diagnostics.Stopwatch]::StartNew()
    $script:UiLastDraw = -1000
    $script:UiLive = $false
    try {
        if (-not [Console]::IsOutputRedirected) {
            Write-Host ''
            Write-Host ''
            $script:UiTop = [Console]::CursorTop - 2
            [Console]::CursorVisible = $false
            $script:UiLive = $true
        }
    } catch { $script:UiLive = $false }
    if ($Count -gt 0) { Update-UiStep -Done 0 -Count $Count -Text '' -Unit $Unit }
}
function Update-UiStep {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([Int64]$Done,[Int64]$Count,[string]$Text='',[string]$Unit='')
    if (-not $script:UiLive) { return }
    $now = $script:UiClock.ElapsedMilliseconds
    if ($Done -lt $Count -and ($now - $script:UiLastDraw) -lt 60) { return }
    $script:UiLastDraw = $now
    try {
        $w = [Console]::WindowWidth
        $pct = if ($Count -gt 0) { [int][Math]::Min(100,(100.0 * $Done / $Count)) } else { 0 }
        $tail = '  {0,3}%  {1}/{2} {3}' -f $pct,$Done,$Count,$Unit
        $barW = [Math]::Max(8,[Math]::Min(28,$w - 24 - $tail.Length - 2))
        $filled = [int]($barW * $pct / 100)
        $used = 24 + $barW + $tail.Length
        [Console]::SetCursorPosition(0,$script:UiTop)
        Write-Host ('  [{0}/{1}] ' -f $script:UiStepNo,$script:UiStepTotal) -NoNewline -ForegroundColor $script:Theme.Color.Muted
        Write-Host ('{0,-14}' -f $script:UiStepName) -NoNewline -ForegroundColor $script:Theme.Color.Subtle
        Write-Host '▶ ' -NoNewline -ForegroundColor $script:Theme.Color.Warning
        Write-Host ($script:Theme.Symbol.BarFull * $filled) -NoNewline -ForegroundColor $script:Theme.Color.Primary
        Write-Host ($script:Theme.Symbol.BarEmpty * ($barW - $filled)) -NoNewline -ForegroundColor $script:Theme.Color.Muted
        Write-Host $tail -NoNewline -ForegroundColor $script:Theme.Color.Muted
        Write-Host (' ' * [Math]::Max(0,$w - 1 - $used))
        $room = [Math]::Max(10,$w - 10)
        $p = $Text
        if ($p.Length -gt $room) { $p = '…' + $p.Substring($p.Length - ($room - 1)) }
        Write-Host ('        ' + $p.PadRight($room)) -ForegroundColor $script:Theme.Color.Muted
    } catch { $script:UiLive = $false }
}
function Show-UiBusy {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Message,[string]$Note='')
    if (-not $script:UiLive) { return }
    try {
        $w = [Console]::WindowWidth
        [Console]::SetCursorPosition(0,$script:UiTop)
        Write-Host ('  [{0}/{1}] ' -f $script:UiStepNo,$script:UiStepTotal) -NoNewline -ForegroundColor $script:Theme.Color.Muted
        Write-Host ('{0,-14}' -f $script:UiStepName) -NoNewline -ForegroundColor $script:Theme.Color.Subtle
        Write-Host '▶ ' -NoNewline -ForegroundColor $script:Theme.Color.Warning
        Write-Host $Message -NoNewline -ForegroundColor $script:Theme.Color.Text
        Write-Host (' ' * [Math]::Max(0,$w - 1 - 24 - $Message.Length))
        Write-Host (('        ' + $Note).PadRight([Math]::Max(0,$w - 2))) -ForegroundColor $script:Theme.Color.Muted
    } catch { $script:UiLive = $false }
}
function Clear-UiLiveLine {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    $w = [Console]::WindowWidth
    [Console]::SetCursorPosition(0,$script:UiTop)
    Write-Host (' ' * [Math]::Max(0,$w - 1)) -NoNewline
    [Console]::SetCursorPosition(0,$script:UiTop + 1)
    Write-Host (' ' * [Math]::Max(0,$w - 1)) -NoNewline
    [Console]::SetCursorPosition(0,$script:UiTop)
}
function Stop-UiStep {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Result)
    $timeText = '{0:N1} sn' -f $script:UiClock.Elapsed.TotalSeconds
    if ($script:UiLive) {
        $script:UiLive = $false
        try {
            Clear-UiLiveLine
            Write-UiStep -No $script:UiStepNo -Name $script:UiStepName -Result $Result -Time $timeText
            [Console]::SetCursorPosition(0,$script:UiTop + 1)
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try { [Console]::CursorVisible = $true } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
    else { Write-UiStep -No $script:UiStepNo -Name $script:UiStepName -Result $Result -Time $timeText }
}
function Stop-UiLiveAbort {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if (-not $script:UiLive) { return }
    $script:UiLive = $false
    try {
        Clear-UiLiveLine
        Write-UiStep -No $script:UiStepNo -Name $script:UiStepName -Result 'başarısız' -Time '' -Failed $true
        [Console]::SetCursorPosition(0,$script:UiTop + 1)
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    try { [Console]::CursorVisible = $true } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}

# ============================================================
# YAZMA MODU BAŞLIĞI (F9 sonrası kalıcı, yanıp sönen)
# ============================================================
$script:WriteModeActive       = $false
$script:WriteModeHeaderHeight = 2
$script:WriteModeHeaderText   = ''
$script:AnsiSupported         = $false

function Test-AnsiSupport {
    [OutputType([bool])]
    [CmdletBinding()]
    param()
    if ([Console]::IsOutputRedirected) { return $false }
    if ($env:WT_SESSION)   { return $true }
    if ($env:TERM_PROGRAM) { return $true }
    try {
        if ([Environment]::OSVersion.Version.Major -ge 10) { return $true }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    return $false
}

function Write-WriteModeHeader {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    try {
        if (-not $script:WriteModeActive) { return }
        $width = [Console]::WindowWidth
        if ($width -le 1) { return }
        $saved = 0
        try { $saved = [Console]::CursorTop } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        [Console]::SetCursorPosition(0, 0)
        $content = " ▶▶ $($script:WriteModeHeaderText) ◀◀ "
        $padLen  = [Math]::Max(0, ($width - 1) - $content.Length)
        $padded  = $content + (' ' * $padLen)
        if ($script:AnsiSupported) {
            $esc  = [char]27
            [Console]::Write("$esc[1;5;37;41m$padded$esc[0m")
        } else {
            [Console]::BackgroundColor = [ConsoleColor]::DarkRed
            [Console]::ForegroundColor = [ConsoleColor]::White
            [Console]::Write($padded)
            [Console]::BackgroundColor = [ConsoleColor]::Black
            [Console]::ForegroundColor = [ConsoleColor]::Gray
        }
        [Console]::SetCursorPosition(0, [Math]::Max($script:WriteModeHeaderHeight, $saved))
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}

function Start-WriteModeHeader {
    [CmdletBinding()]
    param([string]$Text)
    try {
        if ([Console]::IsOutputRedirected) { $script:WriteModeActive = $false; return }
        $script:AnsiSupported         = Test-AnsiSupport
        $script:WriteModeHeaderText   = $Text
        $script:WriteModeHeaderHeight = 2
        Clear-Host
        $script:WriteModeActive = $true
        Write-WriteModeHeader
        try {
            $width = [Console]::WindowWidth
            [Console]::SetCursorPosition(0, 1)
            [Console]::ForegroundColor = [ConsoleColor]::DarkRed
            [Console]::Write(('═' * ($width - 1)))
            [Console]::ForegroundColor = [ConsoleColor]::Gray
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        [Console]::SetCursorPosition(0, $script:WriteModeHeaderHeight)
    } catch {
        $script:WriteModeActive = $false
        Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message)
    }
}

function Update-WriteModeHeader {
    [CmdletBinding()]
    param([string]$Text)
    if (-not $script:WriteModeActive) { return }
    $script:WriteModeHeaderText = $Text
    Write-WriteModeHeader
}

function Clear-WriteModeBody {
    [CmdletBinding()]
    param()
    if (-not $script:WriteModeActive) { Clear-Host; return }
    try {
        if ([Console]::IsOutputRedirected) { return }
        $width  = [Console]::WindowWidth
        $height = [Console]::WindowHeight
        $blank  = ' ' * ($width - 1)
        for ($r = $script:WriteModeHeaderHeight; $r -lt ($height - 1); $r++) {
            [Console]::SetCursorPosition(0, $r)
            [Console]::Write($blank)
        }
        [Console]::SetCursorPosition(0, $script:WriteModeHeaderHeight)
    } catch {
        try { Clear-Host } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
}

function Stop-WriteModeHeader {
    [CmdletBinding()]
    param()
    $script:WriteModeActive = $false
}

# DESTEK KATMANI  PATH / PROJECT
function ConvertTo-FsiNormalizedPath {
    [OutputType([string])]
    [CmdletBinding()]
    param([string]$Path)
    if ($null -eq $Path) {return '\'}
    $clean=$Path.Replace('/','\').Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) {return '\'}
    if (-not $clean.StartsWith('\')) {$clean='\' + $clean}
    $clean=$clean.TrimEnd('\')
    if ($clean.Length -eq 0) {return '\'}
    return $clean
}
function Join-FsiPath {
    [CmdletBinding()]
    param([string]$BasePath,[string]$RelativePath)
    $base=ConvertTo-FsiNormalizedPath $BasePath
    $relative=ConvertTo-FsiNormalizedPath $RelativePath
    if ($relative -eq '\') {return $base}
    if ($base -eq '\') {return $relative}
    return ConvertTo-FsiNormalizedPath (($base.TrimEnd('\')) + '\' + ($relative.TrimStart('\')))
}

# ============================================================
# PROJECT AĞACI
# ============================================================
$script:ProjectRoot=$null
$script:ProjectCurrent=$null

function New-ProjectNode {
    [CmdletBinding()]
    param([string]$Name,[ValidateSet('ROOT','DIR','FILE')][string]$Type,[string]$SourcePath=$null,[Int64]$Length=0,[string]$SourceRoot='')
    return [pscustomobject]@{
        Name=$Name; Type=$Type; SourcePath=$SourcePath; SourceRoot=$SourceRoot; Length=[Int64]$Length
        SourceExists=$false; LastWriteUtc=[datetime]::MinValue; Selected=$false
        TotalCount=0; SelectedCount=0
        Children=New-Object System.Collections.ArrayList; Parent=$null
    }
}
# 2. KULLANICI ARAYÜZÜ
function Update-NodeCountRecursive {
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node) { return }
    if ($Node.Type -eq 'FILE') {
        $Node.TotalCount = 1
        $Node.SelectedCount = if ([bool]$Node.Selected) { 1 } else { 0 }
        return
    }
    $Node.TotalCount = 0
    $Node.SelectedCount = 0
    foreach ($c in @($Node.Children)) {
        Update-NodeCountRecursive $c
        $Node.TotalCount    += [int]$c.TotalCount
        $Node.SelectedCount += [int]$c.SelectedCount
    }
}
function Update-CountsUpward {
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node) { return }
    if ($Node.Type -eq 'FILE') {
        $Node.TotalCount = 1
        $Node.SelectedCount = if ([bool]$Node.Selected) { 1 } else { 0 }
    }
    else {
        $Node.TotalCount = 0
        $Node.SelectedCount = 0
        foreach ($c in @($Node.Children)) {
            $Node.TotalCount    += [int]$c.TotalCount
            $Node.SelectedCount += [int]$c.SelectedCount
        }
    }
    if ($null -ne $Node.Parent) {
        Update-CountsUpward $Node.Parent
    }
}
function Test-CountersConsistency {
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node) { return }
    if ($Node.Type -eq 'FILE') {
        $expectedTotal = 1
        $expectedSelected = if ([bool]$Node.Selected) { 1 } else { 0 }
    } else {
        $expectedTotal = 0
        $expectedSelected = 0
        foreach ($c in @($Node.Children)) {
            Test-CountersConsistency $c
            $expectedTotal    += [int]$c.TotalCount
            $expectedSelected += [int]$c.SelectedCount
        }
    }
    if ([int]$Node.TotalCount -ne $expectedTotal) { Write-Warning ("Drift TotalCount '{0}': beklenen {1}, mevcut {2}" -f $Node.Name,$expectedTotal,$Node.TotalCount) }
    if ([int]$Node.SelectedCount -ne $expectedSelected) { Write-Warning ("Drift SelectedCount '{0}': beklenen {1}, mevcut {2}" -f $Node.Name,$expectedSelected,$Node.SelectedCount) }
}
function Initialize-Project {
    [CmdletBinding()]
    param()
    $script:ProjectRoot=New-ProjectNode -Name 'ARŞİV' -Type ROOT
    $script:ProjectRoot.SourceExists=$true
    $script:ProjectCurrent=$script:ProjectRoot
    $script:CachedSelectedCount=$null
    $script:StatusSummaryText=''
}
function Get-ProjectChild {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([object]$Node=$script:ProjectCurrent)
    if ($null -eq $Node) {return @()}
    return @($Node.Children | Sort-Object @{Expression={if ($_.Type -eq 'DIR') {0} else {1}}},Name)
}
function Find-ProjectChild {
    [CmdletBinding()]
    param([object]$Parent,[string]$Name)
    foreach ($child in @($Parent.Children)) {
        if ([string]::Equals([string]$child.Name,$Name,[StringComparison]::OrdinalIgnoreCase)) {return $child}
    }
    return $null
}
function Get-ProjectRelativePath {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node -or $Node.Type -eq 'ROOT') {return '\'}
    $parts=New-Object System.Collections.Generic.List[string]
    $n=$Node
    while ($null -ne $n -and $n.Type -ne 'ROOT') {$parts.Insert(0,[string]$n.Name);$n=$n.Parent}
    return ('\' + [string]::Join('\',$parts.ToArray()))
}
function ConvertTo-ProjectNormalizedPath {
    [OutputType([string])]
    [CmdletBinding()]
    param([string]$Path)
    if ($null -eq $Path) {return ''}
    $clean=$Path.Trim().Replace('/','\')
    if ($clean.Length -eq 0) {return ''}
    if ([string]::Equals($clean,'Project',[StringComparison]::OrdinalIgnoreCase)) {return ''}
    if ($clean.StartsWith('Project\',[StringComparison]::OrdinalIgnoreCase)) {$clean=$clean.Substring(8)}
    elseif ($clean.StartsWith('\')) {$clean=$clean.TrimStart('\')}
    return $clean.Trim('\')
}
function Get-ProjectPathPart {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([string]$Path)
    $clean=ConvertTo-ProjectNormalizedPath $Path
    if ([string]::IsNullOrWhiteSpace($clean)) {return @()}
    return @($clean -split '\\' | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
}
function Resolve-ProjectNode {
    [CmdletBinding()]
    param([string]$Path,[object]$StartNode=$script:ProjectCurrent)
    $clean=if ($null -eq $Path) {''} else {$Path.Trim()}
    $isRooted=$clean.StartsWith('Project\',[StringComparison]::OrdinalIgnoreCase) -or
               [string]::Equals($clean,'Project',[StringComparison]::OrdinalIgnoreCase) -or
               $clean.StartsWith('\')
    $current=if ($isRooted) {$script:ProjectRoot} else {$StartNode}
    foreach ($part in @(Get-ProjectPathPart $clean)) {
        if ($part -eq '.') {continue}
        if ($part -eq '..') {
            if ($current.Type -eq 'ROOT') {return $null}
            $current=$current.Parent
            continue
        }
        $child=Find-ProjectChild -Parent $current -Name $part
        if ($null -eq $child) {return $null}
        $current=$child
    }
    return $current
}
function Test-ValidProjectName {
    [OutputType([bool])]
    [CmdletBinding()]
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -in @('.','..') -or $Name.EndsWith(' ') -or $Name.EndsWith('.')) {return $false}
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {if ($Name.IndexOf($char) -ge 0) {return $false}}
    $base=$Name.TrimEnd('.').ToUpperInvariant()
    $reserved=@('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
    return $base -notin $reserved
}
# ============================================================
# PROJECT METADATA
# ============================================================
function Update-ProjectMetadataRecursive {
    [CmdletBinding()]
    param([object]$Node=$script:ProjectRoot)
    foreach ($child in @($Node.Children)) {
        if ($child.Type -eq 'FILE') {
            if ([string]::IsNullOrWhiteSpace([string]$child.SourcePath)) {$child.SourceExists=$false;continue}
            if (Test-Path -LiteralPath $child.SourcePath -PathType Leaf) {
                try {
                    $info=Get-Item -LiteralPath $child.SourcePath -Force
                    $child.Length=[Int64]$info.Length
                    $child.LastWriteUtc=$info.LastWriteTimeUtc
                    $child.SourceExists=$true
                } catch {$child.SourceExists=$false}
            } else {$child.SourceExists=$false}
            Update-UiProgress -Increment 1
        } elseif ($child.Type -eq 'DIR') {
            if (-not [string]::IsNullOrWhiteSpace([string]$child.SourcePath)) {
                if (Test-Path -LiteralPath $child.SourcePath -PathType Container) {
                    try {
                        $dirInfo=Get-Item -LiteralPath $child.SourcePath -Force -ErrorAction Stop
                        $child.LastWriteUtc=$dirInfo.LastWriteTimeUtc
                    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
                }
            }
            Update-ProjectMetadataRecursive $child
        }
    }
}
function Get-SourceItemOrThrow {
    [CmdletBinding()]
    param([string]$LiteralPath)
    try { return (Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop) }
    catch {
        if ($_.CategoryInfo.Category -eq 'ObjectNotFound' -or $_.Exception -is [System.Management.Automation.ItemNotFoundException]) {
            throw ('Kaynak bulunamadı: {0}' -f $LiteralPath)
        }
        throw ('Kaynak okunamadı: {0} - {1}' -f $LiteralPath,$_.Exception.Message)
    }
}
function Add-ProjectSourceFile {
    [CmdletBinding()]
    param([object]$Parent,[string]$SourcePath,[string]$SourceRoot,[ref]$AddedNodes)
    $item=Get-SourceItemOrThrow -LiteralPath $SourcePath
    if ($item.PSIsContainer) {throw ('Kaynak dosya değil: {0}' -f $SourcePath)}
    $rp=Get-ReparsePointInfo -Item $item
    if (-not $rp.CanSelect) {
        Write-Host ('⚠ Kaynak atlandı: {0} ({1})' -f $item.FullName,$rp.Display) -ForegroundColor $script:Theme.Color.Warning
        return
    }
    $existing=Find-ProjectChild -Parent $Parent -Name $item.Name
    if ($null -ne $existing) {throw ('Project içinde aynı isim zaten var: {0}' -f $item.Name)}
    $node=New-ProjectNode -Name $item.Name -Type FILE -SourcePath $item.FullName -Length ([Int64]$item.Length) -SourceRoot $SourceRoot
    $node.SourceExists=$true;$node.LastWriteUtc=$item.LastWriteTimeUtc;$node.Selected=$false;$node.Parent=$Parent
    $node.TotalCount=1;$node.SelectedCount=0
    [void]$Parent.Children.Add($node);[void]$AddedNodes.Value.Add($node)
    Update-UiProgress -Increment 1 -IsNew
}
function Add-SourceDirectoryRecursive {
    [CmdletBinding()]
    param([object]$Parent,[string]$SourcePath,[string]$SourceRoot,[ref]$AddedNodes)
    $dirItem=Get-SourceItemOrThrow -LiteralPath $SourcePath
    if (-not $dirItem.PSIsContainer) {throw ('Kaynak klasör değil: {0}' -f $SourcePath)}
    $rp=Get-ReparsePointInfo -Item $dirItem
    if (-not $rp.CanTraverse -or -not $rp.CanSelect) {
        Write-Host ('⚠ Kaynak klasör atlandı: {0} ({1})' -f $dirItem.FullName,$rp.Display) -ForegroundColor $script:Theme.Color.Warning
        return
    }
    $existing=Find-ProjectChild -Parent $Parent -Name $dirItem.Name
    if ($null -ne $existing) {throw ('Project içinde aynı isim zaten var: {0}' -f $dirItem.Name)}
    $newDir=New-ProjectNode -Name $dirItem.Name -Type DIR -SourcePath $dirItem.FullName -SourceRoot $SourceRoot
    $newDir.SourceExists=$true;$newDir.Selected=$false;$newDir.Parent=$Parent
    try { $newDir.LastWriteUtc = $dirItem.LastWriteTimeUtc } catch { $newDir.LastWriteUtc = [datetime]::MinValue }
    [void]$Parent.Children.Add($newDir);[void]$AddedNodes.Value.Add($newDir)
    Update-UiProgress -IncrementDir 1
    try {$children=@(Get-ChildItem -LiteralPath $dirItem.FullName -Force -ErrorAction Stop)}
    catch {throw ('Klasör okunamadı: {0} - {1}' -f $dirItem.FullName,$_.Exception.Message)}
    foreach ($child in $children) {
        $childPath=[string]$child.FullName
        $childRp=Get-ReparsePointInfo -Item $child
        if ($child.PSIsContainer) {
            if (-not $childRp.CanTraverse -or -not $childRp.CanSelect) {
                Write-Host ('⚠ Dallanma atlandı: {0} ({1})' -f $childPath,$childRp.Display) -ForegroundColor $script:Theme.Color.Warning
                continue
            }
            Add-SourceDirectoryRecursive -Parent $newDir -SourcePath $childPath -SourceRoot $SourceRoot -AddedNodes $AddedNodes
        } else {
            if (-not $childRp.CanSelect) {
                Write-Host ('⚠ Dosya atlandı: {0} ({1})' -f $childPath,$childRp.Display) -ForegroundColor $script:Theme.Color.Warning
                continue
            }
            Add-ProjectSourceFile -Parent $newDir -SourcePath $childPath -SourceRoot $SourceRoot -AddedNodes $AddedNodes
        }
    }
}
function Add-MergedSourceDirectory {
    [CmdletBinding()]
    param(
        [object]$Parent,
        [string]$SourcePath,
        [string]$SourceRoot,
        [ref]$AddedNodes,
        [ref]$SkippedCount
    )

    $dirItem=Get-SourceItemOrThrow -LiteralPath $SourcePath
    if (-not $dirItem.PSIsContainer) { throw ('Kaynak klasör değil: {0}' -f $SourcePath) }

    $rp=Get-ReparsePointInfo -Item $dirItem
    if (-not $rp.CanTraverse -or -not $rp.CanSelect) {
        Write-Host ('⚠ Merge atlandı (reparse): {0} ({1})' -f $dirItem.FullName,$rp.Display) -ForegroundColor $script:Theme.Color.Warning
        return
    }

    $existing=Find-ProjectChild -Parent $Parent -Name $dirItem.Name

    if ($null -eq $existing) {
        Add-SourceDirectoryRecursive -Parent $Parent -SourcePath $SourcePath -SourceRoot $SourceRoot -AddedNodes $AddedNodes
        return
    }

    if ($existing.Type -ne 'DIR') {
        Write-Host ('⚠ Merge atlandı (hedefte dosya): {0}' -f $dirItem.Name) -ForegroundColor $script:Theme.Color.Warning
        [void]$SkippedCount.Value++
        Update-UiProgress -Increment 1
        return
    }

    try {$children=@(Get-ChildItem -LiteralPath $dirItem.FullName -Force -ErrorAction Stop)}
    catch {throw ('Klasör okunamadı: {0} - {1}' -f $dirItem.FullName,$_.Exception.Message)}

    foreach ($child in $children) {
        $childPath=[string]$child.FullName
        $childRp=Get-ReparsePointInfo -Item $child

        if ($child.PSIsContainer) {
            if (-not $childRp.CanTraverse -or -not $childRp.CanSelect) {
                Write-Host ('⚠ Merge dallanma atlandı: {0} ({1})' -f $childPath,$childRp.Display) -ForegroundColor $script:Theme.Color.Warning
                [void]$SkippedCount.Value++
                Update-UiProgress -IncrementDir 1
                continue
            }
            Add-MergedSourceDirectory -Parent $existing -SourcePath $childPath -SourceRoot $SourceRoot -AddedNodes $AddedNodes -SkippedCount $SkippedCount
        } else {
            if (-not $childRp.CanSelect) {
                Write-Host ('⚠ Merge dosya atlandı: {0} ({1})' -f $childPath,$childRp.Display) -ForegroundColor $script:Theme.Color.Warning
                [void]$SkippedCount.Value++
                Update-UiProgress -Increment 1
                continue
            }
            $existingFile=Find-ProjectChild -Parent $existing -Name $child.Name
            if ($null -ne $existingFile) {
                [void]$SkippedCount.Value++
                Update-UiProgress -Increment 1
                continue
            }
            Add-ProjectSourceFile -Parent $existing -SourcePath $childPath -SourceRoot $SourceRoot -AddedNodes $AddedNodes
        }
    }
}

function Get-ProjectArchiveFile {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([object]$Node=$script:ProjectRoot)
    $result=New-Object System.Collections.ArrayList
    foreach ($child in @($Node.Children)) {
        if ($child.Type -eq 'FILE') {
            [void]$result.Add([pscustomobject]@{
                SourcePath=[string]$child.SourcePath
                SourceRoot=[string]$child.SourceRoot
                DestinationPath=Get-ProjectRelativePath $child
                Name=[string]$child.Name
                Length=[Int64]$child.Length
                Sha256=''
                SourceExists=[bool]$child.SourceExists
                LastWriteUtc=$child.LastWriteUtc
            })
        } elseif ($child.Type -eq 'DIR') {
            foreach ($f in @(Get-ProjectArchiveFile $child)) {[void]$result.Add($f)}
        }
    }
    return @($result)
}
function Get-ProjectDirectoryPath {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([object]$Node=$script:ProjectRoot)
    $result=New-Object System.Collections.ArrayList
    foreach ($child in @($Node.Children)) {
        if ($child.Type -eq 'DIR') {
            [void]$result.Add((Get-ProjectRelativePath $child))
            foreach ($p in @(Get-ProjectDirectoryPath $child)) {[void]$result.Add($p)}
        }
    }
    return @($result)
}
function Measure-ProjectDirectoryCount {
    <#
    .SYNOPSIS
    PROJECT ağacında bir düğümün altındaki klasörleri özyinelemeli sayar.
    .DESCRIPTION
    Verilen düğümün tüm alt düğümlerini dolaşır ve türü DIR olanları toplar. Düğümün kendisi sayıya dahil edilmez.
    .PARAMETER Node
    Sayımın başlayacağı PROJECT düğümü (genellikle $script:ProjectRoot).
    .OUTPUTS
    System.Int32 - alt klasör sayısı.
    .EXAMPLE
    PS> Measure-ProjectDirectoryCount $script:ProjectRoot
    #>
    [OutputType([int])]
    [CmdletBinding()]
    param([object]$Node)
    $localCount=0
    foreach ($childNode in @($Node.Children)) {
        if ($childNode.Type -eq 'DIR') {$localCount++;$localCount+=Measure-ProjectDirectoryCount $childNode}
    }
    return $localCount
}
function Get-ProjectStatistic {
    <#
    .SYNOPSIS
    PROJECT için özet istatistik üretir.
    .DESCRIPTION
    Önce klasör üst verilerini günceller, sonra dosya sayısını, klasör sayısını, kaynağı mevcut dosyaların toplam baytını ve kaynağı bulunamayan dosya sayısını döndürür.
    .OUTPUTS
    PSCustomObject: FileCount, DirectoryCount, TotalBytes, MissingCount.
    .EXAMPLE
    PS> Get-ProjectStatistic
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param()
    Update-ProjectMetadataRecursive
    $files=@(Get-ProjectArchiveFile)
    [Int64]$bytes=0
    $missing=0
    foreach ($f in $files) {
        if ($f.SourceExists) {$bytes+=[Int64]$f.Length} else {$missing++}
    }
    return [pscustomobject]@{
        FileCount=$files.Count
        DirectoryCount=[int](Measure-ProjectDirectoryCount $script:ProjectRoot)
        TotalBytes=$bytes
        MissingCount=$missing
    }
}

function New-ArchiveVolumeLabel {
    [OutputType([string])]
    [CmdletBinding()]
    param([string]$SessionName)
    if ($SessionName -match '^SESSION_(\d{8}_\d{6})') {$label="MEDYA_$($Matches[1])"}
    else {$label="MEDYA_$SessionName"}
    if ($label.Length -gt 32) {$label=$label.Substring(0,32)}
    return $label
}

# ============================================================
# SESSION / MANIFEST
# ============================================================
function Get-SessionTimestampFromName {
    [CmdletBinding()]
    param([string]$SessionName)
    if ($SessionName -notmatch '^SESSION_(\d{8}_\d{6}_\d{3})$') {throw ('Geçersiz SESSION adı: {0}' -f $SessionName)}
    return $matches[1]
}
function Get-SessionManifestFileName {
    [CmdletBinding()]
    param([string]$SessionName)
    return ('MANIFEST_{0}.txt' -f (Get-SessionTimestampFromName $SessionName))
}
function New-ArchiveSessionIdentity {
    [CmdletBinding()]
    param([string]$VolumePath)
    while ($true) {
        $timestamp=Get-Date -Format 'yyyyMMdd_HHmmss_fff'
        $sessionName="SESSION_$timestamp"
        $manifestFileName="MANIFEST_$timestamp.txt"
        if ([string]::IsNullOrWhiteSpace($VolumePath)) {
            return [pscustomobject]@{SessionName=$sessionName;ManifestFileName=$manifestFileName}
        }
        $candidatePath=Join-Path $VolumePath $sessionName
        if (-not (Test-Path -LiteralPath $candidatePath -PathType Container)) {
            return [pscustomobject]@{SessionName=$sessionName;ManifestFileName=$manifestFileName}
        }
        Start-Sleep -Milliseconds 20
    }
}
function Get-ArchiveManifestIndex {
    [OutputType([hashtable])]
    [CmdletBinding()]
    param([string[]]$Lines)

    $index=@{}

    foreach ($line in @($Lines)) {
        $line=[string]$line

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $line=$line.Trim()

        # Manifest tablo başlığı
        if ($line -match '^No\s*\|\s*SourceRoot\s*\|\s*ArchivePath\s*\|\s*Size\s*\|\s*SHA256$') {
            continue
        }

        # Gerçek "|" karakterine göre 5 kolona ayır
        $parts=@($line -split '\s*\|\s*',5)

        if ($parts.Count -ne 5) {
            continue
        }

        $archivePath=[string]$parts[2].Trim()
        $sizeText=[string]$parts[3].Trim()
        $sha256=[string]$parts[4].Trim().ToUpperInvariant()

        if (
            [string]::IsNullOrWhiteSpace($archivePath) -or
            $sha256 -notmatch '^[A-F0-9]{64}$'
        ) {
            continue
        }

        [Int64]$size=0

        if (
            -not [Int64]::TryParse(
                $sizeText,
                [Globalization.NumberStyles]::Integer,
                [Globalization.CultureInfo]::InvariantCulture,
                [ref]$size
            )
        ) {
            continue
        }

        $archivePath=ConvertTo-FsiNormalizedPath $archivePath

        $index[$archivePath]=[pscustomobject]@{
            Size   = $size
            Sha256 = $sha256
        }
    }

    return $index
}
# ============================================================
# MANIFEST DOSYASI
# ============================================================
function Get-ManifestFileIdentity {
    [CmdletBinding()]
    param([string]$SourcePath,[int]$RecordIndex=0)

    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        throw ('Manifest oluşturulurken kaynak yolu boş: kayıt {0}' -f $RecordIndex)
    }

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw ('Manifest oluşturulurken kaynak dosya bulunamadı: {0}' -f $SourcePath)
    }

    $info=Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
    $hash=(Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
    $infoAfter=Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop

    if ([Int64]$info.Length -ne [Int64]$infoAfter.Length) {
        throw ('Manifest oluşturulurken kaynak dosya boyutu değişti: {0}' -f $SourcePath)
    }

    if ($info.LastWriteTimeUtc -ne $infoAfter.LastWriteTimeUtc) {
        throw ('Manifest oluşturulurken kaynak dosya değişti: {0}' -f $SourcePath)
    }

    return [pscustomobject]@{
        Length = [Int64]$infoAfter.Length
        Sha256 = $hash
    }
}

function Add-ManifestSourceRoot {
    [CmdletBinding()]
    param([System.Collections.Generic.List[string]]$Lines,[string[]]$SourceRoots)

    [void]$Lines.Add('SOURCE ROOTS')
    [void]$Lines.Add('------------------------------------------------------------')

    $rootNo=1
    foreach ($root in @(
        $SourceRoots |
        Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)} |
        Sort-Object -Unique
    )) {
        [void]$Lines.Add("$rootNo. $root")
        $rootNo++
    }
}

function Add-ManifestFileEntry {
    [OutputType([long])]
    [CmdletBinding()]
    param([System.Collections.Generic.List[string]]$Lines,[object[]]$Files)

    [Int64]$totalBytes=0
    $index=0
    $fileCount=[Math]::Max(1,@($Files).Count)

    foreach ($f in @($Files)) {
        $index++
        Update-UiStep -Done ($index-1) -Count $fileCount -Text ([string]$f.SourcePath) -Unit 'dosya'
        $identity=Get-ManifestFileIdentity -SourcePath ([string]$f.SourcePath) -RecordIndex $index

        # MANIFEST gerçek dosya kimliğini burada üretir.
        $f.Length=$identity.Length
        $f.Sha256=$identity.Sha256
        $totalBytes+=$identity.Length

        [void]$Lines.Add(
            ('{0} | {1} | {2} | {3} | {4}' -f
                $index,
                $f.SourceRoot,
                $f.DestinationPath,
                $identity.Length,
                $identity.Sha256
            )
        )

    }

    return [Int64]$totalBytes
}

function Write-ManifestText {
    [CmdletBinding()]
    param([string]$Path,[System.Collections.Generic.List[string]]$Lines)

    $utf8=New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText(
        $Path,
        [string]::Join([Environment]::NewLine,$Lines.ToArray()),
        $utf8
    )
}

function Get-ManifestFileHash {
    [CmdletBinding()]
    param([string]$Path)

    return (
        Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop
    ).Hash.ToUpperInvariant()
}

function New-ManifestFile {
    [CmdletBinding()]
    param(
        [object[]]$Files,
        [string[]]$SourceRoots,
        [object]$Recorder,
        [object]$Media,
        [object]$SelectedSpeed,
        [string]$VolumeLabel,
        [bool]$Multisession,
        [bool]$Finalize,
        [string]$SessionName
    )

    $tempDir=Join-Path $env:TEMP ('OptikMedyaArsiv\_'+[Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $tempDir -Force)
    $manifestFileName=Get-SessionManifestFileName $SessionName
    $manifestPath=Join-Path $tempDir $manifestFileName

    try {
        $recorderName=''
        $firmware=''
        if ($null -ne $Recorder) {
            $recorderName='{0} {1}' -f $Recorder.VendorId,$Recorder.ProductId
            $firmware=[string]$Recorder.ProductRevision
        }

        $lines=New-Object System.Collections.Generic.List[string]
        [void]$lines.Add('OPTİK MEDYA ARŞİV MANİFESTO')
        [void]$lines.Add("Oluşturma Tarihi : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$lines.Add("Uygulama        : $script:AppVersion")
        [void]$lines.Add("Session         : $SessionName")
        [void]$lines.Add("Medya           : $($Media.MediaName)")
        [void]$lines.Add("Kaydedici       : $recorderName")
        [void]$lines.Add("Firmware        : $firmware")
        [void]$lines.Add("Yazma Hızı      : $($SelectedSpeed.X)x")
        [void]$lines.Add("Volume Label    : $VolumeLabel")
        [void]$lines.Add("Çok Oturumlu    : $(Get-BooleanText $Multisession)")
        [void]$lines.Add("Finalize        : $(Get-BooleanText $Finalize)")
        [void]$lines.Add('Dosya Sistemi   : UDF 2.50')
        [void]$lines.Add('Metadata        : Redundant UDF Metadata')
        [void]$lines.Add("Dosya Sayısı    : $(@($Files).Count)")
        [void]$lines.Add('')

        Add-ManifestSourceRoot -Lines $lines -SourceRoots $SourceRoots

        [void]$lines.Add('')
        [void]$lines.Add('FILES')
        [void]$lines.Add('----------------------------------------------------------------------------------------------------------------')
        [void]$lines.Add('No | SourceRoot | ArchivePath | Size | SHA256')

        [Int64]$totalBytes=Add-ManifestFileEntry -Lines $lines -Files $Files
        [void]$lines.Add('')
        [void]$lines.Add("Toplam Veri     : $(Format-Size $totalBytes)")

        Write-ManifestText -Path $manifestPath -Lines $lines
        $manifestHash=Get-ManifestFileHash -Path $manifestPath

        return [pscustomobject]@{
            Path        = $manifestPath
            TempDir     = $tempDir
            TotalBytes  = $totalBytes
            FileName    = $manifestFileName
            SessionName = $SessionName
            Sha256      = $manifestHash
        }
    }
    catch {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }
}
function Get-ManifestIndexCanonicalSha256 {
    [CmdletBinding()]
    param(
        [hashtable]$Index
    )

    $lines = New-Object System.Collections.Generic.List[string]

    foreach ($path in @($Index.Keys | Sort-Object)) {
        $entry = $Index[$path]

        [void]$lines.Add(
            ('{0}|{1}|{2}' -f
                [string]$path,
                [Int64]$entry.Size,
                [string]$entry.Sha256
            )
        )
    }

    $utf8 = New-Object System.Text.UTF8Encoding($false)

    $text = [string]::Join(
        "`n",
        $lines.ToArray()
    )

    $bytes = $utf8.GetBytes($text)

    $sha = $null

    try {
        $sha = [Security.Cryptography.SHA256]::Create()

        return (
            [BitConverter]::ToString(
                $sha.ComputeHash($bytes)
            ).Replace('-','')
        ).ToUpperInvariant()
    }
    finally {
        if ($null -ne $sha) {
            $sha.Dispose()
        }
    }
}


function New-ManifestRecheck {
    [CmdletBinding()]
    param(
        [object[]]$Files
    )

    $recheckIndex = @{}

    foreach ($f in @($Files)) {

        $sourcePath = [string]$f.SourcePath

        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw 'MANIFEST-RECHECK: kaynak yolu boş.'
        }

        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw (
                'MANIFEST-RECHECK: kaynak dosya bulunamadı: {0}' -f
                $sourcePath
            )
        }

        # Hash öncesi durum
        $infoBefore = Get-Item `
            -LiteralPath $sourcePath `
            -Force `
            -ErrorAction Stop

        $hash = (
            Get-FileHash `
                -LiteralPath $sourcePath `
                -Algorithm SHA256 `
                -ErrorAction Stop
        ).Hash.ToUpperInvariant()

        # Hash sonrası durum
        $infoAfter = Get-Item `
            -LiteralPath $sourcePath `
            -Force `
            -ErrorAction Stop

        if ([Int64]$infoBefore.Length -ne [Int64]$infoAfter.Length) {
            throw (
                'MANIFEST-RECHECK: kaynak dosya hash sırasında değişti: {0}' -f
                $sourcePath
            )
        }

        if ($infoBefore.LastWriteTimeUtc -ne $infoAfter.LastWriteTimeUtc) {
            throw (
                'MANIFEST-RECHECK: kaynak dosya hash sırasında değişti: {0}' -f
                $sourcePath
            )
        }

        $archivePath = ConvertTo-FsiNormalizedPath $f.DestinationPath

        $recheckIndex[$archivePath] = [pscustomobject]@{
            Size   = [Int64]$infoAfter.Length
            Sha256 = $hash
        }
    }

    $recheckHash = Get-ManifestIndexCanonicalSha256 `
        -Index $recheckIndex

    return [pscustomobject]@{
        Name      = 'MANIFEST-RECHECK'
        Index     = $recheckIndex
        Sha256    = $recheckHash
        FileCount = $recheckIndex.Count
    }
}
function Remove-TempManifest {
    [CmdletBinding()]
    param([object]$Manifest)
    if ($null -eq $Manifest) {return}
    try {
        if ($Manifest.TempDir -and (Test-Path -LiteralPath $Manifest.TempDir)) {
            Remove-Item -LiteralPath $Manifest.TempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
}

# ============================================================
# WINDOWS IStream KÖPRÜSÜ
# ============================================================
if (-not ('OptikMedyaArsiv.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
namespace OptikMedyaArsiv
{
    public static class NativeMethods
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SHCreateStreamOnFileEx(string pszFile,uint grfMode,uint dwAttributes,bool fCreate,IStream pstmTemplate,out IStream ppstm);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct FILETIME
        {
            public uint dwLowDateTime;
            public uint dwHighDateTime;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct WIN32_FIND_DATA
        {
            public uint dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint dwReserved0;
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string cAlternateFileName;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "FindFirstFileW")]
        static extern IntPtr FindFirstFile(string name,out WIN32_FIND_DATA data);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FindClose(IntPtr handle);
        static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);
        public static uint GetReparseTag(string path)
        {
            string longPath=path;
            if (!longPath.StartsWith(@"\\?\",StringComparison.Ordinal)) {
                if (longPath.StartsWith(@"\\",StringComparison.Ordinal)) longPath=@"\\?\UNC\"+longPath.Substring(2);
                else longPath=@"\\?\"+longPath;
            }
            WIN32_FIND_DATA data;
            IntPtr h=FindFirstFile(longPath,out data);
            if (h==IntPtr.Zero || h==INVALID_HANDLE_VALUE) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            try {
                if ((data.dwFileAttributes & 0x400) == 0) throw new InvalidOperationException("Öğe reparse point değil.");
                return data.dwReserved0;
            } finally {FindClose(h);}
        }
    }
}
'@
}
# ============================================================
# COM IStream KÖPRÜSÜ
# ============================================================
if (-not ('OptikMedyaArsiv.StreamBridge' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OptikMedyaArsiv
{
    public static class StreamBridge
    {
        private static readonly Guid IID_IStream =
            new Guid("0000000C-0000-0000-C000-000000000046");

        public static long CopyToFile(
            object sourceObject,
            string path,
            long expectedBytes
        )
        {
            if (sourceObject == null)
                throw new ArgumentNullException("sourceObject");

            if (expectedBytes <= 0)
                throw new ArgumentOutOfRangeException("expectedBytes");

            IntPtr pUnknown = IntPtr.Zero;
            IntPtr pStream = IntPtr.Zero;
            IntPtr pcbRead = IntPtr.Zero;

            IStream source = null;

            try
            {
                // COM nesnesinin IUnknown adresini al.
                pUnknown = Marshal.GetIUnknownForObject(sourceObject);

                // IUnknown -> IStream
                Guid iid = IID_IStream;

		int hr = Marshal.QueryInterface(
		    pUnknown,
		    ref iid,
		    out pStream
		);

                if (hr < 0)
                    Marshal.ThrowExceptionForHR(hr);

                // IStream RCW oluştur. GetTypedObjectForIUnknown kendi RCW referansını yönetir.
                source = (IStream)
                    Marshal.GetTypedObjectForIUnknown(
                        pStream,
                        typeof(IStream)
                    );

                // Stream başına dön. source RCW artık pStream ile ayrı temizlenecektir.
                source.Seek(
                    0,
                    0,
                    IntPtr.Zero
                );

                byte[] buffer = new byte[1024 * 1024];

                // ULONG* pcbRead için unmanaged 4 byte.
                pcbRead = Marshal.AllocHGlobal(4);

                long totalRead = 0;

                using (
                    FileStream fs = new FileStream(
                        path,
                        FileMode.Create,
                        FileAccess.Write,
                        FileShare.Read,
                        buffer.Length,
                        FileOptions.SequentialScan
                    )
                )
                {
                    while (totalRead < expectedBytes)
                    {
                        int request = (int)Math.Min(
                            buffer.Length,
                            expectedBytes - totalRead
                        );

                        Marshal.WriteInt32(
                            pcbRead,
                            0
                        );

                        source.Read(
                            buffer,
                            request,
                            pcbRead
                        );

                        int bytesRead =
                            Marshal.ReadInt32(pcbRead);

                        if (bytesRead <= 0)
                        {
                            throw new EndOfStreamException(
                                "IMAPI image stream beklenenden önce sona erdi."
                            );
                        }

                        if (bytesRead > request)
                        {
                            throw new InvalidDataException(
                                "IMAPI IStream geçersiz byte sayısı döndürdü."
                            );
                        }

                        fs.Write(
                            buffer,
                            0,
                            bytesRead
                        );

                        totalRead += bytesRead;
                    }

                    fs.Flush(true);
                }

                return totalRead;
            }
            finally
            {
                if (pcbRead != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.FreeHGlobal(pcbRead);
                    }
                    catch {}
                }

                if (source != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(source);
                    }
                    catch {}
                }

                if (pStream != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.Release(pStream);
                    }
                    catch {}
                }

                if (pUnknown != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.Release(pUnknown);
                    }
                    catch {}
                }
            }
        }
    }
}
'@
}
function Get-ReparsePointInfo {
    <#
    .SYNOPSIS
    Bir öğenin reparse point türünü sınıflandırır.
    .DESCRIPTION
    Junction, symlink, WSL symlink, ProjFS ve name-surrogate türleri atlanır (CanTraverse/CanSelect = $false). Cloud Files / OneDrive türü seçilebilir ve gezilebilir kalır. Tag değerleri $script:ReparseTag içindedir.
    .PARAMETER Item
    FileInfo veya DirectoryInfo nesnesi.
    .OUTPUTS
    PSCustomObject: IsReparse, Tag, Kind, CanTraverse, CanSelect, Display, Reason.
    .EXAMPLE
    PS> Get-ReparsePointInfo -Item (Get-Item -LiteralPath C:\Link)
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param([object]$Item)
    if ($null -eq $Item) {return [pscustomobject]@{IsReparse=$false;Tag=[uint32]0;Kind='NONE';CanTraverse=$true;CanSelect=$true;Display='';Reason=''}}
    $isReparse=((([int]$Item.Attributes) -band [int][IO.FileAttributes]::ReparsePoint) -ne 0)
    if (-not $isReparse) {return [pscustomobject]@{IsReparse=$false;Tag=[uint32]0;Kind='NONE';CanTraverse=$true;CanSelect=$true;Display='NORMAL';Reason=''}}
    try {$tag=[uint32][OptikMedyaArsiv.NativeMethods]::GetReparseTag([string]$Item.FullName)}
    catch {return [pscustomobject]@{IsReparse=$true;Tag=[uint32]0;Kind='UNKNOWN REPARSE';CanTraverse=$false;CanSelect=$false;Display='REPARSE';Reason=$_.Exception.Message}}
    # Tag sabitleri $script:ReparseTag içinde tutulur (ondalık yazılmıştır; gerekçe sabit bloğunda).
    switch ($tag) {
        $script:ReparseTag.MountPoint {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='JUNCTION / MOUNT POINT';CanTraverse=$false;CanSelect=$false;Display='JUNCTION';Reason='Yönlendiren junction/mount point.'}}
        $script:ReparseTag.SymbolicLink {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='SYMBOLIC LINK';CanTraverse=$false;CanSelect=$false;Display='SYMLINK';Reason='Yönlendiren symbolic link.'}}
        $script:ReparseTag.WslSymlink {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='WSL SYMLINK';CanTraverse=$false;CanSelect=$false;Display='SYMLINK';Reason='WSL symbolic link.'}}
        $script:ReparseTag.ProjFs {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='PROJFS';CanTraverse=$false;CanSelect=$false;Display='PROJFS';Reason='Projected File System öğesi; güvenlik için atlandı.'}}
        $script:ReparseTag.ProjFsTombstone {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='PROJFS TOMBSTONE';CanTraverse=$false;CanSelect=$false;Display='PROJFS';Reason='Projected File System tombstone.'}}
        default {
            $cloudMask=$script:ReparseTag.CloudFilesMask
            $cloudBase=$script:ReparseTag.CloudFilesBase
            if (($tag -band $cloudMask) -eq $cloudBase) {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='CLOUD FILES';CanTraverse=$true;CanSelect=$true;Display='CLOUD';Reason='Cloud Files / OneDrive reparse point.'}}
            if (($tag -band $script:ReparseTag.NameSurrogate) -ne 0) {return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='NAME SURROGATE';CanTraverse=$false;CanSelect=$false;Display='LINK / REPARSE';Reason='Yönlendiren name-surrogate reparse point.'}}
            return [pscustomobject]@{IsReparse=$true;Tag=$tag;Kind='UNKNOWN REPARSE';CanTraverse=$false;CanSelect=$false;Display='REPARSE';Reason=('Desteklenmeyen reparse tag: 0x{0:X8}' -f $tag)}
        }
    }
}
function Open-ImapiFileStream {
    [CmdletBinding()]
    param([string]$Path,[bool]$ReadOnly=$true)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {throw ('Stream için dosya bulunamadı: {0}' -f $Path)}
    $stream=$null;$mode=if ($ReadOnly) {0x00000020} else {0x00001002}
    try {
        $hr=[OptikMedyaArsiv.NativeMethods]::SHCreateStreamOnFileEx($Path,[uint32]$mode,0,$false,$null,[ref]$stream)
        if ($hr -ne 0) {throw ('SHCreateStreamOnFileEx başarısız: 0x{0:X8}' -f ($hr -band 0xFFFFFFFF))}
        if ($null -eq $stream) {throw ('SHCreateStreamOnFileEx stream nesnesi döndürmedi: {0}' -f $Path)}
        return $stream
    }
    catch {
        if ($null -ne $stream) {Unregister-ComObject $stream}
        throw
    }
}
function Export-ImageStreamToFile {
    [OutputType([long])]
    [CmdletBinding()]
    param(
        [object]$SourceStream,
        [string]$Path,
        [Int64]$ExpectedBytes
    )

    if ($null -eq $SourceStream) {
        throw 'UDF image stream bulunamadı.'
    }

    if ($ExpectedBytes -le 0) {
        throw 'Beklenen image boyutu geçersiz.'
    }

    try {
        $null =[OptikMedyaArsiv.StreamBridge]::CopyToFile(
            $SourceStream,
            $Path,
            $ExpectedBytes
        )
    }
    catch {
        throw (
            'UDF image geçici dosyaya çıkarılamadı: {0}' -f
            $_.Exception.Message
        )
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Geçici image dosyası oluşmadı.'
    }

    $length=(Get-Item -LiteralPath $Path -Force).Length

    if ($length -le 0) {
        throw 'Geçici image dosyası boş oluştu.'
    }

    if ([Int64]$length -ne [Int64]$ExpectedBytes) {
        throw (
            'Geçici image boyutu beklenen değerden farklı: ' +
            'beklenen {0}, oluşan {1}.' -f
            $ExpectedBytes,
            $length
        )
    }

    return [Int64]$length
}
# ============================================================
# UDF
# ============================================================
function Test-Udf250Support {
    [CmdletBinding()]
    param()
    $fsi=$null
    try {
        $fsi=New-Object -ComObject IMAPI2FS.MsftFileSystemImage
        $fileSystemsSupported=[int]$fsi.FileSystemsSupported
        if (($fileSystemsSupported -band $script:FsiFileSystemUDF) -eq 0) {throw 'IMAPI UDF dosya sistemi desteği bildirmiyor.'}
        foreach ($revision in @($fsi.UDFRevisionsSupported)) {if ([int]$revision -eq $script:UdfRevision250) {return $true}}
        $supportedText=@($fsi.UDFRevisionsSupported | ForEach-Object {'0x{0:X}' -f ([int]$_)}) -join ', '
        throw ('UDF 2.50 desteklenmiyor. Desteklenen revisionlar: {0}' -f $supportedText)
    } finally {if ($null -ne $fsi) {Unregister-ComObject $fsi}}
}
function Get-UdfDirectory {
    [CmdletBinding()]
    param([object]$FileSystemImage,[hashtable]$DirectoryMap,[object]$RootDir,[string]$DestinationDirectory)
    $clean=ConvertTo-FsiNormalizedPath $DestinationDirectory
    if ($clean -eq '\') {Write-Output -NoEnumerate -InputObject $RootDir;return}
    $relative=$clean.Trim('\')
    if ($DirectoryMap.ContainsKey($clean)) {Write-Output -NoEnumerate -InputObject $DirectoryMap[$clean];return}
    $parts=@($relative -split '\\' | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
    $current=$RootDir;$built=''
    foreach ($part in $parts) {
        if ($built.Length -gt 0) {$built+='\'}
        $built+=[string]$part
        $key='\' + $built
        if ($DirectoryMap.ContainsKey($key)) {$current=$DirectoryMap[$key];continue}
        $newDir=$null
        try {$newDir=$FileSystemImage.CreateDirectoryItem([string]$part);[void]$current.Add($newDir)}
        catch {throw ('UDF klasör oluşturma başarısız: {0}`r`nHResult: 0x{1:X8}`r`nMesaj: {2}' -f $key,($_.Exception.HResult -band 0xFFFFFFFF),$_.Exception.Message)}
        $DirectoryMap[$key]=$newDir;$current=$newDir
    }
    Write-Output -NoEnumerate -InputObject $current
}
function Get-FsiDirectoryChild {
    [CmdletBinding()]
    param([object]$Directory)
    if ($null -eq $Directory) {throw 'FSI klasör nesnesi yok.'}
    $countBefore=0
    try {$countBefore=[int]$Directory.Count} catch {$countBefore=-1}
    try {
        foreach ($item in $Directory) {
            Write-Output -NoEnumerate -InputObject $item
        }
    }
    catch {
        throw ('FSI klasörü enumerate edilemedi. Count={0}. HResult: 0x{1:X8}. Mesaj: {2}' -f $countBefore,($_.Exception.HResult -band 0xFFFFFFFF),$_.Exception.Message)
    }
}
function Get-FsiImportedDirectoryMap {
    [CmdletBinding()]
    param([object]$Directory,[hashtable]$DirectoryMap)
    foreach ($item in @(Get-FsiDirectoryChild -Directory $Directory)) {
        $path=ConvertTo-FsiNormalizedPath ([string]$item.FullPath)
        $isDirectory=$false
        try {$null=[int]$item.Count;$isDirectory=$true} catch {$isDirectory=$false}
        if ($isDirectory) {
            $DirectoryMap[$path]=$item
            Get-FsiImportedDirectoryMap -Directory $item -DirectoryMap $DirectoryMap
        }
    }
}
function Get-ImportedFsiDirectoryMap {
    [CmdletBinding()]
    param([object]$RootDir)
    $dirs=@{};$dirs['\']=$RootDir
    Get-FsiImportedDirectoryMap -Directory $RootDir -DirectoryMap $dirs
    return $dirs
}
function Add-FsiSourceFile {
    <#
    .SYNOPSIS
    Kaynak dosyayı UDF imajındaki bir klasöre ekler.
    .DESCRIPTION
    Dosya içeriği bir akış olarak bağlanır; akış, imaj oluşturulduktan sonra serbest bırakılmak üzere Streams listesine eklenir. Oluşturma/erişim/değiştirme zamanları kaynaktan kopyalanır; kopyalanamazsa hata Verbose olarak kaydedilir ve işlem sürer.
    .PARAMETER FileSystemImage
    IMAPI2FS dosya sistemi imajı.
    .PARAMETER ParentDir
    Dosyanın ekleneceği imaj klasörü.
    .PARAMETER FileName
    İmaj içindeki dosya adı.
    .PARAMETER SourcePath
    Diskteki kaynak dosya yolu.
    .PARAMETER Streams
    Sonradan serbest bırakılacak akışların listesi.
    .EXAMPLE
    PS> Add-FsiSourceFile -FileSystemImage $fsi -ParentDir $dir -FileName a.txt -SourcePath C:\a.txt -Streams $streams
    #>
    [CmdletBinding()]
    param([object]$FileSystemImage,[object]$ParentDir,[string]$FileName,[string]$SourcePath,[System.Collections.ArrayList]$Streams)
    $sourceInfo=Get-Item -LiteralPath $SourcePath -Force
    $stream=Open-ImapiFileStream $SourcePath
    $fileItem=$null
    try {
        $fileItem=$FileSystemImage.CreateFileItem($FileName)
        $fileItem.Data=$stream
        try {$fileItem.CreationTime=$sourceInfo.CreationTime} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try {$fileItem.LastAccessedTime=$sourceInfo.LastAccessTime} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try {$fileItem.LastModifiedTime=$sourceInfo.LastWriteTime} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        [void]$ParentDir.Add($fileItem)
        [void]$Streams.Add($stream)
        $stream=$null
    } catch {
        throw ('UDF dosya ekleme başarısız: {0}`r`nHResult: 0x{1:X8}`r`nMesaj: {2}' -f $SourcePath,($_.Exception.HResult -band 0xFFFFFFFF),$_.Exception.Message)
    } finally {
        if ($null -ne $fileItem) {Unregister-ComObject $fileItem}
        if ($null -ne $stream) {Unregister-ComObject $stream}
    }
}
function Test-TempImageCapacity {
    <#
    .SYNOPSIS
    TEMP sürücüsünde geçici IMG dosyası için yeterli boş alan olup olmadığını denetler.
    .DESCRIPTION
    Boş alan gereken boyuttan (artı güvenlik payı) azsa hata fırlatır. TEMP sürücüsü belirlenemezse veya sürücü bilgisi okunamazsa yazmayı durdurmaz: durumu loga kaydeder ve `$true döner (bilinçli tasarım; kontrol yapılamadı demektir, alan yeterli demek değildir).
    .PARAMETER RequiredBytes
    IMG dosyası için gereken boyut (bayt).
    .OUTPUTS
    System.Boolean - kontrol geçtiyse veya yapılamadıysa $true.
    .EXAMPLE
    PS> Test-TempImageCapacity -RequiredBytes 4GB
    #>
    [OutputType([bool])]
    [CmdletBinding()]
    param([Int64]$RequiredBytes)

    $tempRoot = [IO.Path]::GetPathRoot($env:TEMP)

    if ([string]::IsNullOrWhiteSpace($tempRoot)) {
        try {
            Add-Content -LiteralPath (Join-Path $script:LogDir 'ui_errors.txt') `
                -Value ('{0} [Test-TempImageCapacity] TEMP sürücü kökü belirlenemedi; kapasite ön kontrolü atlandi.' -f (Get-Date -Format 's')) `
                -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        return $true
    }

    try {
        $drive = New-Object -TypeName System.IO.DriveInfo -ArgumentList $tempRoot
        $required = $RequiredBytes + 100MB
        if ($drive.AvailableFreeSpace -lt $required) {
            throw ('TEMP sürücüsünde yeterli boş alan yok. Gerekli yaklaşık: {0}, mevcut: {1}.' -f (Format-Size $required),(Format-Size $drive.AvailableFreeSpace))
        }
    }
    catch {
        if ($_.Exception.Message -like 'TEMP sürücüsünde yeterli*') { throw }

        # DriveInfo hatasi: yazmayi durdurma degil, gorunur kayit birak.
        try {
            Add-Content -LiteralPath (Join-Path $script:LogDir 'ui_errors.txt') `
                -Value ('{0} [Test-TempImageCapacity] Kapasite on kontrolü yapilamadi: {1}' -f (Get-Date -Format 's'), $_.Exception.Message) `
                -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }

    return $true
}

function Test-ImportedUdfFilesystemIntegrity {
    [CmdletBinding()]
    param([object]$Root)

    if ($null -eq $Root) { throw 'Imported UDF Root nesnesi yok.' }

    $directoryCount = 0
    $fileCount      = 0
    $itemCount      = 0
    $visited        = 0

    $stack = New-Object System.Collections.Stack
    $stack.Push([pscustomobject]@{ Directory = $Root; Path = '\' })

    while ($stack.Count -gt 0) {
        $frame = $stack.Pop()
        $visited++

        if ($visited -gt 100000) {
            throw 'Önceki UDF session ağacı beklenenden büyük; import doğrulaması güvenlik nedeniyle durduruldu.'
        }

        try { $items = @(Get-FsiDirectoryChild -Directory $frame.Directory) }
        catch { throw ('Önceki UDF session klasörü enumerate edilemedi ({0}): {1}' -f $frame.Path, $_.Exception.Message) }

        foreach ($item in @($items)) {
            if ($null -eq $item) {
                throw ('Önceki UDF session içinde geçersiz FSI öğesi: {0}' -f $frame.Path)
            }
            $itemCount++

            try { $fullPath = ConvertTo-FsiNormalizedPath ([string]$item.FullPath) }
            catch { throw ('Önceki UDF session içindeki bir FSI öğesinin yolu okunamadı: {0}' -f $frame.Path) }

            if ([string]::IsNullOrWhiteSpace($fullPath) -or $fullPath -eq '\') {
                throw ('Önceki UDF session içinde geçersiz FSI yolu: {0}' -f $frame.Path)
            }

            $isDirectory = $false
            try { $null = [int]$item.Count; $isDirectory = $true } catch { $isDirectory = $false }

            if ($isDirectory) {
                $directoryCount++
                $stack.Push([pscustomobject]@{ Directory = $item; Path = $fullPath })
            }
            else {
                $fileCount++
            }
        }
    }

    return [pscustomobject]@{
        ItemCount      = $itemCount
        DirectoryCount = $directoryCount
        FileCount      = $fileCount
    }
}

function Initialize-UdfImageFileSystem {
    [CmdletBinding()]
    param([int]$MediaType,[bool]$MediaIsBlank,[bool]$Multisession,[object]$Recorder)

    $fsi=$null;$dataWriter=$null;$success=$false
    try {
        $fsi=New-Object -ComObject IMAPI2FS.MsftFileSystemImage
        $fsi.FileSystemsToCreate=$script:FsiFileSystemUDF
        $fsi.UDFRevision=$script:UdfRevision250

        if ($MediaIsBlank) {
            $fsi.ChooseImageDefaultsForMediaType([int]$MediaType)
            $fsi.FileSystemsToCreate=$script:FsiFileSystemUDF
            $fsi.UDFRevision=$script:UdfRevision250
        }
        else {
            if (-not $Multisession) {throw 'Dolu medya için Çok oturumlu seçeneği AÇIK olmalıdır.'}
            $dataWriter=New-Object -ComObject IMAPI2.MsftDiscFormat2Data
            $dataWriter.Recorder=$Recorder
            $dataWriter.ClientName='OptikMedyaArsiv'
            try {$fsi.MultisessionInterfaces=$dataWriter.MultisessionInterfaces}
            catch {throw ('Bu medya için IMAPI multisession arabirimi kullanılamıyor: {0}' -f $_.Exception.Message)}
            if ($null -eq $fsi.MultisessionInterfaces) {throw 'IMAPI multisession arabirimi alınamadı; önceki session importu güvenilir değil.'}

            Write-Host 'Önceki UDF oturumu içeri aktarılıyor...' -ForegroundColor $script:Theme.Color.Primary
            $fsi.ImportSpecificFileSystem($script:FsiFileSystemUDF)
            $importedUdfRevision=[int]$fsi.UDFRevision
            if ($importedUdfRevision -ne $script:UdfRevision250) {
                throw ('Önceki session UDF 2.50 değil. Algılanan revision: 0x{0:X}. Gerekli revision: 0x{1:X}.' -f $importedUdfRevision,$script:UdfRevision250)
            }
            $importCheck=Test-ImportedUdfFilesystemIntegrity -Root $fsi.Root
            Write-Host ('✓ Önceki session import doğrulandı: {0} öğe, {1} klasör, {2} dosya.' -f $importCheck.ItemCount,$importCheck.DirectoryCount,$importCheck.FileCount) -ForegroundColor $script:Theme.Color.Success
        }

        $root=$fsi.Root
        if ($null -eq $root) {throw 'IMAPI UDF Root nesnesi oluşturulamadı.'}
        $success=$true
        return [pscustomobject]@{FileSystem=$fsi;DataWriter=$dataWriter;Root=$root}
    }
    finally {
        if (-not $success) {
            if ($null -ne $dataWriter) {Unregister-ComObject $dataWriter}
            if ($null -ne $fsi) {Unregister-ComObject $fsi}
        }
    }
}

function Initialize-UdfImageStructure {
    [CmdletBinding()]
    param([object]$FileSystemImage,[object]$Root,[hashtable]$DirectoryMap,[string]$VolumeLabel,[bool]$MediaIsBlank)

    $FileSystemImage.VolumeName=$VolumeLabel
    try {$FileSystemImage.CreateRedundantUdfMetadataFiles=$true}
    catch {throw ('Redundant UDF Metadata etkinleştirilemedi: {0}' -f $_.Exception.Message)}
    if (-not [bool]$FileSystemImage.CreateRedundantUdfMetadataFiles) {throw 'Redundant UDF Metadata etkinleştirilemedi.'}

    if ($MediaIsBlank) {$DirectoryMap['\']=$Root}
    else {
        $importedMap=Get-ImportedFsiDirectoryMap -RootDir $Root
        foreach ($key in @($importedMap.Keys)) {$DirectoryMap[$key]=$importedMap[$key]}
    }
}

function Add-UdfProjectDirectory {
    [CmdletBinding()]
    param([object]$FileSystemImage,[object]$Root,[hashtable]$DirectoryMap,[string]$SessionPath,[string[]]$DirectoryPaths)

    $sessionDir=Get-UdfDirectory -FileSystemImage $FileSystemImage -DirectoryMap $DirectoryMap -RootDir $Root -DestinationDirectory $SessionPath
    $projectDirs=@($DirectoryPaths | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)} | Sort-Object Length)
    $directoryCount=[Math]::Max(1,@($projectDirs).Count)
    $directoryIndex=0

    Start-UiStep -No 2 -Name 'UDF dizin' -Count @($projectDirs).Count -Unit 'klasör'
    foreach ($dirPath in $projectDirs) {
        $directoryIndex++
        $normalizedDir=ConvertTo-FsiNormalizedPath $dirPath
        $sessionDirPath=Join-FsiPath -BasePath $SessionPath -RelativePath $normalizedDir
        if (-not $DirectoryMap.ContainsKey($sessionDirPath)) {
            [void](Get-UdfDirectory -FileSystemImage $FileSystemImage -DirectoryMap $DirectoryMap -RootDir $Root -DestinationDirectory $sessionDirPath)
        }
        Update-UiStep -Done $directoryIndex -Count $directoryCount -Text $sessionDirPath -Unit 'klasör'
    }
    $script:UiDirCount=@($projectDirs).Count
    Stop-UiStep ('{0} klasör' -f @($projectDirs).Count)

    Write-Output -NoEnumerate -InputObject $sessionDir
}

function Add-UdfProjectFile {
    [CmdletBinding()]
    param([object]$FileSystemImage,[object]$Root,[hashtable]$DirectoryMap,[string]$SessionPath,[object[]]$Files,[System.Collections.ArrayList]$Streams)

    $fileIndex=0
    $fileCount=[Math]::Max(1,@($Files).Count)
    foreach ($f in @($Files)) {
        $fileIndex++
        $sourcePath=[string]$f.SourcePath
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {throw ('Kaynak dosya bulunamadı: {0}' -f $sourcePath)}
        $destinationPath=Join-FsiPath -BasePath $SessionPath -RelativePath (ConvertTo-FsiNormalizedPath $f.DestinationPath)
        $parts=@($destinationPath.Trim('\') -split '\\' | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
        if (@($parts).Count -eq 0) {throw ('Geçersiz hedef yol: {0}' -f $destinationPath)}
        $fileName=[string]$parts[@($parts).Count-1]
        $parentPath='\'
        if (@($parts).Count -gt 1) {$parentPath='\'+[string]::Join('\',$parts[0..(@($parts).Count-2)])}
        $parentDir=Get-UdfDirectory -FileSystemImage $FileSystemImage -DirectoryMap $DirectoryMap -RootDir $Root -DestinationDirectory $parentPath
        Update-UiStep -Done $fileIndex -Count $fileCount -Text ([string]$f.DestinationPath) -Unit 'dosya'
        Add-FsiSourceFile -FileSystemImage $FileSystemImage -ParentDir $parentDir -FileName $fileName -SourcePath $sourcePath -Streams $Streams
    }
}

function Add-UdfManifestFile {
    [CmdletBinding()]
    param([object]$FileSystemImage,[object]$SessionDir,[object]$Manifest,[System.Collections.ArrayList]$Streams)

    $manifestStream=Open-ImapiFileStream $Manifest.Path
    $manifestItem=$null
    try {
        $manifestItem=$FileSystemImage.CreateFileItem($Manifest.FileName)
        $manifestItem.Data=$manifestStream
        [void]$SessionDir.Add($manifestItem)
        [void]$Streams.Add($manifestStream)
        $manifestStream=$null
    }
    finally {
        if ($null -ne $manifestItem) {Unregister-ComObject $manifestItem}
        if ($null -ne $manifestStream) {Unregister-ComObject $manifestStream}
    }
}

function Export-UdfImagePackage {
    <#
    .SYNOPSIS
    CreateResultImage sonucunu geçici IMG dosyasına yazar ve boyutunu doğrular.
    .DESCRIPTION
    Sonuç null ise, boyut geçersizse veya oluşan dosya boyutu beklenenden farklıysa hata fırlatır.
    .PARAMETER Result
    IMAPI2FS CreateResultImage() sonucu.
    .PARAMETER SessionName
    Geçici dosya adında kullanılacak SESSION adı.
    .OUTPUTS
    PSCustomObject: ImageBytes, ImageBlocks, BlockSize, TempImagePath, BurnStream.
    .EXAMPLE
    PS> Export-UdfImagePackage -Result $result -SessionName $sessionName
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param([object]$Result,[string]$SessionName)

    $tempImage='';$burnStream=$null;$success=$false
    try {
        if ($null -eq $Result) {throw 'CreateResultImage() null sonuç döndürdü.'}
        [Int64]$imageBlocks=$Result.TotalBlocks
        [Int64]$blockSize=$Result.BlockSize
        if ($blockSize -le 0) {$blockSize=$script:BytesPerBlock}
        [Int64]$imageBytes=$imageBlocks*$blockSize
        if ($imageBlocks -le 0 -or $imageBytes -le 0) {throw 'Geçersiz UDF image boyutu.'}

        [void](Test-TempImageCapacity -RequiredBytes $imageBytes)
        $tempImage=Join-Path $env:TEMP ('OptikMedyaArsiv_'+$SessionName+'.img')
        [void](Export-ImageStreamToFile -SourceStream $Result.ImageStream -Path $tempImage -ExpectedBytes $imageBytes)
        $actualImageBytes=(Get-Item -LiteralPath $tempImage -Force).Length
        if ([Int64]$actualImageBytes -ne $imageBytes) {
            throw ('Geçici image boyutu beklenen değerden farklı: beklenen {0}, oluşan {1}.' -f $imageBytes,$actualImageBytes)
        }
        $burnStream=Open-ImapiFileStream -Path $tempImage -ReadOnly $true
        $success=$true
        return [pscustomobject]@{ImageBytes=[Int64]$actualImageBytes;ImageBlocks=$imageBlocks;BlockSize=$blockSize;TempImagePath=$tempImage;BurnStream=$burnStream}
    }
    finally {
        if (-not $success) {
            if ($null -ne $burnStream) {Unregister-ComObject $burnStream}
            if (-not [string]::IsNullOrWhiteSpace($tempImage)) {
                try {
                    if (Test-Path -LiteralPath $tempImage -PathType Leaf) {
                        Remove-Item -LiteralPath $tempImage -Force -ErrorAction SilentlyContinue
                    }
                } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
        }
    }
}

function New-UdfImagePackage {
    <#
    .SYNOPSIS
    Yazılacak oturum için UDF 2.50 imaj paketini oluşturur.
    .DESCRIPTION
    Dosya sistemini hazırlar, klasörleri, dosyaları ve MANIFEST dosyasını ekler, imajı oluşturup geçici IMG dosyasına aktarır. Başarıda COM nesneleri ve geçici IMG çağırana devredilir; hatada finally bloğunda temizlenir.
    .PARAMETER Files
    Yazılacak dosyalar.
    .PARAMETER DirectoryPaths
    Yazılacak klasör yolları.
    .PARAMETER Manifest
    MANIFEST bilgisi.
    .PARAMETER VolumeLabel
    Disk etiketi.
    .PARAMETER MediaType
    IMAPI fiziksel medya türü.
    .PARAMETER Recorder
    IMAPI2 disc recorder COM nesnesi.
    .PARAMETER Multisession
    Önceki oturumlara ekleme modu.
    .PARAMETER MediaIsBlank
    Medya boşsa $true.
    .PARAMETER SessionName
    SESSION klasör adı.
    .OUTPUTS
    PSCustomObject: FileSystem, Result, ImageBytes, ImageBlocks, BlockSize, SessionName, SessionPath, TempImagePath, BurnStream.
    .EXAMPLE
    PS> New-UdfImagePackage -Files $files -DirectoryPaths $dirs -Manifest $manifest -VolumeLabel $label -MediaType $type -Recorder $recorder -Multisession $false -MediaIsBlank $true -SessionName $sessionName
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param([object[]]$Files,[string[]]$DirectoryPaths,[object]$Manifest,[string]$VolumeLabel,[int]$MediaType,
    [object]$Recorder,[bool]$Multisession,[bool]$MediaIsBlank,[string]$SessionName)

    $fsi=$null;$dataWriter=$null;$result=$null;$burnStream=$null
    $streams=New-Object System.Collections.ArrayList
    $tempImage='';$success=$false

    try {
        Write-Host '  UDF 2.50 hazırlanıyor...' -ForegroundColor $script:Theme.Color.Primary
        $initialized=Initialize-UdfImageFileSystem -MediaType $MediaType -MediaIsBlank $MediaIsBlank -Multisession $Multisession -Recorder $Recorder
        $fsi=$initialized.FileSystem
        $dataWriter=$initialized.DataWriter
        $root=$initialized.Root

        $dirMap=@{}
        Initialize-UdfImageStructure -FileSystemImage $fsi -Root $root -DirectoryMap $dirMap -VolumeLabel $VolumeLabel -MediaIsBlank $MediaIsBlank

        $sessionPath='\'+$SessionName
        $sessionDir=Add-UdfProjectDirectory -FileSystemImage $fsi -Root $root -DirectoryMap $dirMap -SessionPath $sessionPath -DirectoryPaths $DirectoryPaths
        Start-UiStep -No 3 -Name 'Session' -Count @($Files).Count -Unit 'dosya'
        Add-UdfProjectFile -FileSystemImage $fsi -Root $root -DirectoryMap $dirMap -SessionPath $sessionPath -Files $Files -Streams $streams
        Add-UdfManifestFile -FileSystemImage $fsi -SessionDir $sessionDir -Manifest $Manifest -Streams $streams
        Stop-UiStep ([string]$SessionName)
        Start-UiStep -No 4 -Name 'UDF imajı'
        Show-UiBusy 'imaj oluşturuluyor…' 'CreateResultImage 30-90 sn sürebilir, lütfen bekleyin'

        $result=$fsi.CreateResultImage()
        $exported=Export-UdfImagePackage -Result $result -SessionName $SessionName
        $burnStream=$exported.BurnStream
        $tempImage=$exported.TempImagePath
        Stop-UiStep (Format-Size $exported.ImageBytes)

        $success=$true
        return [pscustomobject]@{
            FileSystem=$fsi;Result=$result;ImageBytes=[Int64]$exported.ImageBytes;ImageBlocks=[Int64]$exported.ImageBlocks
            BlockSize=[Int64]$exported.BlockSize;SessionName=$SessionName;SessionPath=$sessionPath;TempImagePath=$tempImage;BurnStream=$burnStream
        }
    }
    finally {
        # Başarılı sonuçta COM nesneleri ve temp IMG dışarıya devredilir.
        # Image üretimi sırasında hata olursa burada temizlenir.
        if (-not $success) {
            if ($null -ne $burnStream) {Unregister-ComObject $burnStream}
            if ($null -ne $result) {Unregister-ComObject $result}
            if ($null -ne $fsi) {Unregister-ComObject $fsi}
            if (-not [string]::IsNullOrWhiteSpace($tempImage)) {
                try {
                    if (Test-Path -LiteralPath $tempImage -PathType Leaf) {
                        Remove-Item -LiteralPath $tempImage -Force -ErrorAction SilentlyContinue
                    }
                } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
        }
        foreach ($stream in @($streams)) {Unregister-ComObject $stream}
        if ($null -ne $dataWriter) {Unregister-ComObject $dataWriter}
    }
}

# ============================================================
# IMAPI2 MEDYA / YAZICI
# ============================================================
function Get-DiscMaster { [CmdletBinding()] param();return (New-Object -ComObject IMAPI2.MsftDiscMaster2)}
function Get-Recorder {
    [OutputType([object[]])]
    [CmdletBinding()]
    param()
    $master=$null;$list=New-Object System.Collections.ArrayList
    try {
        $master=Get-DiscMaster
        foreach ($id in $master) {
            if ([string]::IsNullOrWhiteSpace([string]$id)) {continue}
            $rec=$null
            try {
                $rec=New-Object -ComObject IMAPI2.MsftDiscRecorder2
                $rec.InitializeDiscRecorder($id)
                [void]$list.Add($rec);$rec=$null
            } catch {
                if ($null -ne $rec) {Unregister-ComObject $rec}
            }
        }
        return @($list)
    } finally {Unregister-ComObject $master}
}
function Get-MediaName {
    [OutputType([string])]
    [CmdletBinding()]
    param([int]$MediaType)
    switch ($MediaType) {
        0 {'Bilinmeyen / Medya Yok'}
        1 {'CD-ROM'}
        2 {'CD-R'}
        3 {'CD-RW'}
        4 {'DVD-ROM'}
        5 {'DVD-RAM'}
        6 {'DVD+R'}
        7 {'DVD+RW'}
        8 {'DVD+R DL'}
        9 {'DVD-R'}
        10 {'DVD-RW'}
        11 {'DVD-R DL'}
        12 {'DVD Disk'}
        13 {'DVD+RW DL'}
        14 {'HD DVD-ROM'}
        15 {'HD DVD-R'}
        16 {'HD DVD-RAM'}
        17 {'BD-ROM'}
        18 {'BD-R'}
        19 {'BD-RE'}
        default {"Bilinmeyen medya (0x{0:X})" -f $MediaType}
    }
}
function Convert-SpeedToX {
    [OutputType([double])]
    [CmdletBinding()]
    param([Int64]$SectorsPerSecond,[int]$MediaType)
    if ($SectorsPerSecond -le 0) {return 0.0}
    if ($MediaType -in @(1,2,3)) {return [Math]::Round($SectorsPerSecond/75.0,1)}
    if ($MediaType -in @(17,18,19)) {return [Math]::Round($SectorsPerSecond/2195.0,1)}
    return [Math]::Round($SectorsPerSecond/676.0,1)
}
function Get-MediaInfo {
    <#
    .SYNOPSIS
    IMAPI üzerinden takılı medyanın durumunu okur.
    .DESCRIPTION
    Sürücü hazırlanıyorsa (0xC0AA0205/06/07) zaman aşımına kadar bekleyip yeniden dener. MediaHeuristicallyBlank ve MediaPhysicallyBlank okunamazsa değerleri `$false kalır; boş medya kararı bu durumda yalnızca MediaStatus bitlerine dayanır.
    .PARAMETER Recorder
    IMAPI2 disc recorder COM nesnesi.
    .PARAMETER RetryIntervalSeconds
    Yeniden deneme aralığı (sn).
    .PARAMETER RetryTimeoutSeconds
    Toplam bekleme süresi (sn).
    .OUTPUTS
    PSCustomObject: medya türü, kapasite ve IsBlank / IsAppendable / IsFinalSession / IsFinalized / IsReadOnlyPhysicalType gibi alanlar.
    .EXAMPLE
    PS> Get-MediaInfo -Recorder $recorder -RetryTimeoutSeconds 60
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param([object]$Recorder,[int]$RetryIntervalSeconds=1,[int]$RetryTimeoutSeconds=60)
    $deadline=(Get-Date).AddSeconds($RetryTimeoutSeconds);$attempt=0
    while ($true) {
        $attempt++;$format=$null
        try {
            $format=New-Object -ComObject IMAPI2.MsftDiscFormat2Data
            $format.Recorder=$Recorder
            $mediaType=[int]$format.CurrentPhysicalMediaType
            $mediaStatus=[int]$format.CurrentMediaStatus
            $totalSectors=[Int64]$format.TotalSectorsOnMedia
            $freeSectors=[Int64]$format.FreeSectorsOnMedia
            $speeds=@($format.SupportedWriteSpeeds)
            $heuristicallyBlank=$false;$physicallyBlank=$false
            try {$heuristicallyBlank=[bool]$format.MediaHeuristicallyBlank} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            try {$physicallyBlank=[bool]$format.MediaPhysicallyBlank} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
	    return [pscustomobject]@{
	        MediaType=$mediaType;MediaStatus=$mediaStatus;TotalSectors=$totalSectors;FreeSectors=$freeSectors
	        TotalBytes=$totalSectors*$script:BytesPerBlock;FreeBytes=$freeSectors*$script:BytesPerBlock
	        SupportedSpeeds=$speeds;MediaName=(Get-MediaName $mediaType)
 	        AppSupportedMediaType=($mediaType -in $script:WritableMediaTypes)
    	        MediaHeuristicallyBlank=$heuristicallyBlank;MediaPhysicallyBlank=$physicallyBlank
	        IsBlank=(($mediaStatus -band $script:MediaState.Blank) -ne 0) -or $physicallyBlank -or $heuristicallyBlank
	        IsAppendable=(($mediaStatus -band $script:MediaState.Appendable) -ne 0)
	        IsFinalSession=(($mediaStatus -band $script:MediaState.FinalSession) -ne 0)
	        IsFinalized=(($mediaStatus -band $script:MediaState.Finalized) -ne 0)
        IsReadOnlyPhysicalType=($mediaType -in $script:ReadOnlyPhysicalMediaTypes)
	    }
        } catch {
            $hr=Get-HResultHex $_.Exception
            $retryable=($hr -in $script:ImapiRetryableHResults)
            if (-not $retryable) {throw}
            if ((Get-Date) -ge $deadline) {throw ('Medya belirtilen süre içinde hazır hale gelmedi. Son durum: {0}' -f $hr)}
            if ($attempt -eq 1) {Write-Host 'Medya hazırlanıyor, bekleniyor...' -ForegroundColor $script:Theme.Color.Warning}
            Start-Sleep -Seconds $RetryIntervalSeconds
        } finally {
            if ($null -ne $format) {Unregister-ComObject $format}
        }
    }
}
function Get-MediaStateFlag {
    [CmdletBinding()]
    param([object]$Media)
    if ($null -eq $Media) { throw 'Medya durum bayrakları için medya bilgisi bulunamadı.' }
    $status=[int]$Media.MediaStatus
    return [pscustomobject]@{
        Status=$status; IsBlank=[bool]$Media.IsBlank; IsAppendable=[bool]$Media.IsAppendable; IsFinalSession=[bool]$Media.IsFinalSession; IsFinalized=[bool]$Media.IsFinalized
        IsWriteProtected=(($status -band $script:MediaState.WriteProtected) -ne 0); IsDamaged=(($status -band $script:MediaState.Damaged) -ne 0); IsUnsupported=(($status -band $script:MediaState.UnsupportedMedia) -ne 0)
    }
}
function Test-MediaWritePlan {
    [CmdletBinding()]
    param([object]$Media,[object]$MediaCapabilities,[object]$WriteSettings)
    if ($null -eq $Media) { throw 'Yazma planı doğrulanamadı: medya bilgisi yok.' }
    if ($null -eq $MediaCapabilities) { throw 'Yazma planı doğrulanamadı: medya yetenekleri yok.' }
    if ($null -eq $WriteSettings) { throw 'Yazma planı doğrulanamadı: yazma ayarları yok.' }
    $flags=Get-MediaStateFlag -Media $Media
    if (-not [bool]$MediaCapabilities.CanWrite) { return [pscustomobject]@{Ok=$false;Reason=[string]$MediaCapabilities.Description} }
    if ($flags.IsFinalized -or $flags.IsWriteProtected -or $flags.IsDamaged -or $flags.IsUnsupported) { return [pscustomobject]@{Ok=$false;Reason='Medya yazma başlamadan önce yazılabilir durumunu kaybetti.'} }
    if ($MediaCapabilities.Profile -eq 'BLANK' -and -not $flags.IsBlank) { return [pscustomobject]@{Ok=$false;Reason='Medya durumu değişti: BLANK durumu artık geçerli değil.'} }
    if ($MediaCapabilities.Profile -eq 'APPENDABLE' -and -not ($flags.IsAppendable -or $flags.IsFinalSession)) { return [pscustomobject]@{Ok=$false;Reason='Medya durumu değişti: APPENDABLE/FINAL_SESSION durumu artık geçerli değil.'} }
    if ([bool]$MediaCapabilities.MultisessionRequired -and -not [bool]$WriteSettings.Multisession) { return [pscustomobject]@{Ok=$false;Reason='Bu medya profili için Multisession zorunludur.'} }
    return [pscustomobject]@{Ok=$true;Reason=''}
}
function Get-MediaProfile {
    <#
    .SYNOPSIS
    Medya durumunu bir profil koduna sınıflandırır.
    .DESCRIPTION
    Profil kodları (ör. NO_MEDIA, BLANK, APPENDABLE) belirli bir sırayla denetlenir; sıra önemlidir (BLANK, APPENDABLE denetiminden önce gelir).
    .PARAMETER Media
    Get-MediaInfo çıktısı.
    .OUTPUTS
    System.String - profil kodu.
    .EXAMPLE
    PS> Get-MediaProfile -Media $media
    #>
    [OutputType([string])]
    [CmdletBinding()]
    param(
        [object]$Media
    )

    if ($null -eq $Media) {
        throw 'Medya profili oluşturmak için medya bilgisi bulunamadı.'
    }

    # 0. Sürücüde medya yok — bu, desteklenmeyen medyadan farklı bir durumdur.
    # MediaType=0 IMAPI tarafından "medya yok" olarak raporlanır.
    # Bu durumda MediaStatus=0x0 olur ve Finalized/WriteProtected bitleri anlamsızdır.
    if ([int]$Media.MediaType -eq 0) {
        return 'NO_MEDIA'
    }

    # 1. IMAPI medya durumu: hasarlı
    if (
        ([int]$Media.MediaStatus -band $script:MediaState.Damaged) -ne 0
    ) {
        return 'DAMAGED'
    }

    # 2. IMAPI FINALIZED biti varsa bunu kesin finalized kabul et.
    if ([bool]$Media.IsFinalized) {
        return 'FINALIZED'
    }

    # 3. Salt-okunur fiziksel medya (CD-ROM/DVD-ROM/BD-ROM).
    # 0xA000 gibi durumlarda WriteProtected + UnsupportedMedia birlikte gelebilir.
    # Önce fiziksel readonly tipi ayırt edilir; bu, medya kökeninin finalized olduğunu
    # iddia etmez, yalnızca mevcut medyanın yazılamaz olduğunu doğru sınıflandırır.
    if ([bool]$Media.IsReadOnlyPhysicalType) {
        return 'READ_ONLY'
    }

    # 4. Yazma korumalı yazılabilir medya
    if (
        ([int]$Media.MediaStatus -band $script:MediaState.WriteProtected) -ne 0
    ) {
        return 'WRITE_PROTECTED'
    }

    # 5. Finalize edilmiş / son oturum
    # FINALIZED ile FINAL_SESSION aynı şey değildir.
    # FINALIZED (0x4000) = medya artık yazılamaz.
    # FINAL_SESSION (0x0008) = en fazla bir ek session daha yazılabilir
    # veya medya multisession desteklemiyordur.
    # FINALIZED biti yukarıda zaten ele alındı.
    # 6. Uygulamanın desteklemediği diğer medya türleri
    if (-not [bool]$Media.AppSupportedMediaType) {
        return 'UNSUPPORTED'
    }

    # 7. Boş medya
    # Blank bilgisi, Appendable bilgisinden önceliklidir.
    # Çünkü IMAPI medya durum bitleri birlikte raporlanabilir.
    if ([bool]$Media.IsBlank) {
        return 'BLANK'
    }

    # 8. Yeni session eklenebilen medya
    # APPENDABLE normal çok oturumlu medya durumudur.
    # FINAL_SESSION ise en fazla bir ek session daha yazılabileceğini belirtir.
    # Bu uygulamanın dolu medya yazma modeli her ikisini de APPENDABLE
    # profili üzerinden işler; gerçek FINALIZED durumu ayrıca yukarıda engellenir.
    if ([bool]$Media.IsAppendable -or [bool]$Media.IsFinalSession) {
        return 'APPENDABLE'
    }

    # 9. Boş değil, appendable değil ve finalized değil
    return 'NOT_WRITABLE'
}
function New-MediaWriteCapability {
    [CmdletBinding()]
    param(
        [Alias('Profile')][string]$MediaProfile,
        [bool]$CanWrite,
        [string]$WriteMode,
        [string]$WriteModeText,
        [bool]$MultisessionAvailable,
        [bool]$MultisessionEditable,
        [bool]$MultisessionRequired,
        [bool]$FinalizeAvailable,
        [bool]$FinalizeEditable,
        [bool]$ImportExistingFileSystem,
        [string]$Description
    )

    return [pscustomobject]@{
        Profile                  = $MediaProfile
        CanWrite                 = $CanWrite
        WriteMode                = $WriteMode
        WriteModeText            = $WriteModeText
        MultisessionAvailable    = $MultisessionAvailable
        MultisessionEditable     = $MultisessionEditable
        MultisessionRequired     = $MultisessionRequired
        FinalizeAvailable        = $FinalizeAvailable
        FinalizeEditable         = $FinalizeEditable
        ImportExistingFileSystem = $ImportExistingFileSystem
        Description              = $Description
    }
}

# ============================================================
# Medya yetenekleri — data-driven tablo (Faz 3a)
# Her profil için 11 özellik tek yerde tanımlıdır.
# Yeni profil eklemek için tabloya bir giriş ekleyin.
# ============================================================
$script:MediaCapabilityTable = @{
    'NO_MEDIA' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'Medya Yok'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Sürücüde medya yok. Lütfen boş (BLANK) veya appendable bir optik medya takın.'
    }
    'BLANK' = @{
        CanWrite                 = $true
        WriteMode                = 'BLANK'
        WriteModeText            = 'BLANK'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $true
        FinalizeEditable         = $true
        ImportExistingFileSystem = $false
        Description              = 'Yeni SESSION başlatılacak.'
    }
    'APPENDABLE' = @{
        CanWrite                 = $true
        WriteMode                = 'APPEND_SESSION'
        WriteModeText            = 'APPENDABLE'
        MultisessionAvailable    = $true
        MultisessionEditable     = $false
        MultisessionRequired     = $true
        FinalizeAvailable        = $true
        FinalizeEditable         = $true
        ImportExistingFileSystem = $true
        Description              = 'Yeni SESSION eklenecek.'
    }
    'FINALIZED' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'FINALIZED'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Medya sonlandırılmış.'
    }
    'WRITE_PROTECTED' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'Yazma yapılamaz'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Medya yazmaya karşı korumalı.'
    }
    'DAMAGED' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'Yazma yapılamaz'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Medya hasarlı olarak bildiriliyor.'
    }
    'READ_ONLY' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'SADECE OKUNUR'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Medya salt-okunur. Yeni yazma yapılamaz; sonlandırma kökeni IMAPI tarafından kesin olarak ayrıştırılamıyor.'
    }
    'UNSUPPORTED' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'Yazma yapılamaz'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Bu medya türü uygulama tarafından desteklenmiyor.'
    }
    'NOT_WRITABLE' = @{
        CanWrite                 = $false
        WriteMode                = 'NONE'
        WriteModeText            = 'Yazma yapılamaz'
        MultisessionAvailable    = $false
        MultisessionEditable     = $false
        MultisessionRequired     = $false
        FinalizeAvailable        = $false
        FinalizeEditable         = $false
        ImportExistingFileSystem = $false
        Description              = 'Medya mevcut durumda yeni yazma için uygun değil.'
    }
}

function Get-MediaWriteCapability {
    [CmdletBinding()]
    param([string]$MediaProfile)

    if ([string]::IsNullOrWhiteSpace($MediaProfile)) {
        throw 'Yazma seçeneklerini belirlemek için medya profili bulunamadı.'
    }
    if (-not $script:MediaCapabilityTable.ContainsKey($MediaProfile)) {
        throw ('Tanımsız medya profili: {0}' -f $MediaProfile)
    }

    $entry = $script:MediaCapabilityTable[$MediaProfile].Clone()
    $entry['Profile'] = $MediaProfile
    return New-MediaWriteCapability @entry
}
function New-WriteSettingsSelection {
    [CmdletBinding()]
    param(
        [object]$MediaCapabilities,
        [object]$Settings
    )

    if ($null -eq $MediaCapabilities) {
        throw 'Yazma ayarı seçimi için medya yetenekleri bulunamadı.'
    }

    if ($null -eq $Settings) {
        throw 'Yazma ayarı seçimi için kullanıcı ayarları bulunamadı.'
    }

    $multisession = [bool]$MediaCapabilities.MultisessionRequired

    $finalize = if ([bool]$MediaCapabilities.FinalizeAvailable) {
        [bool]$Settings.Finalize
    }
    else {
        $false
    }

    return [pscustomobject]@{
        Multisession = $multisession
        Finalize     = $finalize

        AutoVerify   = [bool]$Settings.AutoVerify
        Logging      = [bool]$Settings.Logging
        EjectAfter   = [bool]$Settings.EjectAfter
    }
}
function New-WritePlan {
    [CmdletBinding()]
    param([object]$Media,[object]$MediaCapabilities,[object]$Selection)
    if ($null -eq $Media) { throw 'Yazma planı için medya bilgisi bulunamadı.' }
    if ($null -eq $MediaCapabilities) { throw 'Yazma planı için medya yetenekleri bulunamadı.' }
    if ($null -eq $Selection) { throw 'Yazma planı için kullanıcı seçimi bulunamadı.' }
    $flags=Get-MediaStateFlag -Media $Media
    $plan=[pscustomobject]@{
        MediaProfile=[string]$MediaCapabilities.Profile; MediaStatus=[int]$Media.MediaStatus; IsBlank=[bool]$flags.IsBlank; IsAppendable=[bool]$flags.IsAppendable; IsFinalSession=[bool]$flags.IsFinalSession; IsFinalized=[bool]$flags.IsFinalized
        WriteMode=[string]$MediaCapabilities.WriteMode; ImportExistingFileSystem=[bool]$MediaCapabilities.ImportExistingFileSystem; Multisession=[bool]$Selection.Multisession; Finalize=[bool]$Selection.Finalize
        AutoVerify=[bool]$Selection.AutoVerify; Logging=[bool]$Selection.Logging; EjectAfter=[bool]$Selection.EjectAfter; MultisessionRequired=[bool]$MediaCapabilities.MultisessionRequired; FinalizeEditable=[bool]$MediaCapabilities.FinalizeEditable
    }
    $check=Test-MediaWritePlan -Media $Media -MediaCapabilities $MediaCapabilities -WriteSettings $plan
    if (-not $check.Ok) { throw [string]$check.Reason }
    return $plan
}
function New-WriteSettingRow {
    [CmdletBinding()]
    param(
        [string]$Name,
        [string]$Label,
        [bool]$Value,
        [bool]$Editable,
        [bool]$Required=$false
    )

    return [pscustomobject]@{
        Name     = $Name
        Label    = $Label
        Value    = $Value
        Editable = $Editable
        Required = $Required
    }
}

function Get-WriteSettingRow {
    [CmdletBinding()]
    param(
        [object]$MediaCapabilities,
        [object]$Selection
    )

    $rows = New-Object System.Collections.Generic.List[object]

    if ([bool]$MediaCapabilities.MultisessionAvailable) {
        [void]$rows.Add((New-WriteSettingRow `
            -Name 'Multisession' `
            -Label 'Multisession' `
            -Value ([bool]$Selection.Multisession) `
            -Editable ([bool]$MediaCapabilities.MultisessionEditable) `
            -Required ([bool]$MediaCapabilities.MultisessionRequired)))
    }

    if ([bool]$MediaCapabilities.FinalizeAvailable) {
        [void]$rows.Add((New-WriteSettingRow `
            -Name 'Finalize' `
            -Label 'Finalize' `
            -Value ([bool]$Selection.Finalize) `
            -Editable ([bool]$MediaCapabilities.FinalizeEditable)))
    }

    [void]$rows.Add((New-WriteSettingRow -Name 'AutoVerify' -Label 'AutoVerify' -Value ([bool]$Selection.AutoVerify) -Editable $true))
    [void]$rows.Add((New-WriteSettingRow -Name 'Logging'    -Label 'Logging'    -Value ([bool]$Selection.Logging)    -Editable $true))
    [void]$rows.Add((New-WriteSettingRow -Name 'EjectAfter' -Label 'EjectAfter' -Value ([bool]$Selection.EjectAfter) -Editable $true))

    return $rows.ToArray()
}

function Get-WriteSettingEditableIndex {
    [CmdletBinding()]
    param([object[]]$Rows)

    $indexes = New-Object System.Collections.Generic.List[int]

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        if ([bool]$Rows[$i].Editable) {
            [void]$indexes.Add($i)
        }
    }

    if ($indexes.Count -eq 0) {
        throw 'Bu medya için değiştirilebilir yazma ayarı bulunamadı.'
    }

    return $indexes.ToArray()
}

function Get-WriteSettingRowIndex {
    [OutputType([int])]
    [CmdletBinding()]
    param(
        [object[]]$Rows,
        [string]$Name
    )

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        if ([string]::Equals([string]$Rows[$i].Name, $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $i
        }
    }

    return -1
}

function Get-WriteSettingValue {
    [OutputType([bool])]
    [CmdletBinding()]
    param(
        [object[]]$Rows,
        [string]$Name
    )

    $index = Get-WriteSettingRowIndex -Rows $Rows -Name $Name

    if ($index -lt 0) {
        return $false
    }

    return [bool]$Rows[$index].Value
}

function Get-WriteSettingsFlow {
    [CmdletBinding()]
    param(
        [object]$MediaCapabilities,
        [object[]]$Rows
    )

    $flow = New-Object System.Collections.Generic.List[string]

    if ($MediaCapabilities.WriteMode -eq 'APPEND_SESSION') {
        [void]$flow.Add('APPENDABLE')
        [void]$flow.Add('Yeni SESSION')
    }
    else {
        [void]$flow.Add('BLANK')
    }

    [void]$flow.Add('WRITE')

    if (Get-WriteSettingValue -Rows $Rows -Name 'Finalize') {
        [void]$flow.Add('FINALIZE')
    }

    if (Get-WriteSettingValue -Rows $Rows -Name 'AutoVerify') {
        [void]$flow.Add('VERIFY')
    }

    if (Get-WriteSettingValue -Rows $Rows -Name 'Logging') {
        [void]$flow.Add('LOG')
    }

    if (Get-WriteSettingValue -Rows $Rows -Name 'EjectAfter') {
        [void]$flow.Add('EJECT')
    }

    return ($flow -join ' → ')
}

function Show-WriteSettingsScreen {
    [CmdletBinding()]
    param(
        [object]$MediaCapabilities,
        [object[]]$Rows,
        [int[]]$EditableIndexes,
        [int]$SelectedEditableIndex
    )

    Clear-WriteModeBody
    Write-Host ''
    Write-Host 'YAZMA AYARLARI' -ForegroundColor $script:Theme.Color.Primary
    Write-Host '────────────────────────────────────────'
    Write-Host ''

    Write-Host 'YAZMA MODU' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''
    Write-Host $MediaCapabilities.WriteModeText -ForegroundColor $script:Theme.Color.Success
    Write-Host $MediaCapabilities.Description -ForegroundColor $script:Theme.Color.Warning
    Write-Host ''

    Write-Host 'MEDYAYA BAĞLI AYARLAR' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''

    foreach ($row in @($Rows)) {
        if ($row.Name -ne 'Multisession' -and $row.Name -ne 'Finalize') {
            continue
        }

        $prefix = '  '
        $rowIndex = Get-WriteSettingRowIndex -Rows $Rows -Name $row.Name

        if ([bool]$row.Editable -and $EditableIndexes[$SelectedEditableIndex] -eq $rowIndex) {
            $prefix = '► '
        }

        if ([bool]$row.Required) {
            Write-Host (
                '{0}{1,-12}: AÇIK (ZORUNLU)' -f $prefix,$row.Label
            ) -ForegroundColor $script:Theme.Color.Warning
        }
        else {
            Write-Host (
                '{0}{1,-12}: {2}' -f $prefix,$row.Label,(Get-BooleanText $row.Value)
            )
        }
    }

    Write-Host ''
    Write-Host 'KULLANICI TERCİHLERİ' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''

    foreach ($name in @('AutoVerify','Logging','EjectAfter')) {
        $index = Get-WriteSettingRowIndex -Rows $Rows -Name $name

        if ($index -lt 0) {
            continue
        }

        $row = $Rows[$index]
        $prefix = '  '

        if ($EditableIndexes[$SelectedEditableIndex] -eq $index) {
            $prefix = '► '
        }

        Write-Host (
            '{0}{1,-12}: {2}' -f $prefix,$row.Label,(Get-BooleanText $row.Value)
        )
    }

    Write-Host ''
    Write-Host 'YAZMA ÖZETİ' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''
    Write-Host (Get-WriteSettingsFlow -MediaCapabilities $MediaCapabilities -Rows $Rows) -ForegroundColor $script:Theme.Color.Warning
    Write-Host ''
    Write-Host '↑ ↓ Seç    ← → Değiştir    ENTER Onayla    ESC İptal' -ForegroundColor $script:Theme.Color.Muted
}

function Set-WriteSettingRowValue {
    [CmdletBinding()]
    param(
        [object[]]$Rows,
        [int]$RowIndex,
        [bool]$Value
    )

    if ($RowIndex -lt 0 -or $RowIndex -ge $Rows.Count) {
        throw 'Yazma ayarı satır indeksi geçersiz.'
    }

    if (-not [bool]$Rows[$RowIndex].Editable) {
        return
    }

    $Rows[$RowIndex].Value = $Value
}

function Convert-WriteSettingsRowsToSelection {
    [CmdletBinding()]
    param([object[]]$Rows)

    return [pscustomobject]@{
        Multisession = Get-WriteSettingValue -Rows $Rows -Name 'Multisession'
        Finalize     = Get-WriteSettingValue -Rows $Rows -Name 'Finalize'
        AutoVerify   = Get-WriteSettingValue -Rows $Rows -Name 'AutoVerify'
        Logging      = Get-WriteSettingValue -Rows $Rows -Name 'Logging'
        EjectAfter   = Get-WriteSettingValue -Rows $Rows -Name 'EjectAfter'
    }
}

function New-WriteSettingsUiState {
    [CmdletBinding()]
    param([object]$MediaCapabilities, [object]$Selection)
    $rows = @(Get-WriteSettingRow -MediaCapabilities $MediaCapabilities -Selection $Selection)
    $editableIndexes = @(Get-WriteSettingEditableIndex -Rows $rows)
    return [pscustomobject]@{
        Rows=$rows
        EditableIndexes=$editableIndexes
        SelectedEditable=0
    }
}
function Invoke-WriteSettingsInputLoop {
    [CmdletBinding()]
    param([object]$MediaCapabilities, [object]$State)
    $rows = $State.Rows
    $editableIndexes = $State.EditableIndexes
    while ($true) {
        $key = [Console]::ReadKey($true)
        switch ($key.Key) {
            'UpArrow' {
                if ($State.SelectedEditable -gt 0) {
                    $State.SelectedEditable--
                    Show-WriteSettingsScreen -MediaCapabilities $MediaCapabilities -Rows $rows -EditableIndexes $editableIndexes -SelectedEditableIndex $State.SelectedEditable
                }
            }
            'DownArrow' {
                if ($State.SelectedEditable -lt ($editableIndexes.Count - 1)) {
                    $State.SelectedEditable++
                    Show-WriteSettingsScreen -MediaCapabilities $MediaCapabilities -Rows $rows -EditableIndexes $editableIndexes -SelectedEditableIndex $State.SelectedEditable
                }
            }
            'LeftArrow' {
                Set-WriteSettingRowValue -Rows $rows -RowIndex $editableIndexes[$State.SelectedEditable] -Value $false
                Show-WriteSettingsScreen -MediaCapabilities $MediaCapabilities -Rows $rows -EditableIndexes $editableIndexes -SelectedEditableIndex $State.SelectedEditable
            }
            'RightArrow' {
                Set-WriteSettingRowValue -Rows $rows -RowIndex $editableIndexes[$State.SelectedEditable] -Value $true
                Show-WriteSettingsScreen -MediaCapabilities $MediaCapabilities -Rows $rows -EditableIndexes $editableIndexes -SelectedEditableIndex $State.SelectedEditable
            }
            'Spacebar' {
                $index = $editableIndexes[$State.SelectedEditable]
                Set-WriteSettingRowValue -Rows $rows -RowIndex $index -Value (-not [bool]$rows[$index].Value)
                Show-WriteSettingsScreen -MediaCapabilities $MediaCapabilities -Rows $rows -EditableIndexes $editableIndexes -SelectedEditableIndex $State.SelectedEditable
            }
            'Enter' {
                Clear-WriteModeBody
                return Convert-WriteSettingsRowsToSelection -Rows $rows
            }
            'Escape' {
                Clear-WriteModeBody
                return $null
            }
        }
    }
}
function Select-WriteSetting {
    [CmdletBinding()]
    param(
        [object]$MediaCapabilities,
        [object]$Selection
    )
    if ($null -eq $MediaCapabilities) { throw 'Yazma ayarı seçimi için medya yetenekleri bulunamadı.' }
    if ($null -eq $Selection) { throw 'Yazma ayarı seçimi için seçim nesnesi bulunamadı.' }
    if (-not [bool]$MediaCapabilities.CanWrite) { throw 'Yazma yapılamayan medya için yazma ayarı seçilemez.' }
    $state = New-WriteSettingsUiState -MediaCapabilities $MediaCapabilities -Selection $Selection
    Show-WriteSettingsScreen `
        -MediaCapabilities $MediaCapabilities `
        -Rows $state.Rows `
        -EditableIndexes $state.EditableIndexes `
        -SelectedEditableIndex $state.SelectedEditable
    return Invoke-WriteSettingsInputLoop -MediaCapabilities $MediaCapabilities -State $state
}
function Test-MediaReady {
    <#
    .SYNOPSIS
    Medyanın istenen yazma moduna hazır olup olmadığını denetler.
    .DESCRIPTION
    Medya yazmaya uygun değilse (ör. boş, eklenebilir veya son oturum durumlarından hiçbiri yoksa) hata fırlatır.
    .PARAMETER Media
    Get-MediaInfo çıktısı.
    .PARAMETER Multisession
    Önceki oturumlara ekleme (append) modu isteniyorsa $true.
    .EXAMPLE
    PS> Test-MediaReady -Media $media -Multisession $false
    #>
    [CmdletBinding()]
    param([object]$Media,[bool]$Multisession)
    if ($null -eq $Media) {throw 'Medya bilgisi alınamadı.'}
	if ($Media.MediaType -notin $script:WritableMediaTypes) {
	    throw (
        	'Bu medya türü Optik Medya Arşiv Sistemi tarafından desteklenmiyor: {0} (MediaType={1}).' -f
	        $Media.MediaName,
	        $Media.MediaType
	    )
	}
    $state=[int]$Media.MediaStatus
    if (($state -band $script:MediaState.Damaged) -ne 0) {throw 'Medya hasarlı görünüyor.'}
    if (($state -band $script:MediaState.NonEmptySession) -ne 0) {throw 'Medyanın son oturumu tamamlanmamış. IMAPI bu medyayı kullanıma uygun bildirmedi.'}
    if (($state -band $script:MediaState.UnsupportedMedia) -ne 0) {throw 'Medya IMAPI2 tarafından desteklenmiyor.'}
    if (($state -band $script:MediaState.WriteProtected) -ne 0) {throw 'Medya yazma korumalı.'}
    if (($state -band $script:MediaState.Finalized) -ne 0) {throw 'Medya sonlandırılmış; yazılamaz.'}
    if (($state -band $script:MediaState.EraseRequired) -ne 0) {throw 'Medya yeniden yazılabilir ancak önce silinmesi gerekiyor.'}
    if ($Media.IsBlank) {return}

    # APPENDABLE veya FINAL_SESSION durumundaki dolu medya için
    # bu uygulamada yeni SESSION yalnızca multisession ile yazılır.
    if (($Media.IsAppendable -or $Media.IsFinalSession) -and $Multisession) {return}

    if (($Media.IsAppendable -or $Media.IsFinalSession) -and -not $Multisession) {
        throw 'Medya dolu/ek session yazılabilir durumda. Bu medya için Çok oturumlu seçeneği AÇIK olmalıdır.'
    }

    throw ('Medya yazmaya uygun değil. Boş, appendable veya son session yazılabilir durumda değil. MediaStatus=0x{0:X}' -f $state)
}
function Set-SpeedAndGetActual {
    [OutputType([long])]
    [CmdletBinding()]
    param([object]$Format,[Int64]$RequestedSpeed)
    if ($RequestedSpeed -le 0) {return [Int64]$Format.CurrentWriteSpeed}
    [void]$Format.SetWriteSpeed($RequestedSpeed,$false)
    return [Int64]$Format.CurrentWriteSpeed
}
function Select-Speed {
    [CmdletBinding()]
    param([object]$Media)
    $speeds=@($Media.SupportedSpeeds | Sort-Object)
    if ($speeds.Count -eq 0) {throw 'Medya desteklenen yazma hızını bildirmedi.'}
    Write-UiSection 'YAZMA HIZI'
    Write-UiRow -Label 'Medya' -Value (Get-UiMediaText $Media)
    Write-Host ''
    for ($i=0;$i -lt $speeds.Count;$i++) {
        $x=Convert-SpeedToX -SectorsPerSecond ([Int64]$speeds[$i]) -MediaType ([int]$Media.MediaType)
        Write-Host '  ' -NoNewline
        Write-Host ('[{0}]' -f ($i+1)) -NoNewline -ForegroundColor $script:Theme.Color.Warning
        Write-Host ('  ' + ('{0}x' -f $x).PadRight(6)) -NoNewline -ForegroundColor $script:Theme.Color.Text
        Write-Host ('{0:N0} sektör/s' -f [Int64]$speeds[$i]) -ForegroundColor $script:Theme.Color.Muted
    }
    Write-Host '  [Q]  İptal' -ForegroundColor $script:Theme.Color.Muted
    Write-Host ''
    while ($true) {
        Write-Host '  Seçiminiz: ' -NoNewline
        $v=Read-Host
        if ($null -eq $v) {continue}
        if ($v.Trim().ToUpperInvariant() -eq 'Q') {return $null}
        $n=0
        if (-not [int]::TryParse($v,[ref]$n)) {Write-Host '  Geçersiz giriş. Listeden bir numara girin.' -ForegroundColor $script:Theme.Color.Warning;continue}
        if ($n -lt 1 -or $n -gt $speeds.Count) {Write-Host '  Geçersiz seçim. Listede olmayan numara.' -ForegroundColor $script:Theme.Color.Warning;continue}
        $requested=[Int64]$speeds[$n-1]
        Write-Host ''
        return [pscustomobject]@{
            SectorsPerSecond=$requested
            X=(Convert-SpeedToX -SectorsPerSecond $requested -MediaType ([int]$Media.MediaType))
        }
    }
}

# ============================================================
# BURN PROGRESS
# ============================================================
function Initialize-ImapiBurnProgressBridge {
    [CmdletBinding()]
    param()
    $typeName='ImapiBurnProgressBridge'
    $loaded=$false
    try {
        $loaded=$null -ne ([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetType($typeName,$false) } | Where-Object { $null -ne $_ } | Select-Object -First 1)
    } catch {$loaded=$false}
    if ($loaded) {return}

    Add-Type -TypeDefinition @'
using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;

[ComVisible(true)]
[ComImport]
[Guid("2735413C-7F64-5B0F-8F00-5D77AFBE261E")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IImapiDDiscFormat2DataEvents
{
    [DispId(0x200)]
    void Update(
        [In, MarshalAs(UnmanagedType.IDispatch)] object sender,
        [In, MarshalAs(UnmanagedType.IDispatch)] object progress);
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class ImapiBurnProgressBridge : IImapiDDiscFormat2DataEvents, IDisposable
{
    private IConnectionPoint _connectionPoint;
    private int _cookie;
    private bool _disposed;

    private static object GetProperty(object target, string propertyName)
    {
        if (target == null) return null;
        try
        {
            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                null);
        }
        catch
        {
            return null;
        }
    }

    private static long GetInt64(object value)
    {
        if (value == null) return 0;
        try { return Convert.ToInt64(value, CultureInfo.InvariantCulture); }
        catch { return 0; }
    }

    private static string GetActionName(int action)
    {
        switch (action)
        {
            case 0: return "MEDYA KONTROL";
            case 1: return "MEDYA FORMATLANIYOR";
            case 2: return "DONANIM HAZIRLANIYOR";
            case 3: return "LAZER KALİBRASYONU";
            case 4: return "VERİ YAZILIYOR";
            case 5: return "FINALIZE EDİLİYOR";
            case 6: return "TAMAMLANDI";
            case 7: return "DOĞRULANIYOR";
            default: return "BİLİNMEYEN İŞLEM";
        }
    }

    private static string FormatHms(long seconds)
    {
        if (seconds < 0) seconds = 0;
        long h = seconds / 3600;
        long m = (seconds % 3600) / 60;
        long s = seconds % 60;
        return String.Format(CultureInfo.InvariantCulture, "{0:00}:{1:00}:{2:00}", h, m, s);
    }

    private static void WriteConsoleLine(string line)
    {
        lock (typeof(ImapiBurnProgressBridge))
        {
            try
            {
                int width = 0;
                try { width = Console.WindowWidth; } catch { }
                if (width > 1 && line.Length >= width)
                    line = line.Substring(0, width - 1);

                if (width > 1)
                    line = line.PadRight(width - 1);

                Console.Write("\r" + line);
            }
            catch { }
        }
    }

    private static string MakeBar(int percent, int width)
    {
        if (percent < 0) percent = 0;
        if (percent > 100) percent = 100;
        if (width < 8) width = 8;
        int filled = (int)((width * percent) / 100.0);
        int empty = width - filled;
        return new String('█', filled) + new String('░', empty);
    }

    private void Report(object progress)
    {
        try
        {
            int action = (int)GetInt64(GetProperty(progress, "CurrentAction"));
            long elapsed = GetInt64(GetProperty(progress, "ElapsedTime"));
            long remaining = GetInt64(GetProperty(progress, "RemainingTime"));
            long total = GetInt64(GetProperty(progress, "TotalTime"));
            string actionName = GetActionName(action);

            int percent = 0;
            long sectorCount = GetInt64(GetProperty(progress, "SectorCount"));
            long startLba = GetInt64(GetProperty(progress, "StartLba"));
            long lastWrittenLba = GetInt64(GetProperty(progress, "LastWrittenLba"));
            long writtenSectors = lastWrittenLba - startLba;

            if (action == 4 && sectorCount > 0 && writtenSectors >= 0)
            {
                percent = (int)((writtenSectors * 100L) / sectorCount);
                if (percent > 100) percent = 100;
            }
            else if (action >= 6)
            {
                percent = 100;
            }

            string timeText = String.Format(
                CultureInfo.InvariantCulture,
                "Geçen:{0}  Kalan:{1}",
                FormatHms(elapsed),
                FormatHms(remaining));

            if (action == 4)
            {
                string bar = MakeBar(percent, 38);
                WriteConsoleLine(String.Format(
                    CultureInfo.InvariantCulture,
                    "YAZMA [{0}] {1,3}%  {2}  {3}",
                    bar, percent, actionName, timeText));
            }
            else if (action == 5)
            {
                string bar = MakeBar(100, 38);
                WriteConsoleLine(String.Format(
                    CultureInfo.InvariantCulture,
                    "YAZMA [{0}] 100%  {1}  {2}",
                    bar, actionName, timeText));
            }
            else if (action == 6)
            {
                string bar = MakeBar(100, 38);
                WriteConsoleLine(String.Format(
                    CultureInfo.InvariantCulture,
                    "YAZMA [{0}] 100%  {1}",
                    bar, actionName));
            }
            else
            {
                WriteConsoleLine(String.Format(
                    CultureInfo.InvariantCulture,
                    "YAZMA: {0}  {1}",
                    actionName, timeText));
            }
        }
        catch { }
    }

    public static ImapiBurnProgressBridge Register(object format)
    {
        if (format == null) throw new ArgumentNullException("format");
        var bridge = new ImapiBurnProgressBridge();
        try
        {
            var container = (IConnectionPointContainer)format;
            Guid eventGuid = typeof(IImapiDDiscFormat2DataEvents).GUID;
            container.FindConnectionPoint(ref eventGuid, out bridge._connectionPoint);
            if (bridge._connectionPoint == null)
                throw new COMException("IMAPI DDiscFormat2DataEvents connection point bulunamadı.");

            bridge._connectionPoint.Advise(bridge, out bridge._cookie);
            return bridge;
        }
        catch
        {
            bridge.Dispose();
            throw;
        }
    }

    public void Update(object sender, object progress)
    {
        Report(progress);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_connectionPoint != null)
            {
                if (_cookie != 0)
                {
                    try { _connectionPoint.Unadvise(_cookie); } catch { }
                    _cookie = 0;
                }
                try { Marshal.ReleaseComObject(_connectionPoint); } catch { }
                _connectionPoint = null;
            }
        }
        catch { }
        GC.SuppressFinalize(this);
    }

    ~ImapiBurnProgressBridge()
    {
        Dispose();
    }
}
'@
}
function Register-BurnProgress {
    [CmdletBinding()]
    param([object]$Format)
    try {
        Initialize-ImapiBurnProgressBridge
        $bridge=[ImapiBurnProgressBridge]::Register($Format)
        return [pscustomobject]@{
            Subscriber=$bridge
            SourceId=''
            Mode='COM-CONNECTION-POINT'
            Error=''
        }
    } catch {
        $err=Get-FriendlyError $_.Exception
        return [pscustomobject]@{
            Subscriber=$null
            SourceId=''
            Mode='FAILED'
            Error=$err
        }
    }
}
function Unregister-BurnProgress {
    [CmdletBinding()]
    param([object]$Registration)
    if ($null -eq $Registration) {return}

    $subscriber=$null
    $sourceId=''
    try {$subscriber=$Registration.Subscriber} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    try {$sourceId=[string]$Registration.SourceId} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    if ($null -ne $subscriber) {
        try {
            if ($subscriber -is [IDisposable]) {
                $subscriber.Dispose()
                return
            }
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try {Stop-Job -Job $subscriber -ErrorAction SilentlyContinue} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try {Remove-Job -Job $subscriber -Force -ErrorAction SilentlyContinue} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
    if (-not [string]::IsNullOrWhiteSpace($sourceId)) {
        try {Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
}
function Get-ReadOnlyMediaTypeForWritableMedia {
    [OutputType([int])]
    [CmdletBinding()]
    param([int]$InitialMediaType)
    switch ($InitialMediaType) {
        1 { return 1 }   # CD-ROM family for CD media
        2 { return 1 }   # CD-R -> CD-ROM after finalization
        3 { return 1 }   # CD-RW -> CD-ROM/read-only state when closed
        6 { return 4 }   # DVD+R -> DVD-ROM after finalization
        8 { return 4 }   # DVD+R DL -> DVD-ROM after finalization
        9 { return 4 }   # DVD-R -> DVD-ROM after finalization
        11 { return 4 }  # DVD-R DL -> DVD-ROM after finalization
        18 { return 17 } # BD-R -> BD-ROM after finalization
        default { return 0 }
    }
}
function Test-FinalizedReadOnlyState {
    [CmdletBinding()]
    param([object]$Media,[int]$InitialMediaType)
    if ($null -eq $Media -or $InitialMediaType -le 0) { return $false }
    $expectedReadOnlyType=Get-ReadOnlyMediaTypeForWritableMedia -InitialMediaType $InitialMediaType
    if ($expectedReadOnlyType -le 0) { return $false }
    $status=[int]$Media.MediaStatus
    $physicalType=[int]$Media.MediaType
    if ($physicalType -ne $expectedReadOnlyType) { return $false }
    if ([bool]$Media.IsAppendable -or [bool]$Media.IsFinalSession) { return $false }
    $readOnly=($status -band $script:MediaState.WriteProtected) -ne 0
    $unsupported=($status -band $script:MediaState.UnsupportedMedia) -ne 0
    return ($readOnly -or $unsupported)
}
function Get-CurrentFinalizeResult {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Recorder,[bool]$FinalizeRequested,[int]$InitialMediaType=0,[int]$TimeoutSeconds=120,[int]$PollIntervalSeconds=2)
    if ($null -eq $Recorder) { throw 'Finalize sonucu için Recorder nesnesi bulunamadı.' }
    $deadline=(Get-Date).AddSeconds($TimeoutSeconds);$attempt=0
    while ((Get-Date) -lt $deadline) {
        $attempt++
        try {
            $media=Get-MediaInfo -Recorder $Recorder -RetryIntervalSeconds 1 -RetryTimeoutSeconds 5
            $status=[int]$media.MediaStatus
            if (($status -band $script:MediaState.Damaged) -ne 0) { return 'MEDYA HASARLI' }
            if ($FinalizeRequested) {
                # Bazı DVD-R sürücülerinde finalize sonrasında IMAPI,
                # FINALIZED bitini vermek yerine medyayı DVD-ROM + write-protected/unsupported
                # olarak raporlayabilir. Başarılı Write sonrasında appendable değilse bu,
                # fiziksel olarak kapanmış medya durumunun beklenen sonucudur.
                if (($status -band $script:MediaState.Finalized) -ne 0) { return 'SONLANDIRILDI' }
                if (Test-FinalizedReadOnlyState -Media $media -InitialMediaType $InitialMediaType) {
                    return 'SONLANDIRILDI'
                }
                if (($status -band $script:MediaState.UnsupportedMedia) -ne 0) { return 'MEDYA DESTEKLENMİYOR' }
                if (($status -band $script:MediaState.WriteProtected) -ne 0) { return 'MEDYA YAZMA KORUMALI' }
                if ($attempt -eq 1) { Write-Host 'Finalize sonucu bekleniyor...' -ForegroundColor $script:Theme.Color.Warning } else { $elapsed=[int]((Get-Date).Subtract($deadline.AddSeconds(-$TimeoutSeconds)).TotalSeconds); Write-Host ('Finalize bekleniyor... {0}/{1} sn' -f $elapsed,$TimeoutSeconds) -ForegroundColor $script:Theme.Color.Warning }
                Start-Sleep -Seconds $PollIntervalSeconds; continue
            }
            if (($status -band $script:MediaState.Finalized) -ne 0) { return 'OTOMATİK SONLANDIRILDI' }
            if (($status -band $script:MediaState.Appendable) -ne 0) { return 'AÇIK' }
            Start-Sleep -Seconds $PollIntervalSeconds; continue
        } catch {
            $hr=Get-HResultHex $_.Exception
            if ($hr -eq '0xC0AA0205' -or $hr -eq '0xC0AA0207' -or $hr -eq '0xC0AA0202' -or $hr -eq '0xC0AA0206') { Write-Host 'Medya finalize kontrolü için henüz hazır değil; tekrar denenecek...' -ForegroundColor $script:Theme.Color.Warning; Start-Sleep -Seconds $PollIntervalSeconds; continue }
            throw
        }
    }
    if ($FinalizeRequested) { return 'SONLANDIRMA ZAMAN AŞIMI' }
    return 'DOĞRULANAMADI'
}
function Invoke-Burn {
    [CmdletBinding()]
    param([object]$Recorder,[object]$ImageStream,[Int64]$SelectedSpeed,[bool]$FinalizeRequested,[object]$WritePlan=$null)
    if ($null -eq $Recorder) { throw 'Yazma için Recorder nesnesi bulunamadı.' }
    if ($null -eq $ImageStream) { throw 'Yazma için ImageStream bulunamadı.' }
    $format=New-Object -ComObject IMAPI2.MsftDiscFormat2Data
    $progressRegistration=$null
    try {
        $format.Recorder=$Recorder;$format.ClientName='OptikMedyaArsiv';$format.BufferUnderrunFreeDisabled=$false
        if ($null -ne $WritePlan) {
            if ([bool]$WritePlan.Finalize -ne [bool]$FinalizeRequested) { throw 'Yazma planı ile finalize parametresi uyuşmuyor.' }
            $format.ForceMediaToBeClosed=[bool]$WritePlan.Finalize
        } else { $format.ForceMediaToBeClosed=[bool]$FinalizeRequested }
        $actualSpeed=Set-SpeedAndGetActual -Format $format -RequestedSpeed $SelectedSpeed
        $mediaType=[int]$format.CurrentPhysicalMediaType
        $progressRegistration=Register-BurnProgress -Format $format
        $progressSubscriber=if ($null -ne $progressRegistration) {$progressRegistration.Subscriber} else {$null}
        if ($null -eq $progressSubscriber) {
            $progressError=if ($null -ne $progressRegistration -and $progressRegistration.PSObject.Properties.Name -contains 'Error') {[string]$progressRegistration.Error} else {'Bilinmeyen hata.'}
            Write-Warning ('Yazma ilerleme göstergesi başlatılamadı: {0} Fiziksel yazma devam edecek.' -f $progressError)
        } else {
            Write-Host 'Yazma ilerleme olayı bağlandı: IMAPI Update.' -ForegroundColor DarkGreen
        }
        $writeSucceeded=$false;$writeError=''
        try {
            Write-Host '';Write-Host 'FİZİKSEL YAZMA BAŞLADI...' -ForegroundColor $script:Theme.Color.Primary;Write-Host 'Lütfen medya çıkarılmadan bekleyin.' -ForegroundColor $script:Theme.Color.Warning;Write-Host ''
            try {
                [void]$format.Write($ImageStream)
                $writeSucceeded=$true
            }
            catch {
                $writeError=Get-FriendlyError $_.Exception
            }
        } finally {
            Unregister-BurnProgress -Registration $progressRegistration;Write-Host '';if ($writeSucceeded) { Write-Host 'YAZMA: 100% TAMAMLANDI' -ForegroundColor $script:Theme.Color.Success } else { Write-Host 'YAZMA: TAMAMLANAMADI' -ForegroundColor $script:Theme.Color.Danger }
        }
        if (-not $writeSucceeded) {
            return [pscustomobject]@{
                WriteResult='BAŞARISIZ'
                AppliedSpeedX=(Convert-SpeedToX -SectorsPerSecond $actualSpeed -MediaType $mediaType)
                FinalizeResult='YAPILMADI'
                FailureStage='PHYSICAL_STATE_UNKNOWN'
                Reason=('Fiziksel yazma çağrısı tamamlanamadı. Medyanın fiziksel durumu belirsiz. {0}' -f $writeError)
            }
        }

        try {
            $finalizeResult=Get-CurrentFinalizeResult -Recorder $Recorder -FinalizeRequested $FinalizeRequested -InitialMediaType $mediaType
        }
        catch {
            return [pscustomobject]@{
                WriteResult='BAŞARILI'
                AppliedSpeedX=(Convert-SpeedToX -SectorsPerSecond $actualSpeed -MediaType $mediaType)
                FinalizeResult='SONUÇ_BELİRSİZ'
                FailureStage='PHYSICAL_STATE_UNKNOWN'
                Reason=('Fiziksel yazma tamamlandı ancak medya/finalize durumu doğrulanamadı. Medyanın fiziksel durumu belirsiz. {0}' -f (Get-FriendlyError $_.Exception))
            }
        }

        return [pscustomobject]@{
            WriteResult='BAŞARILI'
            AppliedSpeedX=(Convert-SpeedToX -SectorsPerSecond $actualSpeed -MediaType $mediaType)
            FinalizeResult=$finalizeResult
            FailureStage=''
            Reason=''
        }
    } finally { Unregister-BurnProgress -Registration $progressRegistration;Unregister-ComObject $format }
}

# ============================================================
# VERIFY
# ============================================================
function Wait-VolumePath {
    [CmdletBinding()]
    param([object]$Recorder,[int]$TimeoutSeconds=300)
    $deadline=(Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        try {
            foreach ($path in @($Recorder.VolumePathNames)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$path) -and (Test-Path -LiteralPath $path)) {
                    return (([string]$path).TrimEnd('\')+'\')
                }
            }
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        Start-Sleep -Seconds 1
    }
    return $null
}
# DESTEK KATMANI HASH / DOSYA
function Write-ConsoleProgress {
    [CmdletBinding()]
    param(
        [string]$Label,
        [Int64]$CurrentBytes=0,
        [Int64]$TotalBytes=0,
        [string]$Status='',
        [int]$BarWidth=44,
        [bool]$NewLineOnComplete=$false
    )

    try {
        $percent=0
        if ($TotalBytes -gt 0) {
            $percent=[int](($CurrentBytes*100)/$TotalBytes)
            if ($percent -lt 0) {$percent=0}
            if ($percent -gt 100) {$percent=100}
        }

        $filled=[int](($BarWidth*$percent)/100)
        $empty=$BarWidth-$filled
        $bar=($script:Theme.Symbol.BarFull * $filled)+($script:Theme.Symbol.BarEmpty * $empty)

        $line='{0}: {1}/{2} [{3}] {4,3}%' -f `
            $Label,
            (Format-Size $CurrentBytes),
            (Format-Size $TotalBytes),
            $bar,
            $percent

        if (-not [string]::IsNullOrWhiteSpace($Status)) {
            $line+='  '+$Status
        }

        $width=0
        try {$width=[Console]::WindowWidth} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        if ($width -gt 1 -and $line.Length -ge $width) {
            $line=$line.Substring(0,$width-1)
        }

        try {
            [Console]::Write("`r" + $line.PadRight([Math]::Max(0,$width-1)))
        }
        catch {
            Write-Host ("`r"+$line) -NoNewline
        }

        if ($NewLineOnComplete -and $percent -ge 100) {
            Write-Host ''
        }
    }
    catch {
        try {Write-Host $Label -NoNewline} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
}

function Get-FileSha256WithProgress {
    [CmdletBinding()]
    param([string]$Path,[string]$Label='VERIFY',[Int64]$DoneBefore=0,[Int64]$TotalBytes=0)
    $fs=$null;$sha=$null;$buffer=New-Object byte[] ($script:BufferSizeBytes);[Int64]$readTotal=0
    try {
        $fs=New-Object IO.FileStream($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read,1MB,[IO.FileOptions]::SequentialScan)
        $sha=[Security.Cryptography.SHA256]::Create()
        [Int64]$length=$fs.Length
        while (($n=$fs.Read($buffer,0,$buffer.Length)) -gt 0) {
            [void]$sha.TransformBlock($buffer,0,$n,$buffer,0);$readTotal+=$n
            $overall=$DoneBefore+$readTotal
            $pct=0
            if ($TotalBytes -gt 0) {$pct=[int](($overall*100)/$TotalBytes)}
            if ($pct -gt 100) {$pct=100}
            Write-ConsoleProgress -Label $Label -CurrentBytes $overall -TotalBytes $TotalBytes -Status (("dosya: {0}/{1}" -f (Format-Size $readTotal),(Format-Size $length)))
        }
        [void]$sha.TransformFinalBlock($buffer,0,0)
        $hash=([BitConverter]::ToString($sha.Hash).Replace('-','')).ToUpperInvariant()
        return [pscustomobject]@{Sha256=$hash;Length=$length}
    } finally {
        if ($null -ne $sha) {$sha.Dispose()}
        if ($null -ne $fs) {$fs.Dispose()}
    }
}

function Get-SessionPhysicalFile {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([string]$SessionRoot,[string]$ManifestFileName)
    $result=@()
    if (-not (Test-Path -LiteralPath $SessionRoot -PathType Container)) {return @($result)}
    $manifestFullPath=Join-Path $SessionRoot $ManifestFileName
    $result=@(Get-ChildItem -LiteralPath $SessionRoot -Recurse -File -Force -ErrorAction Stop | Where-Object {-not [string]::Equals($_.FullName,$manifestFullPath,[StringComparison]::OrdinalIgnoreCase)})
    return @($result)
}
function Invoke-Verify {
    <#
    .SYNOPSIS
    Yazılan SESSION içeriğini MANIFEST ile karşılaştırarak doğrular.
    .DESCRIPTION
    Klasör ağacını, dosya sayısını, boyutları, dosya SHA256 değerlerini ve fiziksel MANIFEST SHA256 değerini denetler. İlk uyuşmazlıkta nedenle birlikte başarısız sonuç döner.
    .PARAMETER Recorder
    IMAPI2 disc recorder COM nesnesi.
    .PARAMETER SessionName
    Doğrulanacak SESSION adı.
    .PARAMETER Files
    Yazılan dosyalar.
    .PARAMETER Manifest
    Yazılan MANIFEST bilgisi.
    .PARAMETER DirectoryPaths
    Beklenen klasör yolları.
    .OUTPUTS
    PSCustomObject: Ok (bool), Reason (string).
    .EXAMPLE
    PS> Invoke-Verify -Recorder $recorder -SessionName $sessionName -Files $files -Manifest $manifest
    #>
    [OutputType([pscustomobject])]
    [CmdletBinding()]
    param([object]$Recorder,[string]$SessionName,[object[]]$Files,[object]$Manifest,[string[]]$DirectoryPaths=@())
    Write-Host '';Write-Host 'VERIFY: Disk yeniden bağlanıyor...' -ForegroundColor $script:Theme.Color.Primary
    $volumePath=Wait-VolumePath -Recorder $Recorder -TimeoutSeconds $script:VerifyVolumeWaitSeconds
    if ($null -eq $volumePath) {return [pscustomobject]@{Ok=$false;Reason='Disk hacmi Windows tarafından yeniden bağlanmadı.'}}
    $sessionRoot=Join-Path $volumePath $SessionName
    if (-not (Test-Path -LiteralPath $sessionRoot -PathType Container)) {return [pscustomobject]@{Ok=$false;Reason="VERIFY SESSION klasörü bulunamadı: $SessionName"}}

    $expectedDirectorySet=@{}
    foreach ($dirPath in @($DirectoryPaths)) {
        $normalizedDir=ConvertTo-FsiNormalizedPath $dirPath
        if ($normalizedDir -eq '\') {continue}
        $expectedDirectorySet[$normalizedDir]=$true
        $destDir=Join-Path $sessionRoot ($normalizedDir.TrimStart('\').Replace('\',[IO.Path]::DirectorySeparatorChar))
        if (-not (Test-Path -LiteralPath $destDir -PathType Container)) {
            return [pscustomobject]@{Ok=$false;Reason="SESSION içinde klasör yok: $normalizedDir"}
        }
    }

    try {
        $physicalDirectories=@(Get-ChildItem -LiteralPath $sessionRoot -Recurse -Directory -Force -ErrorAction Stop)
    }
    catch {
        return [pscustomobject]@{Ok=$false;Reason=('SESSION klasör ağacı okunamadı: {0}' -f $_.Exception.Message)}
    }

    $physicalDirectorySet=@{}
    foreach ($pd in @($physicalDirectories)) {
        $relativeDir=$pd.FullName.Substring($sessionRoot.Length).TrimStart('\').Replace([IO.Path]::DirectorySeparatorChar,'\')
        $normalizedRelativeDir=ConvertTo-FsiNormalizedPath $relativeDir
        if (-not $expectedDirectorySet.ContainsKey($normalizedRelativeDir)) {
            return [pscustomobject]@{Ok=$false;Reason="SESSION içinde project yapısında olmayan klasör bulundu: $normalizedRelativeDir"}
        }
        $physicalDirectorySet[$normalizedRelativeDir]=$true
    }

    if ($physicalDirectorySet.Count -ne $expectedDirectorySet.Count) {
        return [pscustomobject]@{Ok=$false;Reason=('SESSION klasör sayısı uyuşmuyor. Beklenen: {0}, fiziksel: {1}.' -f $expectedDirectorySet.Count,$physicalDirectorySet.Count)}
    }
    foreach ($expectedDir in @($expectedDirectorySet.Keys)) {
        if (-not $physicalDirectorySet.ContainsKey($expectedDir)) {
            return [pscustomobject]@{Ok=$false;Reason="SESSION içinde beklenen klasör bulunamadı: $expectedDir"}
        }
    }

    $manifestLines=@(Get-Content -LiteralPath $Manifest.Path -Encoding UTF8 -ErrorAction Stop)
    $expected=Get-ArchiveManifestIndex -Lines $manifestLines
    $expectedCount=@($expected.Keys).Count
    $fileCount=@($Files).Count
    if ($expectedCount -ne $fileCount) {return [pscustomobject]@{Ok=$false;Reason=('Manifest dosya sayısı uyuşmuyor. Manifest: {0}, beklenen: {1}.' -f $expectedCount,$fileCount)}}
    [Int64]$totalBytes=0
    foreach ($f in @($Files)) {$totalBytes+=[Int64]$f.Length}
    [Int64]$done=0;$fileIndex=0
    foreach ($f in @($Files)) {
        $fileIndex++;$archivePath=ConvertTo-FsiNormalizedPath $f.DestinationPath
        if (-not $expected.ContainsKey($archivePath)) {return [pscustomobject]@{Ok=$false;Reason="Manifest içinde kayıt yok: $archivePath"}}
        $dest=Join-Path $sessionRoot ($archivePath.TrimStart('\').Replace('\',[IO.Path]::DirectorySeparatorChar))
        if (-not (Test-Path -LiteralPath $dest -PathType Leaf)) {
            Write-ConsoleProgress -Label 'VERIFY - SHA256' -CurrentBytes $done -TotalBytes $totalBytes -Status 'BAŞARISIZ' -NewLineOnComplete $true
            return [pscustomobject]@{Ok=$false;Reason="SESSION içinde dosya yok: $archivePath"}
        }
        $dst=Get-Item -LiteralPath $dest -Force -ErrorAction Stop;$entry=$expected[$archivePath]
        if ([Int64]$dst.Length -ne [Int64]$entry.Size) {
            Write-ConsoleProgress -Label 'VERIFY - SHA256' -CurrentBytes $done -TotalBytes $totalBytes -Status 'BAŞARISIZ' -NewLineOnComplete $true
            return [pscustomobject]@{Ok=$false;Reason="Boyut uyuşmuyor: $archivePath"}
        }
        $hashInfo=Get-FileSha256WithProgress -Path $dest -Label 'VERIFY - SHA256' -DoneBefore $done -TotalBytes $totalBytes
        if ($hashInfo.Sha256 -ne [string]$entry.Sha256) {
            Write-ConsoleProgress -Label 'VERIFY - SHA256' -CurrentBytes $done -TotalBytes $totalBytes -Status 'BAŞARISIZ' -NewLineOnComplete $true
            return [pscustomobject]@{Ok=$false;Reason="SHA256 doğrulama başarısız: $archivePath"}
        }
        $done+=[Int64]$dst.Length
        Write-ConsoleProgress -Label 'VERIFY - SHA256' -CurrentBytes $done -TotalBytes $totalBytes -Status (("dosya: {0}/{1}" -f $fileIndex,$fileCount))
    }
    Write-ConsoleProgress -Label 'VERIFY - SHA256' -CurrentBytes $done -TotalBytes $totalBytes -Status 'tamamlandı' -NewLineOnComplete $true
    $discManifest=Join-Path $sessionRoot $Manifest.FileName
    if (-not (Test-Path -LiteralPath $discManifest -PathType Leaf)) {return [pscustomobject]@{Ok=$false;Reason="$($Manifest.FileName) bulunamadı."}}
    $discManifestHash=Get-FileSha256WithProgress -Path $discManifest -Label 'VERIFY - Manifest' -DoneBefore 0 -TotalBytes 1
    Write-ConsoleProgress -Label 'VERIFY - Manifest' -CurrentBytes 1 -TotalBytes 1 -Status 'tamamlandı' -NewLineOnComplete $true
    if ($discManifestHash.Sha256 -ne [string]$Manifest.Sha256) {return [pscustomobject]@{Ok=$false;Reason='Fiziksel manifest SHA256 doğrulaması başarısız.'}}
    $physical=@(Get-SessionPhysicalFile -SessionRoot $sessionRoot -ManifestFileName $Manifest.FileName)
    $expectedPathSet=@{}
    foreach ($key in @($expected.Keys)) {$expectedPathSet[(ConvertTo-FsiNormalizedPath $key)]=$true}
    foreach ($pf in @($physical)) {
        $relative=$pf.FullName.Substring($sessionRoot.Length).TrimStart('\').Replace([IO.Path]::DirectorySeparatorChar,'\')
        if (-not $expectedPathSet.ContainsKey((ConvertTo-FsiNormalizedPath $relative))) {
            return [pscustomobject]@{Ok=$false;Reason="SESSION içinde manifestte olmayan dosya bulundu: $relative"}
        }
    }
    return [pscustomobject]@{Ok=$true;Reason='VERIFY başarılı.'}
}
function Invoke-Eject {
    [CmdletBinding()]
    param([object]$Recorder)
    try {$Recorder.EjectMedia();return $true} catch {return $false}
}

# ============================================================
# LOG
# ============================================================
function Write-ArchiveLog {
    [CmdletBinding()]
    param([object[]]$Files,[string[]]$SourceRoots,[string]$MediaName,[Int64]$SourceBytes,[Int64]$ImageBytes,[int]$FileCount,
        [double]$SelectedSpeed,[double]$AppliedSpeed,[string]$WriteResult,[string]$Result,[string]$Reason,[string]$RecorderName,
        [string]$Firmware,[string]$VerifyResult,[string]$Stage,[string]$FinalizeResult,[string]$VolumeLabel,[bool]$Multisession,
        [bool]$FinalizeRequested,[string]$SessionName,[string]$FailureStage='',[string]$LogStatus='')
    # Not: Loglama karari cagiran tarafindan verilir (Write-WriteOperationLog / Invoke-WriteFailureHandler).
    # Burada ikinci bir $script:Settings.Logging kontrolu tutmak, kullanici secimini gecersiz kilabilecegi icin kaldirildi.
    try {
        Initialize-AppDirectory;$file=Join-Path $script:LogDir ('Arsiv_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.log')
        $lines=New-Object System.Collections.Generic.List[string]
        [void]$lines.Add('OPTİK MEDYA ARŞİV LOG');[void]$lines.Add("Tarih          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')");[void]$lines.Add("Uygulama       : $script:AppVersion");
        [void]$lines.Add("Session        : $SessionName");[void]$lines.Add("Medya          : $MediaName");[void]$lines.Add("Volume Label   : $VolumeLabel");[void]$lines.Add("Kaydedici      : $RecorderName");[void]$lines.Add("Firmware       : $Firmware");
        [void]$lines.Add("Kaynak Veri    : $(Format-Size $SourceBytes)");[void]$lines.Add("UDF Image      : $(Format-Size $ImageBytes)");[void]$lines.Add("Dosya Sayısı   : $FileCount");[void]$lines.Add("Seçilen Hız    : ${SelectedSpeed}x");[void]$lines.Add("Uygulanan Hız  : ${AppliedSpeed}x");
        [void]$lines.Add("WRITE          : $WriteResult");[void]$lines.Add("VERIFY         : $VerifyResult");[void]$lines.Add("LOG            : $LogStatus");[void]$lines.Add("Çok Oturumlu   : $(Get-BooleanText $Multisession)");[void]$lines.Add("Finalize       : $(Get-BooleanText $FinalizeRequested)");[void]$lines.Add("Finalize Durumu: $(if ($FinalizeRequested) { 'İSTENDİ' } else { 'İSTENMEDİ' })");[void]$lines.Add("Finalize Sonuç : $FinalizeResult");[void]$lines.Add("Aşama          : $Stage");[void]$lines.Add("Hata Aşaması   : $FailureStage");[void]$lines.Add("Sonuç          : $Result");[void]$lines.Add("Açıklama       : $Reason");
        [void]$lines.Add('');[void]$lines.Add('SOURCE ROOTS');[void]$lines.Add('------------------------------------------------------------');$rootNo=1
        foreach ($root in @($SourceRoots | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)} | Sort-Object -Unique) ) {[void]$lines.Add("$rootNo. $root");$rootNo++}
        [void]$lines.Add('');[void]$lines.Add('PROJECT FILES');[void]$lines.Add('----------------------------------------------------------------------------------------------------------------');[void]$lines.Add('SourceRoot | ArchivePath | Size | SHA256 | SourcePath')
        foreach ($f in @($Files)) {[void]$lines.Add(('{0} | {1} | {2} | {3} | {4}' -f $f.SourceRoot,$f.DestinationPath,$f.Length,$f.Sha256,$f.SourcePath))}
        $utf8=New-Object System.Text.UTF8Encoding($false);[IO.File]::WriteAllText($file,[string]::Join([Environment]::NewLine,$lines.ToArray()),$utf8)
        return [pscustomobject]@{Ok=$true;Status='BAŞARILI';Path=$file;Reason='Log oluşturuldu.'}
    } catch {return [pscustomobject]@{Ok=$false;Status='BAŞARISIZ';Path='';Reason=$_.Exception.Message}}
}
# ============================================================
# YAZMA AYARLARI / PRECHECK
# ============================================================
function Reset-BurnSetting {
    [CmdletBinding()]
    param()
    $script:Settings.AutoVerify=$true;$script:Settings.EjectAfter=$true;$script:Settings.Logging=$true
    $script:Settings.Multisession=$false;$script:Settings.Finalize=$true
    $script:UiDirCount=-1
}

function Show-BurnPreflight {
   [CmdletBinding()]
   param([object]$Media,[Int64]$SourceBytes,[Int64]$ImageBytes,[object]$Speed,
    [bool]$VerifyEnabled,[string]$RecorderName,[string]$Firmware,[string]$VolumeLabel,[int]$FileCount,[bool]$Multisession,
    [bool]$FinalizeRequested,[string]$SessionName)
    Write-UiSection 'YAZMA ÖZETİ'
    if ($ImageBytes -le $Media.FreeBytes) {
        Write-Host '  ✓ Kapasite yeterli' -ForegroundColor $script:Theme.Color.Success
    }
    else {
        Write-Host '  ✗ Kapasite yetersiz' -ForegroundColor $script:Theme.Color.Danger
        Write-UiRow -Label 'Image' -Value (Format-Size $ImageBytes) -Note '' -ValueColor Yellow
        Write-UiRow -Label 'Boş alan' -Value (Format-Size $Media.FreeBytes) -Note '' -ValueColor Yellow
        Write-Host '  Yazma başlatılmadı.' -ForegroundColor $script:Theme.Color.Warning
        Wait-ConsoleInput 'Ana ekrana dönmek için Enter'
        return $false
    }
    Write-Host ''
    $mediaText  = Get-UiMediaText $Media
    $mediaNote  = if ($VerbosePreference -eq 'Continue') {'(durum: 0x{0:X})' -f $Media.MediaStatus} else {''}
    $fwNote     = if ([string]::IsNullOrWhiteSpace($Firmware)) {''} else {'(fw {0})' -f $Firmware}
    $usePct     = if ($Media.FreeBytes -gt 0) {100.0 * $ImageBytes / $Media.FreeBytes} else {0.0}
    $useText    = if ($usePct -gt 0 -and $usePct -lt 0.01) {'<%' + (0.01).ToString('N2')} else {'%' + $usePct.ToString('N2')}
    $speedNote  = ''
    try { if ($Speed.PSObject.Properties['SectorsPerSecond']) {$speedNote = '({0:N0} sektör/s)' -f [Int64]$Speed.SectorsPerSecond} } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    Write-UiRow -Label 'Session' -Value ([string]$SessionName)
    Write-UiRow -Label 'Yazıcı' -Value (([string]$RecorderName -replace '\s+',' ').Trim()) -Note $fwNote
    Write-UiRow -Label 'Medya' -Value $mediaText -Note $mediaNote
    Write-UiRow -Label 'Etiket' -Value ([string]$VolumeLabel)
    Write-Host ''
    $contentText = if ($script:UiDirCount -ge 0) {'{0} dosya · {1} klasör · {2}' -f $FileCount,$script:UiDirCount,(Format-Size $SourceBytes)} else {'{0} dosya · {1}' -f $FileCount,(Format-Size $SourceBytes)}
    Write-UiRow -Label 'İçerik' -Value $contentText
    Write-UiRow -Label 'Yazılacak' -Value (Format-Size $ImageBytes) -Note '(UDF yapısı ve manifest dahil)'
    Write-UiRow -Label 'Boş alan' -Value (Format-Size $Media.FreeBytes) -Note ('→  kullanım ' + $useText)
    Write-Host ''
    Write-UiRow -Label 'Hız' -Value ('{0}x' -f $Speed.X) -Note $speedNote
    if ($VerifyEnabled) { Write-UiRow -Label 'Doğrulama' -Value '✓ Açık' -Note '' -ValueColor Green } else { Write-UiRow -Label 'Doğrulama' -Value '✗ Kapalı' -Note '' -ValueColor DarkGray }
    if ($Multisession)  { Write-UiRow -Label 'Çok oturumlu' -Value '✓ Açık' -Note '' -ValueColor Green } else { Write-UiRow -Label 'Çok oturumlu' -Value '✗ Kapalı' -Note '' -ValueColor DarkGray }
    if ($FinalizeRequested) { Write-UiRow -Label 'Sonlandırma' -Value '⚠ Açık' -Note 'disk kapatılır, bir daha yazılamaz' -ValueColor Yellow -NoteColor Yellow } else { Write-UiRow -Label 'Sonlandırma' -Value '✗ Kapalı' -Note '' -ValueColor DarkGray }
    Write-Host ''
    return (Read-UiConfirm 'Yazmayı başlat?')
}
# 5. YAZMA ORKESTRATÖRÜ
function New-WriteOperationResult {
    [CmdletBinding()]
    param(
        [string]$Status,
        [string]$Write,
        [string]$Verify,
        [string]$Log,
        [string]$Finalize,
        [string]$FailureStage = '',
        [string]$Reason,
        [string]$SessionName,
        [string]$VolumeLabel
    )

    return [pscustomobject]@{
        Status      = $Status
        Write       = $Write
        Verify      = $Verify
        Log         = $Log
        Finalize    = $Finalize
        FailureStage= $FailureStage
        Reason      = $Reason
        SessionName = $SessionName
        VolumeLabel = $VolumeLabel
    }
}

function Get-WriteSourceRoot {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([object[]]$Files)

    return @(
        $Files |
            ForEach-Object { [string]$_.SourceRoot } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Show-WriteMediaSummary {
    [CmdletBinding()]
    param(
        [object]$Media,
        [string]$MediaProfile
    )

    Write-Host "Medya tipi : $($Media.MediaName)"

    if ([int]$Media.MediaType -eq 0) {
        Write-Host 'Uygulama   : MEDYA YOK' -ForegroundColor $script:Theme.Color.Warning
    }
    elseif ([bool]$Media.AppSupportedMediaType) {
        Write-Host 'Uygulama   : DESTEKLENİYOR' -ForegroundColor $script:Theme.Color.Success
    }
    elseif ([bool]$Media.IsReadOnlyPhysicalType) {
        Write-Host 'Uygulama   : SALT-OKUNUR MEDYA' -ForegroundColor $script:Theme.Color.Warning
    }
    else {
        Write-Host 'Uygulama   : DESTEKLENMİYOR' -ForegroundColor $script:Theme.Color.Danger
    }

    Write-Host ('Profil     : {0}' -f $MediaProfile)
    Write-Host ('Durum      : 0x{0:X}' -f $Media.MediaStatus)
    Write-Host ('Boş alan   : {0}' -f (Format-Size $Media.FreeBytes))
    Write-Host ('Boş        : {0}' -f (Get-BooleanText $Media.IsBlank))
    Write-Host ('Appendable : {0}' -f (Get-BooleanText $Media.IsAppendable))
    if ([bool]$Media.IsFinalized) {
        $finalizedText = 'AÇIK'
    }
    elseif ([bool]$Media.IsReadOnlyPhysicalType) {
        $finalizedText = 'DOĞRULANAMADI (SADECE OKUNUR)'
    }
    else {
        $finalizedText = 'KAPALI'
    }
    Write-Host ('Finalized  : {0}' -f $finalizedText)
}

function Show-WriteUnavailableMedia {
    [CmdletBinding()]
    param([object]$MediaCapabilities)

    Write-Host ''
    Write-Host 'YAZMA DURUMU' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''
    Write-Host ('Profil     : {0}' -f $MediaCapabilities.Profile)
    Write-Host 'Yazma      : YAPILAMAZ' -ForegroundColor $script:Theme.Color.Danger
    Write-Host ('Açıklama   : {0}' -f $MediaCapabilities.Description) -ForegroundColor $script:Theme.Color.Warning
    if ([string]$MediaCapabilities.Profile -eq 'READ_ONLY') {
        Write-Host 'Not        : Bu medya fiziksel olarak salt-okunur. Uygulama bunu finalize edilmişliği kesin olarak kanıtlanmış bir medya ile eşitlemez.' -ForegroundColor $script:Theme.Color.Muted
    }
    Write-Host ''
}

function Invoke-WriteVerification {
    [CmdletBinding()]
    param(
        [object]$BurnResult,
        [object]$WriteSelection,
        [object]$Recorder,
        [string]$SessionName,
        [object[]]$WriteFiles,
        [object]$Manifest,
        [string[]]$DirectoryPaths=@()
    )

    $result = [pscustomobject]@{
        Verify = 'İSTENMEDİ'
        Reason = ''
    }

    if ([string]$BurnResult.FailureStage -eq 'PHYSICAL_STATE_UNKNOWN') {
        $result.Verify = 'BELİRSİZ'
        $result.Reason = [string]$BurnResult.Reason
        return $result
    }

    if ([string]$BurnResult.WriteResult -eq 'BAŞARISIZ') {
        $result.Reason = [string]$BurnResult.Reason
        return $result
    }

    if ([bool]$WriteSelection.Finalize) {
        if ([string]$BurnResult.FinalizeResult -ne 'SONLANDIRILDI') {
            $result.Reason = "Finalize doğrulaması: $($BurnResult.FinalizeResult)"
            return $result
        }
    }
    else {
        if ([string]$BurnResult.FinalizeResult -notin @('AÇIK','OTOMATİK SONLANDIRILDI')) {
            $result.Reason = "Medya durumu doğrulanamadı: $($BurnResult.FinalizeResult)"
            return $result
        }
    }

    if ([bool]$WriteSelection.AutoVerify) {
        $verify = Invoke-Verify `
            -Recorder $Recorder `
            -SessionName $SessionName `
            -Files $WriteFiles `
            -Manifest $Manifest `
            -DirectoryPaths $DirectoryPaths

        if ($verify.Ok) {
            $result.Verify = 'BAŞARILI'
        }
        else {
            $result.Verify = 'BAŞARISIZ'
            $result.Reason = [string]$verify.Reason
            return $result
        }
    }
    else {
        $result.Verify = 'İSTENMEDİ'
    }

    return $result
}

function Get-WriteOverallResult {
    [CmdletBinding()]
    param(
        [string]$WriteResult,
        [bool]$FinalizeRequested,
        [string]$FinalizeResult,
        [string]$VerifyResult,
        [string]$FailureStage='',
        [string]$Reason
    )

    # Sonuç önceliği bilinçli olarak aşamalıdır:
    # 1) Fiziksel yazma başarısızsa sonraki aşamalar sonucu değiştiremez.
    # 2) Yazma başarılı fakat istenen finalize başarısızsa bu WRITE_FAILED değildir.
    # 3) Finalize başarılıyken verify başarısızsa bu da WRITE_FAILED değildir.
    $result = [pscustomobject]@{
        Status = 'BAŞARISIZ'
        FailureStage = ''
        Reason = $Reason
    }

    if (-not [string]::IsNullOrWhiteSpace($FailureStage)) {
        $result.FailureStage = $FailureStage
        return $result
    }

    if ($WriteResult -eq 'BAŞARISIZ') {
        $result.FailureStage = 'WRITE'
        return $result
    }

    if ($FinalizeRequested) {
        if ($FinalizeResult -ne 'SONLANDIRILDI') {
            $result.FailureStage = 'FINALIZE'
            if ([string]::IsNullOrWhiteSpace($result.Reason)) {
                $result.Reason = "Finalize doğrulaması: $FinalizeResult"
            }
            return $result
        }
    }
    else {
        if ($FinalizeResult -ne 'OTOMATİK SONLANDIRILDI' -and $FinalizeResult -ne 'AÇIK') {
            $result.FailureStage = 'FINALIZE'
            if ([string]::IsNullOrWhiteSpace($result.Reason)) {
                $result.Reason = "Medya durumu doğrulanamadı: $FinalizeResult"
            }
            return $result
        }
    }

    if ($VerifyResult -eq 'BAŞARISIZ') {
        $result.FailureStage = 'VERIFY'
        return $result
    }

    if ($VerifyResult -eq 'İSTENMEDİ') {
        $result.Status = 'BAŞARILI_DOĞRULANMADI'
        return $result
    }

    $result.Status = 'BAŞARILI'
    return $result
}

function Write-WriteOperationLog {
    [CmdletBinding()]
    param(
        [bool]$Enabled,
        [object[]]$Files,
        [object[]]$SourceRoots,
        [object]$Media,
        [Int64]$SourceBytes,
        [Int64]$ImageBytes,
        [double]$SelectedSpeed,
        [double]$AppliedSpeed,
        [string]$WriteResult,
        [string]$OverallResult,
        [string]$Reason,
        [string]$RecorderName,
        [string]$Firmware,
        [string]$VerifyResult,
        [string]$FinalizeResult,
        [string]$FailureStage,
        [string]$VolumeLabel,
        [bool]$Multisession,
        [bool]$FinalizeRequested,
        [string]$SessionName
    )

    if (-not $Enabled) {
        return [pscustomobject]@{Status='KAPALI';Reason=''}
    }

    return Write-ArchiveLog `
        -Files $Files `
        -SourceRoots $SourceRoots `
        -MediaName $Media.MediaName `
        -SourceBytes $SourceBytes `
        -ImageBytes $ImageBytes `
        -FileCount $Files.Count `
        -SelectedSpeed $SelectedSpeed `
        -AppliedSpeed $AppliedSpeed `
        -WriteResult $WriteResult `
        -Result $OverallResult `
        -Reason $Reason `
        -RecorderName $RecorderName `
        -Firmware $Firmware `
        -VerifyResult $VerifyResult `
        -Stage 'Yazma' `
        -FinalizeResult $FinalizeResult `
         -FailureStage $FailureStage `
        -VolumeLabel $VolumeLabel `
        -Multisession $Multisession `
        -FinalizeRequested $FinalizeRequested `
        -SessionName $SessionName
}

function Invoke-WriteFailureHandler {
    <#
    .SYNOPSIS
    Yazma akışında oluşan hatayı raporlar ve standart sonuç nesnesi döndürür.
    .DESCRIPTION
    Hata bilgisini ve aşama sonuçlarını loglar (log sonucu ayrıca raporlanır) ve New-WriteOperationResult ile sonuç nesnesini üretir.
    .PARAMETER Exception
    Yakalanan hata.
    .PARAMETER Media
    Medya bilgisi.
    .PARAMETER Image
    Hazırlanmış UDF imaj paketi.
    .PARAMETER Recorder
    IMAPI2 disc recorder COM nesnesi.
    .PARAMETER WriteSettingsSelection
    Kullanıcının seçtiği yazma ayarları.
    .PARAMETER WriteFiles
    Yazılması planlanan dosyalar.
    .PARAMETER SourceBytes
    Yazılacak toplam kaynak boyutu (bayt).
    .PARAMETER SelectedSpeed
    Seçilen yazma hızı.
    .PARAMETER AppliedSpeed
    Uygulanan yazma hızı.
    .PARAMETER WriteResult
    Yazma aşamasının sonucu.
    .PARAMETER VerifyResult
    Doğrulama aşamasının sonucu.
    .PARAMETER FinalizeResult
    Finalize aşamasının sonucu.
    .PARAMETER FailureStage
    Hatanın oluştuğu aşama.
    .PARAMETER SessionName
    SESSION adı.
    .PARAMETER VolumeLabel
    Disk etiketi.
    .EXAMPLE
    PS> Invoke-WriteFailureHandler -Exception $_.Exception -FailureStage WRITE -SessionName $sessionName
    #>
    [CmdletBinding()]
    param(
        [System.Exception]$Exception,
        [object]$Media,
        [object]$Image,
        [object]$Recorder,
        [object]$WriteSettingsSelection,
        [object[]]$WriteFiles,
        [Int64]$SourceBytes,
        [double]$SelectedSpeed,
        [double]$AppliedSpeed,
        [string]$WriteResult,
        [string]$VerifyResult,
        [string]$FinalizeResult,
        [string]$FailureStage,
        [string]$SessionName,
        [string]$VolumeLabel
    )

    $reason = Get-FriendlyError -Exception $Exception

    if ([string]::IsNullOrWhiteSpace($FailureStage)) {
        if ($WriteResult -eq 'BAŞARISIZ') {
            $FailureStage = 'WRITE'
        }
        elseif ($FinalizeResult -notin @('', 'SONLANDIRILDI', 'OTOMATİK SONLANDIRILDI', 'AÇIK')) {
            $FailureStage = 'FINALIZE'
        }
        elseif ($VerifyResult -eq 'BAŞARISIZ') {
            $FailureStage = 'VERIFY'
        }
        else {
            $FailureStage = 'UNKNOWN'
        }
    }

    Clear-WriteModeBody

    Write-Host ''
    Write-Host 'İŞLEM BAŞARISIZ' -ForegroundColor $script:Theme.Color.Danger
    Write-Host $reason -ForegroundColor $script:Theme.Color.Danger

    $mediaName = ''
    $imageBytes = 0L
    $recorderName = ''
    $firmware = ''

    if ($null -ne $Media) {
        $mediaName = [string]$Media.MediaName
    }

    if ($null -ne $Image) {
        $imageBytes = [Int64]$Image.ImageBytes
    }

    if ($null -ne $Recorder) {
        $recorderName = '{0} {1}' -f $Recorder.VendorId,$Recorder.ProductId
        $firmware = [string]$Recorder.ProductRevision
    }

    $logResult = 'KAPALI'

    $loggingEnabled = if ($null -ne $WriteSettingsSelection) {
        [bool]$WriteSettingsSelection.Logging
    }
    else {
        [bool]$script:Settings.Logging
    }

    if ($loggingEnabled) {
        try {
            $logMultisession = [bool]$script:Settings.Multisession
            $logFinalize = [bool]$script:Settings.Finalize

            if ($null -ne $WriteSettingsSelection) {
                $logMultisession = [bool]$WriteSettingsSelection.Multisession
                $logFinalize = [bool]$WriteSettingsSelection.Finalize
            }

            $log = Write-ArchiveLog `
                -Files $WriteFiles `
                -SourceRoots @(
                    $WriteFiles |
                        ForEach-Object { [string]$_.SourceRoot } |
                        Sort-Object -Unique
                ) `
                -MediaName $mediaName `
                -SourceBytes $SourceBytes `
                -ImageBytes $imageBytes `
                -FileCount $WriteFiles.Count `
                -SelectedSpeed $SelectedSpeed `
                -AppliedSpeed $AppliedSpeed `
                -WriteResult $WriteResult `
                -Result 'BAŞARISIZ' `
                -Reason $reason `
                -RecorderName $recorderName `
                -Firmware $firmware `
                -VerifyResult $VerifyResult `
                -Stage 'Hata' `
                -FinalizeResult $FinalizeResult `
                 -FailureStage $FailureStage `
                -VolumeLabel $VolumeLabel `
                -Multisession $logMultisession `
                -FinalizeRequested $logFinalize `
                -SessionName $SessionName

            $logResult = [string]$log.Status
        }
        catch {
            $logResult = 'BAŞARISIZ'
        }
    }

    Write-Host ('WRITE  : {0}' -f $WriteResult)
    Write-Host ('VERIFY : {0}' -f $VerifyResult)
    Write-Host ('LOG    : {0}' -f $logResult)
    Wait-ConsoleInput 'Ana ekrana dönmek için Enter'

    return New-WriteOperationResult `
        -Status 'FAILURE' `
        -Write $WriteResult `
        -Verify $VerifyResult `
        -Log $logResult `
        -Finalize $FinalizeResult `
        -FailureStage $FailureStage `
        -Reason $reason `
        -SessionName $SessionName `
        -VolumeLabel $VolumeLabel
}

function Clear-WriteResource {
    [CmdletBinding()]
    param(
        [object]$Image,
        [object]$Manifest,
        [object[]]$Recorders
    )

    if ($null -ne $Image -and $null -ne $Image.BurnStream) {
        Unregister-ComObject $Image.BurnStream
    }

    if ($null -ne $Image -and -not [string]::IsNullOrWhiteSpace([string]$Image.TempImagePath)) {
        try {
            Remove-Item -LiteralPath $Image.TempImagePath -Force -ErrorAction SilentlyContinue
        }
        catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }

    if ($null -ne $Image) {
        try { Unregister-ComObject $Image.Result } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        try { Unregister-ComObject $Image.FileSystem } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }

    Remove-TempManifest -Manifest $Manifest

    foreach ($recorder in @($Recorders)) {
        Unregister-ComObject $recorder
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Show-WriteCompletionStatusBlock {
    [CmdletBinding()]
    param(
        [string]$WriteResult,
        [string]$VerifyResult,
        [string]$LogResult,
        [object]$Log,
        [bool]$EjectAfter,
        [object]$Recorder,
        [string]$FinalizeResult,
        [bool]$FinalizeRequested,
        [string]$FailureStage
    )
    Write-Host ''
    $writeColor = if ($WriteResult -eq 'BAŞARILI') { $script:Theme.Color.Success } else { $script:Theme.Color.Danger }
    Write-Host ('WRITE  : {0}' -f $WriteResult) -ForegroundColor $writeColor
    Write-Host ('VERIFY : {0}' -f $VerifyResult)
    if ($FinalizeRequested) {
        Write-Host ('FINALIZE: İSTENDİ → {0}' -f $FinalizeResult)
    }
    else {
        Write-Host ('FINALIZE: İSTENMEDİ → {0}' -f $FinalizeResult)
    }
    Write-Host ('AŞAMA  : {0}' -f $FailureStage)
    Write-Host ('LOG    : {0}' -f $LogResult)
    if ($null -ne $Log -and $Log.Status -eq 'BAŞARISIZ') {
        Write-Host ('⚠ Log oluşturulamadı: {0}' -f $Log.Reason) -ForegroundColor $script:Theme.Color.Warning
    }
    if ($EjectAfter -and $FailureStage -ne 'PHYSICAL_STATE_UNKNOWN') {
        Write-Host ''
        Write-Host 'Medya çıkarılıyor...' -ForegroundColor $script:Theme.Color.Primary
        if (Invoke-Eject $Recorder) {
            Write-Host '✓ Medya çıkarıldı.' -ForegroundColor $script:Theme.Color.Success
        }
        else {
            Write-Host 'Medya otomatik çıkarılamadı.' -ForegroundColor $script:Theme.Color.Warning
        }
    }
    Write-Host ''
}
function Show-WriteCompletionResultBlock {
    [CmdletBinding()]
    param(
        [string]$OverallResult,
        [string]$FailureStage,
        [string]$FinalizeResult,
        [string]$Reason,
        [string]$SessionName,
        [string]$VolumeLabel,
        [Int64]$SourceBytes,
        [Int64]$ImageBytes,
        [int]$FileCount,
        [double]$AppliedSpeed,
        [string]$WriteResult,
        [string]$VerifyResult,
        [string]$LogResult,
        [bool]$FinalizeRequested
    )
    if ($OverallResult -eq 'BAŞARILI') {
        Write-Host 'İŞLEM BAŞARILI' -ForegroundColor $script:Theme.Color.Success
    }
    elseif ($OverallResult -eq 'BAŞARILI_DOĞRULANMADI') {
        Write-Host 'YAZMA BAŞARILI — VERIFY YAPILMADI' -ForegroundColor $script:Theme.Color.Warning
    }
    else {
        Write-Host 'İŞLEM BAŞARISIZ' -ForegroundColor $script:Theme.Color.Danger
        if ($FailureStage -eq 'WRITE') {
            Write-Host 'Sonuç: Fiziksel yazma başarısız.' -ForegroundColor $script:Theme.Color.Danger
        }
        elseif ($FailureStage -eq 'FINALIZE') {
            Write-Host ('Sonuç: Finalize istendi ancak doğrulanamadı ({0}).' -f $FinalizeResult) -ForegroundColor $script:Theme.Color.Danger
        }
        elseif ($FailureStage -eq 'VERIFY') {
            Write-Host 'Sonuç: Yazma tamamlandı ancak VERIFY başarısız.' -ForegroundColor $script:Theme.Color.Danger
        }
        elseif ($FailureStage -eq 'PHYSICAL_STATE_UNKNOWN') {
            Write-Host 'SONUÇ: MEDYANIN FİZİKSEL DURUMU BELİRSİZ.' -ForegroundColor $script:Theme.Color.Danger
            Write-Host 'Aynı medyaya yeniden yazmayın. Medya durumunu ayrıca kontrol edin.' -ForegroundColor $script:Theme.Color.Warning
        }
        elseif (-not [string]::IsNullOrWhiteSpace($Reason)) {
            Write-Host ('Sonuç: {0}' -f $Reason) -ForegroundColor $script:Theme.Color.Danger
        }
    }
    Write-Host "Session : $SessionName"
    Write-Host "Volume  : $VolumeLabel"
    Write-Host ('Veri    : {0}' -f (Format-Size $SourceBytes))
    Write-Host ('Image   : {0}' -f (Format-Size $ImageBytes))
    Write-Host ('Dosya   : {0}' -f $FileCount)
    Write-Host ('Hız     : {0}x' -f $AppliedSpeed)
    Write-Host "WRITE   : $WriteResult"
    Write-Host "VERIFY  : $VerifyResult"
    Write-Host "LOG     : $LogResult"
    if ($FinalizeRequested) {
        Write-Host "Finalize: İSTENDİ → $FinalizeResult"
    }
    else {
        Write-Host "Finalize: İSTENMEDİ → $FinalizeResult"
    }
    if (-not [string]::IsNullOrWhiteSpace($Reason)) {
        Write-Host "Açıklama: $Reason" -ForegroundColor $script:Theme.Color.Warning
    }
}
function Show-WriteCompletionSummary {
    [CmdletBinding()]
    param(
        [string]$WriteResult,
        [string]$VerifyResult,
        [string]$LogResult,
        [object]$Log,
        [bool]$EjectAfter,
        [object]$Recorder,
        [string]$OverallResult,
        [string]$SessionName,
        [Int64]$SourceBytes,
        [Int64]$ImageBytes,
        [int]$FileCount,
        [double]$AppliedSpeed,
        [string]$FinalizeResult,
        [bool]$FinalizeRequested,
        [string]$FailureStage,
        [string]$Reason,
        [string]$VolumeLabel
    )
    Clear-WriteModeBody
    Show-WriteCompletionStatusBlock `
        -WriteResult $WriteResult `
        -VerifyResult $VerifyResult `
        -LogResult $LogResult `
        -Log $Log `
        -EjectAfter $EjectAfter `
        -Recorder $Recorder `
        -FinalizeResult $FinalizeResult `
        -FinalizeRequested $FinalizeRequested `
        -FailureStage $FailureStage
    Show-WriteCompletionResultBlock `
        -OverallResult $OverallResult `
        -FailureStage $FailureStage `
        -FinalizeResult $FinalizeResult `
        -Reason $Reason `
        -SessionName $SessionName `
        -VolumeLabel $VolumeLabel `
        -SourceBytes $SourceBytes `
        -ImageBytes $ImageBytes `
        -FileCount $FileCount `
        -AppliedSpeed $AppliedSpeed `
        -WriteResult $WriteResult `
        -VerifyResult $VerifyResult `
        -LogResult $LogResult `
        -FinalizeRequested $FinalizeRequested
    Wait-ConsoleInput 'Ana ekrana dönmek için Enter'
}

function Get-WriteMediaContext {
    [CmdletBinding()]
    param([object]$Recorder)

    $media = Get-MediaInfo -Recorder $Recorder
    $mediaProfile = Get-MediaProfile -Media $media
    $mediaCapabilities = Get-MediaWriteCapability -MediaProfile $mediaProfile

    $selection = New-WriteSettingsSelection `
        -MediaCapabilities $mediaCapabilities `
        -Settings $script:Settings

    if ($script:WriteModeActive) {
        Update-WriteModeHeader -Text (
            'YAZMA MODU: {0} - {1}' -f
            $mediaCapabilities.WriteModeText,
            $mediaCapabilities.Description
        )
    }

    Show-WriteMediaSummary -Media $media -MediaProfile $mediaProfile

    if (-not [bool]$mediaCapabilities.CanWrite) {
        Show-WriteUnavailableMedia -MediaCapabilities $mediaCapabilities
        Wait-ConsoleInput 'Ana ekrana dönmek için Enter'

        return [pscustomobject]@{
            Cancelled = $true
            Reason = [string]$mediaCapabilities.Description
            Media = $media
            MediaProfile = $mediaProfile
            MediaCapabilities = $mediaCapabilities
            Selection = $null
        }
    }

    Test-MediaReady `
        -Media $media `
        -Multisession $selection.Multisession

    $selection = Select-WriteSetting `
        -MediaCapabilities $mediaCapabilities `
        -Selection $selection

    if ($null -eq $selection) {
        return [pscustomobject]@{
            Cancelled = $true
            Reason = 'Yazma ayarları seçimi kullanıcı tarafından iptal edildi.'
            Media = $media
            MediaProfile = $mediaProfile
            MediaCapabilities = $mediaCapabilities
            Selection = $null
        }
    }

    return [pscustomobject]@{
        Cancelled = $false
        Reason = ''
        Media = $media
        MediaProfile = $mediaProfile
        MediaCapabilities = $mediaCapabilities
        Selection = $selection
    }
}

function Confirm-WritePreflight {
    [OutputType([bool])]
    [CmdletBinding()]
    param(

        [object]$Media,
        [Int64]$SourceBytes,
        [Int64]$ImageBytes,
        [object]$SpeedChoice,
        [bool]$VerifyEnabled,
        [string]$RecorderName,
        [string]$Firmware,
        [string]$VolumeLabel,
        [int]$FileCount,
        [bool]$Multisession,
        [bool]$FinalizeRequested,
        [string]$SessionName
    )

    return [bool](Show-BurnPreflight `
        -Media $Media `
        -SourceBytes $SourceBytes `
        -ImageBytes $ImageBytes `
        -Speed $SpeedChoice `
        -VerifyEnabled $VerifyEnabled `
        -RecorderName $RecorderName `
        -Firmware $Firmware `
        -VolumeLabel $VolumeLabel `
        -FileCount $FileCount `
        -Multisession $Multisession `
        -FinalizeRequested $FinalizeRequested `
        -SessionName $SessionName)
}

function Select-WriteRecorder {
    [CmdletBinding()]
    param([object[]]$Recorders)

    if ($Recorders.Count -eq 0) {
        throw 'Sistemde IMAPI2 tarafından görülen yazıcı bulunamadı.'
    }

    Write-Host ''
    Write-Host 'YAZICI SEÇİMİ' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''

    for ($i = 0; $i -lt $Recorders.Count; $i++) {
        Write-Host (
            '{0}. {1} {2}  Firmware: {3}' -f
            ($i + 1),
            $Recorders[$i].VendorId,
            $Recorders[$i].ProductId,
            $Recorders[$i].ProductRevision
        )
    }

    while ($true) {
        $value = (Read-Host 'Yazıcı numarası (Q=İptal)').Trim()

        if ($value.ToUpperInvariant() -eq 'Q') {
            return $null
        }

        $number = 0
        if ([int]::TryParse($value,[ref]$number) -and $number -ge 1 -and $number -le $Recorders.Count) {
            return $Recorders[$number - 1]
        }

        Write-Host 'Geçersiz seçim.' -ForegroundColor $script:Theme.Color.Warning
    }
}

function New-WritePayload {
    <#
    .SYNOPSIS
    Yazma için gereken MANIFEST ve UDF imaj paketini hazırlar.
    .DESCRIPTION
    Önce MANIFEST dosyasını üretir, ardından yazılacak içeriği imaj paketine dönüştürür. Adımlar arayüzde numaralı adım satırları olarak gösterilir.
    .PARAMETER Files
    Yazılacak dosyalar.
    .PARAMETER Media
    Medya bilgisi.
    .PARAMETER Recorder
    IMAPI2 disc recorder COM nesnesi.
    .PARAMETER SpeedChoice
    Seçilen yazma hızı.
    .PARAMETER WriteSelection
    Yazma ayarları (Multisession, Finalize vb.).
    .PARAMETER SessionName
    SESSION adı.
    .PARAMETER VolumeLabel
    Disk etiketi.
    .EXAMPLE
    PS> New-WritePayload -Files $files -Media $media -Recorder $recorder -SpeedChoice $speed -WriteSelection $selection -SessionName $sessionName -VolumeLabel $label
    #>
    [CmdletBinding()]
    param(
        [object[]]$Files,
        [object]$Media,
        [object]$Recorder,
        [object]$SpeedChoice,
        [object]$WriteSelection,
        [string]$SessionName,
        [string]$VolumeLabel
    )

    $sourceRoots = @(Get-WriteSourceRoot -Files $Files)
    $writeFiles = @($Files)
    $manifest = $null
    $image = $null

    try {
        Start-UiStep -No 1 -Name 'Manifest' -Count @($writeFiles).Count -Unit 'dosya'

        $manifest = New-ManifestFile `
            -Files $writeFiles `
            -SourceRoots $sourceRoots `
            -Recorder $Recorder `
            -Media $Media `
            -SelectedSpeed $SpeedChoice `
            -VolumeLabel $VolumeLabel `
            -Multisession $WriteSelection.Multisession `
            -Finalize $WriteSelection.Finalize `
            -SessionName $SessionName

        Stop-UiStep ('{0} dosya' -f @($writeFiles).Count)

        $manifestLines = @(Get-Content `
            -LiteralPath $manifest.Path `
            -Encoding UTF8 `
            -ErrorAction Stop)

        $manifestIndex = Get-ArchiveManifestIndex -Lines $manifestLines
        $manifestIndexSha256 = Get-ManifestIndexCanonicalSha256 -Index $manifestIndex

        $image = New-UdfImagePackage `
            -Files $writeFiles `
            -DirectoryPaths @(Get-ProjectDirectoryPath) `
            -Manifest $manifest `
            -VolumeLabel $VolumeLabel `
            -MediaType $Media.MediaType `
            -Recorder $Recorder `
            -Multisession $WriteSelection.Multisession `
            -MediaIsBlank $Media.IsBlank `
            -SessionName $SessionName

        if ([Int64]$image.ImageBytes -gt [Int64]$Media.FreeBytes) {
            throw 'UDF image medya kapasitesinden büyük.'
        }

        return [pscustomobject]@{
            Files               = $writeFiles
            SourceRoots         = $sourceRoots
            Manifest            = $manifest
            ManifestIndex       = $manifestIndex
            ManifestIndexSha256 = $manifestIndexSha256
            Image               = $image
            SourceBytes         = [Int64]$manifest.TotalBytes
        }
    }
    catch {
        Stop-UiLiveAbort
        if ($null -ne $image) {
            if ($null -ne $image.BurnStream) {
                try { Unregister-ComObject $image.BurnStream } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$image.TempImagePath)) {
                try {
                    if (Test-Path -LiteralPath $image.TempImagePath -PathType Leaf) {
                        Remove-Item -LiteralPath $image.TempImagePath -Force -ErrorAction SilentlyContinue
                    }
                } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
            if ($null -ne $image.Result) {
                try { Unregister-ComObject $image.Result } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
            if ($null -ne $image.FileSystem) {
                try { Unregister-ComObject $image.FileSystem } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
        }
        Remove-TempManifest -Manifest $manifest
        throw
    }
}

function Confirm-WriteMediaBeforeBurn {
    [CmdletBinding()]
    param(
        [object]$Recorder,
        [object]$InitialMedia,
        [object]$MediaCapabilities,
        [object]$WriteSelection,
        [object]$Image
    )

    Write-Host ''
    Write-Host 'YAZMA ÖNCESİ SON MEDYA KONTROLÜ...' -ForegroundColor $script:Theme.Color.Primary

    $finalMedia = Get-MediaInfo -Recorder $Recorder

    if ([int]$finalMedia.MediaType -ne [int]$InitialMedia.MediaType) {
        throw (
            'Onay sonrası medya değişti. Önceki medya: {0}, mevcut medya: {1}.' -f
            $InitialMedia.MediaName,
            $finalMedia.MediaName
        )
    }

    Test-MediaReady `
        -Media $finalMedia `
        -Multisession $WriteSelection.Multisession

    if ([Int64]$Image.ImageBytes -gt [Int64]$finalMedia.FreeBytes) {
        throw (
            'Onay sonrası medya kapasitesi yetersiz. Image: {0}, Kullanılabilir: {1}.' -f
            (Format-Size $Image.ImageBytes),
            (Format-Size $finalMedia.FreeBytes)
        )
    }

    if ([bool]$finalMedia.IsBlank -ne [bool]$InitialMedia.IsBlank) {
        throw 'Onay sonrası medya durumu değişti. Yazma başlatılmadı.'
    }

    if ([bool]$finalMedia.IsAppendable -ne [bool]$InitialMedia.IsAppendable) {
        throw 'Onay sonrası medya appendable durumu değişti. Yazma başlatılmadı.'
    }

    if ([bool]$finalMedia.IsFinalSession -ne [bool]$InitialMedia.IsFinalSession) {
        throw 'Onay sonrası medya FINAL_SESSION durumu değişti. Yazma başlatılmadı.'
    }

    if ([bool]$finalMedia.IsFinalized) {
        throw 'Onay sonrası medya sonlandırılmış görünüyor. Yazma başlatılmadı.'
    }

    $writePlan = New-WritePlan `
        -Media $finalMedia `
        -MediaCapabilities $MediaCapabilities `
        -Selection $WriteSelection

    Write-Host '✓ Son medya kontrolü başarılı.' -ForegroundColor $script:Theme.Color.Success
    Write-Host (
        '✓ WritePlan: {0} | Multisession={1} | Finalize={2}' -f
        $writePlan.MediaProfile,
        (Get-BooleanText $writePlan.Multisession),
        (Get-BooleanText $writePlan.Finalize)
    ) -ForegroundColor $script:Theme.Color.Success

    return $writePlan
}

function Confirm-ManifestRecheck {
    [CmdletBinding()]
    param(
        [object[]]$WriteFiles,
        [hashtable]$MasterManifestIndex,
        [string]$MasterManifestIndexSha256
    )

    Write-Host ''
    Write-Host 'MANIFEST-RECHECK' -ForegroundColor $script:Theme.Color.Primary
    Write-Host 'Fiziksel yazma öncesi kaynaklar yeniden kontrol ediliyor...' -ForegroundColor $script:Theme.Color.Primary
    Write-Host ''

    $failureText = ''

    try {
        $recheck = New-ManifestRecheck -Files $WriteFiles

        Write-Host ('MASTER MANIFEST SHA256  : {0}' -f $MasterManifestIndexSha256) -ForegroundColor $script:Theme.Color.Muted
        Write-Host ('MANIFEST-RECHECK SHA256 : {0}' -f $recheck.Sha256) -ForegroundColor $script:Theme.Color.Muted
        Write-Host ('MASTER dosya sayısı     : {0}' -f $MasterManifestIndex.Count) -ForegroundColor $script:Theme.Color.Muted
        Write-Host ('RECHECK dosya sayısı    : {0}' -f $recheck.FileCount) -ForegroundColor $script:Theme.Color.Muted

        $same = (
            $recheck.FileCount -eq $MasterManifestIndex.Count -and
            [string]::Equals(
                [string]$recheck.Sha256,
                [string]$MasterManifestIndexSha256,
                [StringComparison]::OrdinalIgnoreCase
            )
        )

        if ($same) {
            Write-Host ''
            Write-Host '✓ MANIFEST-RECHECK BAŞARILI.' -ForegroundColor $script:Theme.Color.Success
            Write-Host 'Kaynaklar ilk MANIFEST ile uyumlu.' -ForegroundColor $script:Theme.Color.Success
            Write-Host ''
            return $true
        }

        $failureText = 'İlk MANIFEST ile MANIFEST-RECHECK aynı değil: kaynak dosyalar değişmiş.'
    }
    catch {
        $failureText = ('MANIFEST-RECHECK tamamlanamadı: {0}' -f $_.Exception.Message)
    }

    # Medya tek seferlik yazılabilir. Image, manifestle tutarlı olduğu kanıtlanamadığı için
    # kullanıcıya "mevcut image ile devam" seçeneği sunulmaz; yazma iptal edilir.
    Write-Host ''
    Write-Host '⚠ KAYNAK DOĞRULANAMADI.' -ForegroundColor $script:Theme.Color.Danger
    Write-Host $failureText -ForegroundColor $script:Theme.Color.Warning
    Write-Host ''
    Write-Host 'Medya tek seferlik yazılabilir olduğu için yazma güvenlik nedeniyle İPTAL EDİLDİ.' -ForegroundColor $script:Theme.Color.Warning
    Write-Host 'Project kaynaklarını kontrol edip (F3 ile yenileyin) yeniden deneyin.' -ForegroundColor $script:Theme.Color.Warning
    Wait-ConsoleInput 'Ana ekrana dönmek için Enter'
    return $false
}

function Invoke-WriteFromProject {
    <#
    .SYNOPSIS
    PROJECT içeriğini diske yazma akışının tamamını yönetir.
    .DESCRIPTION
    Hazırlık, medya denetimi, yazma ayarları, yazma, finalize, doğrulama ve loglama aşamalarını yürütür. WRITE / FINALIZE / VERIFY / LOG sonuçları ayrı ayrı raporlanır. PROJECT boşsa CANCELLED sonucu döner.
    .OUTPUTS
    PSCustomObject: Status, Write, Verify, Log, Finalize, Reason, SessionName, VolumeLabel.
    .EXAMPLE
    PS> Invoke-WriteFromProject
    #>
    [CmdletBinding()]
    param()
    [object[]]$files=@();[object[]]$writeFiles=@();[Int64]$sourceBytes=0;[double]$selectedSpeed=0;[double]$appliedSpeed=0
$writeResult='YAPILMADI'
$verifyResult='YAPILMADI'
$logResult='YAPILMADI'
$log=$null
$reason=''
$finalizeResult='YAPILMADI'
$overallResult='BAŞARISIZ'
$failureStage=''
$manifest=$null;$image=$null;$media=$null;$recorder=$null;$recorders=@();$sessionName='';$sessionVolumePath=$null
$writeSettingsSelection=$null;$writePlan=$null;$volumeLabel=''
    try {
        Reset-BurnSetting
        $totalForProgress = Measure-ProjectFileCount -Node $script:ProjectRoot
        Start-UiProgress -Label 'HAZIRLANIYOR' -Total $totalForProgress
        try {
            Update-ProjectMetadataRecursive
        } finally {
            Stop-UiProgress
        }
        $files=@(Get-ProjectArchiveFile);$projectDirs=@(Get-ProjectDirectoryPath)
      if ($files.Count -eq 0 -and $projectDirs.Count -eq 0) {
	    Write-Host 'Project boş. Önce SOURCE tarafından içerik ekleyin.' -ForegroundColor $script:Theme.Color.Warning

	    return [pscustomobject]@{
	        Status      = 'CANCELLED'
	        Write       = $writeResult
	        Verify      = $verifyResult
	        Log         = $logResult
	        Finalize    = $finalizeResult
	        Reason      = 'Project boş.'
	        SessionName = $sessionName
	        VolumeLabel = $volumeLabel
	    }
	}
        foreach ($f in @($files)) {
            if (-not $f.SourceExists) {throw ('Kaynak dosya mevcut değil: {0}' -f $f.SourcePath)}
            if ([string]$f.Name -match '^MANIFEST_\d{8}_\d{6}_\d{3}\.txt$') {throw ('Project içinde ayrılmış manifest adı kullanılamaz: {0}' -f $f.Name)}
        }


        Write-Host '';Write-Host 'MANIFEST İÇİN KAYNAK KİMLİKLERİ HAZIRLANACAK...' -ForegroundColor $script:Theme.Color.Primary;Write-Host ''
	$sourceBytes=0
	Clear-Host
        Start-WriteModeHeader -Text 'OPTİK MEDYA ARŞİV - YAZMA HAZIRLIĞI'
        Write-Host ''
        Write-Host 'YAZICI SEÇİMİ' -ForegroundColor $script:Theme.Color.Primary
        Write-Host ''

        $recorders = @(Get-Recorder)
        $recorder = Select-WriteRecorder -Recorders $recorders

        if ($null -eq $recorder) {
            Stop-WriteModeHeader
            return New-WriteOperationResult `
                -Status 'CANCELLED' `
                -Write $writeResult `
                -Verify $verifyResult `
                -Log $logResult `
                -Finalize $finalizeResult `
                -Reason 'Yazıcı seçimi kullanıcı tarafından iptal edildi.' `
                -SessionName $sessionName `
                -VolumeLabel $volumeLabel
        }

        Write-Host ''
        Write-Host 'MEDYA KONTROLÜ' -ForegroundColor $script:Theme.Color.Primary
        Write-Host ''

        $mediaContext = Get-WriteMediaContext -Recorder $recorder
        $media = $mediaContext.Media
        $mediaCapabilities = $mediaContext.MediaCapabilities
        $writeSettingsSelection = $mediaContext.Selection

        if ($mediaContext.Cancelled) {
            Stop-WriteModeHeader
            return New-WriteOperationResult `
                -Status 'CANCELLED' `
                -Write $writeResult `
                -Verify $verifyResult `
                -Log $logResult `
                -Finalize $finalizeResult `
                -Reason $mediaContext.Reason `
                -SessionName $sessionName `
                -VolumeLabel $volumeLabel
        }

        # Erken kapasite ön kontrolü: image üretilmeden önce ham veri boyutu medya boş alanıyla karşılaştırılır.
        # (UDF yapısının eklediği küçük fark için image sonrası kontrol ayrıca devam eder.)
        $estimatedBytes=[Int64](($files | Measure-Object -Property Length -Sum).Sum)
        if ($estimatedBytes -gt [Int64]$media.FreeBytes) {
            throw ('Project verisi medya kapasitesinden büyük: {0} > {1}.' -f (Format-Size $estimatedBytes),(Format-Size ([Int64]$media.FreeBytes)))
        }

        if (-not $media.IsBlank) {
            $sessionVolumePath=Wait-VolumePath -Recorder $recorder -TimeoutSeconds $script:VolumeWaitSeconds
            if ($null -eq $sessionVolumePath) {throw 'Mevcut session hacmi Windows tarafından bağlanmadı (120 sn beklendi).'}
        }
        $sessionIdentity=New-ArchiveSessionIdentity -VolumePath $sessionVolumePath;$sessionName=[string]$sessionIdentity.SessionName

       Write-Host ''
	if (Test-Udf250Support) {
	    Write-Host ('  {0} UDF 2.50 destekleniyor' -f $script:Theme.Symbol.Check) -ForegroundColor $script:Theme.Color.Success
	    Write-Host ''
	}

	$speedChoice=Select-Speed -Media $media

	if ($null -eq $speedChoice) {
	    Stop-WriteModeHeader
	    return [pscustomobject]@{
	        Status      = 'CANCELLED'
	        Write       = $writeResult
	        Verify      = $verifyResult
	        Log         = $logResult
	        Finalize    = $finalizeResult
	        Reason      = 'Yazma hızı seçimi kullanıcı tarafından iptal edildi.'
	        SessionName = $sessionName
	        VolumeLabel = $volumeLabel
	    }
	}
        $selectedSpeed=[double]$speedChoice.X
        $volumeLabel=New-ArchiveVolumeLabel -SessionName $sessionName
        $sourceRoots=@($files|ForEach-Object {[string]$_.SourceRoot}|Where-Object {-not [string]::IsNullOrWhiteSpace($_)}|Sort-Object -Unique)
        $writeFiles=@($files)

        $payload = New-WritePayload `
            -Files $files `
            -Media $media `
            -Recorder $recorder `
            -SpeedChoice $speedChoice `
            -WriteSelection $writeSettingsSelection `
            -SessionName $sessionName `
            -VolumeLabel $volumeLabel

        $writeFiles = @($payload.Files)
        $sourceRoots = @($payload.SourceRoots)
        $manifest = $payload.Manifest
        $masterManifestIndex = $payload.ManifestIndex
        $masterManifestIndexSha256 = [string]$payload.ManifestIndexSha256
        $image = $payload.Image
        $sourceBytes = [Int64]$payload.SourceBytes

        $recorderName = '{0} {1}' -f $recorder.VendorId,$recorder.ProductId
        $firmware = [string]$recorder.ProductRevision

        $preflight = Confirm-WritePreflight `
            -Media $media `
            -SourceBytes $sourceBytes `
            -ImageBytes $image.ImageBytes `
            -SpeedChoice $speedChoice `
            -VerifyEnabled $writeSettingsSelection.AutoVerify `
            -RecorderName $recorderName `
            -Firmware $firmware `
            -VolumeLabel $volumeLabel `
            -FileCount $writeFiles.Count `
            -Multisession $writeSettingsSelection.Multisession `
            -FinalizeRequested $writeSettingsSelection.Finalize `
            -SessionName $sessionName
        if (-not $preflight) {
	    Stop-WriteModeHeader
	    return [pscustomobject]@{
	        Status      = 'CANCELLED'
	        Write       = $writeResult
	        Verify      = $verifyResult
	        Log         = $logResult
	        Finalize    = $finalizeResult
	        Reason      = 'Yazma işlemi son kullanıcı onayında iptal edildi.'
	        SessionName = $sessionName
	        VolumeLabel = $volumeLabel
	    }
	}

        $writePlan = Confirm-WriteMediaBeforeBurn `
            -Recorder $recorder `
            -InitialMedia $media `
            -MediaCapabilities $mediaCapabilities `
            -WriteSelection $writeSettingsSelection `
            -Image $image

        $manifestRecheckAccepted = Confirm-ManifestRecheck `
            -WriteFiles $writeFiles `
            -MasterManifestIndex $masterManifestIndex `
            -MasterManifestIndexSha256 $masterManifestIndexSha256

        if (-not $manifestRecheckAccepted) {
            Stop-WriteModeHeader
            return New-WriteOperationResult `
                -Status 'CANCELLED' `
                -Write $writeResult `
                -Verify $verifyResult `
                -Log $logResult `
                -Finalize $finalizeResult `
                -Reason 'MANIFEST-RECHECK sonucunda yazma işlemi iptal edildi.' `
                -SessionName $sessionName `
                -VolumeLabel $volumeLabel
        }

        Clear-WriteModeBody

        $burnResult = Invoke-Burn `
            -Recorder $recorder `
            -ImageStream $image.BurnStream `
            -SelectedSpeed ([Int64]$speedChoice.SectorsPerSecond) `
            -FinalizeRequested $writePlan.Finalize `
            -WritePlan $writePlan

        $writeResult = [string]$burnResult.WriteResult
        $appliedSpeed = [double]$burnResult.AppliedSpeedX
        $finalizeResult = [string]$burnResult.FinalizeResult
        $failureStage = [string]$burnResult.FailureStage

        $verification = Invoke-WriteVerification `
            -BurnResult $burnResult `
            -WriteSelection $writeSettingsSelection `
            -Recorder $recorder `
            -SessionName $sessionName `
            -WriteFiles $writeFiles `
            -Manifest $manifest `
            -DirectoryPaths @(Get-ProjectDirectoryPath)

        $verifyResult = [string]$verification.Verify
        $reason = [string]$verification.Reason

        $overall = Get-WriteOverallResult `
            -WriteResult $writeResult `
            -FinalizeRequested ([bool]$writeSettingsSelection.Finalize) `
            -FinalizeResult $finalizeResult `
            -VerifyResult $verifyResult `
            -FailureStage $failureStage `
            -Reason $reason

        $overallResult = [string]$overall.Status
        $reason = [string]$overall.Reason
        $failureStage = [string]$overall.FailureStage

        try {
            $log = Write-WriteOperationLog `
                -Enabled ([bool]$writeSettingsSelection.Logging) `
                -Files $writeFiles `
                -SourceRoots $sourceRoots `
                -Media $media `
                -SourceBytes $sourceBytes `
                -ImageBytes ([Int64]$image.ImageBytes) `
                -SelectedSpeed $selectedSpeed `
                -AppliedSpeed $appliedSpeed `
                -WriteResult $writeResult `
                -OverallResult $overallResult `
                -Reason $reason `
                -RecorderName $recorderName `
                -Firmware $firmware `
                -VerifyResult $verifyResult `
                -FinalizeResult $finalizeResult `
             -FailureStage $failureStage `
                -VolumeLabel $volumeLabel `
                -Multisession ([bool]$writeSettingsSelection.Multisession) `
                -FinalizeRequested ([bool]$writeSettingsSelection.Finalize) `
                -SessionName $sessionName

            $logResult = [string]$log.Status
        }
        catch {
            $logResult = 'BAŞARISIZ'
        }

        Show-WriteCompletionSummary `
            -WriteResult $writeResult `
            -VerifyResult $verifyResult `
            -LogResult $logResult `
            -Log $log `
            -EjectAfter ([bool]$writeSettingsSelection.EjectAfter) `
            -Recorder $recorder `
            -OverallResult $overallResult `
            -SessionName $sessionName `
            -SourceBytes $sourceBytes `
            -ImageBytes ([Int64]$image.ImageBytes) `
            -FileCount $writeFiles.Count `
            -AppliedSpeed $appliedSpeed `
            -FinalizeResult $finalizeResult `
            -FinalizeRequested ([bool]$writeSettingsSelection.Finalize) `
             -FailureStage $failureStage `
            -Reason $reason `
            -VolumeLabel $volumeLabel

$finalStatus = if ($overallResult -eq 'BAŞARILI' -or $overallResult -eq 'BAŞARILI_DOĞRULANMADI') { 'SUCCESS' } else { 'FAILURE' }
        return New-WriteOperationResult -Status $finalStatus -Write $writeResult -Verify $verifyResult -Log $logResult -Finalize $finalizeResult -FailureStage $failureStage -Reason $reason -SessionName $sessionName `
                -VolumeLabel $volumeLabel
    }
    catch {
        return Invoke-WriteFailureHandler `
            -Exception $_.Exception `
            -Media $media `
            -Image $image `
            -Recorder $recorder `
            -WriteSettingsSelection $writeSettingsSelection `
            -WriteFiles $writeFiles `
            -SourceBytes $sourceBytes `
            -SelectedSpeed $selectedSpeed `
            -AppliedSpeed $appliedSpeed `
            -WriteResult $writeResult `
            -VerifyResult $verifyResult `
            -FinalizeResult $finalizeResult `
            -FailureStage $failureStage `
            -SessionName $sessionName `
                -VolumeLabel $volumeLabel
    }
    finally {
        Stop-WriteModeHeader
        Clear-WriteResource `
            -Image $image `
            -Manifest $manifest `
            -Recorders $recorders
    }
}
# ============================================================
# SOURCE / PROJECT KLAVYE ARAYÜZÜ
# ============================================================
$script:UiMode='SOURCE';$script:SourceIndex=0;$script:SourcePageStart=0;$script:SourceItems=@()
$script:SourceSizeCache=@{};$script:ProjectSizeCache=@{};$script:ProjectIndex=0;$script:ProjectPageStart=0;$script:ProjectItems=@();$script:UiMessage='';$script:UiMessageKind='INFO'
$script:CachedSelectedCount=$null
$script:StatusSummaryText=''
$script:UiProgress=[pscustomobject]@{Active=$false;Label='';Total=0;Current=0;NewAdded=0;DirCount=0;ShowNewAdded=$false;LastDraw=[datetime]::MinValue}

function Get-UiPageSize {
    [OutputType([int])]
    [CmdletBinding()]
    param()
    try {return [Math]::Max(5,[Console]::WindowHeight-14)} catch {return 12}
}
function Write-UiAt {
    <#
    .SYNOPSIS
    Belirli ekran satırına renkli metin yazdırır.
    .PARAMETER Row
    Hedef ekran satırı.
    .PARAMETER Text
    Yazılacak metin.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([int]$Row,[string]$Text,[ConsoleColor]$Color=$script:Theme.Color.Subtle)
    try {
        $width=[Console]::WindowWidth
        if ($width -le 1) {return}
        [Console]::SetCursorPosition(0,[Math]::Max(0,$Row))
        $line=if ($null -eq $Text) {''} else {[string]$Text}
        if ($line.Length -ge $width) {$line=$line.Substring(0,$width-1)}
        [Console]::ForegroundColor=$Color
        [Console]::Write($line.PadRight($width-1))
        [Console]::ForegroundColor=$script:Theme.Color.Subtle
    } catch {Write-Host $Text}
}
function Write-UiProgressLine {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if (-not $script:UiProgress.Active) { return }
    $total = [int]$script:UiProgress.Total
    $cur = [int]$script:UiProgress.Current
    $pct = 0
    if ($total -gt 0) { $pct = [int](($cur * 100) / $total) }
    if ($pct -lt 0) { $pct = 0 }
    if ($pct -gt 100) { $pct = 100 }
    $barWidth = 40
    $filled = [int](($barWidth * $pct) / 100)
    $bar = ($script:Theme.Symbol.BarFull * $filled) + ($script:Theme.Symbol.BarEmpty * ($barWidth - $filled))
    $extra = ''
    if ($script:UiProgress.ShowNewAdded -and $script:UiProgress.NewAdded -gt 0) {
        $extra = '  (+{0} yeni)' -f $script:UiProgress.NewAdded
    }
    $line = '{0} [{1}] {2,3}%  dosya: {3}/{4}  klasör: {5}{6}' -f $script:UiProgress.Label,$bar,$pct,$cur,$total,$script:UiProgress.DirCount,$extra
    $pageSize = Get-UiPageSize
    $statusRow = 8 + $pageSize
    Write-UiAt -Row ($statusRow + 2) -Text ('  ' + $line) -Color ($script:Theme.Color.Primary)
}
function Write-UiProgressText {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Text)
    $pageSize = Get-UiPageSize
    $statusRow = 8 + $pageSize
    Write-UiAt -Row ($statusRow + 2) -Text ('  ' + $Text) -Color ($script:Theme.Color.Primary)
}
function Start-UiProgress {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Label,[int]$Total=0,[bool]$ShowNewAdded=$false)
    $script:UiProgress.Active = $true
    $script:UiProgress.Label = $Label
    $script:UiProgress.Total = $Total
    $script:UiProgress.Current = 0
    $script:UiProgress.NewAdded = 0
    $script:UiProgress.DirCount = 0
    $script:UiProgress.ShowNewAdded = $ShowNewAdded
    $script:UiProgress.LastDraw = [datetime]::MinValue
    Write-UiProgressLine
}
function Update-UiProgress {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([int]$Increment=0,[int]$IncrementDir=0,[switch]$IsNew)
    if (-not $script:UiProgress.Active) { return }
    $script:UiProgress.Current += $Increment
    if ($IsNew) { $script:UiProgress.NewAdded += $Increment }
    $script:UiProgress.DirCount += $IncrementDir
    $now = Get-Date
    if (($script:UiProgress.Current % 100) -eq 0 -or (($now - $script:UiProgress.LastDraw).TotalMilliseconds -gt 200)) {
        $script:UiProgress.LastDraw = $now
        Write-UiProgressLine
    }
}
function Stop-UiProgress {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Message='')
    $script:UiProgress.Active = $false
    $script:UiProgress.Current = 0
    $script:UiProgress.Total = 0
    $script:UiProgress.NewAdded = 0
    $script:UiProgress.DirCount = 0
    $script:UiProgress.ShowNewAdded = $false
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        Set-UiMessage $Message 'INFO'
    }
}
function Measure-ProjectFileCount {
    [OutputType([int])]
    [CmdletBinding()]
    param([object]$Node=$script:ProjectRoot)
    [int]$count = 0
    if ($null -eq $Node) { return 0 }
    foreach ($child in @($Node.Children)) {
        if ($child.Type -eq 'FILE') { $count++ }
        elseif ($child.Type -eq 'DIR') { $count += (Measure-ProjectFileCount -Node $child) }
    }
    return $count
}
function Measure-SourceTreeFileCount {
    [CmdletBinding()]
    param([string]$RootPath)
    [int]$count = 0
    try {
        $root = Get-SourceItemOrThrow -LiteralPath $RootPath
        $stack = New-Object System.Collections.Stack
        $stack.Push($root)
        while ($stack.Count -gt 0) {
            $item = $stack.Pop()
            $rp = Get-ReparsePointInfo -Item $item
            if ($item.PSIsContainer) {
                if (-not $rp.CanTraverse) { continue }
                try {
                    $children = @(Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction Stop)
                    foreach ($c in $children) { $stack.Push($c) }
                } catch { continue }
            } elseif ($rp.CanSelect) { $count++ }
        }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    return $count
}
function Get-UiNamePart {
    [CmdletBinding()]
    param([string]$Name,[bool]$IsDirectory)
    if ($IsDirectory) {return [pscustomobject]@{Name=$Name;Extension='—'}}
    $ext=[IO.Path]::GetExtension($Name);$base=$Name
    if (-not [string]::IsNullOrWhiteSpace($ext)) {$base=$Name.Substring(0,$Name.Length-$ext.Length)}
    $extension='—'
    if (-not [string]::IsNullOrWhiteSpace($ext)) {$extension=$ext}
    return [pscustomobject]@{Name=$base;Extension=$extension}
}
function Get-ItemModeString {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Item)
    if ($null -eq $Item) { return '------' }
    $chars = @('-', '-', '-', '-', '-', '-')
    if ($Item.PSIsContainer) {
        $chars[0] = 'd'
    } else {
        $chars[1] = 'a'
    }
    try {
        $attrs = [int]$Item.Attributes
        if (($attrs -band [int][IO.FileAttributes]::ReadOnly) -ne 0) { $chars[2] = 'r' }
        if (($attrs -band [int][IO.FileAttributes]::Hidden)   -ne 0) { $chars[3] = 'h' }
        if (($attrs -band [int][IO.FileAttributes]::System)   -ne 0) { $chars[4] = 's' }
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    return (-join $chars)
}
function Get-ProjectNodeModeString {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node) { return '------' }
    switch ($Node.Type) {
        'DIR'  { return 'd-----' }
        'FILE' { return '-a----' }
        default { return '------' }
    }
}
function Format-DirDateTime {
    [CmdletBinding()]
    param([datetime]$Time, [bool]$IsUtc = $false)
    if ($null -eq $Time -or $Time -eq [datetime]::MinValue -or $Time.Year -lt 1900) {
        return '—'.PadRight(20)
    }
    $local = if ($IsUtc) { $Time.ToLocalTime() } else { $Time }
    return $local.ToString('dd.MM.yyyy     HH:mm', [Globalization.CultureInfo]::InvariantCulture)
}
function Get-SourceTreeNode {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([string]$RootPath)
    $result=New-Object System.Collections.ArrayList;$stack=New-Object System.Collections.Stack
    $root=Get-SourceItemOrThrow -LiteralPath $RootPath;$stack.Push($root)
    while ($stack.Count -gt 0) {
        $item=$stack.Pop();$rp=Get-ReparsePointInfo -Item $item
        [void]$result.Add([pscustomobject]@{Item=$item;Reparse=$rp})
        if (-not $item.PSIsContainer -or -not $rp.CanTraverse) {continue}
        try {$children=@(Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction Stop)} catch {continue}
        for ($i=$children.Count-1;$i-ge 0;$i--) {$stack.Push($children[$i])}
    }
    return @($result)
}
function Get-SourceFolderSize {
    [CmdletBinding()]
    param([string]$RootPath)
    $total=0L;$stack=New-Object System.Collections.Stack;$root=Get-SourceItemOrThrow -LiteralPath $RootPath
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $item=$stack.Pop();$rp=Get-ReparsePointInfo -Item $item
        if ($item.PSIsContainer) {
            if (-not $rp.CanTraverse) {continue}
            try {$children=@(Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction Stop)} catch {continue}
            foreach ($child in @($children)) {$stack.Push($child)}
        } elseif ($rp.CanSelect) {$total+=[Int64]$item.Length}
    }
    return $total
}
function Get-ProjectFolderSize {
    [CmdletBinding()]
    param([object]$Node)
    [Int64]$total=0
    foreach ($child in @($Node.Children)) {
        if ($child.Type -eq 'FILE') {$total+=[Int64]$child.Length}
        elseif ($child.Type -eq 'DIR') {$total+=Get-ProjectFolderSize $child}
    }
    return $total
}
# 3. SOURCE KATMANI
function Get-SourceUiItem {
    [OutputType([object[]])]
    [CmdletBinding()]
    param([string]$Path=(Get-Location).Path)

    try {
        $children=@(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop)
    }
    catch {
        throw ('SOURCE klasörü okunamadı: {0}' -f $_.Exception.Message)
    }

    $rows=New-Object System.Collections.ArrayList
    foreach ($item in @($children | Sort-Object @{Expression={if ($_.PSIsContainer) {0} else {1}}},Name)) {
        $rp=Get-ReparsePointInfo -Item $item;$parts=Get-UiNamePart -Name $item.Name -IsDirectory $item.PSIsContainer
        $size='—';$status='NORMAL'
        if ($item.PSIsContainer) {
            if ($script:SourceSizeCache.ContainsKey([string]$item.FullName)) {$size=Format-Size ([Int64]$script:SourceSizeCache[[string]$item.FullName])} else {$size='—'}
        } else {$size=Format-Size ([Int64]$item.Length)}
        if ($rp.IsReparse) {$status=$rp.Display}
        $itemType='FILE'
        if ($item.PSIsContainer) {$itemType='DIR'}
        [void]$rows.Add([pscustomobject]@{
            Item=$item;Type=$itemType;Reparse=$rp;Name=$parts.Name;Extension=$parts.Extension
            Size=$size;Status=$status;Selected=$false
            Mode=(Get-ItemModeString -Item $item)
            Tarih=(Format-DirDateTime -Time $item.LastWriteTime -IsUtc $false)
        })
    }
    return @($rows)
}
function Set-SourceLocationAndRefresh {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$SuccessMessage='Klasör açıldı.'
    )

    $oldPath=(Get-Location).Path
    $newPath=[IO.Path]::GetFullPath($TargetPath)

    try {
        # Önce hedef konumun gerçekten okunabildiğini doğrula.
        # Başarısız olursa mevcut konum ve SourceItems aynen korunur.
        $newItems=@(Get-SourceUiItem -Path $newPath)

        # Listeleme başarılı olduktan sonra konumu değiştir.
        Set-Location -LiteralPath $newPath -ErrorAction Stop

        $script:SourceItems=$newItems
        $script:SourceIndex=0
        $script:SourcePageStart=0
        Clear-SourceSizeCache
        $script:StatusSummaryText=''
        Set-UiMessage $SuccessMessage 'INFO'
        return $true
    }
    catch {
        $message=$_.Exception.Message

        # Hedef konuma geçilmiş olabilecek tek noktayı geri al.
        if (-not [string]::Equals((Get-Location).Path,$oldPath,[StringComparison]::OrdinalIgnoreCase)) {
            try {Set-Location -LiteralPath $oldPath -ErrorAction Stop} catch {
                Set-UiMessage ('Konum değişikliği başarısız oldu ve eski konuma dönülemedi: {0}' -f $_.Exception.Message) 'ERROR'
                return $false
            }
        }

        Set-UiMessage ('Konum değiştirilemedi: {0}' -f $message) 'ERROR'
        return $false
    }
}
function Get-ProjectUiItem {
    [OutputType([object[]])]
    [CmdletBinding()]
    param()
    $rows=New-Object System.Collections.ArrayList
    foreach ($node in @(Get-ProjectChild -Node $script:ProjectCurrent)) {
        $parts=Get-UiNamePart -Name $node.Name -IsDirectory ($node.Type -eq 'DIR');$size='—';$status='NORMAL'
        if ($node.Type -eq 'FILE') {$size=Format-Size ([Int64]$node.Length);if (-not $node.SourceExists) {$status='KAYNAK YOK'}}
        elseif ($script:ProjectSizeCache.ContainsKey($node)) {$size=Format-Size ([Int64]$script:ProjectSizeCache[$node])}
        [void]$rows.Add([pscustomobject]@{Node=$node;Type=$node.Type;Name=$parts.Name;Extension=$parts.Extension;Size=$size;Status=$status;Mode=(Get-ProjectNodeModeString -Node $node);Tarih=(Format-DirDateTime -Time $node.LastWriteUtc -IsUtc $true)})
    }
    return @($rows)
}
function Get-UiHeaderLine {
    [OutputType([object[]])]
    [CmdletBinding()]
    param()
    if ($script:UiMode -eq 'SOURCE') {return @('SOURCE',(Get-Location).Path)}
    $path=Get-ProjectRelativePath $script:ProjectCurrent
    $display='ARŞİV'
    if ($path -ne '\') {$display+='\' + $path.TrimStart('\')}
    return @('PROJECT',$display)
}



function Clear-SourceSizeCache {
    [CmdletBinding()]
    param()
    $script:SourceSizeCache=@{}
}

function Clear-ProjectSizeCache {
    [CmdletBinding()]
    param()
    $script:ProjectSizeCache=@{}
}





function Get-SourceSelectionMark {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Row)
    if ([bool]$Row.Selected) { return '[x]' }
    return '[ ]'
}
function Format-FixedWidth {
    [OutputType([string])]
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Text,
        [int]$Width,
        [switch]$RightAlign
    )
    if ($Width -le 0) { return '' }
    $s = if ($null -eq $Text) { '' } else { [string]$Text }
    if ($s.Length -gt $Width) {
        if ($Width -le 1) { $s = $s.Substring(0,1) }
        elseif ($Width -le 2) { $s = $s.Substring(0,$Width) }
        else { $s = $s.Substring(0,$Width-1) + [char]0x2026 }
    }
    if ($RightAlign) { return $s.PadLeft($Width) }
    return $s.PadRight($Width)
}
function Get-UiRowText {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]

    param(
        [object]$Row,
        [int]$No
    )

    $width=$script:UiListWidth
    try {
        $width=[Console]::WindowWidth
    }
    catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    $nameWidth=[Math]::Max($script:UiNameMinWidth,$width-$script:UiColumnReserved)

    $mark='[ ]'
    $icon='[F]'
    $status=[string]$Row.Status

    if ($script:UiMode -eq 'SOURCE') {

        $mark=Get-SourceSelectionMark -Row $Row

        if ($Row.Type -eq 'DIR') {
            $icon='[D]'
        }
        elseif ($Row.Reparse.Display -eq 'CLOUD') {
            $icon='[C]'
        }
        elseif ($Row.Reparse.IsReparse) {
            $icon='[L]'
        }
    }
    else {

        $mark = Get-ProjectSelectionMark -Node $Row.Node

        if ($Row.Type -eq 'DIR') {
            $icon='[D]'
        }
    }

    $nameStr  = Format-FixedWidth -Text $Row.Name      -Width $nameWidth
    $modeStr  = Format-FixedWidth -Text $Row.Mode      -Width 8
    $tarihStr = Format-FixedWidth -Text $Row.Tarih     -Width 20
    $extStr   = Format-FixedWidth -Text $Row.Extension -Width 7
    $sizeStr  = Format-FixedWidth -Text $Row.Size      -Width 14 -RightAlign
    $statStr  = Format-FixedWidth -Text $status        -Width 14

    return (
        ' {0,3} {1} {2} {3} {4} {5} {6} {7}' -f
        $No,
        $mark,
        $icon,
        $modeStr,
        $tarihStr,
        $nameStr,
        $extStr,
        $sizeStr
    ) + ' ' + $statStr
}
function Update-UiItem {
    [CmdletBinding()]
    param()
    try {
        if ($script:UiMode -eq 'SOURCE') {
            # Mevcut seçimleri FullName ile hatırla
            $previousSelection = @{}
            foreach ($old in @($script:SourceItems)) {
                if ([bool]$old.Selected) {
                    $previousSelection[[string]$old.Item.FullName] = $true
                }
            }

            $newItems = @(Get-SourceUiItem)

            # Yeni listede aynı FullName varsa seçimi geri yükle
            foreach ($row in $newItems) {
                if ($previousSelection.ContainsKey([string]$row.Item.FullName)) {
                    $row.Selected = $true
                }
            }

            $max = [Math]::Max(0, $newItems.Count - 1)
            $newIndex = [Math]::Min($script:SourceIndex, $max)
            $newPageStart = [Math]::Max(0, [Math]::Min($script:SourcePageStart, $max))
            $script:SourceItems = $newItems
            $script:SourceIndex = $newIndex
            $script:SourcePageStart = $newPageStart
        }
        else {
            $newItems = @(Get-ProjectUiItem)
            $max = [Math]::Max(0, $newItems.Count - 1)
            $newIndex = [Math]::Min($script:ProjectIndex, $max)
            $newPageStart = [Math]::Max(0, [Math]::Min($script:ProjectPageStart, $max))
            $script:ProjectItems = $newItems
            $script:ProjectIndex = $newIndex
            $script:ProjectPageStart = $newPageStart
        }
        return $true
    }
    catch {
        Set-UiMessage $_.Exception.Message 'ERROR'
        return $false
    }
}
function New-DrawUiContext {
    [CmdletBinding()]
    param()
    $width=$script:UiListWidth
    try {$width=[Console]::WindowWidth} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    $pageSize=Get-UiPageSize
    $header=Get-UiHeaderLine
    if ($script:UiMode -eq 'SOURCE') {$items=@($script:SourceItems);$index=$script:SourceIndex;$pageStart=$script:SourcePageStart}
    else {$items=@($script:ProjectItems);$index=$script:ProjectIndex;$pageStart=$script:ProjectPageStart}
    $total=$items.Count
    $pageCount=[Math]::Max(1,[int][Math]::Ceiling($total/[double]$pageSize))
    $page=[int]([Math]::Floor($pageStart/[double]$pageSize)+1)
    $borderColor = if ($script:UiMode -eq 'SOURCE') { $script:Theme.Color.SourceBorder } else { $script:Theme.Color.ProjectBorder }
    $nameWidth = [Math]::Max($script:UiNameMinWidth,$width-$script:UiColumnReserved)
    return [pscustomobject]@{
        Width=$width; PageSize=$pageSize; Header=$header; Items=$items; Index=$index
        PageStart=$pageStart; Total=$total; PageCount=$pageCount; Page=$page
        BorderColor=$borderColor; NameWidth=$nameWidth; UiMode=$script:UiMode
    }
}
function Write-UiHeader {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([object]$Context)
    Write-UiAt -Row 0 -Text '════════════════════════════════════════════════════════════════════════════════════════════════════════' -Color $Context.BorderColor
    Write-UiAt -Row 1 -Text '                     OPTİK MEDYA ARŞİV SİSTEMİ' -Color ($script:Theme.Color.Primary)
    Write-UiAt -Row 2 -Text ("                                      v{0}" -f $script:AppVersion) -Color ($script:Theme.Color.Info)
    Write-UiAt -Row 3 -Text '════════════════════════════════════════════════════════════════════════════════════════════════════════' -Color $Context.BorderColor
    if ($Context.UiMode -eq 'SOURCE') {
        Write-UiAt -Row 4 -Text ('  {0}' -f $Context.Header[0]) -Color ($script:Theme.Color.SourceBorder)
    } else {
        Write-UiAt -Row 4 -Text ('  {0}' -f $Context.Header[0]) -Color ($script:Theme.Color.Danger)
    }
    if (-not [string]::IsNullOrWhiteSpace($script:StatusSummaryText)) {
        $prefixLen = 2 + ([string]$Context.Header[0]).Length
        $info = '  :  ' + $script:StatusSummaryText
        try {
            $maxCol = [Math]::Max(0, [Console]::WindowWidth - 1)
            if ($prefixLen -lt $maxCol) {
                if (($prefixLen + $info.Length) -ge $maxCol) {
                    $info = $info.Substring(0, $maxCol - $prefixLen)
                }
                [Console]::SetCursorPosition($prefixLen, 4)
                [Console]::ForegroundColor = $script:Theme.Color.Muted
                [Console]::Write($info)
                [Console]::ForegroundColor = $script:Theme.Color.Subtle
            }
        } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
    Write-UiAt -Row 5 -Text ('  {0}' -f $Context.Header[1]) -Color ($script:Theme.Color.Subtle)
    $headerText =
        '  ' +
        'No'.PadLeft(3) + ' ' +
        'Sel'.PadRight(3) + ' ' +
        'Type'.PadRight(4) + ' ' +
        'Mode'.PadRight(8) + ' ' +
        'LastWriteTime'.PadRight(20) + ' ' +
        'Name'.PadRight($Context.NameWidth) + ' ' +
        'Ext'.PadRight(7) + ' ' +
        'Length'.PadLeft(14) + ' ' +
        'Status'.PadRight(14)
    Write-UiAt -Row 6 -Text $headerText -Color ($script:Theme.Color.Muted)
    Write-UiAt -Row 7 -Text '  ──────────────────────────────────────────────────────────────────────────────────────────────────────' -Color $Context.BorderColor
}
function Write-UiPageRow {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([object]$Context)
    for ($r=0;$r-lt $Context.PageSize;$r++) {
        $line=8+$r
        $pos=$Context.PageStart+$r
        if ($pos -lt $Context.Total) {
            $isFull = $false
            $isPartial = $false
            if ($Context.UiMode -eq 'SOURCE') {
                $isFull = [bool]$Context.Items[$pos].Selected
            } else {
                $projState = Get-ProjectSelectionState -Node $Context.Items[$pos].Node
                $isFull = ($projState -eq 'FULL')
                $isPartial = ($projState -eq 'PARTIAL')
            }
            $rowColor=$script:Theme.Color.Subtle
            if ($Context.UiMode -eq 'SOURCE' -and $Context.Items[$pos].Reparse.IsReparse) {
                $rowColor=$script:Theme.Color.Warning
            }
            $prefix=' '
            if ($pos -eq $Context.Index) {$prefix='>'}
            $rowText=$prefix+(Get-UiRowText -Row $Context.Items[$pos] -No ($pos+1))
            if ($pos -eq $Context.Index) {
                try {
                    $oldForeground=[Console]::ForegroundColor
                    $oldBackground=[Console]::BackgroundColor
                    if ($rowText.Length -ge $Context.Width) {$rowText=$rowText.Substring(0,[Math]::Max(0,$Context.Width-1))}
                    [Console]::SetCursorPosition(0,$line)
                    if ($isFull) {
                        [Console]::BackgroundColor=$script:Theme.Color.FullSelectBg
                        [Console]::ForegroundColor=$script:Theme.Color.FullSelectFg
                    } elseif ($isPartial) {
                        [Console]::BackgroundColor=$script:Theme.Color.PartialBg
                        [Console]::ForegroundColor=$script:Theme.Color.SelectionFg
                    } else {
                        [Console]::BackgroundColor=$script:Theme.Color.SelectionBg
                        [Console]::ForegroundColor=$script:Theme.Color.SelectionFg
                    }
                    [Console]::Write($rowText.PadRight([Math]::Max(0,$Context.Width-1)))
                    [Console]::BackgroundColor=$oldBackground
                    [Console]::ForegroundColor=$oldForeground
                } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            }
            elseif ($isFull) {Write-UiAt -Row $line -Text $rowText -Color ($script:Theme.Color.Danger)}
            elseif ($isPartial) {Write-UiAt -Row $line -Text $rowText -Color ($script:Theme.Color.Warning)}
            else {Write-UiAt -Row $line -Text $rowText -Color $rowColor}
        } else {Write-UiAt -Row $line -Text ''}
    }
}
function Write-UiStatusLine {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([object]$Context)
    $statusRow=8+$Context.PageSize
    Write-UiAt -Row $statusRow -Text '  ──────────────────────────────────────────────────────────────────────────────────────────────────────' -Color $Context.BorderColor
    if ($Context.UiMode -eq 'SOURCE') {
        $selectedCount = @($script:SourceItems | Where-Object {[bool]$_.Selected}).Count
    } else {
        if ($null -eq $script:CachedSelectedCount) {
            $script:CachedSelectedCount = @(Get-SelectedProjectNode).Count
        }
        $selectedCount = [int]$script:CachedSelectedCount
    }
    $selectionLabel = 'Seçili'
    Write-UiAt -Row ($statusRow+1) -Text ('  Sayfa {0}/{1}    {2}: {3}    Toplam: {4}' -f $Context.Page,$Context.PageCount,$selectionLabel,$selectedCount,$Context.Total)
    if (-not [string]::IsNullOrWhiteSpace($script:UiMessage)) {
        $msgColor=$script:Theme.Color.Success
        if ($script:UiMessageKind -eq 'ERROR') {$msgColor=$script:Theme.Color.Danger} elseif ($script:UiMessageKind -eq 'WARN') {$msgColor=$script:Theme.Color.Warning}
        Write-UiAt -Row ($statusRow+2) -Text ('  {0}' -f $script:UiMessage) -Color $msgColor
    } else {Write-UiAt -Row ($statusRow+2) -Text ''}
}
function Write-UiFooter {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([object]$Context)
    $statusRow=8+$Context.PageSize
    if ($Context.UiMode -eq 'SOURCE') {
        Write-UiAt -Row ($statusRow+3) -Text '  GEZİNME: ↑↓ seç  ENTER aç  BACKSPACE üst  PGUP/PGDN sayfa  HOME/END baş/son' -Color ($script:Theme.Color.Muted)
        Write-UiAt -Row ($statusRow+4) -Text '  İŞLEM: SPACE seç  F2 boyut  F3 sürücü  F4 yenile  F1 PROJECT''E EKLE  F6 PROJECT  F8 durum  F12 çıkış' -Color ($script:Theme.Color.Muted)
    } else {
        Write-UiAt -Row ($statusRow+3) -Text '  GEZİNME: ↑↓ seç  ENTER aç  BACKSPACE üst  PGUP/PGDN sayfa  HOME/END baş/son' -Color ($script:Theme.Color.Muted)
        Write-UiAt -Row ($statusRow+4) -Text '  İŞLEM: SPACE seç  F2 boyut  F3 yenile  F4 sil  F6 SOURCE  F8 durum  F9 yaz  F12 çıkış' -Color ($script:Theme.Color.Muted)
    }
    Write-UiAt -Row ($statusRow+5) -Text '════════════════════════════════════════════════════════════════════════════════════════════════════════' -Color $Context.BorderColor
}
function Write-UiScreen {
    <#
    .SYNOPSIS
    Ana ekran çizimini yapar.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    $context = New-DrawUiContext
    Write-UiHeader -Context $context
    Write-UiPageRow -Context $context
    Write-UiStatusLine -Context $context
    Write-UiFooter -Context $context
}
function Update-UiSelectionVisible {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:UiMode -eq 'SOURCE') {$index=$script:SourceIndex} else {$index=$script:ProjectIndex}
    $pageSize=Get-UiPageSize;$pageStart=[int]([Math]::Floor($index/[double]$pageSize)*$pageSize)
    if ($script:UiMode -eq 'SOURCE') {$script:SourcePageStart=[Math]::Max(0,$pageStart)} else {$script:ProjectPageStart=[Math]::Max(0,$pageStart)}
}
function Set-UiMessage {
    <#
    .SYNOPSIS
    Durum satırı mesajını ayarlar.
    .PARAMETER Text
    Mesaj metni.
    .PARAMETER Kind
    INFO, WARN veya ERROR.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param([string]$Text,[ValidateSet('INFO','WARN','ERROR')][string]$Kind='INFO')
    $script:UiMessage=$Text;$script:UiMessageKind=$Kind
}
function Invoke-UiFolderEnter {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:UiMode -eq 'SOURCE') {$items=@($script:SourceItems);$index=$script:SourceIndex} else {$items=@($script:ProjectItems);$index=$script:ProjectIndex}
    if ($index -ge $items.Count) {return}
    $row=$items[$index]

    if ($script:UiMode -eq 'SOURCE') {
        if (-not $row.Item.PSIsContainer) {Set-UiMessage 'Dosya açılamaz.' 'WARN';return}
        if (-not $row.Reparse.CanTraverse) {
            $detail=if ($row.Reparse.Reason) {': '+$row.Reparse.Reason} else {''}
            Set-UiMessage ('Bu klasöre girilemez: {0}{1}' -f $row.Reparse.Display,$detail) 'WARN'
            return
        }
        [void](Set-SourceLocationAndRefresh -TargetPath $row.Item.FullName -SuccessMessage 'Klasör açıldı.')
    } else {
        if ($row.Node.Type -ne 'DIR') {Set-UiMessage 'Dosya açılamaz.' 'WARN';return}
        $oldCurrent=$script:ProjectCurrent
        $script:ProjectCurrent=$row.Node
        $script:ProjectIndex=0
        $script:ProjectPageStart=0
        $script:CachedSelectedCount=$null
        if (Update-UiItem) {
            Set-UiMessage 'Klasör açıldı.' 'INFO'
            return
        }
        $script:ProjectCurrent=$oldCurrent
        $script:ProjectIndex=$index
        $script:ProjectPageStart=0
        Update-UiItem
    }
}
function Set-UiParentLocation {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:UiMode -eq 'SOURCE') {
        $path=(Get-Location).Path
        $root=[IO.Path]::GetPathRoot($path)

        if ([string]::Equals($path,$root,[StringComparison]::OrdinalIgnoreCase)) {
            Set-UiMessage 'SOURCE kök sürücüsündesiniz.' 'INFO'
            return
        }

        $childName=[IO.Path]::GetFileName($path.TrimEnd('\'))
        $parent=[IO.Path]::GetDirectoryName($path)

        if ([string]::IsNullOrWhiteSpace($parent)) {
            Set-UiMessage 'Üst klasör bulunamadı.' 'WARN'
            return
        }

        if (-not (Set-SourceLocationAndRefresh -TargetPath $parent -SuccessMessage 'Üst klasöre çıkıldı.')) {return}

        for ($i=0;$i-lt $script:SourceItems.Count;$i++) {
            if ([string]::Equals(
                [string]$script:SourceItems[$i].Item.Name,
                $childName,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                $script:SourceIndex=$i
                break
            }
        }
        Update-UiSelectionVisible
    } elseif ($script:ProjectCurrent.Type -ne 'ROOT') {
        $oldCurrent=$script:ProjectCurrent
        $childName=[string]$oldCurrent.Name
        $parent=$oldCurrent.Parent
        $script:ProjectCurrent=$parent
        $script:ProjectIndex=0
        $script:ProjectPageStart=0
        $script:CachedSelectedCount=$null

        if (-not (Update-UiItem)) {
            # PROJECT ağacı bellekte olduğundan normalde başarısız olmaz;
            # yine de başarısızlıkta eski güvenli konumu aynen geri yükle.
            $script:ProjectCurrent=$oldCurrent
            $script:ProjectIndex=0
            $script:ProjectPageStart=0
            Update-UiItem
            return
        }

        for ($i=0;$i-lt $script:ProjectItems.Count;$i++) {
            if ([string]::Equals(
                [string]$script:ProjectItems[$i].Node.Name,
                $childName,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                $script:ProjectIndex=$i
                break
            }
        }
        Update-UiSelectionVisible
        Set-UiMessage 'Üst klasöre çıkıldı.' 'INFO'
    } else {
        Set-UiMessage 'PROJECT kökündesiniz.' 'INFO'
    }
}
function Get-DriveList {
    [OutputType([object[]])]
    [CmdletBinding()]
    param()
    try {return @(Get-PSDrive -PSProvider FileSystem | Sort-Object Name | ForEach-Object {[pscustomobject]@{Name=$_.Name;Root=$_.Root}})}
    catch {return @()}
}
function Set-SourceDriveLocationSafely {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$TargetPath)

    $oldPath=(Get-Location).Path
    $oldSourceItems=@($script:SourceItems)
    $oldSourceIndex=[int]$script:SourceIndex
    $oldSourcePageStart=[int]$script:SourcePageStart

    try {
        Set-Location -LiteralPath $TargetPath -ErrorAction Stop
        $script:SourceIndex=0
        $script:SourcePageStart=0
        Clear-SourceSizeCache
        $script:StatusSummaryText=''

        if (-not (Update-UiItem)) {
            throw 'Yeni sürücünün SOURCE görünümü yenilenemedi.'
        }

        Set-UiMessage 'Sürücü değiştirildi.' 'INFO'
        return $true
    }
    catch {
        $failureMessage=$_.Exception.Message
        $restoreError=$null

        try {
            if (-not [string]::Equals((Get-Location).Path,$oldPath,[StringComparison]::OrdinalIgnoreCase)) {
                Set-Location -LiteralPath $oldPath -ErrorAction Stop
            }
        }
        catch {
            $restoreError=$_.Exception.Message
        }

        $script:SourceItems=$oldSourceItems
        $script:SourceIndex=$oldSourceIndex
        $script:SourcePageStart=$oldSourcePageStart

        if ($null -ne $restoreError) {
            Set-UiMessage ('Sürücü değiştirilemedi. Eski konuma geri dönülemedi: {0}' -f $restoreError) 'ERROR'
        }
        else {
            Set-UiMessage ('Sürücü değiştirilemedi: {0}' -f $failureMessage) 'ERROR'
        }

        return $false
    }
}

function Select-UiDrive {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    $drives=@(Get-DriveList)
    Clear-Host;Write-Host 'SÜRÜCÜ SEÇ' -ForegroundColor $script:Theme.Color.Primary;Write-Host ''
    if ($drives.Count -eq 0) {Write-Host 'Dosya sistemi sürücüsü bulunamadı.' -ForegroundColor $script:Theme.Color.Warning;[void][Console]::ReadKey($true);return}
    for ($i=0;$i-lt $drives.Count;$i++) {Write-Host ('{0,2}. {1,-4} {2}' -f ($i+1),$drives[$i].Name,$drives[$i].Root)}
    Write-Host '';Write-Host 'Q. Geri'
    while ($true) {
        $v=(Read-Host 'Seçiminiz').Trim().ToUpperInvariant()
        if ($v -eq 'Q') {break}
        $n=0
        if ([int]::TryParse($v,[ref]$n) -and $n -ge 1 -and $n -le $drives.Count) {
            if (Set-SourceDriveLocationSafely -TargetPath $drives[$n-1].Root) {
                break
            }
            Write-Host $script:UiMessage -ForegroundColor $script:Theme.Color.Danger
        } else {Write-Host 'Geçersiz seçim.' -ForegroundColor $script:Theme.Color.Warning}
    }
    Clear-Host
}




function Set-UiSourceSelectionToggle {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:SourceItems.Count -eq 0 -or $script:SourceIndex -ge $script:SourceItems.Count) {return}
    $row=$script:SourceItems[$script:SourceIndex]

    if (-not $row.Reparse.CanSelect) {
        Set-UiMessage ('Seçilemez: {0}' -f $row.Reparse.Display) 'WARN'
        return
    }

    $row.Selected = -not [bool]$row.Selected

    if ($row.Selected) {
        if ($row.Item.PSIsContainer) {
            Set-UiMessage 'Klasör seçildi (F1 ile tüm içeriğiyle aktarılır).' 'INFO'
        } else {
            Set-UiMessage 'Dosya seçildi (F1 ile aktarılır).' 'INFO'
        }
    } else {
        Set-UiMessage 'Seçim kaldırıldı.' 'INFO'
    }
}

function Measure-UiSelectedSize {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:UiMode -eq 'SOURCE') {$items=@($script:SourceItems);$index=$script:SourceIndex}
    else {$items=@($script:ProjectItems);$index=$script:ProjectIndex}
    if ($index -ge $items.Count) {return}
    $row=$items[$index]
    try {
        if ($script:UiMode -eq 'SOURCE') {
            if (-not $row.Item.PSIsContainer) {Set-UiMessage ('Boyut: {0}' -f (Format-Size ([Int64]$row.Item.Length))) 'INFO';return}
            Write-UiProgressText 'HESAPLANIYOR...'
            $size=Get-SourceFolderSize $row.Item.FullName;$script:SourceSizeCache[$row.Item.FullName]=$size
        } else {
            if ($row.Node.Type -ne 'DIR') {Set-UiMessage 'Dosya boyutu zaten listede.' 'INFO';return}
            Write-UiProgressText 'HESAPLANIYOR...'
            $size=Get-ProjectFolderSize $row.Node;$script:ProjectSizeCache[$row.Node]=$size
        }
        if (Update-UiItem) {Set-UiMessage ('Boyut: {0}' -f (Format-Size $size)) 'INFO'}
        else {Set-UiMessage ('Boyut hesaplandı ancak görünüm yenilenemedi: {0}' -f (Format-Size $size)) 'WARN'}
    } catch {Set-UiMessage ('Boyut hesaplanamadı: {0}' -f $_.Exception.Message) 'ERROR'}
}
# 4. PROJECT KATMANI

function Add-ProjectFromSourceSelection {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    $selectedRows = @($script:SourceItems | Where-Object {[bool]$_.Selected})

    if ($selectedRows.Count -eq 0) {
        Set-UiMessage 'Önce SOURCE tarafında seçim yapın.' 'WARN'
        return
    }

    $destination = $script:ProjectCurrent

    $conflicts = New-Object System.Collections.ArrayList
    foreach ($row in $selectedRows) {
        $existing = Find-ProjectChild -Parent $destination -Name $row.Item.Name
        if ($null -ne $existing) {
            [void]$conflicts.Add([pscustomobject]@{
                SourceName = [string]$row.Item.Name
                SourceKind = if ($row.Item.PSIsContainer) { 'Klasör' } else { 'Dosya' }
                ExistingKind = [string]$existing.Type
            })
        }
    }

    $mergeConflicts = $false
    if ($conflicts.Count -gt 0) {
        Clear-Host
        Write-Host ''
        Write-Host 'ÇAKIŞMA TESPİT EDİLDİ' -ForegroundColor $script:Theme.Color.Warning
        Write-Host ''
        Write-Host ('PROJECT içinde {0} öğe zaten mevcut:' -f $conflicts.Count) -ForegroundColor $script:Theme.Color.Warning
        Write-Host ''
        $i = 0
        foreach ($c in $conflicts) {
            $i++
            Write-Host ('  {0}. {1}  ({2})' -f $i, $c.SourceName, $c.SourceKind)
        }
        Write-Host ''
        Write-Host 'A = Çakışanları atla, eksikleri ekle (birleştir)'  -ForegroundColor $script:Theme.Color.Primary
        Write-Host 'H = Tüm işlemi iptal et'                            -ForegroundColor $script:Theme.Color.Primary
        Write-Host ''

        while ($true) {
            $answer = (Read-Host 'Seçiminiz').Trim().ToUpperInvariant()
            if ($answer -eq 'A') { $mergeConflicts = $true; break }
            if ($answer -eq 'H') {
                Set-UiMessage 'Ekleme kullanıcı tarafından iptal edildi.' 'INFO'
                return
            }
        }
    }

    $added = New-Object System.Collections.ArrayList
    $skipped = 0

    # On tarama: toplam dosya sayisini hizli hesapla
    [int]$totalFiles = 0
    Write-UiProgressText 'Seçimler taranıyor...'
    foreach ($row in $selectedRows) {
        if ($row.Item.PSIsContainer) {
            $totalFiles += (Measure-SourceTreeFileCount -RootPath $row.Item.FullName)
        } else {
            $totalFiles += 1
        }
    }
    Start-UiProgress -Label 'EKLENIYOR' -Total $totalFiles -ShowNewAdded $mergeConflicts

    try {
        foreach ($row in $selectedRows) {
            $existing = Find-ProjectChild -Parent $destination -Name $row.Item.Name
            $path = [string]$row.Item.FullName

            if ($null -ne $existing) {
                if ($mergeConflicts -and $row.Item.PSIsContainer -and $existing.Type -eq 'DIR') {
                    Add-MergedSourceDirectory -Parent $destination -SourcePath $path -SourceRoot $path -AddedNodes ([ref]$added) -SkippedCount ([ref]$skipped)
                } else {
                    $skipped++
                }
                continue
            }

            if ($row.Item.PSIsContainer) {
                Add-SourceDirectoryRecursive -Parent $destination -SourcePath $path -SourceRoot $path -AddedNodes ([ref]$added)
            } else {
                Add-ProjectSourceFile -Parent $destination -SourcePath $path -SourceRoot ([IO.Path]::GetDirectoryName($path)) -AddedNodes ([ref]$added)
            }
        }
    }
    catch {
        Stop-UiProgress
        for ($i = $added.Count - 1; $i -ge 0; $i--) {
            if ($null -ne $added[$i].Parent) { [void]$added[$i].Parent.Children.Remove($added[$i]) }
        }
        try { Update-NodeCountRecursive -Node $script:ProjectRoot; Update-CountsUpward $script:ProjectRoot } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        $script:CachedSelectedCount = $null
        Set-UiMessage ('Ekleme başarısız: {0}' -f $_.Exception.Message) 'ERROR'
        return
    }
    Stop-UiProgress

    try {
        $newProjectItems = @(Get-ProjectUiItem)
    }
    catch {
        for ($i = $added.Count - 1; $i -ge 0; $i--) {
            if ($null -ne $added[$i].Parent) { [void]$added[$i].Parent.Children.Remove($added[$i]) }
        }
        try { Update-NodeCountRecursive -Node $script:ProjectRoot; Update-CountsUpward $script:ProjectRoot } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
        $script:CachedSelectedCount = $null
        Clear-ProjectSizeCache
        Set-UiMessage ('PROJECT görünümü hazırlanamadı, ekleme geri alındı: {0}' -f $_.Exception.Message) 'ERROR'
        return
    }

    foreach ($row in @($script:SourceItems)) { $row.Selected = $false }

    Update-NodeCountRecursive -Node $destination
    Update-CountsUpward $destination
    $script:CachedSelectedCount = $null

    Clear-ProjectSizeCache
    $script:ProjectItems = $newProjectItems
    $script:UiMode = 'PROJECT'
    $script:ProjectIndex = 0
    $script:ProjectPageStart = 0
    $script:StatusSummaryText=''
    Clear-Host

    if ($skipped -gt 0) {
        Set-UiMessage ('PROJECT''e {0} öğe eklendi, {1} çakışan atlandı.' -f $added.Count, $skipped) 'INFO'
    } else {
        Set-UiMessage ('PROJECT''e {0} öğe eklendi.' -f $added.Count) 'INFO'
    }
    try { Write-UiScreen } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    if ($added.Count -ge 1000) {
        Start-Sleep -Milliseconds 1500
    }
}
function Set-ProjectSelectionRecursive {
    [CmdletBinding()]
    param([object]$Node,[bool]$Value)
    $Node.Selected=$Value
    if ($Node.Type -eq 'FILE') {
        $Node.TotalCount = 1
        $Node.SelectedCount = if ($Value) { 1 } else { 0 }
        return
    }
    $Node.TotalCount = 0
    $Node.SelectedCount = 0
    foreach ($child in @($Node.Children)) {
        Set-ProjectSelectionRecursive -Node $child -Value $Value
        $Node.TotalCount    += [int]$child.TotalCount
        $Node.SelectedCount += [int]$child.SelectedCount
    }
}

function Get-ProjectSelectionState {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Node)
    if ($null -eq $Node) { return 'NONE' }
    if ($Node.Type -eq 'FILE') {
        if ([bool]$Node.Selected) { return 'FULL' }
        return 'NONE'
    }
    $total = [int]$Node.TotalCount
    $selected = [int]$Node.SelectedCount
    if ($total -eq 0) {
        if ([bool]$Node.Selected) { return 'FULL' }
        return 'NONE'
    }
    if ($selected -eq 0) { return 'NONE' }
    if ($selected -eq $total) { return 'FULL' }
    return 'PARTIAL'
}

function Get-ProjectSelectionMark {
    [OutputType([string])]
    [CmdletBinding()]
    param([object]$Node)
    switch (Get-ProjectSelectionState -Node $Node) {
        'FULL'    { return '[x]' }
        'PARTIAL' { return '[~]' }
        default   { return '[ ]' }
    }
}

function Get-SelectedProjectNode {
    [OutputType([object[]])]
    [CmdletBinding()]
    param()
    $result = New-Object System.Collections.ArrayList
    if ($null -eq $script:ProjectCurrent) { return @() }

    $stack = New-Object System.Collections.Stack
    foreach ($child in @($script:ProjectCurrent.Children)) { $stack.Push($child) }

    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        $state = Get-ProjectSelectionState -Node $node
        if ($state -eq 'FULL') {
            [void]$result.Add($node)
            continue
        }
        if ($state -eq 'PARTIAL' -and $node.Type -eq 'DIR') {
            foreach ($c in @($node.Children)) { $stack.Push($c) }
        }
    }
    return @($result)
}
function Set-ProjectSelectionToggle {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if ($script:ProjectItems.Count -eq 0 -or $script:ProjectIndex -ge $script:ProjectItems.Count) {return}
    $node=$script:ProjectItems[$script:ProjectIndex].Node
    $state=Get-ProjectSelectionState -Node $node
    $newValue=($state -ne 'FULL')
    Set-ProjectSelectionRecursive -Node $node -Value $newValue
    Update-CountsUpward $node
    $script:CachedSelectedCount=$null
    if ($newValue) {Set-UiMessage 'Seçildi.' 'INFO'} else {Set-UiMessage 'Seçim kaldırıldı.' 'INFO'}
}
function Measure-ProjectDeletionScope {
    [CmdletBinding()]
    param([object[]]$SelectedNodes)
    $deleteFiles = 0
    $deleteDirs = 0
    foreach ($node in $SelectedNodes) {
        if ($node.Type -eq 'FILE') { $deleteFiles++; continue }
        $stack = New-Object System.Collections.Stack
        $stack.Push($node)
        while ($stack.Count -gt 0) {
            $n = $stack.Pop()
            if ($n.Type -eq 'FILE') { $deleteFiles++ } else { $deleteDirs++ }
            if ($n.Type -eq 'DIR') {
                foreach ($c in @($n.Children)) { $stack.Push($c) }
            }
        }
    }
    $keepFiles = 0
    $keepDirs = 0
    $allStack = New-Object System.Collections.Stack
    foreach ($child in @($script:ProjectRoot.Children)) { $allStack.Push($child) }
    while ($allStack.Count -gt 0) {
        $n = $allStack.Pop()
        $state = Get-ProjectSelectionState -Node $n
        if ($state -eq 'FULL') { continue }
        if ($n.Type -eq 'FILE') { $keepFiles++ } else { $keepDirs++ }
        if ($n.Type -eq 'DIR') {
            foreach ($c in @($n.Children)) { $allStack.Push($c) }
        }
    }
    $deleteParts = @()
    if ($deleteFiles -gt 0) { $deleteParts += ('{0} dosya' -f $deleteFiles) }
    if ($deleteDirs -gt 0)  { $deleteParts += ('{0} klasör' -f $deleteDirs) }
    $deleteText = if ($deleteParts.Count -gt 0) { $deleteParts -join ', ' } else { '0 öğe' }
    $keepParts = @()
    if ($keepFiles -gt 0) { $keepParts += ('{0} dosya' -f $keepFiles) }
    if ($keepDirs -gt 0)  { $keepParts += ('{0} klasör' -f $keepDirs) }
    $keepText = if ($keepParts.Count -gt 0) { $keepParts -join ', ' } else { '' }
    return [pscustomobject]@{ DeleteText=$deleteText; KeepText=$keepText }
}
function Confirm-ProjectDeletion {
    [CmdletBinding()]
    param([string]$DeleteText, [string]$KeepText)
    Clear-Host
    Write-Host ''
    if (-not [string]::IsNullOrWhiteSpace($KeepText)) {
        Write-Host ('{0} silinecek, {1} korunacak (işaretsiz).' -f $DeleteText, $KeepText) -ForegroundColor $script:Theme.Color.Warning
    } else {
        Write-Host ('{0} silinecek.' -f $DeleteText) -ForegroundColor $script:Theme.Color.Warning
    }
    Write-Host ''
    $answer = (Read-Host 'Onaylıyor musunuz? (E/H)').Trim().ToUpperInvariant()
    return ($answer -in @('E','EVET','Y'))
}
function Remove-ProjectNodesWithRollback {
    <#
    .SYNOPSIS
    Seçili düğümleri PROJECT ağacından çıkarır; hata olursa geri alır.
    .DESCRIPTION
    Her düğümün ebeveyni ve sırası kaydedilir. İşlem sırasında hata çıkarsa düğümler ters sırayla eski yerlerine konur, sayımlar yeniden hesaplanır ve kullanıcıya hata mesajı gösterilir. Sayım yenileme de başarısız olursa bu durum mesaja eklenir. Fonksiyon çıktı üretmez; sonucu arayüz mesajıyla bildirir.
    .PARAMETER SelectedNodes
    Çıkarılacak PROJECT düğümleri.
    .PARAMETER DeleteText
    Kullanıcıya gösterilecek silme kapsamı metni.
    .EXAMPLE
    PS> Remove-ProjectNodesWithRollback -SelectedNodes $nodes -DeleteText '3 öğe'
    #>
    [CmdletBinding()]
    param([object[]]$SelectedNodes, [string]$DeleteText)
    $removed = New-Object System.Collections.ArrayList
    $affectedParents = New-Object System.Collections.ArrayList
    try {
        foreach ($node in @($SelectedNodes)) {
            if ($null -ne $node.Parent) {
                $parent = $node.Parent
                $index = [Array]::IndexOf(@($parent.Children), $node)
                [void]$parent.Children.Remove($node)
                [void]$removed.Add([pscustomobject]@{Node=$node;Parent=$parent;Index=$index})
                if (-not $affectedParents.Contains($parent)) { [void]$affectedParents.Add($parent) }
            }
        }
        foreach ($p in $affectedParents) { Update-CountsUpward $p }
        $script:CachedSelectedCount = $null
        $script:StatusSummaryText=''
        Clear-ProjectSizeCache
        $script:ProjectItems = @(Get-ProjectUiItem)
        $script:ProjectIndex = 0
        $script:ProjectPageStart = 0
        Set-UiMessage ('{0} çıkarıldı.' -f $DeleteText) 'INFO'
        try { Write-UiScreen } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
    catch {
        # Asıl hata mesajı, aşağıdaki iç try/catch bloklarından etkilenmesin diye baştan saklanır.
        $deleteError = $_.Exception.Message
        for ($i = $removed.Count - 1; $i -ge 0; $i--) {
            $entry = $removed[$i]
            $children = @($entry.Parent.Children)
            if ($entry.Index -ge 0 -and $entry.Index -le $children.Count) {
                [void]$entry.Parent.Children.Insert($entry.Index, $entry.Node)
            } else {
                [void]$entry.Parent.Children.Add($entry.Node)
            }
        }
        $countRecoveryError = ''
        try { Update-NodeCountRecursive -Node $script:ProjectRoot } catch { $countRecoveryError = $_.Exception.Message; Write-Verbose ('Rollback sonrası sayımlar yenilenemedi: {0}' -f $countRecoveryError) }
        $script:CachedSelectedCount = $null
        Clear-ProjectSizeCache
        [void](Update-UiItem)
        $rollbackMessage = 'Silme işlemi geri alındı: {0}' -f $deleteError
        # Sayım yenileme de başarısız olduysa kullanıcıya açıkça söylenir; sessizce yutulmaz.
        if (-not [string]::IsNullOrWhiteSpace($countRecoveryError)) { $rollbackMessage += (' (Dikkat: sayımlar yenilenemedi: {0})' -f $countRecoveryError) }
        Set-UiMessage $rollbackMessage 'ERROR'
    }
}
function Remove-ProjectSelectedNode {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    $selectedNodes = @(Get-SelectedProjectNode)
    if ($selectedNodes.Count -eq 0) {
        Set-UiMessage 'Silmek için seçim yapın.' 'WARN'
        return
    }
    $scope = Measure-ProjectDeletionScope -SelectedNodes $selectedNodes
    if (-not (Confirm-ProjectDeletion -DeleteText $scope.DeleteText -KeepText $scope.KeepText)) {
        Set-UiMessage 'Silme iptal edildi.' 'INFO'
        return
    }
    Remove-ProjectNodesWithRollback -SelectedNodes $selectedNodes -DeleteText $scope.DeleteText
}
function Update-ProjectView {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    try {
        $totalForProgress = Measure-ProjectFileCount -Node $script:ProjectRoot
        Start-UiProgress -Label 'YENILENIYOR' -Total $totalForProgress
        try {
            Update-ProjectMetadataRecursive
        } finally {
            Stop-UiProgress
        }
        Update-NodeCountRecursive -Node $script:ProjectRoot
        $script:CachedSelectedCount=$null
        $script:StatusSummaryText=''
        Clear-ProjectSizeCache
        if (Update-UiItem) { Set-UiMessage 'PROJECT yenilendi.' 'INFO' }
    } catch {
        Set-UiMessage ('PROJECT yenilenemedi: {0}' -f $_.Exception.Message) 'ERROR'
    }
}
function Show-UiStatusSummary {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='TUI layer requires direct console control.')]
    param()
    if (-not [string]::IsNullOrWhiteSpace($script:StatusSummaryText)) {
        $script:StatusSummaryText = ''
        Set-UiMessage 'Durum özeti kapatıldı.' 'INFO'
        return
    }
    try {
        Write-UiProgressText 'DURUM HESAPLANIYOR...'
        if ($script:UiMode -eq 'SOURCE') {
            $stats = Get-SourceStatus
        }
        else {
            $stats = Get-ProjectStatistic
        }
        $text = 'Dosya: {0}  Klasör: {1}  Veri: {2}' -f `
            $stats.FileCount,
            $stats.DirectoryCount,
            (Format-Size $stats.TotalBytes)
        if ($script:UiMode -eq 'PROJECT' -and $stats.MissingCount -gt 0) {
            $text += '  Eksik: {0}' -f $stats.MissingCount
        }
        $script:StatusSummaryText = $text
        Set-UiMessage 'Durum özeti açıldı (F8 ile kapat).' 'INFO'
    }
    catch {
        $script:StatusSummaryText = ''
        Set-UiMessage ('Durum bilgisi alınamadı: {0}' -f $_.Exception.Message) 'ERROR'
    }
}
function Get-SourceStatus {
    [CmdletBinding()]
    param()
    $nodes=@(Get-SourceTreeNode -RootPath (Get-Location).Path);$files=@($nodes|Where-Object {-not $_.Item.PSIsContainer -and $_.Reparse.CanSelect});$dirs=@($nodes|Where-Object {$_.Item.PSIsContainer -and $_.Reparse.CanTraverse});[Int64]$bytes=0
    foreach ($n in @($files)) {$bytes+=[Int64]$n.Item.Length}
    return [pscustomobject]@{FileCount=$files.Count;DirectoryCount=$dirs.Count;TotalBytes=$bytes;MissingCount=0}
}
function Switch-UiModeSafely {
    [CmdletBinding()]
    param([ValidateSet('SOURCE','PROJECT')][string]$TargetMode)

    if ($script:UiMode -eq $TargetMode) {return $true}

    try {
        if ($TargetMode -eq 'SOURCE') {
            $newItems=@(Get-SourceUiItem -Path (Get-Location).Path)
            $script:SourceItems=$newItems
            $script:SourceIndex=0
            $script:SourcePageStart=0
            Clear-SourceSizeCache
        } else {
            $newItems=@(Get-ProjectUiItem)
            $script:ProjectItems=$newItems
            $script:ProjectIndex=0
            $script:ProjectPageStart=0
            Clear-ProjectSizeCache
        }

        $script:UiMode=$TargetMode
        $script:CachedSelectedCount=$null
        $script:StatusSummaryText=''
        Clear-Host
        Set-UiMessage ('{0} moduna geçildi.' -f $TargetMode) 'INFO'
        return $true
    }
    catch {
        Set-UiMessage ('{0} moduna geçilemedi: {1}' -f $TargetMode,$_.Exception.Message) 'ERROR'
        return $false
    }
}

function Invoke-UiKey {
    [OutputType([string])]
    [CmdletBinding()]
    param([ConsoleKeyInfo]$Key)
    if ($Key.Key -eq [ConsoleKey]::F12) {
        return 'EXIT'
    }
    if ($Key.Key -eq [ConsoleKey]::Escape) {
        return 'CONTINUE'
    }
    if ($script:UiMode -eq 'SOURCE') {$items=@($script:SourceItems);$count=$items.Count;$index=$script:SourceIndex}
    else {$items=@($script:ProjectItems);$count=$items.Count;$index=$script:ProjectIndex}
    $moveKey=$false
    switch ($Key.Key) {
	UpArrow {if ($count -gt 0) {$index=[Math]::Max(0,$index-1);$moveKey=$true}}
	DownArrow {if ($count -gt 0) {$index=[Math]::Min($count-1,$index+1);$moveKey=$true}}
	PageUp {if ($count -gt 0) {$index=[Math]::Max(0,$index-(Get-UiPageSize));$moveKey=$true}}
	PageDown {if ($count -gt 0) {$index=[Math]::Min($count-1,$index+(Get-UiPageSize));$moveKey=$true}}
        Home {if ($count -gt 0) {$index=0;$moveKey=$true}}
        End {if ($count -gt 0) {$index=$count-1;$moveKey=$true}}
        Enter {Invoke-UiFolderEnter}
        Backspace {Set-UiParentLocation}
        Spacebar {if ($script:UiMode -eq 'SOURCE') {Set-UiSourceSelectionToggle} else {Set-ProjectSelectionToggle}}
	F1 {
	    if ($script:UiMode -eq 'SOURCE') {
	        Add-ProjectFromSourceSelection
	    }
	}
        F2 {Measure-UiSelectedSize}
	F3 {
	    if ($script:UiMode -eq 'SOURCE') {
	        Select-UiDrive
	    }
	    else {
	        Update-ProjectView
	    }
	}
        F4 {
            if ($script:UiMode -eq 'SOURCE') {
                if (Update-UiItem) {Set-UiMessage 'SOURCE yenilendi.' 'INFO'}
            }
            else {Remove-ProjectSelectedNode}
        }
        F6 {
            if ($script:UiMode -eq 'SOURCE') {[void](Switch-UiModeSafely -TargetMode 'PROJECT')}
            else {[void](Switch-UiModeSafely -TargetMode 'SOURCE')}
        }
        F8 {Show-UiStatusSummary}
      F9 {
        if ($script:UiMode -ne 'PROJECT') { return 'CONTINUE' }

        $result = Invoke-WriteFromProject
        [void](Update-UiItem)
        Clear-Host

	    if ($null -ne $result -and $result.PSObject.Properties['Status']) {
        	switch ([string]$result.Status) {
	            'SUCCESS' {
	                if ([string]$result.Verify -eq 'BAŞARILI') {
	                    Set-UiMessage 'Yazma ve doğrulama tamamlandı.' 'INFO'
	                }
	                elseif ([string]$result.Verify -eq 'İSTENMEDİ') {
	                    Set-UiMessage 'Yazma işlemi başarıyla tamamlandı. VERIFY istenmedi.' 'INFO'
	                }
	                else {
	                    Set-UiMessage ('Yazma işlemi başarıyla tamamlandı. VERIFY durumu: {0}.' -f [string]$result.Verify) 'INFO'
	                }
	            }

	            'FAILURE' {
	                $msg=if ([string]$result.Write -eq 'YAPILMADI') {'Yazma başlatılmadı.'} else {'Yazma işlemi tamamlandı ancak sonuç BAŞARISIZ.'}
	                if (-not [string]::IsNullOrWhiteSpace([string]$result.Reason)) {
	                    $msg+=' ' + [string]$result.Reason
	                }
	                Set-UiMessage $msg 'ERROR'
	            }

	            'CANCELLED' {
	                $msg='Yazma işlemi iptal edildi.'
	                if (-not [string]::IsNullOrWhiteSpace([string]$result.Reason)) {
	                    $msg+=' ' + [string]$result.Reason
	                }
	                Set-UiMessage $msg 'WARN'
	            }

	            default {
	                Set-UiMessage 'Yazma sonucu tanınamadı.' 'ERROR'
	            }
	        }
	    }
	    else {
	        Set-UiMessage 'Yazma işlemi başlatılmadı veya iptal edildi.' 'WARN'
	    }
	}
        default {}
    }
    if ($moveKey) {
        if ($script:UiMode -eq 'SOURCE') {$script:SourceIndex=$index} else {$script:ProjectIndex=$index}
        Update-UiSelectionVisible;Set-UiMessage '' 'INFO'
    }
    return 'CONTINUE'
}
function Start-ArchiveShell {
    [CmdletBinding()]
    param()
    Initialize-Project
    $script:UiMode='SOURCE'
    $script:SourceSizeCache=@{}
    $script:ProjectSizeCache=@{}
    $script:SourceIndex=0
    $script:SourcePageStart=0
    $script:ProjectIndex=0
    $script:ProjectPageStart=0
    if (-not (Update-UiItem)) {throw 'Başlangıç SOURCE görünümü hazırlanamadı.'}
    Clear-Host
    try {[Console]::CursorVisible=$false} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    while ($true) {
        try { Write-UiScreen } catch { Write-UiErrorLog -Err $_ -Where 'Write-UiScreen' }
        $key = $null
        try {
            # Buffer'da kalan tuş event'lerini temizle (Wait-ConsoleInput'dan arta kalan)
            while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) }
            $key = [Console]::ReadKey($true)
        } catch {
            Start-Sleep -Milliseconds 100
            continue
        }
        $result = 'CONTINUE'
        try {
            $result = [string]@(Invoke-UiKey -Key $key)[-1]
        } catch {
            Write-UiErrorLog -Err $_ -Where 'Invoke-UiKey'
            try {
                Set-UiMessage ('Beklenmeyen hata: {0}' -f $_.Exception.Message) 'ERROR'
            } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
            $result = 'CONTINUE'
        }
        if ($result -ceq 'EXIT') { break }
        Update-UiSelectionVisible
    }
    try {[Console]::CursorVisible=$true} catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    Clear-Host
}
# ============================================================
# BAŞLAT
# ============================================================
try {
    # UTF-8 konsol encoding: chcp 857'de bulunmayan Turkce karakterler
    # (I, S, G gibi) Console bell'e sebep oluyordu.
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    try { chcp 65001 *> $null } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    Initialize-AppDirectory
    Clear-StaleTempArtifact
    Start-ArchiveShell
}
catch {
    try { Clear-Host } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    Write-Host 'OPTİK MEDYA ARŞİV SİSTEMİ başlatılamadı.' -ForegroundColor $script:Theme.Color.Danger
    Write-Host ''

    Write-Host ('Hata       : {0}' -f $_.Exception.Message) -ForegroundColor $script:Theme.Color.Danger
    $commandName = ''
    try {
        if ($null -ne $_.InvocationInfo.MyCommand) {
            $commandName = $_.InvocationInfo.MyCommand.ToString()
        }
    }
    catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    Write-Host ('Komut      : {0}' -f $commandName) -ForegroundColor $script:Theme.Color.Warning
    Write-Host ('Satır      : {0}' -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor $script:Theme.Color.Warning
    Write-Host ('Dosya      : {0}' -f $_.InvocationInfo.ScriptName) -ForegroundColor $script:Theme.Color.Warning
    Write-Host ''
    Write-Host 'Konum:' -ForegroundColor $script:Theme.Color.Primary
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor $script:Theme.Color.Subtle
    Write-Host ''
    Write-Host 'Stack:' -ForegroundColor $script:Theme.Color.Primary
    Write-Host $_.ScriptStackTrace -ForegroundColor $script:Theme.Color.Muted
    Write-Host ''

    try {
        Write-Host ('HResult: 0x{0:X8}' -f (([int64]$_.Exception.HResult) -band 0xFFFFFFFFL)) -ForegroundColor $script:Theme.Color.Primary
    } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }

    Wait-ConsoleInput
}
finally {
    try { Set-StrictMode -Off } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    if ($null -ne $script:SavedErrorActionPreference) {
        try { $ErrorActionPreference = $script:SavedErrorActionPreference } catch { Write-Verbose ('Sessiz hata: {0}' -f $_.Exception.Message) }
    }
}
