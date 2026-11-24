#requires -Version 5.1

<#!
    GİB Berat Görüntüleme ve XML başlık bilgisi
    Menüden çağrılan tek dosyalı PowerShell launcher.

    Menü:
    1. Berat Görüntüleme Terminal : XML başlık bilgilerini terminalde gösterir.
    2. Berat Görüntüleme Web      : Berat'ı berat.xslt ile varsayılan tarayıcıda gösterir.

    Web çalışma şekli:
    1. HTML arayüzünü geçici klasöre oluşturur.
    2. Varsayılan tarayıcıda açar.
    3. Kullanıcı klasörü seçer; sadece GIB*.xml dosyaları listelenir.
    4. Aynı klasör ağacında bulunan berat.xslt kullanılarak berat görüntülenir.
#>

# Neden: Bu betik etkileşimli bir konsol arayüzüdür; renkli çıktı Windows PowerShell 5.1'de
# yalnızca Write-Host ile verilebilir. Write-Host kullanımı tek yerde (Write-ConsoleText) toplanmıştır.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    'PSAvoidUsingWriteHost', '',
    Justification = 'Etkileşimli konsol arayüzü; renkli çıktı için gereklidir.')]
param()

# M-Files Toolkit ortak kullanici verisi yolu altyapisi
if (-not (Get-Command Get-MFTModuleDataPath -ErrorAction SilentlyContinue)) {
    $mftRootProbe = Get-Item -LiteralPath $PSScriptRoot
    while ($null -ne $mftRootProbe -and -not (Test-Path -LiteralPath (Join-Path $mftRootProbe.FullName 'lib\data.ps1') -PathType Leaf)) { $mftRootProbe = $mftRootProbe.Parent }
    if ($null -eq $mftRootProbe) { throw 'M-Files Toolkit lib\data.ps1 bulunamadi.' }
    . (Join-Path $mftRootProbe.FullName 'lib\data.ps1')
}



Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ModuleConfigPath = Join-Path (Get-MFTModuleDataPath -ModuleId 'gibberatoku' -Category Config) 'gibberatoku_config.json'

# --- Terminal rapor düzeni -------------------------------------------------------------
$script:ReportWidth = 72                                  # Ayraç ve imza satırı genişliği
$script:LabelWidth = 24                                   # En uzun etiket 22 karakter (YEVMİYE MADDESİ SAYISI)
$script:ValueColumn = 2 + $script:LabelWidth + 3          # "  " + etiket + " : " sonrası değer sütunu
$script:TurkishCulture = [Globalization.CultureInfo]::GetCultureInfo('tr-TR')

# Renkler tek yerde tutulur; görünümü değiştirmek için yalnızca burası düzenlenir.
$script:Palette = [pscustomobject]@{
    Frame     = [ConsoleColor]::Magenta     # Dosya adı çerçevesi (menü başlığıyla aynı)
    Section   = [ConsoleColor]::Cyan        # Bölüm başlıkları
    Rule      = [ConsoleColor]::DarkGray    # Ayraç çizgileri ve iki nokta
    Label     = [ConsoleColor]::Gray        # Etiketler
    Value     = [ConsoleColor]::White       # Veri değerleri
    Key       = [ConsoleColor]::Yellow      # Öne çıkan bilgiler (VKN, unvan, tip, dönem)
    Empty     = [ConsoleColor]::DarkGray    # Boş değer
    Missing   = [ConsoleColor]::DarkYellow  # XML içinde bulunamayan bilgi
    Secondary = [ConsoleColor]::DarkGray    # Uzun imza değerleri
    Total     = [ConsoleColor]::Green       # Toplam satırı
    Warning   = [ConsoleColor]::Yellow
    Failure   = [ConsoleColor]::Red
    Success   = [ConsoleColor]::Green
}

function Write-ConsoleText {
    <#
    .SYNOPSIS
        Renkli konsol çıktısının tek noktası.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [Parameter()][ConsoleColor]$Color = [ConsoleColor]::Gray,
        [Parameter()][switch]$NoNewline
    )

    Write-Host -Object $Text -ForegroundColor $Color -NoNewline:$NoNewline
}

function Write-ModuleStatus {
    <#
    .SYNOPSIS
        Çerçeveli durum mesajı yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter()][ConsoleColor]$Color = [ConsoleColor]::Cyan
    )

    Write-ConsoleText -Text ''
    Write-ConsoleText -Text ('-' * 72) -Color DarkGray
    Write-ConsoleText -Text "  $Message" -Color $Color
    Write-ConsoleText -Text ('-' * 72) -Color DarkGray
    Write-ConsoleText -Text ''
}

function Split-TextByWidth {
    <#
    .SYNOPSIS
        Metni verilen genişliğe göre satırlara böler.
    .DESCRIPTION
        Boşluklu metinlerde sözcük sınırından, boşluksuz uzun metinlerde (Base64 imza gibi)
        zorunlu olarak genişlikten keser.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory = $true)][ValidateRange(1, 500)][int]$Width
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $current = ''
    foreach ($token in ($Text -split '\s+')) {
        $word = $token
        while ($word.Length -gt $Width) {
            if ($current.Length -gt 0) { $lines.Add($current); $current = '' }
            $lines.Add($word.Substring(0, $Width))
            $word = $word.Substring($Width)
        }
        if ($current.Length -eq 0) {
            $current = $word
        } elseif (($current.Length + 1 + $word.Length) -le $Width) {
            $current = "$current $word"
        } else {
            $lines.Add($current)
            $current = $word
        }
    }
    if ($current.Length -gt 0) { $lines.Add($current) }
    return $lines
}

function Write-ReportRule {
    <#
    .SYNOPSIS
        Rapor genişliğinde ayraç çizgisi yazar.
    #>
    [CmdletBinding()]
    param([Parameter()][ConsoleColor]$Color = $script:Palette.Rule)

    Write-ConsoleText -Text ('  ' + ('-' * $script:ReportWidth)) -Color $Color
}

function Write-ReportFileHeader {
    <#
    .SYNOPSIS
        Her berat dosyasının başına çerçeveli dosya adı yazar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$FileName)

    Write-ConsoleText -Text ''
    Write-ReportRule -Color $script:Palette.Frame
    Write-ConsoleText -Text "  $FileName" -Color $script:Palette.Value
    Write-ReportRule -Color $script:Palette.Frame
}

function Write-ReportSection {
    <#
    .SYNOPSIS
        Bölüm başlığını ve altına ayraç çizgisini yazar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Title)

    Write-ConsoleText -Text ''
    Write-ConsoleText -Text "  $Title" -Color $script:Palette.Section
    Write-ReportRule
}

function Write-ReportField {
    <#
    .SYNOPSIS
        "ETİKET          : değer" biçiminde, iki noktaları aynı sütunda hizalı satır yazar.
    .DESCRIPTION
        Boş değer EmptyText ile gösterilir. Değer satıra sığmazsa değer sütununa
        hizalı (asılı girintili) şekilde alt satıra sarılır.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
        [Parameter()][switch]$Highlight,
        [Parameter()][string]$EmptyText = '-',
        [Parameter()][ConsoleColor]$EmptyColor = $script:Palette.Empty
    )

    Write-ConsoleText -Text ('  ' + $Label.PadRight($script:LabelWidth) + ' ') -Color $script:Palette.Label -NoNewline
    Write-ConsoleText -Text ': ' -Color $script:Palette.Rule -NoNewline

    $text = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        Write-ConsoleText -Text $EmptyText -Color $EmptyColor
        return
    }

    $valueColor = if ($Highlight) { $script:Palette.Key } else { $script:Palette.Value }
    $availableWidth = $script:ReportWidth + 2 - $script:ValueColumn
    $lines = @(Split-TextByWidth -Text $text -Width $availableWidth)
    $indent = ' ' * $script:ValueColumn

    Write-ConsoleText -Text $lines[0] -Color $valueColor
    for ($index = 1; $index -lt $lines.Count; $index++) {
        Write-ConsoleText -Text ($indent + $lines[$index]) -Color $valueColor
    }
}

function Write-ReportBlock {
    <#
    .SYNOPSIS
        Uzun, boşluksuz değeri (imza gibi) rapor genişliğinde satırlara bölerek yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter()][ConsoleColor]$Color = $script:Palette.Secondary
    )

    foreach ($line in (Split-TextByWidth -Text $Text -Width $script:ReportWidth)) {
        Write-ConsoleText -Text ('  ' + $line) -Color $Color
    }
}

function Get-NodeText {
    <#
    .SYNOPSIS
        Düğümün kırpılmış metnini döndürür; düğüm yoksa boş metin döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter()][AllowNull()][System.Xml.XmlNode]$Node)

    if ($null -eq $Node) { return '' }
    return $Node.InnerText.Trim()
}

function Get-XmlText {
    <#
    .SYNOPSIS
        Düğümün altında XPath ile bulunan ilk öğenin kırpılmış metnini döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()][AllowNull()][System.Xml.XmlNode]$Node,
        [Parameter(Mandatory = $true)][string]$Path
    )

    if ($null -eq $Node) { return '' }
    return Get-NodeText -Node $Node.SelectSingleNode($Path)
}

function Format-XmlDate {
    <#
    .SYNOPSIS
        yyyy-MM-dd tarihini "dd / MM / yyyy" biçimine çevirir.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter()][AllowEmptyString()][string]$Value)

    if (-not $Value) { return '' }
    try {
        $date = [datetime]::ParseExact($Value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        return $date.ToString('dd / MM / yyyy')
    } catch [System.FormatException] {
        # Neden: Beklenmeyen tarih biçimi raporu durdurmamalı; ham değer olduğu gibi gösterilir.
        return $Value
    }
}

function Format-XmlDateRange {
    <#
    .SYNOPSIS
        İki tarihi "başlangıç - bitiş" olarak birleştirir; ikisi de boşsa boş döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()][AllowEmptyString()][string]$Start,
        [Parameter()][AllowEmptyString()][string]$End
    )

    if (-not $Start -and -not $End) { return '' }
    return '{0} - {1}' -f (Format-XmlDate -Value $Start), (Format-XmlDate -Value $End)
}

function ConvertTo-DecimalAmount {
    <#
    .SYNOPSIS
        XML'deki tutarı (1234.56) decimal'e çevirir; çevrilemezse 0 döndürür.
    #>
    [CmdletBinding()]
    [OutputType([decimal])]
    param([Parameter()][AllowEmptyString()][string]$Value)

    $parsed = [decimal]0
    $styles = [Globalization.NumberStyles]::Number
    if ([decimal]::TryParse($Value, $styles, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return [decimal]0
}

function Format-Amount {
    <#
    .SYNOPSIS
        Tutarı Türkçe biçimde (1.234.567,89) yazar.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][decimal]$Amount)

    return $Amount.ToString('N2', $script:TurkishCulture)
}

function Write-XmlDocumentSignatureValue {
    <#
    .SYNOPSIS
        Berat'a konu olan dokümanın imza değerini yazar.
    #>
    [CmdletBinding()]
    param([Parameter()][AllowNull()][System.Xml.XmlNode]$Node)

    Write-ReportSection -Title "BERAT'A KONU OLAN DOKÜMANIN İMZA DEĞERİ"

    # Neden: Base64 değer XML'de satır sonlarıyla bölünmüş olabilir; boşluklar temizlenip yeniden bölünür.
    $value = (Get-NodeText -Node $Node) -replace '\s', ''
    if ([string]::IsNullOrWhiteSpace($value)) {
        Write-ConsoleText -Text '  İmza değeri XML içinde bulunamadı.' -Color $script:Palette.Missing
        return
    }
    Write-ReportBlock -Text $value
}

function Write-GibApprovalInfo {
    <#
    .SYNOPSIS
        GİB onay (karşı imza) bilgilerini yazar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document)

    Write-ReportSection -Title 'GİB ONAY BİLGİLERİ'

    $gibSignature = $Document.DocumentElement.SelectSingleNode(
        ".//*[local-name()='CounterSignature']/*[local-name()='Signature']"
    )
    if ($null -eq $gibSignature) {
        Write-ConsoleText -Text '  GİB onay imzası XML içinde bulunamadı.' -Color $script:Palette.Missing
        return
    }

    $fields = @(
        [pscustomobject]@{ Label = 'Onaylayan'; Path = ".//*[local-name()='X509SubjectName']" }
        [pscustomobject]@{ Label = 'Onay imza zamanı'; Path = ".//*[local-name()='SigningTime']" }
        [pscustomobject]@{ Label = 'Sertifika düzenleyeni'; Path = ".//*[local-name()='IssuerSerial']/*[local-name()='X509IssuerName']" }
        [pscustomobject]@{ Label = 'Sertifika seri numarası'; Path = ".//*[local-name()='IssuerSerial']/*[local-name()='X509SerialNumber']" }
    )
    foreach ($field in $fields) {
        $fieldParams = @{
            Label     = $field.Label
            Value     = Get-XmlText -Node $gibSignature -Path $field.Path
            EmptyText = 'XML içinde bulunamadı.'
            EmptyColor = $script:Palette.Missing
        }
        Write-ReportField @fieldParams
    }

    $signatureValue = (Get-XmlText -Node $gibSignature -Path "./*[local-name()='SignatureValue']") -replace '\s', ''
    Write-ConsoleText -Text ''
    Write-ConsoleText -Text '  GİB onay imza değeri' -Color $script:Palette.Label
    if ([string]::IsNullOrWhiteSpace($signatureValue)) {
        Write-ConsoleText -Text '  XML içinde bulunamadı.' -Color $script:Palette.Missing
        return
    }
    Write-ReportBlock -Text $signatureValue
}

function Write-TaxDetailTable {
    <#
    .SYNOPSIS
        Vergi detayı tablosunu (hesap kodu, adı, borç, alacak) ve toplam satırını yazar.
    #>
    [CmdletBinding()]
    param([Parameter()][AllowNull()][System.Xml.XmlNode]$Accounting)

    if ($null -eq $Accounting) { return }

    $debitXPath = ".//*[local-name()='entryDetail'][*[local-name()='xbrlInfo']/*[local-name()='xbrlInclude']='period_change' and *[local-name()='debitCreditCode']='D']"
    $details = @($Accounting.SelectNodes($debitXPath))
    if ($details.Count -eq 0) { return }

    # Sütun genişlikleri toplamı rapor genişliğine (72) eşittir; başlık ve satırlar aynı biçimi kullanır.
    $rowFormat = '  {0,-12}{1,-31}{2,14}{3,15}'
    $nameLimit = 30

    Write-ReportSection -Title 'VERGİ DETAYI'
    Write-ConsoleText -Text ($rowFormat -f 'HESAP KODU', 'HESAP ADI', 'BORÇ', 'ALACAK') -Color $script:Palette.Label
    Write-ReportRule

    $totalDebit = [decimal]0
    $totalCredit = [decimal]0
    foreach ($detail in $details) {
        $code = Get-XmlText -Node $detail -Path ".//*[local-name()='accountMainID']"
        $name = Get-XmlText -Node $detail -Path ".//*[local-name()='accountMainDescription']"
        $creditXPath = ".//*[local-name()='entryDetail'][*[local-name()='account']/*[local-name()='accountMainID']='$code' and *[local-name()='debitCreditCode']='C']"
        $creditDetail = $Accounting.SelectSingleNode($creditXPath)

        $debit = ConvertTo-DecimalAmount -Value (Get-XmlText -Node $detail -Path ".//*[local-name()='amount']")
        $credit = ConvertTo-DecimalAmount -Value (Get-XmlText -Node $creditDetail -Path ".//*[local-name()='amount']")
        $totalDebit += $debit
        $totalCredit += $credit

        if ($name.Length -gt $nameLimit) { $name = $name.Substring(0, $nameLimit - 1) + '…' }
        $debitColor = if ($debit -eq 0) { $script:Palette.Empty } else { $script:Palette.Value }
        $creditColor = if ($credit -eq 0) { $script:Palette.Empty } else { $script:Palette.Value }

        Write-ConsoleText -Text ('  {0,-12}' -f $code) -Color $script:Palette.Label -NoNewline
        Write-ConsoleText -Text ('{0,-31}' -f $name) -Color $script:Palette.Value -NoNewline
        Write-ConsoleText -Text ('{0,14}' -f (Format-Amount -Amount $debit)) -Color $debitColor -NoNewline
        Write-ConsoleText -Text ('{0,15}' -f (Format-Amount -Amount $credit)) -Color $creditColor
    }

    Write-ReportRule
    $totalLine = $rowFormat -f '', 'TOPLAM', (Format-Amount -Amount $totalDebit), (Format-Amount -Amount $totalCredit)
    Write-ConsoleText -Text $totalLine -Color $script:Palette.Total
}

function Write-XmlXsltAlignedReport {
    <#
    .SYNOPSIS
        Bir berat XML'inin başlık bilgilerini berat.xslt ile aynı bölüm sırasında terminale yazar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$Document,
        [Parameter(Mandatory = $true)][string]$FileName
    )

    $root = $Document.DocumentElement
    $type = Get-XmlText $root ".//*[local-name()='entriesType']"
    $docType = switch ($type) {
        'journal' { 'Yevmiye Defteri' }
        'ledger'  { 'Büyük Defter' }
        default   { 'Belge' }
    }
    $vkn = Get-XmlText $root ".//*[local-name()='context']/*[local-name()='entity']/*[local-name()='identifier'][@scheme='http://www.gib.gov.tr']"
    $entityInfo = $root.SelectSingleNode(".//*[local-name()='entityInformation']")
    $companyName = Get-XmlText $entityInfo ".//*[local-name()='organizationIdentifiers'][*[local-name()='organizationDescription']='Kurum Unvanı']/*[local-name()='organizationIdentifier']"
    $phone = Get-XmlText $entityInfo ".//*[local-name()='entityPhoneNumber']/*[local-name()='phoneNumber']"
    $fax = Get-XmlText $entityInfo ".//*[local-name()='entityFaxNumber']"
    $email = Get-XmlText $entityInfo ".//*[local-name()='entityEmailAddress']"
    $accountant = if ($entityInfo) { $entityInfo.SelectSingleNode(".//*[local-name()='accountantInformation']") } else { $null }
    $accountantName = Get-XmlText $accountant ".//*[local-name()='accountantName']"
    $accountantPhone = Get-XmlText $accountant ".//*[local-name()='accountantContactPhoneNumber']"
    $accountantFax = Get-XmlText $accountant ".//*[local-name()='accountantContactFaxNumber']"
    $accountantEmail = Get-XmlText $accountant ".//*[local-name()='accountantContactEmailAddress']"
    $accountantContract = Get-XmlText $accountant ".//*[local-name()='accountantEngagementTypeDescription']"
    $info = $root.SelectSingleNode(".//*[local-name()='documentInfo']")
    $start = Get-XmlText $info ".//*[local-name()='periodCoveredStart']"
    $end = Get-XmlText $info ".//*[local-name()='periodCoveredEnd']"
    $fiscalStart = Get-XmlText $entityInfo ".//*[local-name()='fiscalYearStart']"
    $fiscalEnd = Get-XmlText $entityInfo ".//*[local-name()='fiscalYearEnd']"
    $contextSegment = $root.SelectSingleNode(".//*[local-name()='context']//*[local-name()='segment']")
    $ettn = Get-XmlText $contextSegment ".//*[local-name()='uniqueID']"
    $ledgerSize = Get-XmlText $contextSegment ".//*[local-name()='measurableQuantity']"
    if ($ledgerSize) { $ledgerSize = "$ledgerSize MB" }
    $accounting = $root.SelectSingleNode(".//*[local-name()='accountingEntries']")

    Write-ReportFileHeader -FileName $FileName

    Write-ReportSection -Title 'MÜKELLEF BİLGİLERİ'
    Write-ReportField -Label 'VKN' -Value $vkn -Highlight
    Write-ReportField -Label 'UNVAN' -Value $companyName -Highlight
    Write-ReportField -Label 'TELEFON' -Value $phone
    Write-ReportField -Label 'FAX' -Value $fax
    Write-ReportField -Label 'E-POSTA' -Value $email

    Write-ReportSection -Title 'MESLEK MENSUBU BİLGİLERİ'
    Write-ReportField -Label 'UNVAN' -Value $accountantName
    Write-ReportField -Label 'TELEFON' -Value $accountantPhone
    Write-ReportField -Label 'FAX' -Value $accountantFax
    Write-ReportField -Label 'E-POSTA' -Value $accountantEmail
    Write-ReportField -Label 'SÖZLEŞME NO' -Value $accountantContract

    Write-ReportSection -Title 'DOKÜMAN BİLGİLERİ'
    Write-ReportField -Label 'DOKÜMAN TİPİ' -Value $docType -Highlight
    Write-ReportField -Label 'OLUŞTURAN' -Value (Get-XmlText $info ".//*[local-name()='creator']")
    Write-ReportField -Label 'DÖNEMİ' -Value (Format-XmlDateRange -Start $start -End $end) -Highlight
    Write-ReportField -Label 'TEKİL NO' -Value (Get-XmlText $info ".//*[local-name()='uniqueID']")
    Write-ReportField -Label 'OLUŞTURMA TARİHİ' -Value (Format-XmlDate -Value (Get-XmlText $info ".//*[local-name()='creationDate']"))
    Write-ReportField -Label 'KAYNAK UYGULAMA' -Value (Get-XmlText $info ".//*[local-name()='sourceApplication']")
    Write-ReportField -Label 'HESAP DÖNEMİ' -Value (Format-XmlDateRange -Start $fiscalStart -End $fiscalEnd)
    Write-ReportField -Label 'ETTN' -Value $ettn
    Write-ReportField -Label 'AÇIKLAMA' -Value (Get-XmlText $info ".//*[local-name()='entriesComment']")
    Write-ReportField -Label 'İLGİLİ DEFTER BOYUTU' -Value $ledgerSize
    Write-ReportField -Label 'YEVMİYE MADDESİ SAYISI' -Value (Get-XmlText $contextSegment ".//*[local-name()='numberOfEntries']")

    Write-TaxDetailTable -Accounting $accounting
    Write-XmlDocumentSignatureValue -Node $root.SelectSingleNode("./*[local-name()='SignatureValue']")
    Write-GibApprovalInfo -Document $Document
}

function Select-ModuleFolder {
    <#
    .SYNOPSIS
        Klasör seçici açar; seçilen yolu döndürür, iptalde $null döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter()][AllowNull()][string]$InitialPath)

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'XML klasörünün konumunu seçin.'
    $dialog.ShowNewFolderButton = $false
    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $dialog.SelectedPath = $InitialPath
    }

    try {
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        return $null
    } finally {
        $dialog.Dispose()
    }
}

function Get-ConfiguredFolder {
    <#
    .SYNOPSIS
        Kayıtlı XML klasörünü ayar dosyasından okur; yoksa $null döndürür.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    if (-not (Test-Path -LiteralPath $script:ModuleConfigPath -PathType Leaf)) { return $null }
    try {
        $settings = Get-Content -LiteralPath $script:ModuleConfigPath -Raw | ConvertFrom-Json
        if ($settings.XmlFolder) { return [string]$settings.XmlFolder }
    } catch {
        # Neden: Bozuk ayar dosyası modülü durdurmamalı; kullanıcı uyarılır ve klasör yeniden seçilir.
        Write-ConsoleText -Text ('Konum ayarı okunamadı: {0}' -f $_.Exception.Message) -Color $script:Palette.Warning
    }
    return $null
}

function Set-ConfiguredFolder {
    <#
    .SYNOPSIS
        Kullanıcıya klasör seçtirir ve ayar dosyasına kaydeder.
    #>
    [CmdletBinding()]
    param()

    $path = Select-ModuleFolder -InitialPath (Get-ConfiguredFolder)
    if (-not $path) {
        Write-ConsoleText -Text 'Konum tanımlama iptal edildi.' -Color $script:Palette.Warning
        return
    }

    try {
        [pscustomobject]@{ XmlFolder = $path } |
            ConvertTo-Json |
            Set-Content -LiteralPath $script:ModuleConfigPath -Encoding UTF8
        Write-ConsoleText -Text ('XML konumu kaydedildi: {0}' -f $path) -Color $script:Palette.Success
    } catch {
        Write-ConsoleText -Text ('Konum kaydedilemedi: {0}' -f $_.Exception.Message) -Color $script:Palette.Failure
    }
    [void](Read-Host 'Modül menüsüne dönmek için Enter')
}

function Show-XmlHeaderReport {
    <#
    .SYNOPSIS
        Seçilen klasördeki GIB*.xml beratlarının başlık bilgilerini terminalde gösterir.
    #>
    [CmdletBinding()]
    param()

    $folder = Select-ModuleFolder -InitialPath (Get-ConfiguredFolder)
    if (-not $folder) {
        Write-ConsoleText -Text 'Klasör seçimi iptal edildi.' -Color $script:Palette.Warning
        return
    }

    $files = @(Get-ChildItem -LiteralPath $folder -File -Filter 'GIB*.xml' -Recurse | Sort-Object Name)
    if ($files.Count -eq 0) {
        Write-ConsoleText -Text 'Seçilen konumda GIB ile başlayan XML bulunamadı.' -Color $script:Palette.Warning
        return
    }
    foreach ($file in $files) {
        try {
            $document = New-Object System.Xml.XmlDocument
            $document.Load($file.FullName)
            Write-XmlXsltAlignedReport -Document $document -FileName $file.Name
        } catch {
            # Neden: Bir dosyanın bozuk olması diğer beratların gösterilmesini engellememeli.
            Write-ConsoleText -Text ('XML okunamadı: {0} — {1}' -f $file.Name, $_.Exception.Message) -Color $script:Palette.Failure
        }
    }
    Write-ConsoleText -Text ''
    [void](Read-Host 'Ana menüye dönmek için Enter')
}

while ($true) {
    Clear-Host
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host '  📚 GİB Berat İşlemleri' -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ''
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host '  1. 🔍 Berat Görüntüleme Terminal'
    Write-Host '  2. 🌐 Berat Görüntüleme Web'
    Write-Host '  3. ⚙️ Konum Tanımlama'
    Write-Host '  Q. 🚪 Modülden çık'
    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host ''

    $mode = $null
    while ($null -eq $mode) {
        $choice = (Read-Host 'Seçiminiz').Trim().ToUpperInvariant()
        switch ($choice) {
            '1' { $mode = 'XmlBaslik' }
            '2' { $mode = 'Berat' }
            '3' { $mode = 'KonumTanimla' }
            'Q' { return }
            default { Write-Host 'Geçersiz seçim. 1, 2, 3 veya Q girin.' -ForegroundColor Yellow }
        }
    }

Clear-Host
Write-Host ""
Write-Host ("-" * 72) -ForegroundColor Magenta
if ($mode -eq 'XmlBaslik') {
    Write-Host '  🔍 Berat Görüntüleme Terminal' -ForegroundColor Cyan
} elseif ($mode -eq 'Berat') {
    Write-Host '  🌐 Berat Görüntüleme Web' -ForegroundColor Cyan
} elseif ($mode -eq 'KonumTanimla') {
    Write-Host '  ⚙️ Konum Tanımlama' -ForegroundColor Cyan
}
Write-Host ("-" * 72) -ForegroundColor Magenta
Write-Host ""
if ($mode -eq 'XmlBaslik') {
    Write-Host '📌 Klasör seçici kayıtlı konumda açılır; klasörü siz seçersiniz.' -ForegroundColor Yellow
    Write-Host '📌 Sonuçlar terminalde gösterilir.' -ForegroundColor Yellow
    Write-Host '📌 Başlıklar web görünümündeki berat.xslt ile uyumlu gösterilir.' -ForegroundColor Yellow
} elseif ($mode -eq 'Berat') {
    Write-Host '📌 Klasör seçici kayıtlı konumda açılır; klasörü siz seçersiniz.' -ForegroundColor Yellow
    Write-Host '📌 Berat arayüzü varsayılan tarayıcıda açılır.' -ForegroundColor Yellow
} elseif ($mode -eq 'KonumTanimla') {
    Write-Host '📌 Bu konum web ve terminal görüntüleme seçeneklerinin klasör seçicisinde başlangıç noktası olur.' -ForegroundColor Yellow
}
Write-Host ""

if ($mode -eq 'XmlBaslik') {
    Show-XmlHeaderReport
    continue
}
if ($mode -eq 'KonumTanimla') {
    Set-ConfiguredFolder
    continue
}
$folder = Select-ModuleFolder -InitialPath (Get-ConfiguredFolder)
if (-not $folder) {
    Write-Host 'Klasör seçimi iptal edildi.' -ForegroundColor Yellow
    continue
}

$xmlFiles = @(Get-ChildItem -LiteralPath $folder -File -Filter 'GIB*.xml' -Recurse | Sort-Object Name)
$xsltFile = Get-ChildItem -LiteralPath $folder -File -Filter 'berat.xslt' -Recurse | Select-Object -First 1
if (-not $xmlFiles.Count) {
    Write-Host 'Seçilen konumda GIB ile başlayan XML bulunamadı.' -ForegroundColor Yellow
    [void](Read-Host 'Modül menüsüne dönmek için Enter')
    continue
}
if (-not $xsltFile) {
    Write-Host 'Seçilen konumda berat.xslt bulunamadı.' -ForegroundColor Red
    [void](Read-Host 'Modül menüsüne dönmek için Enter')
    continue
}

$xmlPayload = foreach ($file in $xmlFiles) {
    [pscustomobject]@{
        name = $file.Name
        contentBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file.FullName))
    }
}
$beratPayload = [pscustomobject]@{
    xmlFiles = @($xmlPayload)
    xsltBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($xsltFile.FullName))
}
$beratPayloadJson = ConvertTo-Json -InputObject $beratPayload -Depth 5 -Compress
$beratPayloadJson = $beratPayloadJson.Replace('</', '<\/')

# Neden: Geçici HTML dosyası uygulamanın kendi Data\Temp\Modules\gibberatoku klasöründe tutulur (klasörü fonksiyon oluşturur).
$tempRoot = Get-MFTModuleTempPath -ModuleId 'gibberatoku'

$htmlPath = Join-Path $tempRoot 'index.html'

$html = @'
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Berat Görüntüleme</title>
<style>
* { box-sizing: border-box; }
body {
    margin: 0;
    font-family: "Segoe UI", Arial, sans-serif;
    background: #f3f4f6;
    color: #222;
}
.ust {
    height: 60px;
    background: #fff;
    border-bottom: 1px solid #ddd;
    display: flex;
    align-items: center;
    padding: 0 20px;
    font-size: 20px;
    font-weight: 600;
}
.ana {
    display: flex;
    height: calc(100vh - 60px);
}
.sol {
    width: 330px;
    background: #fff;
    border-right: 1px solid #ddd;
    display: flex;
    flex-direction: column;
}
.sol-ust {
    padding: 15px;
    border-bottom: 1px solid #ddd;
}
.buton {
    display: block;
    width: 100%;
    padding: 10px;
    background: #2563eb;
    color: white;
    border-radius: 6px;
    text-align: center;
    cursor: pointer;
}
.buton:hover { background: #1d4ed8; }
#klasor { display: none; }
.durum {
    margin-top: 10px;
    font-size: 13px;
    color: #666;
}
.liste {
    flex: 1;
    overflow-y: auto;
    padding: 8px;
}
.dosya {
    padding: 11px 10px;
    margin-bottom: 5px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 13px;
    word-break: break-all;
}
.dosya:hover { background: #eef4ff; }
.dosya.aktif {
    background: #dbeafe;
    color: #1e40af;
    font-weight: 600;
}
.sag {
    flex: 1;
    background: #e5e7eb;
    display: flex;
    flex-direction: column;
    min-width: 0;
}
.sag-ust {
    height: 48px;
    background: white;
    border-bottom: 1px solid #ddd;
    padding: 14px 18px;
    font-size: 14px;
    color: #555;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.goruntu {
    flex: 1;
    padding: 12px;
}
iframe {
    width: 100%;
    height: 100%;
    border: 1px solid #ccc;
    background: white;
    border-radius: 5px;
}
.bos {
    height: 100%;
    display: flex;
    justify-content: center;
    align-items: center;
    color: #777;
    font-size: 15px;
}
</style>
</head>
<body>
<div class="ust">Berat Görüntüleme</div>
<div class="ana">
    <div class="sol">
        <div class="sol-ust">
            <div class="buton">Berat Dosyaları</div>
            <div id="durum" class="durum">Dosyalar yükleniyor.</div>
        </div>
        <div id="liste" class="liste">
            <div class="bos">GIB beratları burada listelenecek.</div>
        </div>
    </div>
    <div class="sag">
        <div id="dosyaAdi" class="sag-ust">Berat seçilmedi.</div>
        <div class="goruntu">
            <iframe id="goruntu"></iframe>
        </div>
    </div>
</div>
<script>
const embeddedData = __BERAT_PAYLOAD__;
const decodeBase64 = value => new TextDecoder("utf-8").decode(
    Uint8Array.from(atob(value), character => character.charCodeAt(0))
);
const xsltFile = decodeBase64(embeddedData.xsltBase64);
const xmlFiles = embeddedData.xmlFiles.map(file => ({
    name: file.name,
    text: decodeBase64(file.contentBase64)
}));
const durum = document.getElementById("durum");
const liste = document.getElementById("liste");
const goruntu = document.getElementById("goruntu");
const dosyaAdi = document.getElementById("dosyaAdi");

durum.textContent = xmlFiles.length + " adet GIB beratı bulundu.";
listeyiGoster();

function listeyiGoster() {
    liste.innerHTML = "";

    if (xmlFiles.length === 0) {
        liste.innerHTML = '<div class="bos">GIB ile başlayan XML bulunamadı.</div>';
        return;
    }

    xmlFiles.forEach(file => {
        const eleman = document.createElement("div");
        eleman.className = "dosya";
        eleman.textContent = file.name;
        eleman.title = file.name;

        eleman.addEventListener("click", function () {
            document.querySelectorAll(".dosya").forEach(x =>
                x.classList.remove("aktif")
            );
            eleman.classList.add("aktif");
            beratiGoster(file);
        });

        liste.appendChild(eleman);
    });
}

async function beratiGoster(xmlFile) {
    dosyaAdi.textContent = xmlFile.name;

    if (!xsltFile) {
        goruntu.srcdoc =
            "<h3 style='color:red'>berat.xslt bulunamadı.</h3>";
        return;
    }

    try {
        const xmlText = xmlFile.text;
        const xsltText = xsltFile;

        const parser = new DOMParser();
        const xml = parser.parseFromString(xmlText, "application/xml");
        const xslt = parser.parseFromString(xsltText, "application/xml");

        if (xml.getElementsByTagName("parsererror").length > 0) {
            throw new Error("XML okunamadı.");
        }

        if (xslt.getElementsByTagName("parsererror").length > 0) {
            throw new Error("berat.xslt okunamadı.");
        }

        const processor = new XSLTProcessor();
        processor.importStylesheet(xslt);

        const sonuc = processor.transformToDocument(xml);
        const serializer = new XMLSerializer();
        const html = serializer.serializeToString(sonuc);

        goruntu.srcdoc = html;
    }
    catch (error) {
        console.error(error);
        goruntu.srcdoc =
            "<h3 style='color:red'>Berat görüntülenemedi.</h3>" +
            "<p>" + escapeHtml(error.message || String(error)) + "</p>";
    }
}

function escapeHtml(value) {
    return value
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}
</script>
</body>
</html>
'@

$html = $html.Replace('__BERAT_PAYLOAD__', $beratPayloadJson)

# UTF-8 olarak kaydet
[System.IO.File]::WriteAllText(
    $htmlPath,
    $html,
    [System.Text.UTF8Encoding]::new($false)
)

# Varsayılan tarayıcıda aç.
Start-Process $htmlPath
Write-ModuleStatus 'İşlem tamamlandı. Seçilen arayüz açıldı; ana menüye dönülüyor.'
Write-Host 'Web ekranını kapattıktan sonra modül menüsüne dönmek için Enter tuşuna basın.' -ForegroundColor Yellow
[void](Read-Host)
Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue
}

