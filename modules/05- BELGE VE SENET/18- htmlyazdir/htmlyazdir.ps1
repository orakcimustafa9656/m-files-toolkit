# --------------------------------------------
# 🚀  Renkli Mesaj Yazdırma Fonksiyonu
# --------------------------------------------
function Yaz-Mesaj {
    param (
        [string]$Mesaj,
        [ConsoleColor]$Renk = 'White',
        [string]$Ikon = "ℹ️"
    )
    Write-Host "$Ikon  $Mesaj" -ForegroundColor $Renk
}

# --------------------------------------------
# 🚀  HTML Dosyalarını Yazdıran Ana Fonksiyon
# --------------------------------------------
function StartProgram {
    # Gerekli .NET tiplerini yükle
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # 📂 Klasör seçimi
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Lütfen HTML dosyalarının bulunduğu klasörü seçin"
    $folderDialog.ShowNewFolderButton = $false

    if ($folderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Yaz-Mesaj "İşlem iptal edildi." Yellow "⚠️"
        return
    }

    $selectedPath = $folderDialog.SelectedPath
    if (-not ('MFilesToolkit.NaturalSort' -as [type])) {
        Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;

namespace MFilesToolkit
{
    public static class NaturalSort
    {
        [DllImport("Shlwapi.dll", CharSet = CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string left, string right);
    }
}
"@
    }

    [System.IO.FileInfo[]]$htmlFiles = @(Get-ChildItem -Path $selectedPath -Filter *.html -File)
    [System.Array]::Sort(
        $htmlFiles,
        [System.Comparison[System.IO.FileInfo]]{
            param($left, $right)
            [MFilesToolkit.NaturalSort]::StrCmpLogicalW($left.Name, $right.Name)
        }
    )

    if ($htmlFiles.Count -eq 0) {
        Yaz-Mesaj "Seçilen klasörde .html dosyası bulunamadı." Yellow "⚠️"
        return
    }

    Yaz-Mesaj "$($htmlFiles.Count) adet HTML dosyası yazdırılıyor..." Green "🖨️"

    foreach ($file in $htmlFiles) {
        Yaz-Mesaj "Yazdırılıyor: $($file.Name)" Cyan "🔸"

        $form = New-Object System.Windows.Forms.Form
        $form.Width = 800
        $form.Height = 600
        $form.WindowState = 'Minimized'

        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.ScriptErrorsSuppressed = $true
        $browser.Dock = 'Fill'
        $form.Controls.Add($browser)

        $browser.Add_DocumentCompleted({
            # Alt çerçevelerin yüklenme olayında yazdırmayı erken başlatma.
            if ($event.Url -ne $this.Url) { return }
            Start-Sleep -Milliseconds 500
            $this.Print()
            $form.Close()
        })

        $browser.Url = "file:///$($file.FullName)"
        $form.ShowDialog()

        Start-Sleep -Seconds 1
    }

    Yaz-Mesaj "Tüm işlemler tamamlandı." Green "✅"
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  🌍 HTML Klasör Yazdırma Aracı" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencere ile bir klasör seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçilen klasördeki HTML dosyaları yazdırılır." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    StartProgram
    Write-Host ""
    Write-Host ("-" * 52) -ForegroundColor DarkGray
    Write-Host "  🔚 [Q/Esc] Ana menüye dön  |  [D] Yeniden çalıştır" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor DarkGray
    Write-Host ""
    do {
        $keyInfo = [System.Console]::ReadKey($true)
        $next = $keyInfo.Key.ToString().ToUpper()
    } while ($next -ne 'Q' -and $next -ne 'D' -and $next -ne 'ESCAPE')
    if ($next -eq 'Q' -or $next -eq 'ESCAPE') { break }
} while ($true)
