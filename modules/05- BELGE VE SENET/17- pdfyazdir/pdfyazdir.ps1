# --------------------------------------------
# 🚀 Klasör Seçim Penceresi Fonksiyonu
# --------------------------------------------
function Select-FolderDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Lütfen klasör seçin"
    $fbd.ShowNewFolderButton = $false
    $result = $fbd.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fbd.SelectedPath
    }
    else {
        return $null
    }
}

# --------------------------------------------
# 🚀 Seçilen klasördeki PDF dosyalarını tespit eder ve Adobe Acrobat varsa sessiz modda,
#    yoksa varsayılan yazıcı ile yazdırma işlemi gerçekleştirir.
# --------------------------------------------
function Print-PdfFiles {
    param (
        [string]$FolderPath
    )
    $files = Get-ChildItem -Path $FolderPath -File
    if ($files.Count -eq 0) {
        Write-Host "📁 Bu klasörde dosya bulunamadı." -ForegroundColor Yellow
        return
    }
    $acrobatPaths = @(
        "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "$env:ProgramFiles(x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    )
    $acrobatPath = $acrobatPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($acrobatPath) {
        Write-Host "Adobe Acrobat bulundu: $acrobatPath" -ForegroundColor Green
    }
    else {
        Write-Host "Adobe Acrobat bulunamadı, varsayılan uygulama ile yazdırılacak." -ForegroundColor Yellow
    }
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

    [System.IO.FileInfo[]]$pdfFiles = @(Get-ChildItem -Path $FolderPath -Filter "*.pdf" -File)
    [System.Array]::Sort(
        $pdfFiles,
        [System.Comparison[System.IO.FileInfo]]{
            param($left, $right)
            [MFilesToolkit.NaturalSort]::StrCmpLogicalW($left.Name, $right.Name)
        }
    )
    if ($pdfFiles.Count -eq 0) {
        Write-Warning "⚠️ Bu klasörde PDF dosyası bulunamadı."
        return
    }
	# 🏁 Adobe Acrobat varsa, önceden gizli şekilde yeni oturumla başlat
	if ($acrobatPath) {
	    Start-Process -FilePath $acrobatPath -ArgumentList "/n" -WindowStyle Hidden
	    Start-Sleep -Milliseconds 500
	}

	# 🖨️ PDF dosyalarını sırayla yazdır
	foreach ($file in $pdfFiles) {
	    try {
        Write-Host "🖨️ Yazdırılıyor: $($file.Name)" -ForegroundColor Green
        if ($acrobatPath) {
            Start-Process -FilePath $acrobatPath -ArgumentList "/h /t `"$($file.FullName)`"" -NoNewWindow
        }
        else {
            Start-Process -FilePath $file.FullName -Verb Print
        }
        # Gerekirse aşağıdaki satırı aktif et
            Start-Sleep -Milliseconds 300
    }
    catch {
        Write-Warning "HATA: $($file.Name) yazdırılamadı. $_"
    }
}
	Write-Host "`n✅ Tüm yazdırma komutları gönderildi." -ForegroundColor Green
	# 🧹 Acrobat işlemlerini kapat
	if ($acrobatPath) {
    	    Get-Process AcroRd32 -ErrorAction SilentlyContinue | ForEach-Object {
                $_.CloseMainWindow() | Out-Null
                Start-Sleep -Milliseconds 500
                if (!$_.HasExited) {
                    $_ | Stop-Process -Force
                }
            }
	}
}  

# --------------------------------------------
# 🚀 Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    # Klasör seçimi
    $selectedPath = Select-FolderDialog
    if (-not $selectedPath) {
        Write-Host "İşlem iptal edildi." -ForegroundColor Yellow
        return $false
    }
    Write-Host "📂 Seçilen klasör: $selectedPath" -ForegroundColor Cyan
    # PDF dosyalarını yazdır
    Print-PdfFiles -FolderPath $selectedPath
    return $true
}

# --------------------------------------------
# 🚀 Başlık yazdırma fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  🅿️ PDF Klasör Yazdırma' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencerede bir PDF klasörü seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçilen klasördeki PDF dosyaları yazdırılır." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    [void](StartProgram)
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
