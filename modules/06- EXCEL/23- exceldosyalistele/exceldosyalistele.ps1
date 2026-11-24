# --------------------------------------------
# 🚀  Klasör Seçim Penceresi Fonksiyonu
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
# 🚀  Excel'e Dosya Listesi Aktarma Fonksiyonu
# --------------------------------------------
function Export-FilesToExcel {
    param (
        [string]$FolderPath,
        [System.IO.FileInfo[]]$Files
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.UserControl = $true
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    $sheet.Name = "Dosyalar"

    $headers = @("Dosya Adı", "Boyut (KB)", "Oluşturulma", "Son Değiştirilme")
    for ($i=0; $i -lt $headers.Length; $i++) {
        $cell = $sheet.Cells.Item(1, $i+1)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.ColorIndex = 36
        $cell.Font.ColorIndex = 1
        $cell.Borders.LineStyle = 1
        $cell.HorizontalAlignment = -4108
    }

    $row = 2
    foreach ($file in $Files) {
        $sheet.Cells($row, 1).Value2 = $file.Name
        $sheet.Cells($row, 2).Value2 = [double][math]::Round($file.Length / 1KB, 2)
        $sheet.Cells($row, 3).Value2 = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $sheet.Cells($row, 4).Value2 = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

        for ($col=1; $col -le 4; $col++) {
            $sheet.Cells($row, $col).Borders.LineStyle = 1
        }
        $row++
    }

    $lastDataRow = $row - 1
    # Sadece boyut sütununa sayı formatı uygula
    $sheet.Range("B2:B$lastDataRow").NumberFormat = "#,##0.00"
    $sheet.Columns("A:D").AutoFit()

    $totalSizeKB = [math]::Round(($Files | Measure-Object -Property Length -Sum).Sum / 1KB, 2)
    $footerText = "Toplam Dosya Sayısı: $($Files.Count), Toplam Boyut: $totalSizeKB KB"
    $sheet.Cells($row + 1, 1).Value2 = [string]$footerText
    $sheet.Range("A$($row + 1):D$($row + 1)").Merge()
    $sheet.Cells($row + 1, 1).Font.Italic = $true
    $sheet.Cells($row + 1, 1).HorizontalAlignment = -4131

    Write-Host "Excel dosyası başarıyla oluşturuldu." -ForegroundColor Green

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --------------------------------------------
# 🚀  Ana Program Döngüsü
# --------------------------------------------
function StartProgram {
    $selectedPath = Select-FolderDialog
    if (-not $selectedPath) {
        Write-Host "İşlem iptal edildi." -ForegroundColor Yellow
        return $false
    }
    Write-Host "📂 Seçilen klasör: $selectedPath" -ForegroundColor Cyan
    $files = Get-ChildItem -Path $selectedPath -File
    if ($files.Count -eq 0) {
        Write-Host "📁 Bu klasörde dosya bulunamadı." -ForegroundColor Yellow
        return
    }
    Export-FilesToExcel -FolderPath $selectedPath -Files $files
    return $true
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  📒 Excel Dosya Listesi' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencerede bir klasör seçin." -ForegroundColor Yellow
    Write-Host "📌 Klasör içindeki tüm dosya adlarını Excel dosyasına aktarır." -ForegroundColor Yellow
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
