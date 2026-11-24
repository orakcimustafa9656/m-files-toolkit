# --------------------------------------------
# 🎨 Yükleme Efekti
# --------------------------------------------

function Show-ConsoleProgress {
    param(
        [int]$current,
        [int]$total,
        [int]$step,
        [string]$message = "Yükleniyor"
    )
    $frames = @("|", "/", "-", "\")
    $percent = [math]::Round(($current / $total) * 100)
    $frame = $frames[$step % $frames.Length]
    Write-Host -NoNewline "`r$message... %$percent $frame  "
}

# --------------------------------------------
# 🧠 Excel İşleme Fonksiyonu
# --------------------------------------------
function Process-ExcelFile {
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    Add-Type -AssemblyName PresentationFramework
    $openDlg = New-Object Microsoft.Win32.OpenFileDialog
    $openDlg.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx"
    $openDlg.Title  = "İşlenecek Excel dosyasını seçin"
    if ($openDlg.ShowDialog() -ne $true) { return $false }

    $inputPath = $openDlg.FileName
    $folderPath = [System.IO.Path]::GetDirectoryName($inputPath)
    $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $excelApp = $null
    $workbook = $null
    $worksheet = $null
    $filteredWb = $null
    $filteredWs = $null
    $nonFilteredWb = $null
    $nonFilteredWs = $null
    $filteredTempPath = $null
    $nonFilteredTempPath = $null

    Write-Host "📂 Seçilen dosya: $inputPath"
    Write-Host "📂 Dosya açılıyor, lütfen bekleyin..."

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $workbook   = $excelApp.Workbooks.Open($inputPath)
        $worksheet  = $workbook.Sheets.Item(1)
        $usedRange  = $worksheet.UsedRange
        $rowCount   = $usedRange.Rows.Count

        $filteredWb    = $excelApp.Workbooks.Add()
        $filteredWs    = $filteredWb.Sheets.Item(1)
        $nonFilteredWb = $excelApp.Workbooks.Add()
        $nonFilteredWs = $nonFilteredWb.Sheets.Item(1)

        $filteredRow    = 1
        $nonFilteredRow = 1
        $errors         = @()

        # Başlık satırını kopyala ve görünür yap
        $headerRow = $worksheet.Rows.Item(1)
        $headerRow.Copy($filteredWs.Rows.Item($filteredRow)) | Out-Null
        $headerRow.Copy($nonFilteredWs.Rows.Item($nonFilteredRow)) | Out-Null
        $filteredWs.Rows.Item($filteredRow).EntireRow.Hidden = $false
        $nonFilteredWs.Rows.Item($nonFilteredRow).EntireRow.Hidden = $false
        $filteredRow++
        $nonFilteredRow++

        for ($i = 2; $i -le $rowCount; $i++) {
            try {
                $row = $worksheet.Rows.Item($i)
                $hidden1 = $row.EntireRow.Hidden
                $hidden2 = $row.Hidden
                $isVisible = -not ($hidden1 -or $hidden2)

                $targetSheet = if ($isVisible) { $filteredWs } else { $nonFilteredWs }
                $targetRow   = if ($isVisible) { $filteredRow } else { $nonFilteredRow }

                $row.Copy($targetSheet.Rows.Item($targetRow)) | Out-Null
                $targetSheet.Rows.Item($targetRow).EntireRow.Hidden = $false

                if ($isVisible) { $filteredRow++ } else { $nonFilteredRow++ }

                # İlerlemeyi göster (her 5 satırda)
                if ($i % 5 -eq 0 -or $i -eq $rowCount) {
                    Show-ConsoleProgress -current $i -total $rowCount -step $i -message "İşleniyor"
                }
            } catch {
                $errors += "• Satır $i kopyalanırken hata: $_"
            }
        }

        $filteredPath    = [System.IO.Path]::Combine($folderPath, "$baseName-filtreli.xlsx")
        $nonFilteredPath = [System.IO.Path]::Combine($folderPath, "$baseName-filtresiz.xlsx")
        $suffix = [guid]::NewGuid().ToString('N')
        $filteredTempPath = [System.IO.Path]::Combine($folderPath, ".$baseName-filtreli-$suffix.xlsx")
        $nonFilteredTempPath = [System.IO.Path]::Combine($folderPath, ".$baseName-filtresiz-$suffix.xlsx")

        $filteredWb.SaveAs($filteredTempPath)
        $nonFilteredWb.SaveAs($nonFilteredTempPath)

        $filteredWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($filteredWb)
        $filteredWb = $null
        $nonFilteredWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($nonFilteredWb)
        $nonFilteredWb = $null
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        $workbook = $null
        foreach ($item in @($filteredWs, $nonFilteredWs, $worksheet)) {
            if ($item) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        }
        $filteredWs = $null
        $nonFilteredWs = $null
        $worksheet = $null
        $excelApp.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
        $excelApp = $null

        # İki yeni dosya da başarıyla kaydedildikten sonra hedeflerle değiştir.
        Move-Item -LiteralPath $filteredTempPath -Destination $filteredPath -Force -ErrorAction Stop
        $filteredTempPath = $null
        Move-Item -LiteralPath $nonFilteredTempPath -Destination $nonFilteredPath -Force -ErrorAction Stop
        $nonFilteredTempPath = $null

        # Son mesajı temiz satırda göster
        Write-Host "`r✅ İşlem başarıyla tamamlandı!           " -ForegroundColor Green

        if ($errors.Count -gt 0) {
            Write-Host "⚠️ İşlem tamamlandı ancak bazı hatalar oluştu:" -ForegroundColor Yellow
            $errors | ForEach-Object { Write-Host $_ -ForegroundColor DarkYellow }
        }

        Write-Host "📄 Filtreli dosya: $filteredPath"    -ForegroundColor Green
        Write-Host "📄 Filtresiz dosya: $nonFilteredPath" -ForegroundColor Green
        return $true

    } catch {
        Write-Host "`r❌ Beklenmeyen bir hata oluştu: $_" -ForegroundColor Red
        return $false
    } finally {
        foreach ($item in @($filteredWb, $nonFilteredWb, $workbook)) {
            if ($null -ne $item) {
                try { $item.Close($false) } catch {}
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) } catch {}
            }
        }
        if ($excelApp) { try { $excelApp.Quit() } catch {} }
        foreach ($item in @($filteredWs, $nonFilteredWs, $worksheet, $excelApp)) {
            if ($null -ne $item) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) } catch {}
            }
        }
        foreach ($tempPath in @($filteredTempPath, $nonFilteredTempPath)) {
            if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Show-ConsoleProgress {
    param(
        [int]$current,
        [int]$total,
        [int]$step,
        [string]$message = "İşleniyor"
    )
    $frames = @("|", "/", "-", "\")
    $percent = [math]::Round(($current / $total) * 100)
    $frame = $frames[$step % $frames.Length]
    Write-Host -NoNewline "`r$message... %$percent $frame  "
}



# --------------------------------------------
# 🚀 Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  🔍 Excel Filtreli Satırları Parçalama' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencerede bir Excel dosyası seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçtiğiniz dosyadaki filtreli satırları ayrı dosyaya diğerlerini ayrı dosyaya böler." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    $selectionMade = Process-ExcelFile
    if (-not $selectionMade) { Write-Host "⚠️ İşlem iptal edildi." -ForegroundColor Yellow }
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
