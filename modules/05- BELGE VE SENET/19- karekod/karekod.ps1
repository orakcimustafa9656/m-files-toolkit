function Start-KarekodOkuyucu {
    $ErrorActionPreference = "Stop"
    
    # Add-Type'ları sadece ilk çağrıda yükle
    if (-not $global:KarekodTypesLoaded) {
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms
        $global:KarekodTypesLoaded = $true
    }

    $root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $zxingDll = Get-ChildItem -Path $root -Recurse -Filter "zxing.dll" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $zxingDll) {
        Write-Host "❌ ZXing kütüphanesi bulunamadı: $root" -ForegroundColor Red
        return
    }
    Add-Type -Path $zxingDll.FullName

    Clear-Host
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  📷 JPG/PNG Karekod Okuyucu ve Excel Aktarıcı" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencerede JPG/PNG klasörü seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçilen klasördeki dosyalar Excel'de raporlanır." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "📁 Karekod içeren görsellerin bulunduğu klasörü seçin"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or
        [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
        Write-Host "⚠️ İşlem iptal edildi." -ForegroundColor Yellow
        return
    }
    $inputDir = $dialog.SelectedPath
    Write-Host "`n📂 Seçilen klasör: $inputDir" -ForegroundColor Gray

    $imgFiles = Get-ChildItem -Path $inputDir -Include *.jpg, *.jpeg, *.png -Recurse | Sort-Object Name
    if ($imgFiles.Count -eq 0) {
        Write-Host "⚠️ Seçilen klasörde uygun görsel dosyası bulunamadı." -ForegroundColor Yellow
        return
    }

    $results = @()
    $i = 0

    Write-Host "🔍 Görsellerde karekod aranıyor..."
    $reader = New-Object ZXing.BarcodeReader
    foreach ($file in $imgFiles) {
        $imgPath = $file.FullName
        $original = $null
        $bitmap = $null

        try {
            # Bitmap kilidini çözmek için kopyasını kullan
            $original = [System.Drawing.Bitmap]::FromFile($imgPath)
            $bitmap = New-Object System.Drawing.Bitmap($original)
            $decoded = $reader.DecodeMultiple($bitmap)

            if ($decoded) {
                $texts = $decoded | ForEach-Object { $_.Text }
                $text = ($texts -join "`n")
                $status = "✅ $($texts.Count) karekod okundu"
            }
            else {
                $text = ""
                $status = "❌ Karekod bulunamadı"
            }
        }
        catch {
            $text = ""
            $status = "⚠️ Hata: $_"
        }
        finally {
            if ($bitmap) { $bitmap.Dispose() }
            if ($original) { $original.Dispose() }
        }

        $results += [PSCustomObject]@{
            Dosya   = $file.Name
            Durum   = $status
            Karekod = $text
        }

        $i++
        Write-Host -NoNewline "."
        if ($i % 50 -eq 0) { Write-Host "" }
    }

    Write-Host "`n🔎 Toplam $i dosya işlendi."

    $excel = $null
    $workbook = $null
    $sheet = $null
    $excelReportCreated = $false
    try {
        Write-Host "📊 Excel'de sonuçlar açılıyor..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true

        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)

        $sheet.Cells.Item(1, 1) = "Dosya Adı"
        $sheet.Cells.Item(1, 2) = "Durum"
        $sheet.Cells.Item(1, 3) = "Karekod"
        $sheet.Rows.Item(1).Font.Bold = $true

        $row = 2
        foreach ($item in $results) {
            $dosya = $item.Dosya
            $durum = $item.Durum
            $karekodlar = $item.Karekod -split "`n"

            foreach ($karekod in $karekodlar) {
                $sheet.Cells.Item($row, 1) = $dosya
                $sheet.Cells.Item($row, 2) = $durum
                $sheet.Cells.Item($row, 3) = $karekod
                $row++
            }
        }

        $sheet.Columns.Item(1).AutoFit()
        $sheet.Columns.Item(2).AutoFit()
        $sheet.Columns.Item(3).ColumnWidth = 100
        $sheet.Columns.Item(3).WrapText = $true
        $excelReportCreated = $true
    }
    catch {
        Write-Host "❌ Excel açılırken hata oluştu: $_" -ForegroundColor Red
    }
    finally {
        if (-not $excelReportCreated) {
            if ($workbook) { try { $workbook.Close($false) } catch {} }
            if ($excel) { try { $excel.Quit() } catch {} }
            foreach ($comObject in @($sheet, $workbook, $excel)) {
                if ($comObject) {
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
                }
            }
        }
    }
    if ($excelReportCreated) {
        Write-Host "`n✅ Rapor Excel'de açıldı. Kaydetmeyi unutmayın!" -ForegroundColor Green
    }
}

# 🚀 Ana Döngü
do {
    Clear-Host
    Start-KarekodOkuyucu
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
