function Read-ToolkitLine {
    param([string]$Prompt)

    if ($Prompt) { Write-Host $Prompt -NoNewline }
    $buffer = [System.Text.StringBuilder]::new()
    $cursor = 0
    $lineStartX = [System.Console]::CursorLeft
    $lineStartY = [System.Console]::CursorTop
    $width = [System.Console]::BufferWidth
    if ($width -lt 1) { $width = 80 }
    $renderedLength = 0

    while ($true) {
        $key = [System.Console]::ReadKey($true)
        $changed = $false
        switch ($key.Key) {
            Enter {
                $absolute = $lineStartX + $buffer.Length
                [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
                [System.Console]::WriteLine()
                return $buffer.ToString()
            }
            Escape {
                $absolute = $lineStartX + $buffer.Length
                [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
                [System.Console]::WriteLine()
                throw [System.OperationCanceledException]::new('Kullanıcı Esc ile işlemi iptal etti.')
            }
            Left {
                if ($cursor -gt 0) { $cursor--; $changed = $true }
            }
            Right {
                if ($cursor -lt $buffer.Length) { $cursor++; $changed = $true }
            }
            Home {
                if ($cursor -gt 0) { $cursor = 0; $changed = $true }
            }
            End {
                if ($cursor -lt $buffer.Length) { $cursor = $buffer.Length; $changed = $true }
            }
            Backspace {
                if ($cursor -gt 0) {
                    [void]$buffer.Remove($cursor - 1, 1)
                    $cursor--
                    $changed = $true
                }
            }
            Delete {
                if ($cursor -lt $buffer.Length) {
                    [void]$buffer.Remove($cursor, 1)
                    $changed = $true
                }
            }
            default {
                if (-not [char]::IsControl($key.KeyChar)) {
                    [void]$buffer.Insert($cursor, $key.KeyChar)
                    $cursor++
                    $changed = $true
                }
            }
        }

        if ($changed) {
            [System.Console]::SetCursorPosition($lineStartX, $lineStartY)
            [System.Console]::Write($buffer.ToString())
            $clearCount = $renderedLength - $buffer.Length
            if ($clearCount -gt 0) { [System.Console]::Write((' ' * $clearCount)) }
            $absolute = $lineStartX + $cursor
            [System.Console]::SetCursorPosition(($absolute % $width), ($lineStartY + [int][Math]::Floor($absolute / $width)))
            $renderedLength = $buffer.Length
        }
    }
}


# --------------------------------------------
# 🚀 Excel dosyasını seç ve aç, Excel objelerini döndür
# --------------------------------------------
function Load-ExcelFile {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"

    $objFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $objFileDialog.Filter = "Excel Dosyalari (*.xlsx;*.xls)|*.xlsx;*.xls|Tum Dosyalar (*.*)|*.*"
    $objFileDialog.Title = "Bolmek Istediginiz Excel Dosyasini Secin"
    $objFileDialog.RestoreDirectory = $true

    if ($objFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $script:FileSelectionCancelled = $true
        return $null
    }

    $fullPath = $objFileDialog.FileName
    $outputDirectory = Split-Path -Path $fullPath -Parent

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Open($fullPath)

    return @{ Excel = $Excel; Workbook = $Workbook; OutputDir = $outputDirectory }
}

# --------------------------------------------
# 🚀 Sayfa seçimi yap (çok sayfa varsa kullanıcıdan numara al)
# --------------------------------------------
function Select-Worksheet {
    param (
        $Workbook
    )

    if ($Workbook.Sheets.Count -gt 1) {
        Write-Host "`nSecilen Excel dosyasinda birden fazla sayfa bulundu:" -ForegroundColor Yellow
        for ($i = 1; $i -le $Workbook.Sheets.Count; $i++) {
            Write-Host "$i. $($Workbook.Sheets.Item($i).Name)" -ForegroundColor Cyan
        }

        $sheetIndex = 0
        do {
            $input = Read-ToolkitLine -Prompt "Islemek istediginiz sayfanin numarasini girin (orn: 1): "
            if ([int]::TryParse($input, [ref]$sheetIndex) -and $sheetIndex -ge 1 -and $sheetIndex -le $Workbook.Sheets.Count) {
                break
            } else {
                Write-Host "Gecersiz giris. Lutfen gecerli bir sayfa numarasi girin." -ForegroundColor Yellow
            }
        } while ($true)
        return $Workbook.Sheets.Item($sheetIndex)
    } else {
        return $Workbook.Sheets.Item(1)
    }
}

# --------------------------------------------
# 🚀 Başlık satırı var mı sor
# --------------------------------------------
function Ask-HeaderPresence {
    do {
        $rawHeaderInput = Read-ToolkitLine -Prompt "Ilk satir baslik satiri mi? (E/H): "
        if ($rawHeaderInput -eq 'E' -or $rawHeaderInput -eq 'e') {
            return $true
        } elseif ($rawHeaderInput -eq 'H' -or $rawHeaderInput -eq 'h') {
            return $false
        } else {
            Write-Host "Gecersiz giris. Lutfen 'E' veya 'H' girin." -ForegroundColor Yellow
        }
    } while ($true)
}

# --------------------------------------------
# 🚀 Parça başına satır sayısını sor
# --------------------------------------------
function Ask-SplitRowCount {
    do {
	[int]$count = 0 
        $input = Read-ToolkitLine -Prompt "Her bir parcanin kac satir olmasini istersiniz? (Pozitif tam sayi): "
        if ([int]::TryParse($input, [ref]$count) -and $count -gt 0) {
            return $count
        } else {
            Write-Host "Gecersiz giris. Lutfen pozitif bir tam sayi girin." -ForegroundColor Yellow
        }
    } while ($true)
}

# --------------------------------------------
# 🚀 Çıktı dosya adı ön eki sor
# --------------------------------------------
function Ask-OutputPrefix {
    $outputPrefix = Read-ToolkitLine -Prompt "Cikti dosya adi on eki girin (orn: BolunmusVeri_). Bos birakirsaniz 'Parca_' kullanilir.: "
    if ([string]::IsNullOrWhiteSpace($outputPrefix)) {
        $outputPrefix = "Parca_"
    }
    return $outputPrefix
}

# --------------------------------------------
# 🚀 İşlem başlasın mı onayı al
# --------------------------------------------
function Confirm-StartProcessing {
    do {
        $confirmStartInput = Read-ToolkitLine -Prompt "Dosya Parcala islemi baslasin mi? (E/H): "
        if ($confirmStartInput -eq 'E' -or $confirmStartInput -eq 'e') {
            return $true
        } elseif ($confirmStartInput -eq 'H' -or $confirmStartInput -eq 'h') {
            Write-Host "İşlem iptal edildi." -ForegroundColor Yellow
            return $false
        } else {
            Write-Host "Gecersiz giris. Lutfen 'E' veya 'H' girin." -ForegroundColor Yellow
        }
    } while ($true)
}

# --------------------------------------------
# 🚀 Excel dosyasını parçala
# --------------------------------------------
function Split-ExcelFile {
    param (
        $Sheet,
        $lastUsedRow,
        $lastUsedColumn,
        $hasHeader,
        $parcaSatirSayisi,
        $outputPrefix,
        $outputDirectory,
        $Excel
    )

    $headerOffset = if ($hasHeader) { 1 } else { 0 }
    $originalDataStartRow = if ($hasHeader) { 2 } else { 1 }
    $veriSatirSayisi = $lastUsedRow - $headerOffset
    $dosyaSayisi = [math]::Ceiling($veriSatirSayisi / $parcaSatirSayisi)
    $paddingLength = $dosyaSayisi.ToString().Length
    $fileNumberFormat = "D$paddingLength"

    Write-Host "`nToplam Satir Sayisi: $lastUsedRow"
    Write-Host "Toplam Sutun Sayisi: $lastUsedColumn"
    Write-Host "Olusturulacak Dosya Sayisi: $dosyaSayisi`n"

    for ($i = 0; $i -lt $dosyaSayisi; $i++) {
        $baslangicSatiri = ($i * $parcaSatirSayisi) + $originalDataStartRow
        $bitisSatiri = [math]::Min($baslangicSatiri + $parcaSatirSayisi - 1, $lastUsedRow)

        $yeniWorkbook = $Excel.Workbooks.Add()
        $yeniSheet = $yeniWorkbook.Sheets.Item(1)
        $currentPasteRow = 1

        if ($hasHeader) {
            $headerRange = $Sheet.Range($Sheet.Cells(1, 1), $Sheet.Cells(1, $lastUsedColumn))
            $headerRange.Copy() | Out-Null
            $yeniSheet.Paste($yeniSheet.Cells($currentPasteRow, 1)) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerRange) | Out-Null
            $currentPasteRow++
        }

        $range = $Sheet.Range($Sheet.Cells($baslangicSatiri, 1), $Sheet.Cells($bitisSatiri, $lastUsedColumn))
        $range.Copy() | Out-Null
        $yeniSheet.Paste($yeniSheet.Cells($currentPasteRow, 1)) | Out-Null

        $usedRange = $yeniSheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null

        $fileNumber = ($i + 1).ToString($fileNumberFormat)
        $yeniDosyaAdi = "${outputPrefix}${fileNumber}.xlsx"
        $savePath = Join-Path $outputDirectory $yeniDosyaAdi
        $yeniWorkbook.SaveAs($savePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
        $yeniWorkbook.Close($false)

        Write-Host " - '$yeniDosyaAdi' basariyla olusturuldu. (Satirlar: $baslangicSatiri - $bitisSatiri)" -ForegroundColor Green

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($yeniSheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($yeniWorkbook) | Out-Null
    }

    Write-Host "`nTum parcalama islemi tamamlandi!`n" -ForegroundColor Green
}

# --------------------------------------------
# 🚀 Excel COM nesnelerini temizle
# --------------------------------------------
function Cleanup-Excel {
    param (
        $Workbook,
        $Excel
    )
    if ($Workbook) {
        try { $Workbook.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch {}
    }
    if ($Excel) {
        try { $Excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Host "Excel kapatildi ve bellek temizlendi." -ForegroundColor DarkGray
}

# --------------------------------------------
# 🚀 Ana Program Fonksiyonu (Fonksiyonları çağırır)
# --------------------------------------------
function StartProgram {
    try {
        $excelData = Load-ExcelFile
        if (-not $excelData) { return }  # Dosya seçimi iptal edildi

        $Excel = $excelData.Excel
        $Workbook = $excelData.Workbook
        $outputDirectory = $excelData.OutputDir

        $Sheet = Select-Worksheet -Workbook $Workbook

        $lastUsedCell = $Sheet.Cells.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeLastCell)
        $lastUsedRow = $lastUsedCell.Row
        $lastUsedColumn = $lastUsedCell.Column
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($lastUsedCell) | Out-Null

        $hasHeader = Ask-HeaderPresence
        $parcaSatirSayisi = Ask-SplitRowCount
        $outputPrefix = Ask-OutputPrefix

        if (-not (Confirm-StartProcessing)) {
            return
        }

        Split-ExcelFile -Sheet $Sheet -lastUsedRow $lastUsedRow -lastUsedColumn $lastUsedColumn -hasHeader $hasHeader `
            -parcaSatirSayisi $parcaSatirSayisi -outputPrefix $outputPrefix -outputDirectory $outputDirectory -Excel $Excel

    } catch {
        if ($_.Exception -is [System.OperationCanceledException]) { throw }
        Write-Host "Hata: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        Cleanup-Excel -Workbook $Workbook -Excel $Excel
    }
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host '  ✂️ Excel Dosya Parçalama' -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencerede bir Excel dosyası seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçtiğiniz dosya belirttiğiniz satır sayısına göre bölünecektir." -ForegroundColor Yellow
    Write-Host "📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz." -ForegroundColor Yellow
    Write-Host ""
}

# --------------------------------------------
# 🚀 Ana Döngü
# --------------------------------------------
do {
    Clear-Host
    Show-Header
    $script:FileSelectionCancelled = $false
    try {
        StartProgram
    }
    catch [System.OperationCanceledException] {
        Write-Host '⚠️ İşlem iptal edildi.' -ForegroundColor Yellow
    }
    if ($script:FileSelectionCancelled) { Write-Host "⚠️ İşlem iptal edildi." -ForegroundColor Yellow }
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
