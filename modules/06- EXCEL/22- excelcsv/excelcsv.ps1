# --------------------------------------------
# 🚀  Bilgilendirici Renkli Mesaj Yazdırma
# --------------------------------------------
function Yaz-Mesaj {
    param (
        [string]$Mesaj,
        [string]$Renk = "White",
        [string]$Ikon = "ℹ️"
    )
    Write-Host "$Ikon  $Mesaj" -ForegroundColor $Renk
}

# --------------------------------------------
# 🚀  Program Döngüsü
# --------------------------------------------
function StartProgram {
    # Gerekli Assembly'yi yükle (sadece bir kez)
    Add-Type -AssemblyName System.Windows.Forms

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*"
    $OpenFileDialog.Title = "Bir veya birden fazla Excel dosyası seçin"
    $OpenFileDialog.Multiselect = $true

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $SecilenDosyalar = $OpenFileDialog.FileNames

        # Excel uygulamasını başlat
        $Excel = $null
        $failedCount = 0
        try {
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Excel.DisplayAlerts = $false

            foreach ($DosyaYolu in $SecilenDosyalar) {
                $Workbook = $null
                $Worksheet = $null
                try {
                    $Workbook = $Excel.Workbooks.Open($DosyaYolu)
                    $Worksheet = $Workbook.Worksheets.Item(1)

                    $CsvYolu = Join-Path (Split-Path $DosyaYolu) ("$([IO.Path]::GetFileNameWithoutExtension($DosyaYolu))_$($Worksheet.Name).csv")
                    $Workbook.SaveAs($CsvYolu, 6)  # 6 = xlCSV

                    Yaz-Mesaj "Dönüştürüldü: $CsvYolu" "Green" "✅"
                }
                catch {
                    $failedCount++
                    Yaz-Mesaj "❌ Hata oluştu: $DosyaYolu" "Red" "❌"
                    Yaz-Mesaj $_.Exception.Message "DarkRed" "⚠️"
                }
                finally {
                    if ($Worksheet) {
                        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null } catch {}
                    }
                    if ($Workbook) {
                        try { $Workbook.Close($false) } catch {}
                        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch {}
                    }
                    $Worksheet = $null
                    $Workbook = $null
                }
            }

            if ($failedCount -eq 0) {
                Yaz-Mesaj "✅ Tüm dosyalar başarıyla dönüştürüldü!" "Cyan" "🏁"
            }
            else {
                Yaz-Mesaj "$failedCount dosya dönüştürülemedi; hata mesajlarını kontrol edin." "Yellow" "⚠️"
            }
        }
        finally {
            if ($Excel) {
                try { $Excel.Quit() } catch {}
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch {}
                $Excel = $null
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }
    }
    else {
        Yaz-Mesaj "Dosya seçilmedi." "Yellow" "⚠️"
        return $false
    }

    return $true
}

# --------------------------------------------
# 🚀  Başlık Yazdırma Fonksiyonu
# --------------------------------------------
function Show-Header {
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host "  🔨 Excel → CSV Dönüştürme Aracı" -ForegroundColor Cyan
    Write-Host ("-" * 52) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "📌 Açılan pencere ile bir veya daha fazla Excel dosyası seçin." -ForegroundColor Yellow
    Write-Host "📌 Seçilen dosyalar .CSV formatına dönüştürülür." -ForegroundColor Yellow
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
