$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
                # Only call this for COM references owned by this script and no longer used.
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
            }
        }
        catch {
            # Cleanup should continue even if one COM reference fails to release.
        }
    }
}

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

function Write-ModuleStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ConsoleColor]$Color = 'Cyan'
    )

    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host "  $Message" -ForegroundColor $Color
    Write-Host ("-" * 72) -ForegroundColor DarkGray
}

function Read-ConfirmationInput {
    param([Parameter(Mandatory)][string]$Prompt)

    Write-Host $Prompt -NoNewline
    $buffer = [System.Text.StringBuilder]::new()
    while ($true) {
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            Enter {
                [System.Console]::WriteLine()
                return $buffer.ToString()
            }
            Escape {
                [System.Console]::WriteLine()
                return $null
            }
            Backspace {
                if ($buffer.Length -gt 0) {
                    [void]$buffer.Remove($buffer.Length - 1, 1)
                    [System.Console]::Write("`b `b")
                }
            }
            default {
                if (-not [char]::IsControl($key.KeyChar)) {
                    [void]$buffer.Append($key.KeyChar)
                    [System.Console]::Write($key.KeyChar)
                }
            }
        }
    }
}

function Show-ModuleHeader {
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host '  🖨️ Excel Puantaj Yazdırma' -ForegroundColor Magenta
    Write-Host ("-" * 72) -ForegroundColor Magenta
    Write-Host ""
    Write-Host '📌 Açılan pencerede Excel dosyalarının bulunduğu klasörü seçin.' -ForegroundColor Yellow
    Write-Host '📌 Uygun dosyalar varsayılan yazıcıya gönderilir.' -ForegroundColor Yellow
    Write-Host '📌 Veri girişi sırasında [Esc] ile işlemi iptal edip ana menüye dönebilirsiniz.' -ForegroundColor Yellow
    Write-Host ""
}

function Show-ModuleExitChoice {
    Write-Host ""
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host '  🔚 [Q/Esc] Ana menüye dön  |  [D] Yeniden çalıştır' -ForegroundColor Cyan
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host ""
    do {
        $key = [System.Console]::ReadKey($true).Key
    } while ($key -notin @([ConsoleKey]::Q, [ConsoleKey]::D, [ConsoleKey]::Escape))
    return $key
}

function Invoke-ExcelPuantajPrint {

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = 'Excel dosyalarının bulunduğu klasörü seçin'
$folderDialog.ShowNewFolderButton = $false

try {
    $dialogResult = $folderDialog.ShowDialog()
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host '⚠️ İşlem iptal edildi.' -ForegroundColor Yellow
        return
    }
    $folder = $folderDialog.SelectedPath
}
finally {
    $folderDialog.Dispose()
}

if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
    Write-Host "Klasör bulunamadı: $folder" -ForegroundColor Red
    Write-ModuleStatus 'İşlem tamamlanamadı. Ana menüye dönülüyor.' 'Yellow'
    return
}

# Windows Explorer-style natural filename sorting (e.g. File2 before File10).
# StrCmpLogicalW is the Windows shell's logical string comparison function.
if (-not ('ExplorerFileNameComparer' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public sealed class ExplorerFileNameComparer : IComparer<FileInfo>
{
    [DllImport("Shlwapi.dll", CharSet = CharSet.Unicode)]
    private static extern int StrCmpLogicalW(string psz1, string psz2);

    public int Compare(FileInfo x, FileInfo y)
    {
        if (Object.ReferenceEquals(x, y)) return 0;
        if (x == null) return -1;
        if (y == null) return 1;
        int result = StrCmpLogicalW(x.Name, y.Name);
        if (result != 0) return result;
        return StringComparer.OrdinalIgnoreCase.Compare(x.Name, y.Name);
    }
}
'@
}

$files = @(Get-ChildItem -LiteralPath $folder -File |
    Where-Object {
        $_.Extension -in @('.xls', '.xlsx', '.xlsm', '.xlsb') -and
        $_.Name -notlike '~$*'
    })

if ($files.Count -gt 1) {
    [Array]::Sort([System.IO.FileInfo[]]$files, [ExplorerFileNameComparer]::new())
}

if ($files.Count -eq 0) {
    Write-Host 'Bu klasörde Excel dosyası bulunamadı.' -ForegroundColor Yellow
    Write-ModuleStatus 'İşlem tamamlandı. Ana menüye dönülüyor.'
    return
}

Write-Host "Bulunan Excel dosyası: $($files.Count)"
$confirmation = Read-ConfirmationInput 'Bu dosyalar varsayılan yazıcıya gönderilsin mi? (E/H) '
if ($confirmation -notmatch '^(E|EVET)$') {
    Write-Host '⚠️ İşlem iptal edildi.' -ForegroundColor Yellow
    return
}

$excel = $null
$workbooks = $null
$printed = 0
$failed = New-Object 'System.Collections.Generic.List[string]'
$fatalError = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $workbooks = $excel.Workbooks

    $total = $files.Count
    $index = 0

    foreach ($file in $files) {
        $index++
        $workbook = $null
        $worksheets = $null

        Write-Host "[$index/$total] Isleniyor: $($file.Name)"

        try {
            # Force-disable VBA macros immediately before opening; do not update external links.
            $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
            try {
                $workbook = $workbooks.Open($file.FullName, 0, $true) # UpdateLinks=0, ReadOnly=$true
            }
            finally {
                # Restore default automation security; it is set to ForceDisable again before each open.
                try { $excel.AutomationSecurity = 1 } catch {} # msoAutomationSecurityLow
            }

            $worksheets = $workbook.Worksheets
            $sheetCount = $worksheets.Count

            for ($i = 1; $i -le $sheetCount; $i++) {
                $worksheet = $null
                $pageSetup = $null
                try {
                    $worksheet = $worksheets.Item($i)

                    # xlSheetVisible = -1. Hidden and very-hidden sheets are left alone.
                    if ($worksheet.Visible -eq -1) {
                        $pageSetup = $worksheet.PageSetup
                        $pageSetup.Zoom = $false
                        $pageSetup.FitToPagesWide = 1
                        $pageSetup.FitToPagesTall = 1
                    }
                }
                finally {
                    Release-ComObject $pageSetup
                    Release-ComObject $worksheet
                    $pageSetup = $null
                    $worksheet = $null
                }
            }

            # Sends the workbook's print job to the Windows default printer.
            $workbook.PrintOut()
            $printed++
            Write-Host '  Yazdirma istegi Excel tarafindan gonderildi.' -ForegroundColor Green
        }
        catch {
            $message = '{0}: {1}' -f $file.Name, $_.Exception.Message
            [void]$failed.Add($message)
            Write-Host "  HATA: $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch {}
            }
            Release-ComObject $worksheets
            Release-ComObject $workbook
            $worksheets = $null
            $workbook = $null
        }
    }
}
catch {
    $fatalError = $_.Exception.Message
    Write-Host "KRITIK HATA: Excel baslatilamadi veya genel islem hatasi olustu: $fatalError" -ForegroundColor Red
}
finally {
    # Release child COM references before quitting the dedicated Excel instance.
    Release-ComObject $workbooks
    $workbooks = $null

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
        $excel = $null
    }

    # Finalize any remaining RCWs after explicit COM cleanup.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ''
if ($null -ne $fatalError) {
    Write-Host 'İşlem tamamlanamadı. Yukarıdaki kritik hatayı kontrol edin.' -ForegroundColor Red
    Write-ModuleStatus 'İşlem tamamlanamadı. Ana menüye dönülüyor.' 'Yellow'
    return
}

Write-Host "Excel tarafından yazdırma isteği gönderilen dosya: $printed"
Write-Host "Hata alan dosya: $($failed.Count)"

if ($failed.Count -gt 0) {
    Write-Host ''
    Write-Host 'Hata listesi:' -ForegroundColor Yellow
    foreach ($item in $failed) { Write-Host " - $item" }
}

Write-ModuleStatus 'İşlem sona erdi. Ana menüye dönülüyor.'
}

do {
    Clear-Host
    Show-ModuleHeader
    Invoke-ExcelPuantajPrint
    $next = Show-ModuleExitChoice
} while ($next -eq [ConsoleKey]::D)

