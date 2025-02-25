# author: eterna1_0blivion
$version = 'v0.3.0'

# Некоторые переменные для облегчения работы 
$theme = '$Host.UI.RawUI.BackgroundColor = "Black"; $Host.UI.RawUI.ForegroundColor = "Gray"; Clear-Host'
$exit = 'Read-Host -Prompt "Press Enter to exit"; Break'

# Устанавливаем заголовок консоли, меняем тему и выводим первую строку
$Host.UI.RawUI.WindowTitle = "Icons Extractor ($version)"
Invoke-Expression $theme
Write-Host "`nScript running..." -ForegroundColor White

# Вызов ExtractIconEx при помощи C#
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
public class IconExtractor
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);

    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);
    
    public static Icon Extract(string path, int index)
    {
        IntPtr[] largeIcon = new IntPtr[1];
        ExtractIconEx(path, index, largeIcon, null, 1);
        if (largeIcon[0] == IntPtr.Zero) return null;
        Icon icon = (Icon)Icon.FromHandle(largeIcon[0]).Clone();
        DestroyIcon(largeIcon[0]); // Освобождаем handle
        return icon;
    }
}
"@ -Language CSharp -ReferencedAssemblies "System.Drawing.Common"

# Основаня функция
function Get-Icons {
    param (
        [string]$filePath,
        [string]$outputFolder,
        [int]$maxIcons,
        [ValidateSet("Output", "Verbose", "Debug")]
        [string]$logLevel
    )

    # Создаём выходные папки, если их нет
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }

    # Предварительная проверка первых трёх индексов
    $hasIcons = $false
    for ($i = 0; $i -lt [math]::Min(3, $maxIcons); $i++) {
        try {
            $icon = [IconExtractor]::Extract($filePath, $i)
            if ($null -ne $icon) {
                $hasIcons = $true
                break
            }
        }
        catch {
            if ($logLevel -eq "Debug") { Write-Debug "Error checking index $i for $filePath : $_" }
        }
    }

    # Если иконок нет на первых трёх индексах, пропускаем файл
    if (-not $hasIcons) {
        Write-Output "Skipped: $filePath (no icons detected)"
        return
    }

    # Полная обработка файла
    $extractedCount = 0
    for ($i = 0; $i -lt $maxIcons; $i++) {
        try {
            $icon = [IconExtractor]::Extract($filePath, $i)
            if ($null -ne $icon) {
                $iconPath = Join-Path -Path $outputFolder -ChildPath "$(Split-Path -Leaf $filePath)_icon_$i.ico"
                try {
                    $fileStream = [System.IO.File]::OpenWrite($iconPath)
                    $icon.Save($fileStream)
                    $fileStream.Close()
                    $extractedCount++
                    Write-Output "Saved: $iconPath"
                }
                catch { Write-Warning "Failed to save icon $i for $filePath : $_" }
            }
            elseif ($logLevel -eq "Debug") { Write-Debug "No icon at index $i for $filePath" }
        }
        catch { Write-Warning "Error extracting icon $i from $filePath : $_" }
    }
    Write-Output "Extracted $extractedCount icons from $filePath"
}

# Настройки
$sourcePath = "C:"
$sourceExtensions = @('dll', 'exe', 'mun')
$iconsLimit = 512
$logLevel = "Debug" # ["Output"/"Verbose"/"Debug"]

# Для правильной работы отладки
if ($logLevel = "Debug") { Set-PSDebug -Trace 0 } else { Set-PSDebug -Off }

# Подготовка выходных папок и лог-файла
$baseOutputPath = "$PSScriptRoot\out"
$logFile = "$PSScriptRoot\log.txt"

# Очистка логов и выходных папок
if (Test-Path $logFile) { Remove-Item $logFile -Force }
if (Test-Path $baseOutputPath) {
    Get-ChildItem -Path $baseOutputPath -Recurse -File | Remove-Item -Force -ErrorAction SilentlyContinue
}

# Определение нужных файлов
$sourceFilePaths = New-Object System.Collections.Generic.List[string]
foreach ($extension in $sourceExtensions) {
    Write-Output "Scanning for `'.$extension`' files..." | Tee-Object -FilePath $logFile -Append
    (Get-ChildItem -Path $sourcePath -Filter "*.$extension" -Recurse -Force -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName) | ForEach-Object { $sourceFilePaths.Add($_) }
}
Write-Output "Found $($sourceFilePaths.Count) files to process." | Tee-Object -FilePath $logFile -Append

# Обработка файлов
foreach ($path in $sourceFilePaths) {
    if ($logLevel -ne "Output") { Write-Output "Processing: $path" | Tee-Object -FilePath $logFile -Append }
    $extension = [System.IO.Path]::GetExtension($path).TrimStart('.')
    $outputPath = "$baseOutputPath\$extension"
    Get-Icons -filePath $path -outputFolder $outputPath -maxIcons $iconsLimit -logLevel $logLevel
}


# Notification of successfully finished work
Write-Host "`nThe script completed successfully." -ForegroundColor Green
Invoke-Expression $exit
