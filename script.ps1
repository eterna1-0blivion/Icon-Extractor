# author: eterna1_0blivion
$version = 'v0.4.0'

# Некоторые переменные для облегчения работы 
$theme = '$Host.UI.RawUI.BackgroundColor = "Black"; $Host.UI.RawUI.ForegroundColor = "Gray"; Clear-Host'
$exit = 'Read-Host -Prompt "Press Enter to exit"; Break'

# Устанавливаем заголовок консоли, меняем тему и выводим первую строку
$Host.UI.RawUI.WindowTitle = "Icons Extractor ($version)"
Invoke-Expression $theme
Write-Host "`nScript running..." -ForegroundColor White


# Настройки
$sourcePath = "C:"
$sourceExtensions = @('dll', 'exe', 'mun')
$parallelThreads = 16
$iconsLimit = 512
$logLevel = "Debug" # ["Output"/"Verbose"/"Debug"]

# Подготовка выходных папок и лог-файла
$baseOutputPath = "$PSScriptRoot\out"
$logFile = "$PSScriptRoot\log.txt"

# Для правильной работы отладки
if ($logLevel -eq "Debug") { Set-PSDebug -Trace 0 } else { Set-PSDebug -Off }

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

# Параллельная обработка файлов
$sourceFilePaths | ForEach-Object -Parallel {
    # Определяем класс IconExtractor внутри блока
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
            DestroyIcon(largeIcon[0]);
            return icon;
        }
    }
"@ -Language CSharp -ReferencedAssemblies "System.Drawing.Common"

    # Определяем функцию Get-Icons внутри блока
    function Get-Icons {
        param (
            [string]$filePath,
            [string]$outputFolder,
            [int]$maxIcons,
            [ValidateSet("Output", "Verbose", "Debug")]
            [string]$logLevel,
            [string]$logFile
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
                if ($logLevel -eq "Debug") {
                    Write-Warning "Error checking index $i for $filePath : $_" |
                    Tee-Object -FilePath $logFile -Append
                }
            }
        }

        # Если иконок нет на первых трёх индексах, пропускаем файл
        if (-not $hasIcons) {
            Write-Output "Skipped: $filePath (no icons detected)" | Tee-Object -FilePath $logFile -Append
            return
        }

        # Полная обработка файла с динамической остановкой
        $extractedCount = 0
        $consecutiveNulls = 0
        for ($i = 0; $i -lt $maxIcons; $i++) {
            try {
                $icon = [IconExtractor]::Extract($filePath, $i)
                if ($null -ne $icon) {
                    $consecutiveNulls = 0
                    $iconPath = Join-Path -Path $outputFolder -ChildPath "$(Split-Path -Leaf $filePath)_icon_$i.ico"
                    try {
                        $fileStream = [System.IO.File]::OpenWrite($iconPath)
                        $icon.Save($fileStream)
                        $fileStream.Close()
                        $extractedCount++
                        Write-Output "Saved: $iconPath" | Tee-Object -FilePath $logFile -Append
                    }
                    catch {
                        Write-Warning "Failed to save icon $i for $filePath : $_" |
                        Tee-Object -FilePath $logFile -Append
                    }
                }
                else {
                    $consecutiveNulls++
                    if ($logLevel -eq "Debug") {
                        Write-Output "No icon at index $i for $filePath" | 
                        Tee-Object -FilePath $logFile -Append
                    }
                    if ($consecutiveNulls -ge 3) {
                        if ($logLevel -eq "Verbose" -or $logLevel -eq "Debug") {
                            Write-Output "Stopped at index $i for $filePath (3 consecutive nulls)" | 
                            Tee-Object -FilePath $logFile -Append
                        }
                        break
                    }
                }
            }
            catch {
                Write-Warning "Error extracting icon $i from $filePath : $_" | Tee-Object -FilePath $logFile -Append
                $consecutiveNulls++
                if ($consecutiveNulls -ge 3) {
                    if ($logLevel -eq "Verbose" -or $logLevel -eq "Debug") {
                        Write-Output "Stopped at index $i for $filePath (3 consecutive nulls after error)" | 
                        Tee-Object -FilePath $logFile -Append
                    }
                    break
                }
            }
        }
        Write-Output "Extracted $extractedCount icons from $filePath" | Tee-Object -FilePath $logFile -Append
    }

    # Передаём внешние переменные
    $logFile = $using:logFile
    $baseOutputPath = $using:baseOutputPath
    $iconsLimit = $using:iconsLimit
    $logLevel = $using:logLevel

    $path = $_
    $extension = [System.IO.Path]::GetExtension($path).TrimStart('.')
    $outputPath = "$baseOutputPath\$extension"
    Get-Icons -filePath $path -outputFolder $outputPath -maxIcons $iconsLimit -logLevel $logLevel -logFile $logFile
} -ThrottleLimit $parallelThreads


# Уведомляем пользователя о завершении работы (Нажми Enter)
Write-Host "`nThe script completed successfully." -ForegroundColor Green | Tee-Object -FilePath $logFile -Append
Invoke-Expression $exit
