# author: eterna1_0blivion
$version = 'v0.4.6'

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
$baseOutputPath = "$PSScriptRoot\out"
$logFile = "$PSScriptRoot\log.txt"
$lockObject = [Object]::new()


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
"@ -Language CSharp -ReferencedAssemblies "System.Drawing.Common" -ErrorAction SilentlyContinue

    # Функция для синхронизированной записи в лог и консоль
    function Write-Log {
        param (
            [string]$Message,
            [ValidateSet("Output", "Warning")]
            [string]$Type = "Output",
            [string]$LogFile,
            [object]$LockObject
        )
        [System.Threading.Monitor]::Enter($LockObject)
        try {
            if ($Type -eq "Output") {
                Write-Output $Message | Tee-Object -FilePath $LogFile -Append
            }
            elseif ($Type -eq "Warning") {
                Write-Output "WARNING: $Message" -ForegroundColor Yellow | Tee-Object -FilePath $LogFile -Append
            }
        }
        finally {
            [System.Threading.Monitor]::Exit($LockObject)
        }
    }

    # Определяем функцию Get-Icons внутри блока
    function Get-Icons {
        param (
            [string]$filePath,
            [string]$outputFolder,
            [int]$maxIcons,
            [ValidateSet("Output", "Verbose", "Debug")]
            [string]$logLevel,
            [string]$logFile,
            [object]$lockObject
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
                    Write-Log -Message "Error checking index $i for $filePath : $_" -Type "Warning" -LogFile $logFile -LockObject $lockObject
                }
            }
        }

        # Если иконок нет на первых трёх индексах, пропускаем файл
        if (-not $hasIcons) {
            Write-Log -Message "Skipped: $filePath (no icons detected in first 3 indexes)" -Type "Output" -LogFile $logFile -LockObject $lockObject
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
                        Write-Log -Message "Saved: $iconPath" -Type "Output" -LogFile $logFile -LockObject $lockObject
                    }
                    catch {
                        Write-Log -Message "Failed to save icon $i for $filePath : $_" -Type "Warning" -LogFile $logFile -LockObject $lockObject
                    }
                }
                else {
                    $consecutiveNulls++
                    if ($logLevel -eq "Debug") {
                        Write-Log -Message "No icon at index $i for $filePath" -Type "Output" -LogFile $logFile -LockObject $lockObject
                    }
                    if ($consecutiveNulls -ge 3) {
                        if ($logLevel -eq "Verbose" -or $logLevel -eq "Debug") {
                            Write-Log -Message "Stopped at index $i for $filePath (3 consecutive nulls)" -Type "Output" -LogFile $logFile -LockObject $lockObject
                        }
                        break
                    }
                }
            }
            catch {
                Write-Log -Message "Error extracting icon $i from $filePath : $_" -Type "Warning" -LogFile $logFile -LockObject $lockObject
                $consecutiveNulls++
                if ($consecutiveNulls -ge 3) {
                    if ($logLevel -eq "Verbose" -or $logLevel -eq "Debug") {
                        Write-Log -Message "Stopped at index $i for $filePath (3 consecutive nulls after error)" -Type "Output" -LogFile $logFile -LockObject $lockObject
                    }
                    break
                }
            }
        }
        Write-Log -Message "Extracted $extractedCount icons from $filePath" -Type "Output" -LogFile $logFile -LockObject $lockObject
    }

    # Передаём внешние переменные
    $logFile = $using:logFile
    $baseOutputPath = $using:baseOutputPath
    $iconsLimit = $using:iconsLimit
    $logLevel = $using:logLevel
    $lockObject = $using:lockObject

    $path = $_
    $extension = [System.IO.Path]::GetExtension($path).TrimStart('.')
    $outputPath = "$baseOutputPath\$extension"
    Get-Icons -filePath $path -outputFolder $outputPath -maxIcons $iconsLimit -logLevel $logLevel -logFile $logFile -lockObject $lockObject
} -ThrottleLimit $parallelThreads


# Уведомляем пользователя о завершении работы (Нажми Enter)
Write-Host "`nThe script completed successfully." -ForegroundColor Green | Tee-Object -FilePath $logFile -Append
Invoke-Expression $exit
