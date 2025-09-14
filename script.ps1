# author: eterna1_0blivion
$version = 'v0.6.2-dev'
$version = 'v0.6.2-dev'

# Некоторые пред-установки
$theme = '$Host.UI.RawUI.BackgroundColor = "Black"; $Host.UI.RawUI.ForegroundColor = "Gray"; Clear-Host'
$exit = 'Read-Host -Prompt "Press Enter to exit"; Break'

# Устанавливаем заголовок консоли, меняем тему и выводим первую строку
$Host.UI.RawUI.WindowTitle = "Icons Extractor ($version)"
Invoke-Expression $theme
Write-Host "`nScript running..." -ForegroundColor White

# Настройки
$sourcePath = "C:"
$sourceExtensions = @('dll', 'exe', 'mun')
$iconsLimit = 512
$threads = 0 # [если 0 - автоопределение]
$logLevel = "Debug" # ["Output"/"Verbose"/"Debug"]
$baseOutputPath = "$PSScriptRoot\out"
$logFile = "$PSScriptRoot\log.txt"
$instanceName = "1.5a"  # Имя инстанса Everything (для альфы)

# Динамическое определение кол-ва параллельных потоков
# Если не указан -Threads, определяем по числу ядер CPU
if ($threads -eq 0) {
    $parallelThreads = [Math]::Max(1, [Environment]::ProcessorCount / 2)
    Write-Host "Auto-detected threads: $parallelThreads (based on CPU cores)" -ForegroundColor Yellow
} else {
    $parallelThreads = $threads
    Write-Host "Using user-specified threads: $parallelThreads" -ForegroundColor Yellow
}

# Очистка логов и выходных папок
if (Test-Path $logFile) { Remove-Item $logFile -Force }
if (Test-Path $baseOutputPath) {
    Get-ChildItem -Path $baseOutputPath -Recurse -File | Remove-Item -Force -ErrorAction SilentlyContinue
}

# Создаем concurrent структуры для параллельной работы
$logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$processed = [System.Collections.Concurrent.ConcurrentDictionary[string, bool]]::new()

# Определяем класс IconExtractor вне параллельного блока
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

# Определение нужных файлов (Параллельный поиск + интеграция Everything)
# Определение нужных файлов (Параллельный поиск + интеграция Everything)
$sourceFilePaths = New-Object System.Collections.Generic.List[string]

# Проверяем наличие и работоспособность Everything CLI с указанным инстансом
try {
    if (& es -instance $instanceName ext:exe -n 1) {
        $useEverything = $true
        Write-Output "Using Everything ($instanceName) for fast file search." | Tee-Object -FilePath $logFile -Append
    }
} catch {
    $useEverything = $false
    Write-Output "Everything CLI or instance not available, falling back to Get-ChildItem." | Tee-Object -FilePath $logFile -Append
}

if ($useEverything) {
    foreach ($extension in $sourceExtensions) {
        Write-Output "Scanning for '.$extension' files with Everything..." | Tee-Object -FilePath $logFile -Append
        # Поиск файлов с помощью Everything CLI
        & es -instance $instanceName -p $sourcePath ext:$extension | ForEach-Object { $sourceFilePaths.Add($_) }
    }
} else {
    foreach ($extension in $sourceExtensions) {
        Write-Output "Scanning for '.$extension' files..." | Tee-Object -FilePath $logFile -Append
        (Get-ChildItem -Path $sourcePath -Filter "*.$extension" -Recurse -Force -ErrorAction SilentlyContinue | 
        Select-Object -ExpandProperty FullName) | ForEach-Object { $sourceFilePaths.Add($_) }
    }
}

Write-Output "Found $($sourceFilePaths.Count) files to process." | Tee-Object -FilePath $logFile -Append

# Параллельная обработка файлов
$sourceFilePaths | ForEach-Object -Parallel {
    # Определяем глобальные переменные
    $script:logQueue = $using:logQueue
    $script:processed = $using:processed
    $script:baseOutputPath = $using:baseOutputPath
    $script:iconsLimit = $using:iconsLimit
    $script:logLevel = $using:logLevel

    # Функция для записи в лог (параллельная, с использованием очереди)
    function Write-Log {
        param (
            [string]$Message,
            [ValidateSet("Output", "Warning")]
            [string]$Type = "Output"
        )
        if ($Type -eq "Output") {
            Write-Output $Message; $script:logQueue.Enqueue($Message)
        }
        elseif ($Type -eq "Warning") {
            Write-Warning $Message; $script:logQueue.Enqueue("WARNING: $Message")
        }
    }

    # Определяем функцию Get-Icons внутри блока
    function Get-Icons {
        param (
            [string]$filePath
        )

        # Предотвращаем одновременную обработку одного и того же файла
        if (!$script:processed.TryAdd($filePath.ToLower(), $true)) {
            Write-Log -Message "Skipping already processed file: $filePath" -Type "Warning"
            return
        }

        # Создаём выходные папки, если их нет
        $outputFolder = "$script:baseOutputPath\$([System.IO.Path]::GetExtension($filePath).TrimStart('.'))"
        if (-not (Test-Path $outputFolder)) {
            New-Item -ItemType Directory -Path $outputFolder | Out-Null
        }

        # Полная обработка файла с динамической остановкой
        $extractedCount = 0
        $consecutiveNulls = 0
        for ($i = 0; $i -lt $script:iconsLimit; $i++) {
            try {
                # Пробуем вытащить иконку из файла
                $icon = [IconExtractor]::Extract($filePath, $i)
                if ($null -ne $icon) {
                    $consecutiveNulls = 0
                    $iconPath = Join-Path -Path $outputFolder -ChildPath "$(Split-Path -Leaf $filePath)_icon_$i.ico"
                    # Основной процесс параллельной обработки файлов
                    try {
                        $fileStream = [System.IO.File]::OpenWrite($iconPath)
                        $icon.Save($fileStream)
                        $fileStream.Close()
                        $extractedCount++
                        Write-Log -Message "Saved: $iconPath" -Type "Output"
                    }
                    catch {
                        Write-Log -Message "Failed to save icon $i for $filePath : $_" -Type "Warning"
                    }
                }
                # Накопление счётчика "не нахождения" иконок
                else {
                    $consecutiveNulls++
                    if ($script:logLevel -eq "Debug") {
                        Write-Log -Message "No icon at index $i for $filePath" -Type "Output"
                    }
                    # Если не находим иконок три раза подряд, прекращаем обработку файла
                    if ($consecutiveNulls -ge 3) {
                        if ($script:logLevel -eq "Verbose" -or $script:logLevel -eq "Debug") {
                            Write-Log -Message "Stopped at index $i for $filePath (3 consecutive nulls)" -Type "Output"
                        }
                        break
                    }
                }
            }
            # В случае трех ошибок в извлечении иконок, так же прекращаем обработку
            catch {
                Write-Log -Message "Error extracting icon $i from $filePath : $_" -Type "Warning"
                $consecutiveNulls++
                if ($consecutiveNulls -ge 3) {
                    if ($script:logLevel -eq "Verbose" -or $script:logLevel -eq "Debug") {
                        Write-Log -Message "Stopped at index $i for $filePath (3 consecutive nulls after error)" -Type "Output"
                    }
                    break
                }
            }
        }
        # При завершении обработки файла, сообщаем о кол-ве извлеченных иконок
        Write-Log -Message "Extracted $extractedCount icons from $filePath" -Type "Output"
    }

    $path = $_
    Get-Icons -filePath $path
} -ThrottleLimit $parallelThreads

# Сливаем логи из очереди в файл
$msg = $null
while ($logQueue.TryDequeue([ref]$msg)) {
    $msg | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Уведомляем пользователя о завершении работы (Нажми Enter)
Write-Host "`nThe script completed successfully." -ForegroundColor Green
Invoke-Expression $exit