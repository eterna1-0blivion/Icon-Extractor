# author: eterna1_0blivion
$version = 'v0.2.8'

# Some variables for easy invocation
$theme = '$Host.UI.RawUI.BackgroundColor = "Black"; $Host.UI.RawUI.ForegroundColor = "Gray"; Clear-Host'
$exit = 'Read-Host -Prompt "Press Enter to exit"; Break'

# Set title, theme and display a greeting
$Host.UI.RawUI.WindowTitle = "Icons Extractor ($version)"
Invoke-Expression $theme
Write-Host "`nScript running..." -ForegroundColor White

# Add a C# code
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
public class IconExtractor
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);
    
    public static Icon Extract(string path, int index)
    {
        IntPtr[] largeIcon = new IntPtr[1];
        ExtractIconEx(path, index, largeIcon, null, 1);
        if (largeIcon[0] == IntPtr.Zero) return null;
        return (Icon)Icon.FromHandle(largeIcon[0]).Clone();
    }
}
"@ -Language CSharp -ReferencedAssemblies "System.Drawing.Common"

# Main function
function Get-Icons {
    param (
        [string]$filePath,
        [string]$outputFolder,
        [int]$maxIcons,
        [switch]$debugMode
    )

    # Check folder exist
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }
    
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
            elseif ($debugMode) { Write-Debug "No icon at index $i for $filePath" }
        }
        catch { Write-Warning "Error extracting icon $i from $filePath : $_" }
    }
    Write-Output "Extracted $extractedCount icons from $filePath"
}

# Settings
$sourcePath = "C:"
$sourceExtensions = @('dll', 'exe', 'mun')
$iconsLimit = 512
$debug = $False

# Process output folder
$baseOutputPath = "$PSScriptRoot\out"
if (Test-Path $baseOutputPath) {
    Get-ChildItem -Path $baseOutputPath -Recurse -File | Remove-Item -Force -ErrorAction SilentlyContinue
}
else {
    New-Item -ItemType Directory -Path $baseOutputPath | Out-Null
}

# Process file paths
$sourceFilePaths = New-Object System.Collections.Generic.List[string]
foreach ($extension in $sourceExtensions) {
    Write-Output "Scanning for `'.$extension`' files..."
    (Get-ChildItem -Path $sourcePath -Filter "*.$extension" -Recurse -Force -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName) | ForEach-Object { $sourceFilePaths.Add($_) }
}
Write-Output "Found $($sourceFilePaths.Count) files to process."

# Main
foreach ($path in $sourceFilePaths) {
    Write-Verbose "Processing: $path"
    $extension = [System.IO.Path]::GetExtension($path).TrimStart('.')
    $outputPath = "$baseOutputPath\$extension"
    Get-Icons -filePath $path -outputFolder $outputPath -maxIcons $iconsLimit -debugMode $debug
}


# Notification of successfully finished work
Write-Host "`nThe script completed successfully." -ForegroundColor Green
Invoke-Expression $exit
