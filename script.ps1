# author: eterna1_0blivion
$version = 'v0.2.0'

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
        [int]$maxIcons = 512
    )

    if (!(Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }
    
    for ($i = 0; $i -lt $maxIcons; $i++) {
        try {
            $icon = [IconExtractor]::Extract($filePath, $i)
            if ($null -ne $icon) {
                $iconPath = Join-Path -Path $outputFolder -ChildPath "$(Split-Path -Leaf $filePath)_icon_$i.ico"
                $fileStream = [System.IO.File]::OpenWrite($iconPath)
                $icon.Save($fileStream)
                $fileStream.Close()
                Write-Output "Saved: $iconPath"
            }
        }
        catch {
            Write-Output "Failed to extract from: $filePath (Index: $i)"
        }
    }
}


# Custom settings
$sourcePath = "C:"
$sourceExtensions = @(
    'dll',
    'exe',
    'mun'
)

# Auto variables
$sourceFilePaths = New-Object System.Collections.Generic.List[string]
foreach ($extension in $sourceExtensions) {
    (Get-ChildItem -Path "$sourcePath" -Filter "*.$extension" -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName) | ForEach-Object { $sourceFilePaths.Add($_) }
}

# Use a function
foreach ($path in $sourceFilePaths) {
    $extension = $path -replace '^.*(?=.{3}$)'
    $outputPath = "$PSScriptRoot\out\$extension" ?? "$PSScriptRoot\out\other"
    Get-Icons -filePath $path -outputFolder $outputPath
}



# Notification of successfully finished work
Write-Host "`nThe script completed successfully." -ForegroundColor Green
#Invoke-Expression $exit
