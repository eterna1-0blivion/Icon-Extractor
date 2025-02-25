# author: eterna1_0blivion
$version = 'v0.1.0'

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
    
    public static Icon Extract(string file, int index)
    {
        IntPtr[] largeIcon = new IntPtr[1];
        ExtractIconEx(file, index, largeIcon, null, 1);
        if (largeIcon[0] == IntPtr.Zero) return null;
        return (Icon)Icon.FromHandle(largeIcon[0]).Clone();
    }
}
"@ -Language CSharp -ReferencedAssemblies "System.Drawing.Common"

# Some variables for easy invocation
$theme = '$Host.UI.RawUI.BackgroundColor = "Black"; $Host.UI.RawUI.ForegroundColor = "Gray"; Clear-Host'
$exit = 'Read-Host -Prompt "Press Enter to exit"; Break'
$debug = "$True"

# Set title, theme and display a greeting
$Host.UI.RawUI.WindowTitle = "Icons Extractor ($version)"
Invoke-Expression $theme
Write-Host "`nScript running..." -ForegroundColor White


# Custom settings
$sourcePath = "C:"
$sourceExtensions = "exe"

# Auto variables
$sourceFiles = Get-ChildItem -Path "$sourcePath" -Filter "*.$sourceExtensions" -Recurse -Force | Select-Object -ExpandProperty FullName
$targetFolder = "$PSScriptRoot\out\$sourceExtensions"

# Debug mode
foreach ($file in $sourceFiles) {
    if ($debug) { Write-Host "$file" }
}

# Main function
function extractIcons {
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

# Use a function
foreach ($file in $sourceFiles) {
    extractIcons -filePath $file -outputFolder $targetFolder
}
