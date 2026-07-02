[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ToolSlug,

    [ValidateSet('Compact', 'LargeUI', 'Step1', 'Step2', 'Step3', 'Panel', 'PanelRibbon')]
    [string]$Variant = 'Compact',

    [string]$WindowTitleContains,

    [string]$OutputDirectory = 'MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

if (-not ('MhtToolUiCapture.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MhtToolUiCapture {
    public static class NativeMethods {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);
    }
}
"@
}

function Get-WindowTitle {
    param([System.IntPtr]$Handle)

    $buffer = New-Object System.Text.StringBuilder 512
    [void][MhtToolUiCapture.NativeMethods]::GetWindowText($Handle, $buffer, $buffer.Capacity)
    return $buffer.ToString()
}

function Find-WindowHandle {
    param([string]$TitleContains)

    $needle = if ($null -eq $TitleContains) { '' } else { $TitleContains.Trim().ToLowerInvariant() }
    if ([string]::IsNullOrWhiteSpace($needle)) {
        return [System.IntPtr]::Zero
    }

    $match = [System.IntPtr]::Zero
    $callback = [MhtToolUiCapture.NativeMethods+EnumWindowsProc]{
        param($hWnd, $lParam)

        if (-not [MhtToolUiCapture.NativeMethods]::IsWindowVisible($hWnd)) {
            return $true
        }

        $title = Get-WindowTitle -Handle $hWnd
        if (-not [string]::IsNullOrWhiteSpace($title) -and $title.ToLowerInvariant().Contains($needle)) {
            $script:match = $hWnd
            return $false
        }

        return $true
    }

    [void][MhtToolUiCapture.NativeMethods]::EnumWindows($callback, [System.IntPtr]::Zero)
    return $script:match
}

function Get-OutputFileName {
    param(
        [string]$Slug,
        [string]$Kind
    )

    switch ($Kind) {
        'Compact' { return 'ui-{0}.png' -f $Slug }
        'LargeUI' { return 'ui-{0}-LargeUI.png' -f $Slug }
        'Step1' { return 'ui-{0}-step1.png' -f $Slug }
        'Step2' { return 'ui-{0}-step2.png' -f $Slug }
        'Step3' { return 'ui-{0}-step3.png' -f $Slug }
        'Panel' { return 'panel-{0}.png' -f $Slug }
        'PanelRibbon' { return 'panel-{0}.png' -f $Slug }
        default { throw 'Unsupported screenshot variant.' }
    }
}

$windowHandle = if ([string]::IsNullOrWhiteSpace($WindowTitleContains)) {
    [MhtToolUiCapture.NativeMethods]::GetForegroundWindow()
} else {
    Find-WindowHandle -TitleContains $WindowTitleContains
}

if ($windowHandle -eq [System.IntPtr]::Zero) {
    throw 'No visible window matched the requested capture target.'
}

$rect = New-Object MhtToolUiCapture.NativeMethods+RECT
if (-not [MhtToolUiCapture.NativeMethods]::GetWindowRect($windowHandle, [ref]$rect)) {
    throw 'Could not read the target window bounds.'
}

$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top
if ($width -le 0 -or $height -le 0) {
    throw 'The target window has invalid bounds for screenshot capture.'
}

$outputRoot = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path (Get-Location) $OutputDirectory
}

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
$outputPath = Join-Path $outputRoot (Get-OutputFileName -Slug $ToolSlug -Kind $Variant)

$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

try {
    $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
    $bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
}
finally {
    $graphics.Dispose()
    $bitmap.Dispose()
}

Write-Host ('Saved screenshot: {0}' -f $outputPath)
