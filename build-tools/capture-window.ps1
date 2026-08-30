Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinCap {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
Add-Type -AssemblyName System.Drawing

$hwnd = 0
$procs = Get-Process -Name '报告生成工具' -ErrorAction SilentlyContinue
foreach ($p in $procs) {
    if ($p.MainWindowHandle -ne 0) { $hwnd = $p.MainWindowHandle }
}
if ($hwnd -eq 0) { Write-Output 'NO_HWND'; exit }

$rect = New-Object WinCap+RECT
[WinCap]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
$w = $rect.Right - $rect.Left
$h = $rect.Bottom - $rect.Top
$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$dc = $g.GetHdc()
[WinCap]::PrintWindow($hwnd, $dc, 2) | Out-Null
$g.ReleaseHdc($dc)
$g.Dispose()
$out = 'C:\文件\工作工具台\报告生成工具\artifacts\screenshots\exe-window.png'
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Output ("SAVED {0} {1}x{2}" -f $out, $w, $h)
