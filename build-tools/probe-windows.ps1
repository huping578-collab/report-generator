$ErrorActionPreference = 'SilentlyContinue'
$procs = Get-Process | Where-Object { $_.ProcessName -like '*报告生成工具*' -or $_.ProcessName -like '*python*' }
$found = $false
foreach ($p in $procs) {
    if ($p.MainWindowHandle -ne 0) {
        $found = $true
        Write-Output ("PROC name={0} id={1} hwnd={2} responding={3}" -f $p.ProcessName, $p.Id, $p.MainWindowHandle, $p.Responding)
    }
}
if (-not $found) {
    Write-Output 'NO_WINDOWED'
}
