$ErrorActionPreference = 'SilentlyContinue'
$exe = 'C:\文件\工作工具台\报告生成工具\dist\报告生成工具\报告生成工具.exe'
if (Test-Path $exe) {
    Start-Process -FilePath $exe -WorkingDirectory 'C:\文件\工作工具台\报告生成工具\dist\报告生成工具'
} else {
    Start-Process -FilePath 'C:\文件\工作工具台\报告生成工具\.venv\Scripts\python.exe' -ArgumentList 'app.py' -WorkingDirectory 'C:\文件\工作工具台\报告生成工具'
}
Start-Sleep -Seconds 10
$procs = Get-Process | Where-Object { $_.ProcessName -like '*报告生成工具*' -or $_.ProcessName -eq 'python' }
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
