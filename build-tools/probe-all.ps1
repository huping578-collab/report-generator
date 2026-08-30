$procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('python', 'pythonw', '报告生成工具') }
if ($procs) {
    foreach ($p in $procs) {
        Write-Output ("PROC name={0} id={1} hwnd={2} title={3}" -f $p.ProcessName, $p.Id, $p.MainWindowHandle, $p.MainWindowTitle)
    }
} else {
    Write-Output 'NO_PROCS'
}
