$ErrorActionPreference = 'SilentlyContinue'
$procs = Get-Process -Name 'python','pythonw','报告生成工具' -ErrorAction SilentlyContinue
if ($procs) {
    foreach ($p in $procs) {
        Write-Output ("PROC name={0} id={1} hwnd={2} responding={3}" -f $p.ProcessName, $p.Id, $p.MainWindowHandle, $p.Responding)
    }
} else {
    Write-Output 'NO_PROCS'
}
