$ErrorActionPreference = 'SilentlyContinue'
$procs = Get-Process -Name '报告生成工具' -ErrorAction SilentlyContinue
if ($procs) {
    foreach ($p in $procs) {
        Write-Output ("PROC id={0} title={1} responding={2} hwnd={3}" -f $p.Id, $p.MainWindowTitle, $p.Responding, $p.MainWindowHandle)
    }
} else {
    Write-Output 'NO_PROCESS'
}
