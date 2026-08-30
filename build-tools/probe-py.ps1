$procs = Get-Process -Name py -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
if ($procs) {
    foreach ($p in $procs) {
        Write-Output ("PROC name={0} id={1} title={2} responding={3}" -f $p.ProcessName, $p.Id, $p.MainWindowTitle, $p.Responding)
    }
} else {
    Write-Output 'NO_WINDOWED_PY'
}
