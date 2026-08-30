$ErrorActionPreference = 'SilentlyContinue'
$exe = 'C:\文件\工作工具台\报告生成工具\dist\报告生成工具\报告生成工具.exe'
Start-Process -FilePath $exe
Start-Sleep -Seconds 8
$procs = Get-Process | Where-Object { $_.ProcessName -like '*报告生成工具*' }
if ($procs) {
    $procs | ForEach-Object { Write-Output ("PROC name={0} id={1} title={2} responding={3}" -f $_.ProcessName, $_.Id, $_.MainWindowTitle, $_.Responding) }
} else {
    Write-Output 'NO_PROCESS'
}
