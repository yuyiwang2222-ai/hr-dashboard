param(
    [string]$PythonPath = "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
    [string]$ProjectRoot = "C:\Users\chiehyi\OneDrive - 大豐環保科技股份有限公司\文件\Claude\Agents-人力變化"
)

$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $ProjectRoot 'scripts\run_scheduled_monday.py'
if (-not (Test-Path $scriptPath)) {
    throw "找不到腳本: $scriptPath"
}

if (-not (Test-Path $PythonPath)) {
    throw "找不到 Python: $PythonPath"
}

$prepareAction = New-ScheduledTaskAction -Execute $PythonPath -Argument "`"$scriptPath`" --phase prepare"
$emailAction = New-ScheduledTaskAction -Execute $PythonPath -Argument "`"$scriptPath`" --phase email"

$prepareTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 9:00AM
$emailTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 9:30AM

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName 'HRDashboard-Monday-Prepare' -Action $prepareAction -Trigger $prepareTrigger -Settings $settings -Description 'Weekly HR dashboard prepare phase' -Force | Out-Null
Register-ScheduledTask -TaskName 'HRDashboard-Monday-Email' -Action $emailAction -Trigger $emailTrigger -Settings $settings -Description 'Weekly HR dashboard email phase' -Force | Out-Null

Write-Host '已建立排程：HRDashboard-Monday-Prepare (每週一 09:00)'
Write-Host '已建立排程：HRDashboard-Monday-Email (每週一 09:30)'
