$ErrorActionPreference = 'Stop'

$taskNames = @(
    'HRDashboard-Monday-Prepare',
    'HRDashboard-Monday-Email'
)

foreach ($taskName in $taskNames) {
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "已移除排程：$taskName"
    }
}
