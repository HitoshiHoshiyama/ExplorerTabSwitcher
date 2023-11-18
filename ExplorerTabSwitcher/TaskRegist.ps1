Param([Parameter(Mandatory=$false)][string]$Exedir = $PSScriptRoot)

$ExePath = Convert-Path (Join-Path $Exedir "ExplorerTabSwitcher.exe")
try {
    $TaskAction = New-ScheduledTaskAction -Id "ExplorerTabSwitcher" -Execute $ExePath
}
catch {
    Write-Host エラー発生のため、タスクを登録しませんでした。
    exit
}
$TaskTrigger = New-ScheduledTaskTrigger -AtLogOn
$TaskOptions = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -Hidden

Write-Host タスク:ExplorerTabSwitcher の実行ファイル[$ExePath]

$IsExist = Get-ScheduledTask | Select-String ExplorerTabSwitcher
if(-not [string]::IsNullOrEmpty($IsExist))
{
    . .\TaskRemove.ps1 -NoWait
}

Register-ScheduledTask -TaskName "ExplorerTabSwitcher" -Action $TaskAction -Trigger $TaskTrigger -Settings $TaskOptions -RunLevel "Highest" -Description "エクスプローラのタブをひとつのウィンドウにまとめます。"
Write-Host タスク:ExplorerTabSwitcher を登録しました。

Start-ScheduledTask -TaskName "ExplorerTabSwitcher"
Write-Host タスク:ExplorerTabSwitcher を開始しました。

Get-ScheduledTask -TaskName "ExplorerTabSwitcher"

$KeyIn = Read-Host "何かキーを押すと終了します。"
Write-Host $KeyIn