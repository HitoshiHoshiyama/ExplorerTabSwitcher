Param([Parameter(Mandatory=$false)][switch]$NoWait)

$IsExist = Get-ScheduledTask | Select-String ExplorerTabSwitcher
if(-not [string]::IsNullOrEmpty($IsExist))
{
    $Task = Get-ScheduledTask -TaskName ExplorerTabSwitcher
    if(($Task).State -eq 'Running')
    {
        write-Host タスク:ExplorerTabSwitcher を停止します。
        Stop-ScheduledTask -TaskName "ExplorerTabSwitcher"
    }
    write-Host タスク:ExplorerTabSwitcher を削除します。
    Unregister-ScheduledTask -TaskName "ExplorerTabSwitcher" -Confirm:$false
}else {
    write-Host タスク:ExplorerTabSwitcher は登録されていません。
}

if(-not $NoWait)
{
    $KeyIn = Read-Host "何かキーを押すと終了します。"
    Write-Host $KeyIn
}