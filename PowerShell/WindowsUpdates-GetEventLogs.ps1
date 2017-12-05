$date = Get-Date
$date = $date.AddDays(-5)
$entries = @(
    "GDLA-LT-170714A"
    "gdl-dc-01",
    "gdl-vhost"
)

$Search = Get-EventLog -LogName System -Source "Microsoft-Windows-WindowsUpdateClient" -ComputerName $entries
$Search | where {$_.InstanceId -eq 19} | 
        Select-Object MachineName,
        @{ name = "Date"; expression = {$_.TimeGenerated} },
        @{ name = "Operation"; expression = { $_.Message.split(":")[0]}},
        @{ name = "Title"; expression = { $_.Message.split(":")[2]}} | 
        Where-Object { $_.Date -gt $date } |
        Export-Csv -NoType "$Env:userprofile\Desktop\WindowsUpdates.csv"