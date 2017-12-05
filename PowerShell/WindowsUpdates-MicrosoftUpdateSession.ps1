$entries = @(
    "gdl-dc-01.ad.unosquare.com"
    # "GDLA-LT-170714A.ad.unosquare.com",
    # "gdl-vhost.ad.unosquare.com"
)
foreach($entry in $entries)
{
    $result = Invoke-Command -ComputerName $entry -ScriptBlock {
        $Session = New-Object -ComObject "Microsoft.Update.Session"
        $Searcher = $Session.CreateUpdateSearcher()
        $historyCount = $Searcher.GetTotalHistoryCount()
        
        $search = $Searcher.QueryHistory(0, $historyCount)
        $search
    }

    $date = Get-Date
    $date = $date.AddDays(-10)

    $result | Select-Object Date,
        @{
            name = "Operation"; 
            expression = {
                switch ($_.operation) {
                    1 {"Installation"};
                    2 {"Uninstallation"};
                    3 {"Other"}
                }
            }
        },   
        @{
            name = "Status"; 
            expression = {
                switch ($_.resultcode) {       
                    1 {"In Progress"}; 
                    2 {"Succeeded"}; 
                    3 {"Succeeded With Errors"};       
                    4 {"Failed"}; 
                    5 {"Aborted"}
                }
            }
        } |
        Where-Object { $_.Date -gt $date } |
        Export-Csv -NoType "$Env:userprofile\Desktop\WindowsUpdates_$entry.csv"
}