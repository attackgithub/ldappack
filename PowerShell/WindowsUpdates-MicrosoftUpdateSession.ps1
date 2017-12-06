$entries = @(
    # "gdl-dc-01",
    "jac-dc-01",
    "allius-dc-01",
    "leon-dc-01",
    "unocorp-dc-01",
    "unocorp-dc-02"
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
        } ,title |
        Where-Object { $_.Date -gt $date } |
        Export-Csv -NoType "$Env:userprofile\Desktop\Windows Updates Logs\WindowsUpdates_$entry.csv"
}

# 1 - When is no data show message check server
# 2 - Check when is local execute local command