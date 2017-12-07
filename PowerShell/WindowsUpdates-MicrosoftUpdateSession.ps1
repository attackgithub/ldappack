$entries = @(
    "gdl-dc-01",
    "jac-dc-01",
    "allius-dc-01",
    "leon-dc-01",
    "unocorp-dc-01",
    "unocorp-dc-02"
)
$sc = { 
    $Session = New-Object -ComObject "Microsoft.Update.Session"
    $Searcher = $Session.CreateUpdateSearcher()
    $historyCount = $Searcher.GetTotalHistoryCount()

    $search = $Searcher.QueryHistory(0, $historyCount)
    $search 
}
$hostname = hostname

foreach($entry in $entries)
{
    $remote = $entry -notmatch $hostname

    if($remote) {
        #If remote computer
        $result = Invoke-Command -ComputerName $entry -ScriptBlock $sc
    } else { 
        #Localhost
        Get-Date
        $result = Invoke-Command -ScriptBlock $sc
    }

    $date = Get-Date
    $date = $date.AddDays(-10)

    $output = $result | Select-Object Date,
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
        Where-Object { $_.Date -gt $date } 

    if($output -eq $null)
    {
        $obj = @{ 'Message' = "There is no updates information. Please check.";}
        $output = New-Object -TypeName PSObject -Property $obj
    }

    $output | Export-Csv -NoType "$Env:userprofile\Desktop\Windows Updates Logs\WindowsUpdates_$entry.csv"
}