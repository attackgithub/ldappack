# Array of servers
$entries = @(
    "gdl-dc-01",
    "jac-dc-01",
    "allius-dc-01",
    "leon-dc-01",
    "unocorp-dc-01",
    "unocorp-dc-02"
)
# Script Block that its going to be executed in each server
$sc = { 
    # Creates a Microsoft Update Session
    $Session = New-Object -ComObject "Microsoft.Update.Session" 
    $Searcher = $Session.CreateUpdateSearcher()
    $historyCount = $Searcher.GetTotalHistoryCount()
    # Gets the information of the installed updates
    $search = $Searcher.QueryHistory(0, $historyCount)
    # Gets the information of the available updates
    $searchresult = $Searcher.Search("IsInstalled=0")

    # Initialize an array
    $updates = @()

    foreach($Update in $searchresult.Updates){
        $updates += [pscustomobject]@{
            Title = $Update.Title
            Date = $Update.LastDeploymentChangeTime
            IsDownloaded = $Update.IsDownloaded
            IsInstalled = $Update.IsInstalled 
            Url = $($Update.MoreInfoUrls)
        } 
    }

    $search
    $updates
}

# Gets the hostname
$hostname = hostname 
# Iterates through the array of servers
foreach($entry in $entries)
{
    # Checks if the entry is a ipaddress
    $ipaddress = [bool]($entry -as [ipaddress])
    # Checks if the entry is localhost
    $remote = $entry -notmatch $hostname 

    # Executes the Script Block in each server
    if($ipaddress){
        # If Ipaddress
        # (Get-Credential).Password | ConvertFrom-SecureString | Out-File "$Env:userprofile\Desktop\pass.txt"
        $User = $entry+"\"+$env:username
        $File = "$Env:userprofile\Desktop\pass.txt"
        $MyCredential = New-Object -TypeName System.Management.Automation.PSCredential `
         -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)
        
        # Adds the ipaddress to the trustedhosts
        Set-Item WSMan:\localhost\Client\TrustedHosts –Value $entry -Concatenate -Force
        Restart-Service WinRM
        $result = Invoke-Command -ComputerName $entry -Credential (Get-Credential -Credential $MyCredential) -ScriptBlock $sc
    }
    elseif($remote) {
        # If remote computer
        $result = Invoke-Command -ComputerName $entry -ScriptBlock $sc
    } else { 
        # Localhost
        $result = Invoke-Command -ScriptBlock $sc
    }
    
    $date = Get-Date
    $date = $date.AddDays(-10)
    $result
    # Selects the important information and switches the number of operation and resultcode to text
    $output = $result | Select-Object Title, Date,
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
        Where-Object { $_.Date -gt $date -and $_.Status -ne $null} | Format-List

    $AvailableUpdates = $result | Select-Object Title, KB, Date, MaxDownloadSize, IsDownloaded, IsInstalled, Url | Where-Object { $_.IsInstalled -eq $false}  #Selects the updates that are not installed
    
    if($output -eq $null)
    {
        $obj = @{ 'Message' = "There is no updates information. Please check.";}
        $output = New-Object -TypeName PSObject -Property $obj
    }

    if($AvailableUpdates -eq $null)
    {
        $obj = @{ 'Message' = "There is no available updates.";}
        $AvailableUpdates = New-Object -TypeName PSObject -Property $obj
    }
    
    # Sends the output to a file
    $output | Out-File "$Env:userprofile\Desktop\Windows Updates Logs\WindowsUpdates_$entry.csv" -width 300
    $AvailableUpdates | Out-File  "$Env:userprofile\Desktop\Windows Updates Logs\$entry.AvailableWindowsUpdates.csv" -width 120
}