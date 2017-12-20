# Array of servers.
$entries = @( 
    "jac-vhost.ad.unosquare.com",
    "leon-vhost.ad.unosquare.com",
    "allius-vhost.ad.unosquare.com",
    "gdl-vhost.ad.unosquare.com",
    "gdl-devlab.ad.unosquare.com"
)

# Script Block that its going to be executed in each server.
$sc = { 
    # Adds windows.serverbackup snap-in to the current session.
    add-pssnapin windows.serverbackup 
    # Gets the history of the backup operations.
    $Summary = Get-WBSummary 
    $Summary
}

# Gets the hostname
$hostname = hostname 

foreach($entry in $entries)
{
    # Checks if the entry is localhost
    $remote = $entry -notmatch $hostname 
    
    # Executes the Script Block in each server
    if($remote) {
        # If remote computer
        $result = Invoke-Command -ComputerName $entry -ScriptBlock $sc
    } else { 
        # Localhost
        $result = Invoke-Command -ScriptBlock $sc
    }

    $date = Get-Date
    $date = $date.AddDays(-10)

    # Selects the important information
    $output = $result | Select-Object PSComputerName, 
        NextBackupTime, LastSuccessfulBackupTime, LastSuccessfulBackupTargetLabel, LastBackupTime | 
        Where-Object { $_.LastBackupTime -gt $date } 

    # Sends the output to a file
    $output | Out-File -NoType "$Env:userprofile\Desktop\Backups Logs\BackupLog_$entry.txt" -width 300
}