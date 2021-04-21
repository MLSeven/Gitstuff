#Read events -d 403
#get message
#Match message for encodedcommand
#decode

Function Read-PSLog {
    param (
        $EventId=403,
        $exportedLog = ""
    )

    if ($exportedLog) {
        $Events = Get-WinEvent -Path $exportedLog | where-object {$_.id -eq $EventId}
    }
    else {
        $Events = Get-EventLog -log "Windows Powershell" -InstanceId $EventId
    }

    write-host "Extracted ID 403 events : count $($Events.count)"

    $eventData = foreach ($e in $Events) {
        if ($exportedLog) {
            $output = [PSCustomObject]@{computer=$e.MachineName;index=$e.recordId;Time=$e.TimeCreated;UTCTime=$e.TimeCreated.ToUniversalTime();host="";command="";encodedcommand=""}
        }
        else {
            $output = [PSCustomObject]@{computer=$e.MachineName;index=$e.index;Time=$e.TimeWritten;UTCTime=$e.TimeWritten.ToUniversalTime();host="";command="";encodedcommand=""}
        }
        if ($e.message -match "HostName=(.*)\r") {
            $output.host=$matches[1]
        }
        if ($e.message -match "HostApplication=(.*)\r") {
            $output.command=$matches[1]
            if ($output.command -match "-encodedcommand (\S*)") {
                #Base64 encoded command
                $output.encodedcommand=[System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($matches[1]))
            }
        }
        $output
    }
    $eventData
}

function Get-AgentInstall {
    param (
        $eventData
    )

    $eventData = $eventData | Sort-Object -property Index
    write-host "Processing Event Data for Agent Install"

    $fragments = foreach ($e in $eventData) {
        if ($e.command -match ",(\[.*\)\))\)") {
            invoke-expression $matches[1]
        }
    }
    [string]::join('',$fragments)
}

function Read-Setuplog {
    param (
        [Datetime]$startDate = [Datetime]::now.Date,
        [string]$path = "C:\Windows\Panther\setup.etl"
    )

    write-host "Reading Log file $path - filtering on events after $($startDate)"

    # Load the setup event log and filter on StartDate (Defaults to today)    
    $events = Get-WinEvent -Path $path -Oldest | where-object {$_.TimeCreated -ge $startDate.Date}
    $events
}
