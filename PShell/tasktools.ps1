# Some Useful Powershell snippets for debuging Morpheus Tasks
#
# Note if executed directly from GitHub then Morpheus Variable Substitutiin will not apply

# Return the process hierarchy. Returns an array of processes from child-parent order 
function Get-ProcessTree {
<#
.SYNOPSIS
Get all processeses in the Process Tree starting at the current process and working back through the parents

.DESCRIPTION
Uses win32_process CimInstance to retrieve the process tree from the current process through the ancestors

Examples:
Get-ProcessTree

.PARAMETER pid
ProcessId which is the leaf process

.OUTPUTS
array of processes child - parent - grandparent etc

#>    
    param (
        [int]$ProcessId=$PId
    )

    $tree = do {
        $currProcess = Get-CimInstance -class win32_Process -filter "ProcessId=$processId" -ErrorAction SilentlyContinue
        if ($?) {
            #Process exists - get Parent ID return current process and loop
            $processId = $currProcess.ParentProcessID
            $currProcess
        } else {
            # No Process
            $processId = $null
        }
    } while ($processId)
    #force an array to be returned
    return $tree
}


Function Get-TaskInfo {
<#
.SYNOPSIS
Get some basic information (user, computer enviornment etc) avout the task

.DESCRIPTION
Get some basic information (user, computer enviornment etc) avout the task

Examples:
Get-TaskInfo

.PARAMETER morpheus
Optional PSCustomObject containing Morpheus specific properties which will be returned with the Task info

.PARAMETER AsJson
Switch parameter. Return the data as json rather than PSCustomObject

.OUTPUTS
PSCustomObject containing the task properties

#>    
    param (
        $morpheus=$null,
        [switch]$AsJson
    )

    # LoginId is useful for tracing in Security log
    $loginId = Get-CimInstance -class win32_process -Filter "ProcessId =$PID" | Get-CimAssociatedInstance -ResultClassName win32_logonsession 

    # Windows Security and Identity object for current process
    $userIdentity =  [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $userPrincipal = [System.Security.Principal.WindowsPrincipal]$UserIdentity
    $adminElevated=$UserPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    #Create a Process Hierarchy 
    $tree = get-processTree -processId $PID 
    $child = $null
    foreach ($p in $tree) {
        $root = [PSCustomObject]@{pid=$p.ProcessId;name=$p.name;child=$child}
        $child = $root
    }

    $TaskInfo = [PSCustomObject]@{
        hostName = [Environment]::MachineName;
        morpheus = $morpheus;
        OS = [Environment]::OSVersion;
        envUserName = [Environment]::UserName;
        CurrentDirectory = [Environment]::CurrentDirectory;
        interactive = [Environment]::UserInteractive;
        userIdentity = $userIdentity | Select-Object -property Name,AuthenticationType,IsAuthenticated,IsSystem,ImpersonationLevel;
        elevated = $adminElevated;
        UTCStart = [DateTime]::now.toUniversalTime().toString();
        processId = $pid;
        eVars = [Environment]::GetEnvironmentVariables();
        cmdLine = [Environment]::CommandLine;
        processTree = $root;
        loginId = $loginId | Select-Object -prop LogonId,LogonType,StartTime,AuthenticationPackage
    }
    if ($AsJson) {return $TaskInfo | ConvertTo-Json -Depth 10} else {return $TaskInfo}
}