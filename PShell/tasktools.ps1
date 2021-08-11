# Some Useful Powershell snippets for debuging Morpheus Tasks

function Get-ProcessTree {
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

# Get all Proccesses in the current Session
#$sessionId= (Get-Process -Id $PID).SessionId 
#$sessionProcesses = Get-Process | Where-Object {$_.SessionId -eq $sessionId} | Select-Object -property Id,Name,SessionId

$loginId = Get-CimInstance -class win32_process -Filter "ProcessId =$PID" | Get-CimAssociatedInstance -ResultClassName win32_logonsession 
$loginProcesses = $loginId | Get-CimAssociatedInstance -ResultClassName win32_process 

# Windows Security and Identity object for current process
$UserIdentity =  [System.Security.Principal.WindowsIdentity]::GetCurrent()
$UserPrincipal = [System.Security.Principal.WindowsPrincipal]$UserIdentity
$AdminElevated=$UserPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$t = get-processTree -processId $PID | Select-object -prop id,name

$PSEnv = [PSCustomObject]@{
    hostName = [Environment]::MachineName;
    OS = [Environment]::OSVersion;
    envUser = [Environment]::UserName;
    CurrentDirectory = [Environment]::CurrentDirectory;
    interactive = [Environment]::UserInteractive;
    user = $UserIdentity | Select-Object -property Name,AuthenticationType,IsAuthenticated,IsSystem,ImpersonationLevel;
    elevated = $AdminElevated;
    UTCStart = [DateTime]::now.toUniversalTime().toString();
    pId = $pid;
    environment = [Environment]::GetEnvironmentVariables();
    cmdLine = [Environment]::CommandLine;
    processTree = $t;
    loginId = $loginId | Select-Object -prop LogonId,LogonType,StartTime,AuthenticationPackage;
    loginProcesses = $loginProcesses | Select-Object -prop ProcessId,Name,CreationDate,ParentProcessId
}

$json = $PSEnv | convertto-json -depth 3
$json