# Some Useful Powershell snippets for debuging Morpheus Tasks

# Return the process hierarchy. Returns an array of processes from child-parent order 
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

# Use Morpheus Variable to determine the context
$morpheusNullString="null"
if ("<%=instance.name %>" -eq $MorpheusNullString) {$context="Server"} else {$context="Instance"}

# LoginId is useful for tracing in Security log
$loginId = Get-CimInstance -class win32_process -Filter "ProcessId =$PID" | Get-CimAssociatedInstance -ResultClassName win32_logonsession 

# Windows Security and Identity object for current process
$userIdentity =  [System.Security.Principal.WindowsIdentity]::GetCurrent()
$userPrincipal = [System.Security.Principal.WindowsPrincipal]$UserIdentity
$adminElevated=$UserPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

#Proccess Hierarchy
$tree = get-processTree -processId $PID 

$child = $null
foreach ($p in $tree) {
  $root = [PSCustomObject]@{pid=$p.ProcessId;name=$p.name;child=$child}
  $child = $root
}

$PSEnv = [PSCustomObject]@{
    hostName = [Environment]::MachineName;
    context= $context;
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

$json = $PSEnv | convertto-json -depth 3
$json