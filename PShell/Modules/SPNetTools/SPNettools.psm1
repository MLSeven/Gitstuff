<# ===========================================================================================
Powershell Library of Network Functions

S Potts
==============================================================================================
#>

# Script Environment
$ModuleFile = Get-Item $myinvocation.Mycommand.Path
$ModuleRoot = Split-Path -Path $myinvocation.Mycommand.Path -Parent
$ModuleName = $ModuleFile.BaseName
$ModuleLib = Join-Path $ModuleRoot -ChildPath "Lib"
$OSVersion = [Environment]::OSVersion.Version

# Global Definitions


# Informational

Write-Verbose "Loading Module $($ModuleName)"
Write-Verbose "OSVersion $($OSVersion.toString())"

<# ===========================================================================================
Check and Load any Dependent modules required by this module Here

ABORT Module if any Modules are Missing
==============================================================================================
#>

#$RequiredModules = @("DNSClient","DNSServer")
$RequiredModules=@()
$AvailableModules = $(Get-Module -ListAvailable).Name

$Available = @($RequiredModules | Foreach-Object {$AvailableModules -Match $_})
If ($RequiredModules.Count -NE $Available.Count)
{
	Write-Warning "$($ModuleName): At least 1 required Module is Missing - Aborting Import-Module"
	Exit
}
else
{
	$RequiredModules | Foreach-Object {Import-Module -Name $_ -Verbose}
}

<# ===========================================================================================
Check and Dot Source Any Common PShell Functions and Libraries here.

IncludeScripts is an Array of full Pathnames to the Powershell Script Files
==============================================================================================
#>

# Automatically include all Scripts in $ModuleLib
if (Test-Path -Path $ModuleLib) {Get-Childitem $ModuleLib -Filter *.ps1 | Foreach-Object {. $_.fullname}}

$IncludeScripts = @("\\ukmwemrpt1\decom\PShellLib\NetUtil_lib.ps1")
$IncludeScripts | Foreach-Object {if (Test-Path -Path $_) {Write-Verbose "Including Lib $_" ; . $_} else {Write-Warning "Cannot locate Library script $_"}}


<# ===========================================================================================
Module Code - These fuctions are Exported from the Module
==============================================================================================
#>
Function Split-List
<#
.SYNOPSIS
Splits a list into slices annd returns an object identifying the begin and end index of 
each slice

.DESCRIPTION
Takes in a list of objects and divides it up into a specified number of slices, or by specifying
a MaxSliceSize the object will be divided up into slices based on this value8

Examples:
$test = Split-List -List $List -Slices 3

Returns an index with the start and end position of each slice
.PARAMETER List
Array of objects to be sliced

.PARAMETER Slices
[Int] The number of slices required

.OUTPUTS
Array of Index items showing the Index Number, Begin and End positions
relative to the Input object

#>
{
	Param 
	(
		$List,
		[Int]$Slices=1,
		[Int]$MaxSliceSize=0
	)
	
	$Index = New-Object System.Collections.ArrayList
	#Check the parameters to see how the list is to be sliced up
	If ($MaxSliceSize -gt 0)
	{
		#A maximum slice size has been specified
		$Slices = [Math]::Truncate($List.Count/$MaxSliceSize) + 1
		$SliceSize = [Math]::Truncate($list.count / $Slices)
	}
	elseif ($Slices -GT 1) 
	{
		$SliceSize = [Math]::Truncate($list.count / $Slices)
	}
	else
	{
		#If neither $Slice or $MaxSliceSize are specified return the whole list as 1 slice
		$Slices = 1
		$SliceSize = $list.count
	}
	
	$Begin=0
	$End = $Begin + $SliceSize -1
	#Process first n-1 slices - these will all be equal size
	For ($Slice=0; $Slice -LT $Slices-1; $Slice ++)
	{
		[Void]$Index.Add([PSCustomObject]@{Slice=$Slice;Begin=$begin;End=$End})
		$Begin=$End+1
		$End = $Begin + $SliceSize -1
	}
	#Last Slice is all the way to the end of the list
	$End = $list.count -1
	[Void]$Index.Add([PSCustomObject]@{Slice=$Slice;Begin=$begin;End=$End})
	Return $Index
} # End of Split-List


Function ConvertTo-DecimalIP
<#
.SYNOPSIS
Connverts a Dotted Notation IPAddress into an Integer representation

.DESCRIPTION
Returns the Interger representation of the IPAddress

Examples:
$IntIP = ConvertTo-DecimalIP

.PARAMETER IPAddress
[String] IPAddress in Dotted Notation

.OUTPUTS
[Uint32] Unsigned Integer representation of IPAddress
#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [String]$IPAddress
	)
	
	$oIPAddress = [Net.IPAddress]::Parse($IPAddress)
	if ($oIPAddress)
	{
		$i=3
		[Uint32]$DecimalIP = 0
		$oIPAddress.GetAddressBytes() | foreach-object {$DecimalIP += $_ * [Math]::Pow(256, $i); $i-- }
		Return $DecimalIP
	}
	else
	{
		Return $Null
	}
} # End of ConvertTo-DecimalIP

Function ConvertTo-DottedIP
<#
.SYNOPSIS
Connverts an Unsigned Integer to Dotted Notation IPAddress

.DESCRIPTION
Returns the Dotted Notation IPAddress

Examples:
$IP= ConvertTo-DottedIP

.PARAMETER IPNumber
[String] IPAddress in Dotted Notation

.OUTPUTS
IP Address in Dotted Notation
#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [uInt32]$IPNumber
	)

	for ($i = 0; $i -lt 4; $i++)
	{
		$Octet = $($IPNumber % 256).ToString()
		$IPNumber = [Math]::Floor($IPNumber / 256)
		if ($i -eq 0) {$DottedIP = $octet} else {$DottedIP = $Octet + "." + $DottedIP}
	}
	$DottedIP
} # End of ConvertTo-DottedIP


Function Get-IPSubnet
<#
.SYNOPSIS
Takes a CIDR string and converts to an array of IP addresses

.DESCRIPTION
Takes a CIDR string and converts to an array of IP addresses

Examples:
Get-IPSubnet -CIDR 10.18.192.0/24

.PARAMETER CIDR
Subnet in CIDR Notation

.OUTPUTS
[String[]] String containing the IP Addresses in Dotted Notation
#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [String]$CIDR,
		[Switch]$AsRange
	)

	$SNStart = $CIDR.Split("/")[0]
	$Mask = $CIDR.Split("/")[1]
	$iSNStart = Convertto-DecimalIP $SNStart
	$iSNSize = [Math]::Pow(2,32-$Mask)-1
	Write-Verbose "iSNStart: $iSNStart iSNSize: $iSNSize"
	if ($AsRange)
	{
		$iSNEnd = $iSNStart+$iSNSize
		$SubnetList="$($SNStart)-$(Convertto-DottedIP $iSNEnd)"
	}
	else
	{
		$SubnetList = for ($IPNum=$iSNStart; $IPNum -le ($iSNStart+$iSNSize); $IPNum++) {Convertto-DottedIP $IPNum}
	}
	$SubnetList
} # End of Get-IPSubnet


Function Get-NBTStat
<#
.SYNOPSIS
Returns NBTStat data for and IPAddress

.DESCRIPTION
Peforms a NBTStat -A <IPAddress> and returns the captured output in a PSObject

Examples:
Get-NBTStat -IPAddress "156.71.11.11"

.PARAMETER IPAddress
IP Address to query

.OUTPUTS
Returns a PSSobject with the following Properties
TestedIP - The IPAddress Tested
NBTNode - $True if the IPAddress is a NetBios node
NBTBName - Netbios Name
NBTDomain - Netbios Domain
MAC - Netbios listner MAC Address
#>
{
	Param ($IPAddress)
	
	$Sb = {Param ($IP) ; nbtstat -A $IP}
	$startTime = get-date
	$Lines = Invoke-Command -Scriptblock $Sb -ArgumentList $IPAddress
	$elapsed = "{0:F2}" -f $(New-Timespan -Start $startTime).TotalSeconds
	$NBTOut = [String]::Join("`n",$Lines)
	$NamePattern = "^\s*(\S+)\s*\<00.*"
	$MACPattern = ""
	
	if ($NBTOut -match "Host not found")
	{
		$status = [PSCustomObject][Ordered]@{TestedIP=$IPAddress;NBTNode=$false;NBTName="";NBTDomain="";MAC="";elapsed=$elapsed}
	}
	else
	{
		if ($NBTOut -Match "\s*([A-Z0-9_-]+)\s*\<00.*UNIQUE") {$Name = $matches[1]} else {$Name = ""}
		if ($NBTOut -Match "\s*([A-Z0-9_-]+)\s*\<00.*GROUP") {$Domain = $matches[1]} else {$Domain = ""}
		if ($NBTOut -Match ".=\s([A-F0-9-]+)\s*") {$MAC = $matches[1]} else {$MAC = ""}
		
		#Write-Host "NBTstat returns $($Domain)\$($Name)"
		$status = [PSCustomObject][Ordered]@{TestedIP=$IPAddress;NBTNode=$true;NBTName=$Name;NBTDomain=$domain;MAC=$mac;elapsed=$elapsed}
	}
	$status
} # End of Get-NBTStat


Function Start-NBTScan
<#
.SYNOPSIS
Uses the NBTScan Utility to to a fast NBTStat -A on a list of IPAddresses
A New process is executed and StdOut is captured and returned

.DESCRIPTION
Returns NBT Netbios information from UDP port 137 via utility NBTScan.
The utility must exist in the same directory as the Powershel Module

Examples:
Start-NBTScan -IPAddress @("156.71.11.11","10.54.2.3")

.PARAMETER IPAddress
Array of IP address strings

.PARAMETER Arguments 
Optional parameters for the NBTScan command - changing these will change the output

.OUTPUTS
Returns the StdOut content
#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)][String[]]$IPAddress,
		[Parameter()][String]$Arguments="-m -n"
	)
	
	$NBTScanEXE = Join-Path -Path $Script:ModuleRoot -ChildPath "nbtscan-1.0.35.exe"
	Write-Verbose "EXE Path $($NBTScanEXE)"
	Write-Verbose "$($IPAddress.Count) - $($IPAddress.gettype())"
	$AddressList = [String]::Join(" ",$IPAddress)
	$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
	$ProcessInfo.FileName = $NBTScanEXE
	$ProcessInfo.RedirectStandardError = $true
	$ProcessInfo.RedirectStandardOutput = $true
	$ProcessInfo.UseShellExecute = $false
	$ProcessInfo.Arguments = "$($Arguments) $($AddressList)"
	Write-Verbose "$($Arguments) $($AddressList)"
	$P = New-Object System.Diagnostics.Process
	$P.StartInfo = $ProcessInfo
	$Started = $P.Start()
	if ($Started)
	{
		Write-Verbose "$($P.ProcessName) Launched OK"
		$stdout = $P.StandardOutput.ReadtoEnd()
		$stderr = $P.StandardError.ReadtoEnd()
		$P.WaitForExit()
		$ExitCode = $P.ExitCode
		If ($ExitCode -ne 0)
		{
			Write-Error "NBTScan returned error Code $ExitCode"
			$NBTStat=@()
		}
		else
		{
			$Output = $stdout -Split "[\r\n]+" | Foreach-Object {$_.Trim()} | Where-object {$_.Length -gt 0}
			Write-Verbose "NBTscan Completed OK and returned $($Output.Count) Non-Blank lines"
			$output | foreach-object {Write-verbose "$_"}
		}
		$P.Close()
		$P.Dispose()
		$NBTStat = @()
		$IPPattern = "(?<IP>(?:\d{1,3}\.){3}(?:\d{1,3}))"
		#HostPattern matches and captures <Domain>\<Node> - Optionally Domain may be blank
		$HostPattern = "(?<Domain>[a-zA-Z0-9\-_]*)\\(?<Node>[a-zA-Z0-9\-_]+)"
		$MACPattern = "(?<MAC>(?:[0-9A-Fa-f]{2}:){5}(?:[0-9A-Fa-f]{2}))"
		Foreach ($row in $Output)
		{
			$Fields = $Row -Split "\s+" 
			if ($Fields.Count -gt 0)
			{
				#Start checking the NBTScan record - Should be <IPAddress>  <Domain\None>  <MACAddress> <Other details>
				$NBT = [PSCustomObject][Ordered]@{TestedIP="";NBTNode="";NBTName="";NBTDomain="";MAC=""}
				#IP Address - if no match then record is malformed so discard
				if ($Fields[0] -Match $IPPattern) {$NBT.TestedIP = $Matches.IP} else {$NBT.TestedIP = $Null}
				#DOMAIN\NodeName
				if ($Fields[1] -Match $HostPattern) 
				{
					$NBT.NBTName = $matches.Node
					$NBT.NBTDomain = $Matches.Domain
				}
				else
				{
					#Malformed record - dont return 
					$NBT.TestedIP = $Null
				}
				
				#Check for Match against MAC address and use if it matches the RegEx
				if ($Fields[2] -Match $MACPattern) {$NBT.MAC = $Matches.MAC.ToUpper().Replace(":","-")}
				# Not interested in any other data
				
				if ($NBT.TestedIP) 
				{
					$NBT.NBTNode = $True
					$NBTStat += $NBT
				}
			}
		}
	}
	Else
	{
		Write-Warning "Process $NBTScanEXE failed to start"
		$NBTStat=@()
	}
	
	if ($NBTStat) {Write-Output $NBTstat}
} # End of Get-NBTScan


Function Ping-Async 
<#
.SYNOPSIS
Performs an Async Ping test against an array of IPAddresses

.DESCRIPTION
Performs a multithreaded Ping test on an array of IPAddress strings.

Examples:
$results = Ping-Async -IPList $IP -Timeout 4000 -Maxthreads 100

This performs a ping test on the IPAddresses in $IP. The timeout is set to 4000 ms
and the funtion will split the list up into a maximum of 100 Async threads

.PARAMETER IPList
List of IPAddress Strings

.PARAMETER Timeout
Ping Timeout in Milliseconds (Default 3000)

.PARAMETER MaxThread
Maximum async Threads

.OUTPUTS
PSCustomObject containing the results

#>
{
	[CmdletBinding()]
	param 
	(	
		[parameter(Position=0)][String[]]$IPList=@(),
		[parameter(Position=1)][String[]]$CIDRList=@(),
		[Int32]$Timeout=4000,
		[Int32]$MaxThreads=100
	)

	if ($CIDRList) 
	{
		Write-Host "Ping-Async: CIDR Mode: Converting Subnets to IPList"
		$IPList = $CIDRList | Foreach-Object {Get-IPSubnet -CIDR $_}
	}
	Write-Host "Ping-Async: IPList contains $($IPList.Count) Addresses. Splitting into Slices based on MaxTreads"
	$SavedEP = $ErrorActionPreference
	#$ErrorActionPreference = "SilentlyContinue"
	#Slice up the IPList to build a job queue
	$JobQueue = Split-List -List $IPList -MaxSliceSize $MaxThreads
	$JobQueue | Foreach-Object {Write-Verbose "Slice $($_.Slice) Begin $($_.Begin) End $($_.End)"}
	#Initialise Pinger
	$PingOptions = New-Object System.Net.NetworkInformation.PingOptions
	$encoder = [system.text.encoding]::ASCII
	$buffer = $encoder.GetBytes('acbdefghijklmnopqrstuvwxyz1234567890acbdefghijklmnopqrstuvwxyz1234567890')
	$PingResult = New-Object System.Collections.ArrayList
	Foreach ($Job in $JobQueue)
	{
		$IPSlice = $IPList[$Job.Begin..$job.End]
		Write-Verbose "Preparing Parsed Ping List for Slice $($Job.Slice) - Slice Size $($IPSlice.Count)"
		$PingThreads = $IPSlice | Where-Object {[System.Net.IPAddress]::TryParse($_,[ref]$Null)} |
			Foreach-Object {
				Write-Verbose "Initiating Ping on $_";
				$pinger = New-Object System.Net.NetworkInformation.Ping;
				[PSCustomobject]@{PingIP=$_;PingTask = $Pinger.SendPingAsync([System.Net.IPAddress]::Parse($_),$Timeout,$buffer,$PingOptions)}
			}
		Write-Verbose "$($PingThreads.Count) - Job $($Job.Slice)"
		#Wait for Async Ping threads 
		Try 
		{
			Write-Verbose "Waiting on Slice $($Job.Slice)"
			[void][Threading.Tasks.Task]::WaitAll($PingThreads.PingTask,$Timeout)
			Write-Verbose "All Tasks Complete - Slice $($Job.Slice)"
		} 
		Catch 
		{}

		ForEach ($Thread in $PingThreads)
		{
			
			$Task = $Thread.PingTask
			If ($Task.IsFaulted) 
			{
				$Result = $Task.Exception.InnerException.InnerException.Message
				$IPAddress = $Null
			} 
			Else
			{
				$Result = $Task.Result.Status.toString()
				$IPAddress = $Task.Result.Address.ToString()
				$RTT = $Task.Result.RoundtripTime
			}
			$Task.Dispose()
			$Task=$Null
			[VOID]$PingResult.Add([PSCustomobject]@{PingIP=$Thread.PingIP;ReplyAddress=$IPAddress;Status=$Result;RoundtripTime = $RTT})
		}
		#Return result
	}
	$ErrorActionPreference = $SavedEP
	Return $PingResult
} # End of Ping-Async 

Function Test-TCPPort 
<#
.SYNOPSIS
Test a list of IPAddress to see if a list of TCP Ports is open

.DESCRIPTION
Performs a multithreaded Port Scan on a list ip IP Addresses

Examples:
$results = Test-TCPPort -IPList $IP -Timeout 10000 -Maxthreads 100 -Portlist=@(22,139)

This performs a scan of ports 22 and 139 in the list of IPAddresses

.PARAMETER IPList
List of IPAddress Strings

.PARAMETER Timeout
Port Timeout in Milliseconds (Default 10000)

.PARAMETER MaxThread
Maximum async Threads

.PARAMETER Portlist
Array of port numbers

.OUTPUTS
PSCustomObject containing the results

#>
{
	[CmdletBinding()]
	param 
	(	[parameter(ValueFromPipeline=$True)][String[]]$IPList,
		[Int32]$Timeout=10000,
		[Int32]$MaxThreads=100,
		[Int32[]]$Portlist
	)

	Write-Verbose "IPAddress List Size $($IPList.count)"
	#Slice up the IPList to build a job queue
	$JobQueue = Split-List -List $IPList -MaxSliceSize $MaxThreads
	$JobQueue | Foreach-Object {Write-Verbose "Slice $($_.Slice) Begin $($_.Begin) End $($_.End)"}

	$PortStatus = @()
	#$sockets=@()
	foreach ($Port in $PortList)
	{
		Write-Verbose "Preparing to test port $Port"
		Foreach ($Job in $JobQueue)
		{
			$Sockets=@()
			$IPSlice = $IPList[$Job.Begin..$job.End]
			Write-Verbose "Preparing IPAddess List for Slice $($Job.Slice) - Slice Size $($IPSlice.Count)"
			Foreach ($IP in $IPSlice)
			{
				$Socket = New-Object Net.Sockets.TcpClient;
				$Sockets += [PSCustomobject]@{IP=$IP;Socket=$Socket;Task = $Socket.ConnectAsync($IP,$Port)}
			}
			While (($Sockets.Task | Where-object {-NOT $_.isCompleted}).count -gt 0)
			{
				Try 
				{
					Write-Verbose "$([Datetime]::now) Waiting on Task Threads TimeOut ..."
					[void][Threading.Tasks.Task]::WaitAll($Sockets.Task,$timeout)
					$sockets.Task | group-object -Prop Status | Foreach-Object {Write-Verbose "$($_.Name) - $($_.Count)"}
					
				} 
				Catch 
				{ 
					write-verbose "$([Datetime]::now) All-Complete"
					$sockets.Task | group-object -Prop Status | Foreach-Object {Write-Verbose "Final $($_.Name) - $($_.Count)"}
				}
			}
			#Process first Slice
			
			Foreach ($S in $Sockets)
			{
				Write-Verbose "Port Status for $($S.IP) : $($S.Socket.Connected)"
				$PortStatus += [PSCustomObject][ordered]@{IP=$S.IP;Port=$Port;Connected=$S.Socket.Connected;Status=$S.Task.Status;Message=$S.Task.Exception.Innerexception.Message}
				if ($S.Task.isCompleted)
				{
					$S.task.Dispose()
					$S.Socket.Close()
					$S.Socket.Dispose()
				}
			}
			
		}
	}
	Return $PortStatus

} # End of Test-TCPPort

<#
.SYNOPSIS
Converts from a SID into a Security Principal

.DESCRIPTION
Converts a SID string or ByteArray into NT Style Object

Examples:

	ConvertFrom-SID -Sid $SIDString

.PARAMETER SID
SID Identifier to Translate

.PARAMETER FromByte
SID Format is Byte String (Default is [String])

.OUTPUTS
Windows security Principal
#>
Function ConvertFrom-SID

{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] $SID,
		[Switch]$FromByte
	)

	if($FromByte)
	{
		$ID = New-Object System.Security.Principal.SecurityIdentifier($sid,0)
	}
	else
	{
		$ID = New-Object System.Security.Principal.SecurityIdentifier($sid)
	}
	Trap [system.exception]
	{
		continue
	}

	if ($ID)
	{
		$User = $ID.Translate([System.Security.Principal.NTAccount])
		$IDName = $User.Value
	}
	else
	{
		$IDName = $null
	}
	$IDName
} # End of ConvertFrom-SID


Function Test-AdminAccess
<#
.SYNOPSIS
Checks a Server to see if the user (or Supplied credentials) Has Admin Access

.DESCRIPTION
Checks a Server to see if the user (or Supplied credentials) Has Admin Access

Examples:
$Status = Test-AdminAccess -Server Name -Credentials $User


.PARAMETER Server
Server Name or IPAddress

.PARAMETER Credential
Credentials to use for the Connection

.OUTPUTS
[PSCustomObject] containing test results
#>
{
	[CmdletBinding()]
	param( 
	[Parameter(Position=0, Mandatory=$True)]$Server, 
	[Parameter(Position=1, Mandatory=$false)]$Credential
	)
	
	if ($PSBoundParameters.ContainsKey("Credential")) {$User = $Credential.Username} else {$User = "$($ENV:USERDOMAIN)\$($Env:Username)"}
	$Results = [PSCustomObject][Ordered]@{TestedHost=$Server;PingOK=$False;Credentials=$User;WMIHost="NotTested";WMIAccess="NotTested";AdminShare="NotTested"}
	Write-Verbose "Checking Admin Access to Server $Server with Credentials $($USER)"
	$save = $erroractionpreference
	$erroractionpreference = "Silentlycontinue"
	if ($(Test-connection -computer $server -Count 1 -quiet))
	{
		$Results.PingOK = $True
		if ($PSBoundParameters.ContainsKey("Credential"))
		{
			$wmitest = Get-WMICustom -class win32_computersystem -Computer $Server -Credential $Credential -Timeout 10
		}
		else
		{
			$wmitest = Get-WMICustom -class win32_computersystem -Computer $Server -Timeout 10
		}
		if ($($WMITest | Select-Object -First 1).Gettype().Name -ne "ErrorRecord")
		{
			$Results.WMIHost = $wmitest.Name
			$Results.WMIAccess = $True
		}
		Else
		{
			$Results.WMIHost = ""
			$Results.WMIAccess = $False
		}
		if ($PSBoundParameters.ContainsKey("Credential"))
		{
			$FileTest = Connect-WSHNetDrive -Path "\\$Server\Admin$" -Credential $Credential
		}
		else
		{
			$FileTest = Connect-WSHNetDrive -Path "\\$Server\Admin$"
		}
		if ($Filetest)
		{
			$Results.AdminShare=$true
			$Null=Disconnect-WSHNetDrive -Path "\\$Server\Admin$"
		}
		else
		{
			$Results.AdminShare = $False
		}
	}
	$erroractionpreference = $save
	write-verbose "Finished testing $Server"
	$Results
} # End of Test-AdminAccess


<#
.SYNOPSIS
Returns members of a Local group on Server

.DESCRIPTION
remotely returns the members of a Windows LocalGroup 

Examples:

Get-Localgroupmembers -Server test1 -Localgroup Administrators

.PARAMETER Server
Server Name

.PARAMETER LocalGroup
LocalGroup name to evaluate

.OUTPUTS
PSCustomobject containing the group members

#>
Function Get-LocalGroupMembers
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0)] [String]$Server="$Env:Computername",
		[Parameter(Position=1)] [String]$LocalGroup="Administrators"
	)


	# List local group members on the local or a remote computer  
	if (Test-Connection -Computername $server -Quiet -count 1)
	{
		if($([ADSI]::Exists("WinNT://$($Server)/$($LocalGroup),group")))
		{
			$ReturnData = @()
			$group = [ADSI]("WinNT://$($Server)/$($LocalGroup),group")  
			$members = @()  
			write-verbose "Checking server $Server..."
			Foreach ($item in $Group.Members())  
			{  
				$AdsPath = $item.GetType().InvokeMember("Adspath", 'GetProperty', $null, $item, $null)  

				# Domain members will have an ADSPath like WinNT://DomainName/UserName.  
				# Local accounts will have a value like WinNT://DomainName/ComputerName/UserName.  

				$PathParts = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries)  
				$name = $PathParts[-1]  
				$domain = $PathParts[-2]  
				$class = $item.GetType().InvokeMember("Class", 'GetProperty', $null, $item, $null)
				if ($name -Match "S-1-5") 
				{
					$GrpName = ConvertFrom-SID -SID $name
					if ($GrpName)
					{
						$Name = $GrpName.Split("\")[1]
						$Domain = $GrpName.Split("\")[0]
						$Class="TranslatedSID"
					}
				}
				$ReturnData += [PSCustomObject][Ordered]@{Server=$Server;LocalGroup=$LocalGroup;Member=$name;Domain=$domain;Class=$class}
			}
		}
		else
		{
			Write-Verbose "Failed to access $LocalGroup on Server $Server"
			$ReturnData += [PSCustomObject][Ordered]@{Server=$Server;LocalGroup=$LocalGroup;Member="WinNT provider not found";Domain="";Class=""}
		}
	}
	else
	{
		write-verbose "Server $Server is not responding"
		$ReturnData += [PSCustomObject][Ordered]@{Server=$Server;LocalGroup=$LocalGroup;Member="Server Not Responding";Domain="";Class=""}
	}
	$ReturnData
}  


Function Test-Credential
<#
.SYNOPSIS
	Takes a PSCredential object and validates it against the domain (or local machine, or ADAM instance).

.PARAMETER cred
	A PScredential object with the username/password you wish to test. Typically this is generated using the Get-Credential cmdlet. Accepts pipeline input.

.PARAMETER context
	An optional parameter specifying what type of credential this is. Possible values are 'Domain','Machine',and 'ApplicationDirectory.' The default is 'Domain.'

.OUTPUTS
	A boolean, indicating whether the credentials were successfully validated.

#>
{
	[CmdletBinding()]
	Param
	(
		[parameter(Mandatory=$true,ValueFromPipeline=$true)][System.Management.Automation.PSCredential]$credential,
		[parameter()][validateset('Domain','Machine','ApplicationDirectory')][string]$context = 'Domain'
	)
	
	Begin
	{
		Add-Type -assemblyname system.DirectoryServices.accountmanagement
		if ($context -Match 'Domain')
		{
			$DomainName = $Credential.GetNetworkCredential().Domain
			$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::$context,$DomainName) 
		}
		else
		{
			$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::$context) 
		}
	}
	process 
	{
		$Valid = $DS.ValidateCredentials($credential.UserName, $credential.GetNetworkCredential().password)
		Write-Verbose "Authenticating Credential $($credential.UserName) against $Context Context : Validated $Valid"
		$Valid
	}
} #End of Test-Credential


Function Get-WMINet
<#
.SYNOPSIS
Using WMI get the Remote computers Network Adapter config

.DESCRIPTION
Uses WinNT provider to add an Object from Domain into LocalGroup on Server

Examples:
Using WMI get the Remote computers Network Adapter config

.PARAMETER Server
Server Name

.PARAMETER Credential
Credential Object

.OUTPUT
WMI Network Adapter Config
#>
{
	[CmdletBinding()]
	Param
	(
		[String]$Server,
		$credential
	)
	
	if ($PSBoundParameters.ContainsKey("Credential"))
	{
		$comp = get-wmiobject -Credential $Credential -Class Win32_ComputerSystem -Computer $Server
		$net = get-wmiobject -Credential $Credential -query "Select * from win32_networkadapterconfiguration where IPEnabled='True'" -computer $server | Select-object -property Description,MACAddress,IPAddress,IPSubnet,DefaultIPGateway,DHCPEnabled,DNSDomain,DNSDomainSuffixSearchOrder,DNSServerSearchOrder,WINSPrimaryServer,WINSSecondaryServer
	}
	else
	{
		$comp = get-wmiobject -Class Win32_ComputerSystem -Computer $Server
		$net = get-wmiobject -query "Select * from win32_networkadapterconfiguration where IPEnabled='True'" -computer $server | Select-object -property Description,MACAddress,IPAddress,IPSubnet,DefaultIPGateway,DHCPEnabled,DNSDomain,DNSDomainSuffixSearchOrder,DNSServerSearchOrder,WINSPrimaryServer,WINSSecondaryServer
	}
	#Fix some of the Network Properties so they display correctly
	foreach ($nic in $net)
	{
		$nic.IPAddress = [System.String]::Join("`n",$nic.IPAddress)
		$nic.IPSubnet = [System.String]::Join("`n",$nic.IPSubnet)
		$nic.DefaultIPGateway = if ($nic.DefaultIPGateway) {[System.String]::Join("`n",$nic.DefaultIPGateway)}
		$nic.DNSServerSearchOrder = if ($nic.DNSServerSearchOrder) {[System.String]::Join("`n",$nic.DNSServerSearchOrder)}
		$nic.DNSDomainSuffixSearchOrder = if ($nic.DNSDomainSuffixSearchOrder) {[System.String]::Join("`n",$nic.DNSDomainSuffixSearchOrder)}
	}
	$Return = [PSCustomObject]@{Question=$Server;Comp=$Comp;Net=$Net}
	return $Return	
} #End Of Get-WMINet

Function Add-LocalGroupMember
<#
.SYNOPSIS
Adds and Object from the Domain into a local group on Remote Server

.DESCRIPTION
Uses WinNT provider to add an Object from Domain into LocalGroup on Server

Examples:
Add-LocalGroupMember -Server test1 -Domain ASTRAZENECA -Object XAZ-Users -LocalGroup Administrators

.PARAMETER Server
Server Name
.PARAMETER Object
Group or UserName to add
.PARAMETER Domain
Domain containing Object
.PARAMETER LocalGroup
Name of the Local Group on Server (Default Administrators)
#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] $Server,
		[Parameter(Position=1,Mandatory=$true)] [String]$Domain,
		[Parameter(Position=2,Mandatory=$true)] [String]$Object,
		[Parameter(Position=3)] [String]$LocalGroup="Administrators"
	)
	
	# Get the Path to the Domain Object to add to the Local group
	
	$Obj = [ADSI]"WinNT://$Domain/$Object"
	if ($Obj.Path)
	{
		Foreach ($item in $Server)
		{
			$Lg = [ADSI]"WinNT://$Item/$LocalGroup,group"
			if ($LG.Path)
			{
				#Local Group object connected
				$Lg.PSBase.Invoke("Add",$Obj.Path)
				If ($?)
				{
					Write-Host "Added $Domain\$Object to Localgroup $LocalGroup on $Item"
				}
				else
				{
					Write-Host "FAILED TO ADD $Domain\$Object to Localgroup $LocalGroup on $Item"
				}
			}
			Else
			{
				Write-Host "Failed to locate Local Group Object  $Item\$LocalGroup."
			}
			$Lg = $Null
		}
	}
	Else
	{
		Write-Host "Failed to locate $Domain\$Object"
	}
} #End of Add-LocalGroupMember


Function Resolve-NameLookup
<#
	.SYNOPSIS
		Performs a Fwd or Rev Lookup Asynchronously 

	.DESCRIPTION
		Performs a Fwd or Rev Lookup Asynchronously

	.PARAMETER Computername
		List of computernames or IPAddresses to resolve

	.NOTES

	.OUTPUT
		{PSCustomObject]

	.EXAMPLE

#>
{
	[cmdletbinding()]
	param 
	(
		[parameter(Mandatory=$True,Position=0)][String[]]$Computername,
		[Int32]$Timeout=4000,
		[Int32]$MaxThreads=100
	)
	
	Write-Verbose "Computername Item List Size $($Computername.count)"
	$SavedEP = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
	#Slice up the IPList to build a job queue
	$JobQueue = Split-List -List $Computername -MaxSliceSize $MaxThreads
	$JobQueue | Foreach-Object {Write-Verbose "Slice $($_.Slice) Begin $($_.Begin) End $($_.End)"}
	
	$Results = @()
	Foreach ($Job in $JobQueue)
	{
		$Computerlist = $Computername[$Job.Begin..$job.End]
		$SliceTasks = @()
		ForEach ($Computer in $Computerlist) 
		{
			If (Test-IPString -IPString $Computer)
			{
				Write-Verbose "Reverse Lookup IPAddress $Computer"
				$SliceTasks += [pscustomobject]@{Computername = $Computer; Task = [system.net.dns]::GetHostEntryAsync($Computer)}
			} 
			Else
			{
				Write-Verbose "Forward Lookup Hostname $Computer"
				$SliceTasks += [pscustomobject]@{Computername = $Computer; Task = [system.net.dns]::GetHostAddressesAsync($Computer)}
			}
		}
		Try 
		{
			[void][Threading.Tasks.Task]::WaitAll($SliceTasks.Task)
		} 
		Catch {}
		Foreach ($Task in $SliceTasks)
		{
			Write-Verbose "Processing Result for $($Task.Computername)"
			If ($Task.Task.IsFaulted) 
			{
				$Res = $Task.Task.Exception.InnerException.Message
			} 
			Else 
			{
				If ($Task.Task.Result.IPAddressToString)
				{
					$Res = $Task.Task.Result.IPAddressToString
				} 
				Else 
				{
					$Res = $Task.Task.Result.HostName
				}
			}
			$Results += [PSCustomObject][Ordered]@{Computername = $Task.Computername; Result = $Res}
		}
	}
	$ErrorActionPreference = $SavedEP
	Return $Results
}

Function Get-FWStatus
<#
	.SYNOPSIS
		Use WinRM to get Windows FW Status 

	.DESCRIPTION
		Use WinRM to get Windows FW Status y

	.PARAMETER Computername
		List of computernames to check

	.NOTES

	.OUTPUT
		Firewall Status

	.EXAMPLE

#>
{
	[cmdletbinding()]
	param 
	(
		$Computername
	)
	
	$FWStatus = Invoke-Command -Scriptblock {Get-NetFirewallProfile -Policystore Activestore | Select-Object -Property PSComputername,Profile,Enabled} -Computername $Computername
	$FWStatus
}

	

<# ===========================================================================================
End of Module Functions

Define and Export Module Members below
==============================================================================================
#>

$ModuleFunctionNames = Get-Content $ModuleFile | Where-Object {$_ -Match "^\s*Function\s+(\S+)" } | Foreach-Object {$Matches[1]}
Export-Modulemember -Function $ModuleFunctionNames
