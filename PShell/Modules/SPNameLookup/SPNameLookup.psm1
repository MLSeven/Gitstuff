<# ===========================================================================================
Powershell Library of IP Testing AND Name Lookup Functions

For perfomance reasons the Netbios Node type should be set to p-node
to prevent lengthly time-outs on WINS Lookups

S Potts
==============================================================================================
#>

# Script Environment
$ModuleFile = Get-Item $myinvocation.Mycommand.Path
$ModuleRoot = Split-Path -Path $myinvocation.Mycommand.Path -Parent
$ModuleName = $ModuleFile.BaseName
$OSVersion = [Environment]::OSVersion.Version

# Global Definitions

$ZoneFile = "\\ukmwemrpt1\decom\DNS\Zonelist.xml"

# Informational

Write-Verbose "Loading Module $($ModuleName)"
Write-Verbose "OSVersion $($OSVersion.toString())"

<# ===========================================================================================
Check and Load any Dependent modules required by this module Here

ABORT Module if any Modules are Missing
==============================================================================================
#>

#$RequiredModules = @("DNSClient","DNSServer","SPNetTools")
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
$IncludeScripts = @()
$IncludeScripts | Foreach-Object {if (Test-Path -Path $_) {Write-Verbose "Including Lib $_" ; . $_} else {Write-Warning "Cannot locate Library script $_"}}


<# ===========================================================================================
Module Code - These fuctions are Exported from the Module
==============================================================================================
#>


$NSlookupScript =
{
	param ($Hostnames,$Zones,$DNSServer)
	
	If ($Zones)
	{
		#Non-Qualified names try each Zone in $Zones
		$ZoneRecs = $Zones |
			Foreach-Object {$ZoneName = $_; $Hostnames | Foreach-Object {Resolve-DNSname -name "$($_).$($ZoneName)" -server $DNSServer -QuickTimeout -DNSOnly -ErrorAction "SilentlyContinue"}}
	}
	else
	{
		#FQDN Searches - No Zone names
		$ZoneRecs = $Hostnames | Foreach-Object {Resolve-DNSname -name $_ -server $DNSServer -QuickTimeout -DNSOnly -ErrorAction "SilentlyContinue"}
	}
	$Lookup = Switch ($Zonerecs) 
			{ 
				{$_.Type -eq "A"} {[PSCustomObject][Ordered]@{Host=$($_.Name.Split(".")[0]).ToUpper();FQDN=$_.Name.ToUpper();IPAddress=$_.IP4Address;NameHost="";QueryType=$_.Type.ToString()}}
				{$_.Type -eq "CNAME"} {[PSCustomObject][Ordered]@{Host=$($_.Name.Split(".")[0]).ToUpper();FQDN=$_.Name.ToUpper();IPAddress="";NameHost=$_.Namehost.ToUpper();QueryType=$_.Type.ToString()}}
				default {}
			}
	Return $Lookup
}

$WINSlookupScript =
{
	param ($Hostnames,$WINSServer)
	
	#NQDN WINS Searches - No Zone names
	$WINSRecs = $Hostnames | Foreach-Object {Resolve-DNSName -Server $WINSServer -LlmnrNetbiosOnly -QuickTimeOut -Name $_  -ErrorAction SilentlyContinue}
	$Lookup = Switch ($WinsRecs) 
			{ 
				{$_.Type -eq "A"} {[PSCustomObject][Ordered]@{Host=$($_.Name.Split(".")[0]).ToUpper();FQDN=$_.Name.ToUpper();IPAddress=$_.IP4Address;NameHost="";QueryType="NetBIOS"}}
				default {}
			}
	Return $Lookup
}


Function Test-Zonelist
<#
.SYNOPSIS
Load the AZ DNS structure from the XML Zonelist and test Name Servers are OK

.DESCRIPTION
Test DNS and WINS Servers are fuctionsing OK

Examples:
Test-DNSServer -Zonelist $Zonelist 

.PARAMETER Zonelist
XML Structure for the DNS Zones and Servers to query

.OUTPUT
An Updated XML structure containing the results of Name Server Tests

#>
{
	[CmdletBinding()]
	param
	( 
		[Parameter(Position=0)][String]$ZoneFile=$ZoneFile
	)
	
	$Zonelist = New-Object -Type System.XML.XMLDocument
	$Zonelist.Load($ZoneFile)
	
	# check the WINS NameServers - Non-qualified names ONLY
	Foreach ($NS in $Zonelist.AZDNS.WINS.NS)
	{
		Write-Verbose "Testing WINS Server $($NS.name) $($NS.IPAddress)"
		$Result = Test-NetConnection -Computer $NS.Name
		if ($Result.PingSucceeded)
		{
			$NS.IPAddress = $Result.RemoteAddress.ToString()
			if (Resolve-DNSName -Name $NS.name -Server $NS.IPAddress -LlmnrNetbiosOnly -erroraction "Silentlycontinue") {$NS.Tested ="OK"} else {$NS.Tested="WINSError"}
		}
		else
		{
			$NS.Tested = "NoPing"
		}
	}
	
	# check the Preferred DNS NameServers
	
	Foreach ($NS in $Zonelist.AZDNS.DNS.NS)
	{
		Write-Verbose "Testing Preferred DNS Server $($NS.fqdn) $($NS.IPAddress)"
		$Result = Test-NetConnection -Computer $NS.fqdn
		if ($Result.PingSucceeded)
		{
			$NS.IPAddress = $Result.RemoteAddress.ToString()
			if (Resolve-DNSName -Name $NS.fqdn -Server $NS.IPAddress -DNSOnly -erroraction "Silentlycontinue") {$NS.Tested ="OK"} else {$NS.Tested="DNSError"}
		}
		else
		{
			$NS.Tested = "NoPing"
		}
	}	
	
	Foreach ($Domain in $ZoneList.AZDNS.Domain)
	{
		Write-Verbose "Testing $($Domain.Name)"
		foreach ($NS in $Domain.NS)
		{
			$Result = Test-NetConnection -Computer $NS.fqdn
			if ($Result.PingSucceeded)
			{
				$NS.IPAddress = $Result.RemoteAddress.ToString()
				if (Resolve-DNSName -Name $NS.fqdn -Server $NS.IPAddress -DNSOnly -erroraction "Silentlycontinue") {$NS.Tested ="OK"} else {$NS.Tested="DNSError"}
			}
			else
			{
				$NS.Tested = "NoPing"
			}
		}
	}
	#Then check if each Zone is resolving
	Foreach ($domain In $Zonelist.AZDNS.Domain)
	{
		#Make sure there is a valid Name Server
		$NSrec = $Domain.NS | Where-object {$_.Tested -eq "OK"} | select-object -First 1
		if ($NSRec)
		{
			Write-Verbose "Checking Zones "
			foreach ($zone in $Domain.Zone)
			{
				write-verbose "Trying zone $($Zone.Name)"
				$Zonetest = Resolve-DNSname -name $zone.Name -server $NSRec.IPAddress -ErrorAction Silentlycontinue -QuickTimeout -DNSOnly
				if ($ZoneTest) {$Zone.Active="True"} else {$Zone.Active="False";Write-Host "$($Zone.Name) Inactive"}
			}
		}
		else
		{
			Write-Verbose "No Name server for Domain $($Domain.Name)"
		}
	}
	$zonelist.AZDNS.lastTested = $(Get-date -Format s).ToString()
	Return $Zonelist
}


Function Get-NSLookup
<#
.SYNOPSIS
Takes a list of hostnames (FQDN and NQDNs allows) and performs an NSLookup using
Resolve-DNSName. The list is split into FQDNs and NQDNs. FQDNs are queried directly
NQDNs are searched against all zones using the Zonelist

.DESCRIPTION
Peforms multiple NSLookups (Resolve-DNSName Fuction) against multiple DNS Zones
returning DNS records

Examples:
Do-NSLookup -HostName $Hosts -Zonelist $Zonelist 

.PARAMETER Hostname
Unqualifiled Host name(s) to resolve
.PARAMETER Ping
Optionally Ping IPaddresses
.PARAMETER Zonelist
XML Structure for the DNS Zones and Servers to query

.OUTPUTS
[PSCustomObject] containing the Fwd Name data 
#>
{	
	[CmdletBinding()]
	Param
	(
		[parameter(Position=1)][string[]]$Hostname,
		[Switch]$Ping,
		[XML]$Zonelist=$Zonelist
	)
	
	$begin = Get-date
	$HostPattern = "^([A-Za-z0-9]+(?:[A-Za-z0-9_\-]+)*)(?:\.[A-Za-z0-9]+(?:[A-Za-z0-9]+)*)*$"
	$Hostnames = $Hostnames | Sort-Object -Unique | Where-Object ($_ -Match $HostPattern)
	#Split the Qualified and Non Qualified names into 2 lists
	$FQDNs = @($Hostname | Where-object {$_ -Match "\."})
	$NQDNs = @($Hostname | Where-object {$_ -NotMatch "\."})
	
	Write-Host "Starting NSLookup on $($FQDNs.count) FQDNs and $($NQDNs.count) NQDNs"
	
	$FQDNAnswer=@()
	$NQDNAnswer=@()
	if ($FQDNs.Count -gt 0) 
	{
		$DNSServers =  @($Zonelist.AZDNS.DNS.NS | Where-object {$_.Tested -eq "OK"}  | select-object -First 1)
		If ($DNSServers.Count -ne 0)
		{
			Foreach ($NS in $DNSServers)
			{
				#Write-Host "Dispatching FQDN DSN Lookup for $($FQDNs.Count) Hosts Server $NS..."
				$FQDNAnswer = @(Invoke-Command -Scriptblock $Script:NSLookupScript -Argumentlist $FQDNs,@(),$NS.IPAddress)
				#$Param = @{Hostnames=$FQDNs;Zones=@();DNSServer=$NS.IPAddress}
				#$DNSJobs += Start-RunspaceJob -jobName "FQDN Lookup Slice $($Slice)" -RunspacePool $RSpool -ScriptBlock $NSLookupScript -Parameter $Param
			}
		}
		else
		{
			Write-Warning "Warning: None of the DNS Servers in XML Zonelist are available"
		}
	}
	
	if ($NQDNs.Count -Gt 0)
	{
		Foreach ($domain In $Zonelist.AZDNS.Domain)
		{
			#Make sure there is a valid Name Server
			$DNSServers =  @($Domain.NS | Where-object {$_.Tested -eq "OK"}  | select-object -First 1)
			#Write-Host "Domain $($Domain.Name) has $($DNSServers.Count) Name Servers available"
			if ($DNSServers.Count -ne 0)
			{
				Foreach ($NS in $DNSServers)
				{
					#Write-Host "Dispatching NQDN DSN Lookup $($NQDNs.Count) hosts against Zone Group $($Domain.Name) ..."
					$Zones = @($Domain.Zone | Where-Object {$_.active -match "True"} | Foreach-Object {$_.Name})
					$NQDNAnswer += Invoke-Command -Scriptblock $Script:NSLookupScript -Argumentlist $NQDNs,$Zones,$NS.IPAddress
				}
			}
			else
			{
				Write-Warning "No DNS Name server for Domain $($Domain.Name) in XML Zonelist available "
			}
		}
		$WINSServers =  @($Zonelist.AZDNS.WINS.NS | Where-object {$_.Tested -eq "OK"} | select-object -First 1)
		if ($WINSServers.Count -ne 0)
		{
			# Test all the NQDNs directly against WINS
			Foreach ($NS in $WINSServers)
			{
				#Write-Host "WINS Lookup for $($NQDNs.Count) Hosts Server $NS.IPAddress..."
				$NQDNAnswer += Invoke-Command -Scriptblock $Script:WINSLookupScript -Argumentlist $NQDNs,$NS.IPAddress
			}
		}
		Else
		{
			Write-Warning "Warning: None of the WINS Servers in XML Zonelist are available"
		}
	}
	$NSLookup = @($FQDNAnswer+$NQDNAnswer | Sort-Object -Prop FQDN, IPAddress -Unique)
	
	if ($Ping)
	{
		$IP = $NSLookup | Where-Object {$_.IPAddress -ne ""} | Select-Object -Property IPAddress -Unique | Foreach-Object {$_.IPAddress}
		$PingData = Ping-Async -IPList $IP -MaxThreads 200
		$PingIndex = $PingData | Group-Object -Prop PingIP -AsHashTable -AsString
		$NSLookup | %{Add-Member -InputObject $_ -Type NoteProperty -Name PingData -Value $PingIndex.Item($_.IPAddress) -Force}
	}
	$runTime = "{0:F3}" -f $(New-TimeSpan -Start $begin).Totalseconds 
	Write-Host "Get-NSLookup Returned $($NSLookup.count) NameRecords in $Runtime Seconds"
	Return $NSLookup
} #End Get-NSLookup


Function Get-ReverseLookup
<#
.SYNOPSIS
Performs a reverse lookup on a list of IP addresses

.DESCRIPTION
Peforms a PTR DNS Lookup on  <IPAddress> and returns the output

Examples:
Get-ReverseLookup -IPAddress "156.71.11.11"

.PARAMETER IPAddress
IP Address to look up

.OUTPUTS

#>
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Position=0,Mandatory=$true)][String[]]$IPAddress,
		[XML]$Zonelist=$Zonelist
	)
	
	#Use the XML Zonelist $Zonelist to get the DNS Server to use for the Reverse Lookup
	$DNSServer =  $Zonelist.AZDNS.DNS.NS | Where-object {$_.Tested -eq "OK"}  | select-object -First 1
	if ($DNSServer) 
	{
		Write-Output "Starting Reverse Lookup on $($IPAddress.Count) Addresses @ $([Datetime]::now)"
		$PTRResults = $IPAddress | 
		Foreach-Object {$IP=$_;Resolve-DNSName -Name $_ -Type "PTR" -Server $DNSServer.IPaddress -DNSOnly -ErrorAction "SilentlyContinue" |
		Select-Object -Property @{Name="Question";Expression={$IP}},Name,@{Name="PTRHost";Expression={$_.NameHost.ToUpper()}},QueryType}
		Write-Output "Discovered $($PTRResults.Count) PTR Records from Reverse Lookups @ $([Datetime]::now)"
	}
	Return $PTRResults
} #End of Get-ReverseLookup


Function Get-DNSRecord
<#
.SYNOPSIS
Get DNS Resource Record for FQDN Using Get-DNSServerResourceRecord

.DESCRIPTION
Looks  up the DNS Resource record for FQDN

Examples:
Get-DNSRecord -FQDN test.emea.astrazeneca.net -Type A

.PARAMETER FQDN
Host name to look up

.OUTPUTS
DNS Resource record

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [String]$FQDN,
		[Parameter(Position=1,Mandatory=$true)] [String]$Type="A"
	)
	
	$FQDNRegex = "^(?<Host>\w+)\.(?<Zone>.*)$"
	
	$CheckFQDN = $FQDN -Match $FQDNRegex
	If ($CheckFQDN)
	{
		$Hostname = $Matches.Host
		$Zone = $Matches.Zone
	
		$SOA = Resolve-DNSName -Name $Zone -Type "SOA" -DNSOnly | Where-Object {$_.Type -eq "SOA"}
		if ($SOA)
		{
			$DNSServer = $SOA.PrimaryServer
			$DNSRec = Get-DNSServerResourceRecord -Name $hostname -Zone $Zone -Computername $DNSServer
		}
	}
	else
	{
		Write-Host "$($FQDN) is Not a valid FQDN"
	}
	$DNSRec
}

Function Remove-DNSRecord
<#
.SYNOPSIS
Remove a DNS Resource Record for FQDN Using Remove-DNSServerResourceRecord

.DESCRIPTION
Looks  up the DNS Resource record for FQDN

Examples:
Get-DNSRecord -FQDN test.emea.astrazeneca.net -Type A

.PARAMETER FQDN
Host name to look up

.OUTPUTS
DNS Resource record

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [String]$FQDN,
		[Parameter(Position=1,Mandatory=$true)] [String]$Type="A"
	)
	
	$FQDNRegex = "^(?<Host>\w+)\.(?<Zone>.*)$"
	
	$CheckFQDN = $FQDN -Match $FQDNRegex
	If ($CheckFQDN)
	{
		$Hostname = $Matches.Host
		$Zone = $Matches.Zone
	
		$SOA = Resolve-DNSName -Name $Zone -Type "SOA" -DNSOnly | Where-Object {$_.Type -eq "SOA"}
		if ($SOA)
		{
			$DNSServer = $SOA.PrimaryServer
			$DNSRec = Get-DNSServerResourceRecord -Name $hostname -Zone $Zone -Computername $DNSServer
			if ($DNSrec) {$DNSRec | remove-DNSServerResourceRecord -Confirm -Zone $Zone -Computername $DNSServer}
		}
	}
	else
	{
		Write-Host "$($FQDN) is Not a valid FQDN"
	}
	$DNSRec
}

<# ===========================================================================================
End of Module Functions

Define and Export Module Members below
==============================================================================================
#>

if (-NOT $Zonelist) 
{
# Load the Default Zonelist
	Write-Host "Loading Global DNS ZoneList from XML file $($ZoneFile)"
	[XML]$Zonelist = Get-Content $ZoneFile
	$DateTest=0
	If ([DateTime]::TryParse($Zonelist.AZDNS.LastTested,[Ref]$DateTest)) 
	{
		$LastTested = [DateTime]$Zonelist.AZDNS.LastTested
		$age = $([DateTime]::Now - [DateTime]$Zonelist.AZDNS.LastTested).TotalHours
		$Retest = ($Age -gt 5.0) 
	}
	else
	{
		$Retest = $true
	}
	if ($Retest) 
	{
		#Check DNS Servers and update XML structure
		$Zonelist = Test-Zonelist -Zonefile $ZoneFile
		$Zonelist.Save($ZoneFile)
	}
}

$ModuleFunctionNames = Get-Content $ModuleFile | Where-Object {$_ -Match "^\s*Function\s+(\S+)" } | Foreach-Object {$Matches[1]}
Export-Modulemember -Function $ModuleFunctionNames

# Uncomment this line to Export all Functions including those Dot Sourced from Libraries
# Export-Modulemember -Function "*"

Export-Modulemember -Variable @("Zonelist","Zonefile")