<# ===========================================================================================
SPADModule

Active Directory Related Powershell Functions and Scriptblocks

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

#$ZoneFile = "\\ukmwemrpt1\decom\DNS\Zonelist.xml"

# Informational

Write-Verbose "Loading Module $($ModuleName)"
Write-Verbose "OSVersion $($OSVersion.toString())"

<# ===========================================================================================
Check and Load any Dependent modules required by this module Here

ABORT Module if any Modules are Missing
==============================================================================================
#>

#$RequiredModules = @("SPNetTools")
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

# Optionally Specify additional Libraries as as Array of Pathnames in $IncludeScripts
$IncludeScripts = @()
$IncludeScripts | Foreach-Object {if (Test-Path -Path $_) {Write-Verbose "Including Lib $_" ; . $_} else {Write-Warning "Cannot locate Library script $_"}}

<# ===========================================================================================
Module Code - These fuctions are Exported from the Module
==============================================================================================
#>


Function Get-ADSubnets
<#
.SYNOPSIS
Return the Sites and Subnets from the Active Directory

.DESCRIPTION
Searches teh Actioved Direrctory and returns all the Site and Subnet objects as a 
PSCustomObject

Examples:
$Sites = Get-ADSubnets

.PARAMETER Forest
Specify the Forest name. Defaults to the current Foresrt


.OUTPUTS
[PSCustomObject] containing Site Data

#>
{

	Param
	(
		[String]$Forest
	)
	
	$SiteData = @()
	if ($Forest)
	{
		$RootPath = "LDAP://$($Forest)/RootDSE"
	}
	else
	{
		$Rootpath = "LDAP://RootDSE"
	}
	$RootDSE=New-Object System.DirectoryServices.DirectoryEntry($RootPath)
	$SearchRoot = "LDAP://" + $RootDSE.configurationNamingContext
	$AllSites = Query-AD -SearchRoot $SearchRoot -Query "(ObjectClass=Site)"

	$SiteInfo = @()
	
	Foreach ($Site in $AllSites)
	{
		#Get the SubNet CIDR from the SiteObjectbl Property
		$Name=$site.Name.ToUpper()
		If ($Site.Location -Match "\/")
		{
			$Delim = "\/"
		}
		Elseif ($Site.Location -Match ",")
		{
			$Delim = ","
		}
		Else 
		{
			$Delim = $Null
		}
		If ($Delim)
		{
			$Locations = $Site.Location.Split($Delim)
			$C=$Locations[0] #Country
			$S=$Locations[1] #Site
			$L=$Locations[2] #Location
		}
		Else
		{
			$C = ""
			$S = ""
			$L = ""
		}

		$Subnets = if ($Site.Siteobjectbl) {$Site.Siteobjectbl | Foreach-Object {$_.Split(',')[0].Split("=")[1]}} else {$Subnets=@()}
		Foreach ($Subnet in $Subnets)
		{
			#Use the CIDR to calculate the IPAddress Ranges and Integer ranges for the Subnet
			$SubNetStart = $Subnet.Split("/")[0]
			$Mask = $Subnet.Split("/")[1]
			$iSNStart = ConvertTo-DecimalIP $SubNetStart
			[Int]$iSize = [Math]::Pow(2,32-$Mask) -1
			$iSNEnd=$iSNStart + $iSize
			$SubNetEnd = ConvertTo-DottedIP $iSNEnd
			$SiteInfo += [PSCustomObject]@{
				Site=$Name
				Country=$C
				SiteName=$S
				Location=$L
				Path=$Site.Location
				whenChanged=$Site.WhenChanged
				CIDR=$Subnet
				FirstIPAddress=$SubNetStart
				LastIPAddress=$SubNetEnd
				iSNStart=$iSNStart
				iSNEnd=$iSNEnd
				iSize=$iSize}
		}
	}
	
	Return $SiteInfo
} # End of Get-ADSubnets


Function Find-ADObject
<#
.SYNOPSIS
Search AD via Global Catalog for objects matching Name

.DESCRIPTION
Searchj for object of the ObjectClass specified where cn or SAMAccountname match Name
Wildcards are allowed in Name

Examples:
Find-ADObject -Name Kvtl* -Objecttype User

.PARAMETER Name
Object name used to match cn or SAMAccountName attributes

.PARAMETER ObjectType
User (Default) Group or Computer

.PARAMETER $Forest
Forest name in Dns format (ie medimmune.com)

.OUTPUTS
[DirectoryEntry] LDAP objects matching

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] [String]$Name,
		[ValidateSet("User","Computer","Group")][String]$ObjectType="User",
		[String]$Forest
	)

	if ($Forest)
	{
		$ForestRoot = [String]::Join(",",$($forest.Split(".") | % {"DC="+$_}))
	}
	else
	{
		$ForestRoot = ([ADSI]"LDAP://RootDSE").rootDomainNamingContext
	}
	#Global Catalog Search
	$Searchroot = new-object System.DirectoryServices.DirectoryEntry("GC://" +$ForestRoot)
	$filter="(&(objectclass=$ObjectType)(|(cn=$Name)(SAMAccountName=$Name)))"
	$props = @("distinguishedName","SAMAccountName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($Searchroot,$filter,$props)
	$GCObj = $Searcher.FindAll()
	if ($GCObj)
	{
		$ADObj = $GCObj | Foreach-Object {[ADSI]"LDAP://$($_.Properties.distinguishedname)"}
	}
	else
	{
		$ADObj = $Null
	}
	$ADObj	
} # End Of Find-ADObject


<#
.SYNOPSIS
Search the Forest for computer accounts matching Name via the Global Catalog

.DESCRIPTION
Looks  up Computer objects in the forest returning key properties 

Examples:
Get-ADComputerProperties -Name Server1 

.PARAMETER Name
Object Name
.PARAMETER Forest
Forest name in Dotted Notation
.OUTPUTS
[PSCustoObject] containing the Computer Account properties

#>
Function Get-ADComputerProperties
{
	Param($Name, $Forest)

	#Search for a single object from the Forest Root via the Global Catalog
	#Return the LDAP Directory Object and not the one from the GC so all attributes are available if needed

	$PropertyList = @("distinguishedName","cn","dNSHostName","userAccountControl","pwdLastSet","operatingSystem","operatingSystemVersion","operatingSystemServicePack")

	if ($Forest)
	{
		$root = [String]::Join(",",$($forest.Split(".") | % {"DC="+$_}))
	}
	else
	{
		$root = ([ADSI]"LDAP://RootDSE").rootDomainNamingContext
	}
	#Search Global Catalog
	$Searchroot = new-object System.DirectoryServices.DirectoryEntry("GC://" +$root)
	$filter = "(&(objectclass=computer)(cn=$name))"
	$props = @("distinguishedName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($searchroot,$filter,$props)
	$GCObj = $Searcher.Findall()

	if ($GCObj) 
	{
		#At least One Object returned
		#Write-Verbose "Getting LDAP Object for $($Name) - GCObjects = $($GCObj.Count)"
		$LDAPObj = $GCObj | foreach-object {$LDAPPath = $_.Path -Replace "GC:","LDAP:"; new-object System.DirectoryServices.DirectoryEntry($LDAPPath)}
		$ADInfo = @()
		Foreach ($DN in $LDAPObj)
		{
			$CompInfo = New-Object -type PSObject
			#$CompObj = New-Object -Type PSObject -Property @{TestedIP=$IPAddress;NBTNode=$false;NBTName="";NBTDomain="";MAC=""}
			foreach ($Property in $PropertyList)
			{
				if ($Property -eq "pwdLastSet")
				{
					$Value = [Datetime]::Fromfiletime($DN.ConvertLargeIntegertoInt64($DN.PwdLastSet.value))
				}
				elseif ($DN.Properties.Item($Property).count -gt 1)
				{
					$Value =[String]::Join(";",$DN.Properties.Item($Property).Value)
				}
				else
				{
					$Value = $DN.Properties.Item($Property).Value
				}
				Add-Member -InputObject $CompInfo -Type Noteproperty -Name $Property -Value $Value
			}
			$ADInfo += $CompInfo
		}
	}
	Return $ADInfo
} # End Of Get-ADComputerProperties



Function get-FQDNfromDN
<#
.SYNOPSIS
Construct an FQDN from the DN

.DESCRIPTION
Construct an FQDN from the DN

.PARAMETER DN
Distinguished Name

.OUTPUTS
DNS Style FQDN

#>
{
	Param
	(
		$Dn
	)
	
	$DNParts= $dn.toUpper().Split(",")
	$NQDN = $DNParts[0].Split("=")[1]
	$Zone = [String]::Join(".",($DNParts | Where-Object {$_ -Match "DC="} | Foreach-Object {$_.Split("=")[1]}))
	$FQDN="$($NQDN).$($Zone)"
	Return $FQDN
} # End Of Get-FQDNFromDN

Function Query-AD
<#
.SYNOPSIS
Searches the Forest to find Objects

.DESCRIPTION
Uses the Global Catalog to perform an AD Search for Objects matching Name
Matching Items are returned via LDAP Paths

Examples:
$match = Find-ADObject -Name ksx* 

.PARAMETER SearchRoot
Relative DistinguishedName indicating where to root the Search

.PARAMETER Query
Query String

.PARAMETER AttribList
AttribList

.PARAMETER Scope
Scope


.OUTPUTS
Array of Direcory Entry objects

#>
{
	[CmdletBinding()]
	Param
	(
		[String]$SearchRoot="",
		[string]$Query, 
		[String]$AttribList="", 
		[String]$Scope="SubTree"
	)
	
	$SearchResults = @()
	$ADEntry = New-Object System.DirectoryServices.DirectoryEntry
	
	if ($SearchRoot -eq "")
	{
		$RootDSE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
		$SearchRoot = "LDAP://" + $RootDSE.DefaultNamingContext
	}
	
	$ADSearcher = New-Object system.DirectoryServices.DirectorySearcher
	if ($AttribList -ne "") {$scrap = $AttribList.Split(",") | ForEach-Object{$ADSearcher.PropertiesToLoad.Add($_)}}
	$ADSearcher.SearchRoot = $SearchRoot
	$ADSearcher.PageSize=1000
	$ADSearcher.Filter = $Query
	Write-Verbose "Query $Query"
	$Results = $ADSearcher.Findall()
	
	Foreach ($Result in $Results)
	{
		$ADObj = New-Object -TypeName System.Object
		#Add-Member -InputObject $ADObj -Type NoteProperty -name ADSPath -value $Result.Path
		Foreach ($item in $result.properties.propertynames)
		{
			#Write-Host "Adding Property $Item - $($Result.properties.item($item))"
			if ($Result.properties.item($item).count -eq 1)
			{
				add-member -inputObject $ADObj -Type NoteProperty -Name $item -value $Result.properties.item($item)[0]
			}
			else
			{
				add-member -inputObject $ADObj -Type NoteProperty -Name $item -value $Result.properties.item($item)
			}			
		}
		$SearchResults += $ADObj
	}
	$SearchResults
} # End of Query-AD


Function Get-GroupMembers
<#
.SYNOPSIS
Enumerate Group Members

.DESCRIPTION
Enumerate Group Members

Examples:
$member = Get-GroupMembers -DN $Group

.PARAMETER DN
Common Name or SAMAccountName

.OUTPUTS
Group Members

#>
{
	Param ($DN)

	$chunk = 999
	$count = 0

	$lowRange = -1 * ($Chunk + 1)
	$members = @()

	$script:finished = $false
	:GetChunk do
	{
		$count = $count + 1
		$lowRange = $lowRange + $chunk + 1
		$highrange = $lowRange + $chunk
		$range = "member;range=$lowRange-$highRange"
		$status = $DN.getInfoEx(@($range),0)
		Trap [system.exception]
		{
			$script:finished = $true
			write-Verbose "Exception Raised $script:finished"
			continue
		}
		If (-not $script:finished)
		{
			$slice = $dn.Get("Member")
			Write-verbose "Slice count $($slice.count)"
			if ($count -eq 1) {$members = $slice} else {$members = $members + $slice}
		}
	} until ($script:finished)

	Write-verbose "Member Count $($members.count)"
	$members
} # End of Get-GroupMembers


Function Get-TokenGroups
<#
.SYNOPSIS
Returns the Tokengroups attribute for a given user and translates
The SIDS to Group names

.DESCRIPTION
TokenGroups is a quick way of seeing the Groups a user is a member of without
having to process the nested Groups

Examples:
$groups = Get-TokenGrouops -DN $Group

.PARAMETER DN
DN of the User object

.OUTPUTS
Translated Tokens (Domain\Name)

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Position=0,Mandatory=$true)] $DN
	)

	$dn.GetInfoEX(@("tokengroups"),0)
	$tokens = $dn.getex("tokengroups")

	$grps = $Tokens | foreach-object {ConvertFrom-SID -sid $_ -Frombyte}
	$grps
} # End of Get-Tokengroups


<# ===========================================================================================
End of Module Functions

Define and Export Module Members below
==============================================================================================
#>

$ModuleFunctionNames = Get-Content $ModuleFile | Where-Object {$_ -Match "^\s*Function\s+(\S+)" } | Foreach-Object {$Matches[1]}
Export-Modulemember -Function $ModuleFunctionNames

# Uncomment this line to Export all Functions including those Dot Sourced from Libraries
# Export-Modulemember -Function "*"

