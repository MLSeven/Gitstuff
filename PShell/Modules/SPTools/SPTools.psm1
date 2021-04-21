<# ===========================================================================================
Powershell Module

Name		:SPTools.psm1


S Potts

Module Header - Executed on Import
==============================================================================================
#>
# Script Environment
$ModuleFile = Get-Item $myinvocation.Mycommand.Path
$ModuleRoot = Split-Path -Path $myinvocation.Mycommand.Path -Parent
$ModuleName = $ModuleFile.BaseName
$OSVersion = [Environment]::OSVersion.Version

Write-Verbose "Loaded Module $($ModuleName)"
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
$IncludeScripts = @()
$IncludeScripts | Foreach-Object {if (Test-Path -Path $_) {Write-Verbose "Including Lib $_" ; . $_} else {Write-Warning "Cannot locate Library script $_"}}


<# ===========================================================================================
Inculde Functions Difined AND Exported by This Module
==============================================================================================
#>

Function Get-Clipboard
<#
.SYNOPSIS
Returns the contents of the clipboard in various formats

.DESCRIPTION
Returns the contents of the clipboard in various formats. Options include
	Text (Default)
	String
	CSV
	XML

Examples:
$Clip = Get-Clipboard -Text

Returns the clipboard as an array of Strings

.PARAMETER Text
(Default) Array of String objects (blank lines removed).

.PARAMETER Raw
Single String objected caontaining the unmodified Text

.PARAMETER CSV
CSV Object - First row contains Headers - delimiter specified by parameter -Deliminator

.PARAMETER Deliminator
Single String objected caontaining the unmodified TextDeliminator characher for CSV Object
Default is TAB

.PARAMETER XML
XML Document

.OUTPUTS
Depends on Parameter

#>
{
	[CmdletBinding()]
	Param 
	(
		[ValidateSet("Text","Raw","CSV","XML")][String]$Type="Text",
		[String]$Delim="`t"
	)	
	
	if([threading.thread]::CurrentThread.GetApartmentState() -eq 'MTA')
	{
		Write-Warning "MTA Threading Detected !!!"
	}
	
	Add-Type -Assembly PresentationCore
	$Clip = [Windows.Clipboard]::GetText()
	if ($Clip)
	{
		Switch ($Type)
		{
			"Text"
			{
				$Output = $Clip -Split '\r\n' | Where-Object {(($_.Trim()).Length -gt 0)} | Foreach-Object {$_.Trim()}
			}
			"Raw"
			{
				$Output = $Clip
			}
			"CSV"
			{
				$Lines = $Clip -Split '\r\n' | Where-Object {(($_.Trim()).Length -gt 0)} | Foreach-Object {$_.Trim()}
				$Output = $Lines | Convertfrom-CSV -Delim $delim 
			}
			"XML"
			{
				$Output = New-Object -Type System.XML.XMLDocument
				try
				{
					$Output.LoadXML($Clip)
				}
				catch
				{
					Write-Warning "Clipboard does not contain a Well formed XML document"
					$output = $Null
				}
			}	
		}
	}
	Else
	{
		Write-Warning "Nothing available from the Clipboard - Try again !!!"
		$output= $Null
	}
	$Output
}


Function Get-ScriptFunctions
<#
.SYNOPSIS
Get a list of Function names from a Visual Basic Script

.DESCRIPTION

Get a list of Function names from a Visual Basic Script

.PARAMETER ScriptName
(Default) Array of String objects (blank lines removed).

.PARAMETER Recurse
Single String objected caontaining the unmodified Text

.Parameter Language
Degfault =PS, can ve VB

.OUTPUTS
Depends on Parameter

#>
{
	[CmdletBinding()]
	Param
	(
		[String[]]$ScriptName,
		[String]$language="PS",
		[Switch]$Recurse
	)
	
	if ($Language -EQ "PS")
	{
		$Pat = "^\s*Function\s+(\w+-\w+)"
	}
	else
	{
		$Pat = "^\s*(?i:(?:(?:PRIVATE|PUBLIC)\s+)?(?:STATIC\s+)?(?:SUB|FUNCTION|PROPERTY))\s+([a-zA-Z0-9_\-]+)?(?:\(|$)"
	}
	
	if ($Recurse)
	{
		$Results = Get-Childitem -Recurse $Scriptname | Select-String -Pattern $Pat -AllMatches
	}
	else
	{
		$Results = Select-String -Path $ScriptName -Pattern $pat -AllMatches
	}
	$Output = $Results | Foreach-Object{[PSCustomObject][Ordered]@{Filename=$_.Filename;Path=$_.Path;LineNumber=$_.LineNumber;Function=$_.Matches.Groups[1].Value}}
	$Output
}

<#
.SYNOPSIS
Exports an XML version of A PSObject

.DESCRIPTION
Exports an XML version of A PSObject

.PARAMETER Object
ObjectName

.PARAMETER Path
XML File Path

.PARAMETER DocumentName
XML File Name

.PARAMETER ObjectName
Display name for the object

.OUTPUTS

#>
Function Export-PSObject

{
	[CmdletBinding()]
	param( 
	[Parameter(Position=0, Mandatory=$True)]$Object,
	[Parameter(Position=1, Mandatory=$True)][String]$Path, 
	[Parameter(Position=2, Mandatory=$false)][String]$DocumentName,
	[Parameter(Position=3, Mandatory=$false)][String]$ObjectName
	)
	
	$xml = New-Object System.Xml.XmlTextWriter($Path,$null)
	$Xml.Formatting = "indented"
	$xml.Indentation = 1
	$xml.IndentChar = "`t"
	$xml.WriteStartDocument()
	# set XSL statements
	#$xml.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")
 
	$xml.WriteComment("XML Export for Object name $($ObjectName)")
	$xml.WriteStartElement($DocumentName)
	$xml.WriteAttributeString("ExportedOn", $(Get-Date))
 
	# Export the Objects
	Foreach ($Item in $Object)
	{
		Write-XMLElements -XMLWriter $xml -Object $Item -Name $ObjectName
	}

	# close the "machines" node:
	$xml.WriteEndElement()
	 
	# finalize the document:
	$xml.WriteEndDocument()
	$xml.Flush()
	$xml.Close()
}
 
 
Function Write-XMLElements

{
	[CmdletBinding()]
	param( 
	[Parameter(Position=0, Mandatory=$True)]$XMLWriter,
	[Parameter(Position=1, Mandatory=$True)]$Object,
	[Parameter(Position=2, Mandatory=$True)]$Name
	)
	
	Foreach ($Item in $object)
	{
		Write-Host "Starting Element $($Name)"
		$XMLWriter.WriteStartElement($Name)
		Foreach ($Property in $Item.PSObject.Properties)
		{
			If ($Property.Value.Gettype().Name -match "Object")
			{
				Write-Host "Nested Object Element $($Property.Name)"

				Write-XMLElements -XMLWriter $XMLWriter -Object $Property.Value -Name $Property.Name
			}
			Else
			{
				Write-Host "Object Element $($Property.Name) = $($Property.Value)"
				$XMLWriter.WriteElementString($Property.Name,$Property.Value)
			}
		}
		Write-Host "Ending Element $($name)"
		$XMLWriter.WriteEndElement()
	}
}



<# ===========================================================================================
End of Module Functions

Define and Export Module Members below
==============================================================================================
#>

$ModuleFunctionNames = Get-Content $ModuleFile | Where-Object {$_ -Match "^\s*Function\s+(\S+)" } | Foreach-Object {$Matches[1]}
Export-Modulemember -Function $ModuleFunctionNames

# Uncomment this line to Export all Functions including those Dot Sourced from Libraries
# Export-Modulemember -Function "*"
