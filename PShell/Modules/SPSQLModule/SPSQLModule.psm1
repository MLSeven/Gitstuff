<# ===========================================================================================
Powershell Library of SQL Related Functions

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

$DBstr = "Server=UKMCWRAASD01;Database=Discovery;Trusted_Connection=yes"
$EOSL = "Server=SESKSHRSQLDEV01;Database=EOSL;Trusted_Connection=yes"

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
Get-Childitem $ModuleLib -Filter *.ps1 | Foreach-Object {. $_.fullname}

# Optionally Specify additional Libraries as as Array of Pathnames in $IncludeScripts
$IncludeScripts = @()
$IncludeScripts | Foreach-Object {if (Test-Path -Path $_) {Write-Verbose "Including Lib $_" ; . $_} else {Write-Warning "Cannot locate Library script $_"}}

<# ===========================================================================================
Module Code - These fuctions are Exported from the Module
==============================================================================================
#>

Function Invoke-Sqlcmd 
<#
.SYNOPSIS
Generic function for executing SQL commands

.DESCRIPTION
Generic function for executing SQL commands

.PARAMETER ConnectionString
DB Connectionstring
.PARAMETER ServerInstance
Server Name (if not using ConnectionString)
.PARAMETER Database
Database Name (if not using ConnectionString)
.PARAMETER Query
SQL Query to be executed
.PARAMETER Username
User Name (if not using ConnectionString)
.PARAMETER Password
Password (if not using ConnectionString)
.PARAMETER QueryTimeout
Timeout (default 600s)
.PARAMETER ConnectionTimeout
ConnectionTimeout (Default 15sec)
.PARAMETER InputFile
SQL Scripot file
.PARAMETER ExecOnly
Executes query but does not expect any results
.PARAMETER AS
Objects can be returned AS

Dataset
DataTable
DataRow

.OUTPUTS
Returns objects in Format specified by Parameter AS (Default DataRow)
if -EXECOnly the no data is returned

#>
{ 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$false)] [string]$ConnectionString, 
    [Parameter(Position=1, Mandatory=$false)] [string]$ServerInstance, 
    [Parameter(Position=2, Mandatory=$false)] [string]$Database, 
    [Parameter(Position=3, Mandatory=$false)] [string]$Query, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=5, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$QueryTimeout=600, 
    [Parameter(Position=7, Mandatory=$false)] [Int32]$ConnectionTimeout=15, 
    [Parameter(Position=8, Mandatory=$false)] [ValidateScript({test-path $_})] [string]$InputFile,
	[Parameter(Position=9, Mandatory=$false)] [Switch]$ExecOnly,
    [Parameter(Position=10, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As="DataRow" 
    ) 
 
 	#Create a connection object
    $conn=new-object System.Data.SqlClient.SQLConnection 
	#Connect to DB
	if ($Connectionstring)
	{
		$conn.ConnectionString=$ConnectionString
	}
	else
	{
		if ($Username) 
		{
			$ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout
		}
		else 
		{
			$ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout
		}
		$conn.ConnectionString=$ConnectionString 
     }
	 
    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller
    if ($PSBoundParameters.Verbose) 
    { 
        $conn.FireInfoMessageEventOnUserErrors=$true 
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {Write-Verbose "$($_)"} 
        $conn.add_InfoMessage($handler) 
    } 

    $conn.Open()
	$ReturnedData = @()
	#Execute SQL Commands
	if ($InputFile) 
    {
		#T-SQL Command Script - Split into Batches using GO operator
        #$filePath = $(resolve-path $InputFile).path 
        $TSQL =  [System.IO.File]::ReadAllText($InputFile)
		#Divide the SQL into Batches
		$Batches = $TSQL -Split "\s+GO\s+"
		$BatchID = 0
		Foreach ($Batch in $Batches)
		{
			if ($batch -NotMatch "\S")
			{
				#batch contains only white space - no action to take
				write-host -Foregroundcolor magenta "Batch $BatchID contains no SQL - Skipping .. "
			}
			else
			{
				Write-Host -Foregroundcolor magenta "--- Starting Batch $BatchID ----------------------------------"
				Write-Host -Foregroundcolor magenta "Executing T-SQL Commands for Batch $BatchID "
				Write-Host -Foregroundcolor magenta ""
				if ($ExecOnly)
				{
					#Use -ExecOnly to run SQL Scripts that do not return any Data
					$ds=$Null
					Write-Host -Foregroundcolor magenta "Invoke-SQLCmd -ExecOnly Batch will NOT return Data"
					$cmd=new-object system.Data.SqlClient.SqlCommand($Batch,$conn)
					$cmd.CommandTimeout=$QueryTimeout 
					$Rowcount = $Cmd.ExecuteNonQuery()
					Write-Host "ExecOnly Batch $($BatchID) Query results : $($Rowcount) Rows affected"
				}
				Else
				{
					#Script produces output so capture
					Write-Host -Foregroundcolor magenta "Invoke-SQLCmd DataSet Mode - Batch will return Data if applicable"
					$cmd=new-object system.Data.SqlClient.SqlCommand($Batch,$conn)
					$cmd.CommandTimeout=$QueryTimeout
					$ds=New-Object system.Data.DataSet
					$ds.DataSetName = "Batch $BatchID"
					$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 
					[void]$da.fill($ds)
					if ($ds.Tables.Count -gt 0)
					{
						$ReturnedData += $ds
					}
				}
			}
			Write-Host -Foregroundcolor magenta "--- End Of Batch $BatchID ----------------------------------------"
			$BatchID += 1
			
		}	
    }
	else
	{
		#Execute the SQL Query String passed as a Parameter
		if ($query)
		{
			Write-Host -Foregroundcolor green "Invoke-SQLCmd Command Line Query Mode"
			Write-Host -Foregroundcolor green "-------------------------------------"
			write-Host -Foregroundcolor green $Query
			Write-Host -Foregroundcolor green "-------------------------------------"
			$cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn) 
			$cmd.CommandTimeout=$QueryTimeout
			if ($ExecOnly)
			{
				$Rowcount = $Cmd.ExecuteNonQuery()
				Write-Host "Invoke-SQLcmd -ExecOnly CmdLine Query : $($Rowcount) Rows affected"
			}
			Else
			{
				$ds=New-Object system.Data.DataSet
				$ds.DataSetName = "CmdLine Query"
				$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 
				[void]$da.fill($ds)
				$ReturnedData += $ds
			}
		}
	}
    $conn.Close()
	
	if ($ReturnedData) 
	{
		#Analyse the returned Data
		Write-Host ""
		$DSCount = $ReturnedData.Count
		Write-Host "Invoke-SQLCMD returned $DSCount DataSets"
		Foreach ($DataSet in $ReturnedData)
		{
			Write-Host "Returned Dataset $($DataSet.DataSetName)"
			Write-Host "DataSet contains $($DataSet.Tables.Count) Tables"
			$DataSet.Tables | Foreach-Object {Write-Host "Dataset Table $($_.Tablename) has $($_.Rows.count) Rows"}
			Write-Host ""
		}
		if ($DSCount -eq 1 -AND $ReturnedData[0].Tables.Count -eq 1)
		{
			Return $ReturnedData[0].Tables.Item(0)
		}
		else
		{
			# Multiple datasets returned - pass back everything
			Return $ReturnedData
		}
	}
	else
	{
		Write-Host "Invoke-SQLCMD has no data to return..."
		Return $null
	}	
} #End of Invoke-Sqlcmd


Function Import-Clipboard
<#
.SYNOPSIS
Takes the contents of the Clipboard and Imports this into a CSV. 
The CSV is then uploaded to the SQL Table. Required Get-Clipboard Function

.DESCRIPTION
Takes the contents of the Clipboard and Imports this into a CSV.

Import-Clipboard -DBStr $DBStr -Tablename tblTest -Append

Will try and append the clipboard to table tblTest (formats must be the same)

.PARAMETER DBStr
DB Connection string

.PARAMETER TableName
DB TableName

.PARAMETER BatchID
Optional BatchID

.PARAMETER Append
Overwrite or Append to table

.OUTPUTS
Returns no Data

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Mandatory=$true)] [String]$DbStr, 
		[Parameter(Mandatory=$true)] [String]$TableName,
		[String]$batchID = [DateTime]::Now.Ticks.ToString(),
		[Switch]$Append
	)

	$CSV = Get-Clipboard -Type CSV
	If ($CSV) 
	{
		Write-Host "Pasted $($CSV.Count) Rows from Clipboard - importing to SQL ..."
		if ($Append)
		{
			Importto-SQL -DBStr $DBStr -CSVobject $CSV -TableName $Tablename -BatchID $BatchID 
		}
		else
		{
			Importto-SQL -DBStr $DBStr -CSVobject $CSV -TableName $Tablename -BatchID $BatchID -AddId -ForceNew
		}
	}
} # End of Import-Clipboard


Function ImportTo-SQL
<#
.SYNOPSIS
Performs a Bulk import of a DataTable object into a SQL Table

.DESCRIPTION
Performs a Bulk import of a DataTableobject into a SQL Table

Example 

Importto-SQL -DBStr $DBstr -Tablename tblTest -DTObject $Data

Appends the contents of the DataTable object into table tblTest

Importto-SQL -DBStr $DBstr -Tablename tblTest -CSVObject $CSV -Force

Appends the contents of the DataTable object into table tblTest

.PARAMETER DBStr
Source Object to be used to create the table

.PARAMETER Tablename
SQL Table name

.PARAMETER DTObject
DataTable Object containing the data

.PARAMETER ForceNewTable
SWITCH parameter - Forces existing table to be dropped and recreated

.PARAMETER AddID

Adds ID and BatchID columns for managing updates

.OUTPUTS
No Outputs

#>
{
	[CmdletBinding()]
	Param 
	(
		[String]$DbStr, 
		[PSCustomObject[]]$CSVObject=$Null,
		[System.Data.DataTable]$DTObject=$Null,
		[String]$TableName,
		[String]$batchID = [DateTime]::Now.Ticks.ToString(),
		[Switch]$ForceNewTable,
		[Switch]$AddID
	)
	
	if ($PSBoundParameters.ContainsKey("CSVObject"))
	{
		Write-Host "CSVObject Will be Imported"
		if ($CSVObject.GetType().IsArray) 
		{
			$InputType = $CSVObject[0].GetType().Name
			Write-Host "CSVObject is an Array of $($CSVObject[0].GetType().Name) Object containing $($CSVObject.Count) Items"
			$DTObject = $CSVObject | ConvertTo-DataTable
			Write-Host "CSVObject converted to DataTable Object containing $($DTObject.Rows.Count) Items"
		}
		else
		{
			$InputType = $CSVObject.GetType().Name
			Write-Host "CSVObject is a $($CSVObject.GetType().Name) Object containing $($CSVObject.Count) Items"
		}
	}
	elseif ($PSBoundParameters.ContainsKey("DTObject"))
	{
		Write-Host "DTObject Will Be imported"
		Write-Host "DTObject is a $($DTObject.GetType().Name) Object containing $($DTObject.Rows.Count) Row Items"
		$InputType = $DTObject.GetType().Name
	}
	Else
	{
		Write-Warning "No Objects specified - Use -CSVObject or -DTBoject to specify objects to import"
		return
	}
	
	
	if ($ForceNewTable) 
	{
		Write-Warning "Create New Table Mode Selected"
		Write-Host ""
		Write-Warning "WARNING - Any existing table will be Dropped ...."
		Write-Host ""
		Write-Host "Analysing Data Object to assess maximum Column Sizes ..."
		Write-Host ""
		Write-Host "Input Type: $InputType"
		if ($InputType -Match "PSCustomObject")
		{
			$Columns = Analyse-Object -InputObject $CSVObject
		}
		else
		{
			$Columns = Analyse-Object -InputObject $DTObject
		}
		Write-Host "Maximum Column sizes Detected"
		Write-Host "Dropping and Creating New SQL Table .."
		if ($AddID)
		{
			Write-Host "Adding Identity Column ID .."
			$NewTab = Create-SQLTableDDL -InputObject $DTObject -Tablename $TableName -Columndefs $Columns -AddID
		}
		else
		{
			$NewTab = Create-SQLTableDDL -InputObject $DTObject -Tablename $TableName -Columndefs $Columns
		}
		Write-Host ""
		#$NewTab | Foreach-Object {Write-Host -Foregroundcolor Green $_}
		Write-Host ""
		Write-Host "Running SQL to Create Table ... "
		Write-Host ""
		$Null=Invoke-SQLCMD -ConnectionString $DBStr -Query $NewTab -Verbose
		Write-Host "Table $($TableName) Created - Preparing Data Table ..."
		Write-Host ""
	}
	if ($TableName.Split(".").Count -gt 1)
	{
		$Owner = $TableName.Split(".")[0]
		$Table = $TableName.Split(".")[1]
	}
	else
	{
		$Owner = "dbo"
		$Table = $TableName
	}
	#Check if the table has a BatchID Column
	$GetCols = @"
	SELECT column_name 
	FROM information_schema.columns 
	WHERE table_name = '$($Table)' AND table_schema='$($Owner)' 
	ORDER BY ordinal_position
"@
	
	Write-Host "ImportTo-SQL: Checking if Table has BatchId"
	$TableCols = Invoke-SQLCmd -ConnectionString $DBStr -Query $getCols
	
	if ($TableCols)
	{
		#Table Exists
		$TableHasBatchID = ($($TableCols | Where-Object {$_.Column_Name -eq "BatchID"} | Measure-Object).Count -eq 1)
	
		if ($TableHasBatchID)
		{
			#Batch ID requires adding to this Import - Generate one from the Current Time
			#By Default, if not supplied then $batchID = [DateTime]::Now.Ticks.ToString()
			#Important -Table must have been created with BatchID column otherwise it will fail
			if (-NOT $DTObject.Columns.Contains("BatchID"))
			{
				Write-Host ""
				Write-Host "Adding BatchID $BatchID to the Data... "
				Write-Host ""
				$null = $DTObject.Columns.Add("BatchID","System.String")
			}	
			else
			{
				Write-Host ""
				Write-Host "BatchID Column already Exists in DataTable"
				Write-Host ""
			}
			Write-Host "Updating BatchID with value $BatchID"
			Write-Host
			$DTObject | Foreach-Object {$_.BatchID = $BatchID}
		}
		$ColNames = $DTObject.Columns | Foreach-object {$_.ColumnName}
		#$DT.Columns | % {Write-Host $_.ColumnName}
		Write-Host "DataObject Created with $($DT.Rows.Count) Rows and $($DTObject.Columns.count) Columns"
		Write-Host ""
		Write-Host "Starting Bulk Import ..."
		# If $AddID is specified we have to Explicitly Map Columns from the DataTable object
		
		BulkTransfer-ToSQL -DBStr $DBStr -TableName $TableName -TableData $DTObject -SourceCol $ColNames -TargetCol $ColNames
	}
	else
	{
		Write-Warning "Table $($TableName) Does not exists in the Database"
	}
}




Function ImportFrom-CSVFile
<#
.SYNOPSIS
Performs a Bulk import of a CSV object into a SQL Table

.DESCRIPTION
Performs a Bulk import of a CSV object into a SQL Table

Example 

ImportCSVto-SQL -DBStr $DBstr -Tablename tblTest -CSVObject $Data

Appends the contents of the CSV object into table tblTest

.PARAMETER DBStr
Database Connection string

.PARAMETER Tablename
SQL Table name

.PARAMETER CSVFile
CSV File Name

.PARAMETER ForceNew
SWITCH parameter - Forces existing table to be dropped and recreated

.PARAMETER ImportScript

Optionally runs the Import script file

.OUTPUTS
No Outputs

#>
{
	[CmdletBinding()]
	Param
	(
		[String]$CSVFile,
		[String]$Tablename,
		[String]$DBStr,
		[String]$ImportScript,
		[Switch]$ForceNew
	)

	if ($PSBoundParameters.Containskey("DBConnection") -AND $PSBoundParameters.Containskey("CSVFile") -AND $PSBoundParameters.Containskey("ImportTable"))
	{
		if (Test-Path $CSVfile) 
		{
			Write-Host "Importing File (Default Encoding Comma Separated) $($CSVFile)"
			Write-Host ""
			$CSV = Import-CSV -Encoding Default $CSVFile
			Write-Host "CSVObject Contains $($CSV.Count) Rows"
			Write-Host "Beginning Bulk Import to SQL Server via $($DBConnection)"
			if ($ForceNew)
			{
				$Status= importto-sql -CSVObject $CSV -table $ImportTable -dbstr $DBConnection -AddID -ForceNewTable
			}
			else
			{
				$Status= importto-sql -CSVObject $CSV -table $ImportTable -dbstr $DBConnection
			}
			if ($PSBoundParameters.Containskey("ImportScript"))
			{
				Write-Host "Invoking Import Script $($ImportScript) ..."
				$status = invoke-sqlcmd -InputFile $ImportScript -ConnectionString $DBstr
			}
		}
		Else
		{
			Write-Warning "Cannot Locate CSV file $($CSVObject)"
		}
	}
	Else
	{
		Write-Warning "You MUST supply the following parameters"
		Write-Warning "-DBConnection <DBString>"
		Write-Warning "-CSVFile <Path to CSV file>"
		Write-Warning "-ImportTable <import tablename"
	}
}
# End of ImportCSVto-SQL

Function Invoke-SNowQuery
<#
.SYNOPSIS
Perfoms an ODBC Query against ServiceNow

.DESCRIPTION
Performs a Bulk import of a CSV object into a SQL Table

Example 

Invoke-SNowQuery -Query "Select * from discovery_credentials"

.PARAMETER SNowConn
[System.Data.Odbc.OdbcConnection] Connection object

.PARAMETER DSName
DATA Source Nname

.PARAMETER Query
SQL Query

.PARAMETER Timeout
Query Tinmeout in Seconds

.OUTPUTS
DataRow Object containing the Dataset returned from ServiceNow

#>
{
	[CmdletBinding()]
	Param
	(
		$SNowConn,
		$DSName="SNowQry",
		[int]$Timeout=200,
		$Query,
		[ValidateSet("DataSet", "DataTable")] [string]$As="DataTable"
	)
	
	#Check if ODBC Source is installed
	$ODBCSrc = get-item "HKLM:\SOFTWARE\ODBC\ODBC.INI\ServiceNow" -ErrorAction "SilentlyContinue"
	if ($ODBCSrc)
	{
		$stime = get-date
		write-host "Openning ServiceNow Connection @ $($stime)"
		write-host ""
		Write-verbose $query
		$SNowConn = New-Object System.Data.Odbc.OdbcConnection
		$SNowConn.ConnectionString = "DSN=ServiceNow;UID=svc-reporting;pwd=A5tr0z"
		$SNowConn.ConnectionTimeout=20
		$SNowConn.Open()
		if ($SNowConn.State -Eq "Open")
		{
			Write-Host -Foregroundcolor Green "ODBC Connection Opened to DSN ServiceNow"
			$cmd = New-Object System.Data.Odbc.OdbcCommand
			$cmd.CommandText = $Query
			$cmd.CommandTimeout=$Timeout
			$cmd.Connection = $SNowConn
			$SNowDA= New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
			$SNowDS = New-Object System.Data.DataSet
			$SnowDS.DataSetName=$DSName
			#$SNowdt= New-Object System.Data.DataTable
			[void]$SNowDA.Fill($SNowDS)
			$SNowConn.Close()
			$etime = $(new-timespan -start $stime).totalseconds
			#We have DataSet containing the results - check the -As parameter to se how to return the data
			if ($As -eq "DataSet")
			{
				write-host -Foregroundcolor Green "Snow Query complete: Elapsed Time $($etime) Seconds Returning DataSet"
				Write-Output $SNowDS
			}
			else #DataTable
			{
				#$SNowDT= New-Object System.Data.DataTable
				$SnowDT = $SNowDS.Tables | Select-Object -First 1
				$SnowDT.TableName = "$($DSName) Table"
				write-host -Foregroundcolor Green "Snow Query complete: Elapsed Time $($etime) Seconds Returning DataTable with $($SnowDT.Rows.count) Rows"
				Write-Output -NoEnumerate $SnowDT
			}
		}
		Else
		{
			Write-Warning "Failed to open ODBC connection to DSN ServiceNow"
			Write-Output $Null
		}
	}
	else
	{
		Write-Warning "ServiceNow ODBC Drivers do NOT appear to be installed"
		Write-Output $Null
	}
} # End of Invoke-SNowQuery


<# ===========================================================================================
End of Module Functions

Define and Export Module Members below
==============================================================================================
#>

$ModuleFunctionNames = Get-Content $ModuleFile | Where-Object {$_ -Match "^\s*Function\s+(\S+)" } | Foreach-Object {$Matches[1]}
Export-Modulemember -Function $ModuleFunctionNames

# Uncomment this line to Export all Functions including those Dot Sourced from Libraries
# Export-Modulemember -Function "*"

# Export Variables
Export-Modulemember -Variable @("DBstr","EOSL")
