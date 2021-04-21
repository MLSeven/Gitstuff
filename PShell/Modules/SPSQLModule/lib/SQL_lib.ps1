<# ===========================================================================================
Library Scripots used by SPSQLModule

SP 23 June 2017
==============================================================================================
#>

Function ConvertTo-DataTable
<#
.SYNOPSIS
Converts the Pipeline into a DataTable Object

.DESCRIPTION
Converts the Pipeline into a DataTable Object


.PARAMETER InputObject (From Pipeline)
Source Object to be converted


.OUTPUTS
Datatable Object

#>
{
	[CmdletBinding()] 
	Param
	(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject
	) 

	Begin 
	{
		$types = @('System.Boolean','System.Byte[]','System.Byte','System.Char','System.Datetime','System.Decimal','System.Double','System.Guid','System.Int16','System.Int32','System.Int64','System.Single','System.UInt16','System.UInt32','System.UInt64')
		$dt = new-object Data.datatable
		$First = $true
	}
	Process 
	{
		foreach ($object in $InputObject) 
		{ 
			$DR = $DT.NewRow()   
			foreach($property in $object.PsObject.get_properties()) 
			{
				#From the first object create the DataColumn Object
				if ($first) 
				{
					$Col =  new-object Data.DataColumn
					$Col.ColumnName = $property.Name.ToString()
					if ($property.value) 
					{ 
						if ($property.value -isnot [System.DBNull]) 
						{
							if ($types -Contains $property.TypeNameOfValue)
							{
								$ColType = $property.TypeNameOfValue
							}
							else
							{
								$ColType = 'System.String'
							}
							$Col.DataType = [System.Type]::GetType($ColType)
						} 
					} 
					$DT.Columns.Add($Col) 
				}
				#Obtain the Row values
				if ($property.Gettype().IsArray)
				{ 
					$DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
				}
				else 
				{ 
					$DR.Item($property.Name) = $property.value 
				} 
			}
		$DT.Rows.Add($DR)   
		$First = $false 
		}
	}
	End 
	{
		#Return $DT
		Write-Output @(,($dt)) 
	}
} # End of ConvertTo-DataTable


Function Analyse-Object
<#
.SYNOPSIS
Takes a PSObject and returns the maximum data size of each property

.DESCRIPTION
Takes a PSObject and returns the maximum data size of each property.
This data is used to create a SQL table based on the object

.PARAMETER InputObject
Source Object to be analysed

.OUTPUTS
Array of PSCustomObjects detailing Name, Type and Max Size

#>
{
	[CmdletBinding()] 
	param
	(
		[Object]$InputObject
	) 
	
	if ($InputObject)
	{
		#ObjType could be an Array of Objects - use Select-Object to get the first one and test that
		if ($InputObject.gettype().IsArray)
		{
			$ObjType = ($InputObject | Select-Object -First 1).Gettype().Name
			Write-Verbose "InputObject is an Array of Type $ObjType"
		}
		else
		{
			$ObjType = $InputObject.Gettype().Name
			Write-Verbose "InputObject is of Type $ObjType"
		}
	}
	else
	{	
		$ObjType = "Empty"
		Write-Verbose "InputObject is Empty"
	}
	
	$ColData = @()
	
	Switch ($ObjType)
	{
		"PSCustomObject"
		{
			Write-Host "Analysing PSCustomObject Object for Column Name, Type and Maximum Size"
			#CSV Object: Take Names from first Row - default all types to String. Cycle through each row to find Maximum Sizes
			$ColData = $InputObject | select-Object -First 1 | Foreach-Object {$_.PSObject.Properties | Foreach-Object {[PSCustomObject]@{name=$_.Name;Type="String";Size=0}}}
			$InputObject | Foreach-Object {$i=0;$_.PSObject.Properties | Foreach-Object {if ($_.Value.Length -gt $ColData[$i].Size) {$ColData[$i].Size = $_.Value.Length}; $i += 1}}
		}
		"DataSet"
		{
			Write-Host "Analysing DataSet Object $($InputObject.DataSetName) (Table(0) Only) for Column Name, Type and Maximum Size"
			Foreach ($Column in $InputObject.Tables.item(0).Columns)
			{
				$Name = $Column.Columnname
				$Type = $Column.DataType.Name
				if ($Type -eq "String")
				{
					$Size = ($InputObject.Tables[0].Item($Name) | foreach-Object {$_.length} | Measure-Object -Maximum).Maximum
				}
				else
				{
					$Size = -1
				}
				$ColData += [PSCustomObject]@{Name=$Name;Type=$type;MaxSize=$Size}
			}
		}
		"DataTable"
		{
			Write-Host "Analysing DataTable Object $($InputObject.TableName) for Column Name, Type and Maximum Size"
			Foreach ($Column in $InputObject.Columns)
			{
				$Name = $Column.Columnname
				$Type = $Column.DataType.Name
				if ($Type -eq "String")
				{
					$Size = ($InputObject.Item($Name) | foreach-Object {$_.length} | Measure-Object -Maximum).Maximum
					Write-Host "Column $name Size $Size"
				}
				else
				{
					$Size = -1
				}
				$ColData += [PSCustomObject]@{Name=$Name;Type=$type;MaxSize=$Size}
			}
		}		
		"Empty"
		{
			$ColData = @()
		}
	}
	Write-Host "Returning Column Data ..."
	Write-Output $ColData
} # End of Analyse-Object


Function Create-SQLTableDDL
<#
.SYNOPSIS
Takes the Column Data specification and Creates a Simple SQL table

.DESCRIPTION
Takes the Column Data specification and Creates a Simple SQL table

.PARAMETER InputObject
Source Object to be used to create the table. Can be CSV or DataTable

.PARAMETER Tablename
SQL Table name which may be in the format Schema.Table. If schema is omitted dbo is assumed

.PARAMETER AddID
Adds ID and BatchID columns to the table

.PARAMETER ColStats
Used to specify the size of the table columns

.OUTPUTS
SQL CREATE Table statement

#>
{
	[CmdletBinding()] 
	param 
	(
		$InputObject,
		[String]$TableName,
		[switch]$AddID,
		$ColumnDefs
	)
	
	#Go through the InputObject and determine the max size of each Column
	
	if (-NOT $ColumnDefs) {$ColumnDefs = Analyse-Object -InputObject $InputObject}
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
	#$ColumnDefs describes the table schema
	
	if ($ColumnDefs)
	{
		Write-Host "Data Object has $($ColumnDefs.count) Columns"
		$ColumnDefs | Foreach-Object {Write-Host "$($_.Name) : Type $($_.Type) : Size $($_.Size)"}
		
		$SQL = @("IF OBJECT_ID('[$($Owner)].[$($Table)]','U') IS NOT NULL DROP TABLE [$($Owner)].[$($Table)];")
		$SQL += "CREATE TABLE [$($Owner)].[$($Table)]"
		$SQL +="("
		if ($AddID) 
		{
			$SQL +="[ID] Int IDENTITY(1,1),"
			$SQL +="[BatchID] [nvarchar] (20) NOT NULL,"
		}
		Foreach ($Col in $ColumnDefs)
		{
			$Size = $Col.MaxSize
			if ($Size -le 255)  {$ColSize = "(255)"}
			elseif ($Size -le 1023) {$ColSize = "(1023)"}
			else {$ColSize = "(Max)"}

			switch ($Col.Type)
			{
				"DateTime" {$SQL += "[$($Col.Name)] [DATETIME] NULL,"}
				"String" {$SQL += "[$($Col.Name)] [nvarchar] $($ColSize) NULL,"}
				default {$SQL += "[$($Col.Name)] [nvarchar] $($ColSize) NULL,"}
			}
		}
		#$Remove the comma from the last column
		$SQL[-1] = $SQL[-1] -replace ",",""
		$SQL += ");"
	
	$SQLScript = [String]::Join("`n",$SQL)
	$SQLScript
	}
} # End of Create-SQLTableDDL



# Retired

Function Create-SQLTableFromPSObject
<#
.SYNOPSIS
Takes a PSObject and returns SQL statements to Create a table

.DESCRIPTION
Takes a PSObject and returns SQL statements to Create a table

.PARAMETER InputObject
Source Object to be used to create the table

.PARAMETER Tablenmae
SQL Table name

.PARAMETER AddID
Adds ID and BatchID columns to the table

.PARAMETER ColStats
Used to specify the size of the table columns

.OUTPUTS
SQL CREATE Table statement

#>
{
	[CmdletBinding()] 
	param 
	(
		[PSObject]$InputObject,
		[String]$TableName,
		[switch]$AddID,
		$ColStats
	)
	
	#Go through the InputObject and determine the max size of each Column
	
	if ($ColStats -eq $Null) {$ColStats = Analyse-Object -InputObject $InputObject}
	# Get the Column Names from the first object 
	If ($InputObject.count -gt 1) {$row = $InputObject | select-Object -first 1} else {$row = $InputObject}
	$ColCount = $($Row.PSObject.Properties | Measure-Object).Count
	if ($ColStats -eq $Null) {$ColStats = Analyse-Object -InputObject $InputObject}
	
	
	$SQL = @("IF OBJECT_ID('[$($TableName)]','U') IS NOT NULL DROP TABLE $($TableName)")
	$SQL += "CREATE TABLE [$Tablename]"
	$SQL +="("
	if ($AddID) 
	{
		$SQL +="[ID] Int IDENTITY(1,1),"
		$SQL +="[BatchID] [nvarchar] (20) NOT NULL,"
	}
	$col = 0
	Foreach ($Prop in $InputObject.PSObject.Properties)
	{
		$Col += 1
		$Size = $ColStats.Item($Prop.Name)
		if ($Size -le 255)  {$ColSize = "[nvarchar] (255)"}
		elseif ($Size -le 1023) {$ColSize = "[nvarchar] (1023)"}
		else {$ColSize = "[nvarchar] (Max)"}
		#Write-Host "$Size - col $Colsize"
		if ($Col -LT $ColCount)
		{
			$SQL += "[$($Prop.Name)] $ColSize NULL,"
		}
		Else
		{
			$SQL += "[$($Prop.Name)] $ColSize NULL"
		}
	}
	$SQL += ")"
	
	$SQLScript = [String]::Join("`n",$SQL)
	$SQLScript
} # End of Create-SQLTableFromPSObject


Function BulkTransfer-ToSQL
<#
.SYNOPSIS
Performs a Bulk Update into a SQL table

.DESCRIPTION
Performs a Bulk Update into a SQL table. The input object must be a correctly formed DataTable object

.PARAMETER DBStr
Source Object to be used to create the table

.PARAMETER TableData
Data to Bulk Load as a DataTable

.PARAMETER Tablename
SQL Table name

.PARAMETER SourceCol
This array specifies which columns from the TableData are to be imported

.PARAMETER TargetCol
This array specifies the Column names in the Target Table 

.OUTPUTS
The number of Rows Uploaded

#>
{
	[CmdletBinding()]
	Param 
	(
		[Parameter(Mandatory=$true)] [String]$DbStr, 
		[Parameter(Mandatory=$true, ValueFromPipeline=$True)] $TableData, 
		[Parameter(Mandatory=$true)] [String]$TableName,
		$SourceCol,
		$TargetCol
	)

	Write-Host "Initiating Function BulkTransfer-ToSQL"
	Write-Host ""
	$DBConn = New-Object -Type 'System.Data.SqlClient.SqlConnection' -ArgumentList $DBStr
	$DBConn.Open()
	$bc = New-Object -Type 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList $DBconn
	$bc.DestinationTableName = $TableName
	$RowCount = $TableData.Rows.Count
	
	if ($SourceCol)
	{
		if ($SourceCol.GetType().IsArray) 
		{
			Write-host "Column Mappings have been specified"
			Write-Host ""
			for ($i=0; $i -lt $SourceCol.count; $i++)
			{
				Write-Host "Mapping $($SourceCol[$i]) to $($TargetCol[$i])"
				$null=$bc.ColumnMappings.Add($SourceCol[$i], $TargetCol[$i])
			}
		}
	}
	Write-Host ""
	Write-Host "Beginning Bulk Transfer to SQL Server ..."
	Write-Host ""
	$Stats = $bc.WriteToServer($TableData)
	if ($?) 
		{Write-Host -Foregroundcolor Green  "Bulk Data Transfer to table $($TableName) completed OK - $($RowCount) Rows Transferred"} 
	else 
		{Write-Warning "Error occurred during Bulk Transfer to table $($TableName)"}
	$DBConn.Close()
	Write-Host "Exiting Function BulkTransfer-ToSQL"
	Write-Output $stats
} # End of BulkTransfer-ToSQL