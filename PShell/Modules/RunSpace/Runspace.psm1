<#
	.SYNOPSIS
		Create a new runspace pool
	.DESCRIPTION
		This function creates a new runspace pool. This is needed to be able to run code multi-threaded.
	.EXAMPLE
		$pool = New-RunspacePool

		Description
		-----------
		Create a new runspace pool with default settings, and store it in the pool variable.
	.EXAMPLE
		$pool = New-RunspacePool -Snapins 'vmware.vimautomation.core'

		Description
		-----------
		Create a new runspace pool with the VMWare PowerCli snapin added, and store it in the pool variable.

#>	
function New-RunspacePool
{
	[CmdletBinding()]
	param
	(
		# The minimum number of concurrent threads to be handled by the runspace pool. The default is 1.
		[Parameter(HelpMessage='Minimum number of concurrent threads')]
		[ValidateRange(1,65535)]
		[int32]$minRunspaces = 1,
 
		# The maximum number of concurrent threads to be handled by the runspace pool. The default is 15.
		[Parameter(HelpMessage='Maximum number of concurrent threads')]
		[ValidateRange(1,65535)]
		[int32]$maxRunspaces = 15,
 
		# Using this switch will set the apartment state to MTA.
		[Parameter()]
		[switch]$MTA,
 
		# Array of snapins to be added to the initial session state of the runspace object.
		[Parameter(HelpMessage='Array of SnapIns you want available for the runspace pool')]
		[string[]]$Snapins,
 
		# Array of modules to be added to the initial session state of the runspace object.
		[Parameter(HelpMessage='Array of Modules you want available for the runspace pool')]
		[string[]]$Modules,
 
		# Array of functions to be added to the initial session state of the runspace object.
		[Parameter(HelpMessage='Array of Functions that you want available for the runspace pool')]
		[string[]]$Functions,
 
		# Array of variables to be added to the initial session state of the runspace object.
		[Parameter(HelpMessage='Array of Variables you want available for the runspace pool')]
		[string[]]$Variables
	)

	# create the initial session state
	$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
 
	# add any snapins to the session state object
	if($Snapins)
	{
		foreach ($snapName in $Snapins)
		{
			try
			{
				$iss.ImportPSSnapIn($snapName,[ref]'') | Out-Null
				Write-Verbose "Imported $snapName to Initial Session State"
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
		}
	}
 
	# add any modules to the session state object
	if ($Modules)
	{
		foreach ($module in $Modules)
		{
			try
			{
				$iss.ImportPSModule($module) | Out-Null
				Write-Verbose "Imported $module to Initial Session State"
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
		}
	}
 
	# add any functions to the session state object
	if ($Functions)
	{
		foreach ($func in $Functions)
		{
			try
			{
				$thisFunction = Get-Item -LiteralPath "function:$func"
				[String]$functionName = $thisFunction.Name
				[ScriptBlock]$functionCode = $thisFunction.ScriptBlock
				$iss.Commands.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $functionName,$functionCode))
				Write-Verbose "Imported $func to Initial Session State"
				Remove-Variable thisFunction, functionName, functionCode
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
		}
	}
 
	# add any variables to the session state object
	if ($Variables)
	{
		foreach ($var in $Variables)
		{
			try
			{
				$thisVariable = Get-Variable $var
				$iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $thisVariable.Name, $thisVariable.Value, ''))
				Write-Verbose "Imported $var to Initial Session State"
			}
			catch
			{
				Write-Warning $_.Exception.Message
			}
		}
	}
 
	# create the runspace pool
	$runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($minRunspaces, $maxRunspaces, $iss, $Host)
	Write-Verbose 'Created runspace pool'
 
	# set apartmentstate to MTA if MTA switch is used
	if($MTA)
	{
		$runspacePool.ApartmentState = 'MTA'
		Write-Verbose 'ApartmentState: MTA'
	}
	else 
	{
		Write-Verbose 'ApartmentState: STA'
	}
 	# open the runspace pool
	$runspacePool.Open()
	Write-Verbose 'Runspace Pool Open'
 	# return the runspace pool object
	Write-Output $runspacePool
}

<#
	.SYNOPSIS
		Create a new runspace job.
	.DESCRIPTION
		This function creates a new runspace job, executed in it's own runspace (thread).
	.EXAMPLE
		Start-RunspaceJob -JobName 'Inventory' -ScriptBlock $code -Parameters $parameters

		Description
		-----------
		Execute code in $code with parameters from $parameters in a new runspace (thread).
	.OUTPUT
	
	[PSCustomObject]RunSpaceJob used to communicate with the RunSpace via this Module
#>
function Start-RunspaceJob
{
	[CmdletBinding()]
	param
	(
		# Optionally give the job a name.
		[Parameter()][string]$JobName,
		# The code you want to execute.
		[Parameter(Mandatory = $true)][ScriptBlock]$ScriptBlock,
		# A working runspace pool object to handle the runspace job.
		[Parameter(Mandatory = $true)][System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
		# Hashtable of parameters to add to the runspace scriptblock.
		[Parameter()][HashTable]$Parameters
	)

	$runspace = [System.Management.Automation.PowerShell]::Create()
	$runspace.RunspacePool = $RunspacePool
 
	# add the scriptblock to the runspace
	$runspace.AddScript($ScriptBlock) | Out-Null
 
	# if any parameters are given, add them into the runspace
	if($parameters)
	{
		foreach ($parameter in ($Parameters.GetEnumerator()))
		{
			$runspace.AddParameter("$($parameter.Key)", $parameter.Value) | Out-Null
		}
	}
 
	# invoke the runspace and Return the RunspaceJob custom object
	$RunSpaceJob = [PSCustomObject]@{JobName=$JobName;ID=$Runspace.InstanceID;Start=$(Get-Date);End=$null;Elapsed=0;Runspace=$RunSpace;Task=$Runspace.BeginInvoke()}

	Write-Verbose 'Task invoked in runspace $($RunSpaceJob.ID)' 
	Write-Output $RunSpaceJob
}


<#
	.SYNOPSIS
		Receive data back from a runspace job.
	.DESCRIPTION
		This function checks for completed runspace jobs, and retrieves the return data.
	.EXAMPLE
		Receive-RunspaceJob -Wait

		Description
		-----------
		Will wait until all runspace jobs are complete and retrieve data back from all of them.
	.EXAMPLE
		Receive-RunspaceJob -JobName 'Inventory'

		Description
		-----------
		Will get data from all completed jobs with the JobName 'Inventory'.
	.NOTES

#>
function Receive-RunspaceJob
{
	[CmdletBinding()]
	param
	(
		# Runspace Object(s) Created by Start-RunSpaceJob
		[Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$true)]$RunSpaceJob
	)
	
	Begin {}
	Process
	{
		$returnData = $RunSpaceJob |
			Where-Object {$_.Task.isCompleted} | 
			Foreach-Object {$_.Runspace.Endinvoke($_.Task)}
		if ($returnData) {Write-Output $returnData}
	}
	end {}
}

<#
	.SYNOPSIS
		Wait for a collection of Runspace Jobs ro complete
	.DESCRIPTION
		
	.EXAMPLE
		Wait-RunspaceJob -Wait

		Description
		-----------
		Will wait until all runspace jobs are complete and retrieve data back from all of them.
	.EXAMPLE
		Wait-RunspaceJob -JobName 'Inventory'

	.NOTES

#>
Function Wait-RunspaceJob

{
	[CmdletBinding()]
	param
	(
		# Runspace Object(s) Created by Start-RunSpaceJob
		[Parameter(Mandatory=$True,Position=0,ValuefromPipeline=$true)]$RunSpaceJob,
		[Parameter(HelpMessage='Set a total timeout before exiting this function. 0 means no timeout')][int]$TimeOut = 0,
		[Parameter(HelpMessage='Set a time in Seconds to sleep between checks')][int]$SleepInterval = 5,
		[Parameter()][switch]$ShowProgress
	)
	
	Begin
	{
		write-Verbose "Begin block - Initialising"
		$startTime = Get-Date
		if ($PSBoundParameters.ContainsKey("RunSpaceJob"))
		{
			$PipeLine = $false
			$filteredRunspaces = $RunSpaceJob.Clone()
		}
		Else
		{
			$Pipeline = $True
			$filteredRunspaces = @() 
		}
		Write-Verbose "Input expected down Pipeline? : $Pipeline"
	}
	
	Process
	{
		write-Verbose "Process block"
		if ($Pipeline) 
		{
			Write-Verbose "Adding Job from Pipeline ..."
			$filteredRunspaces += $RunSpaceJob
		}
	}

	End
	{
		write-Verbose "In END block"
		write-Host "Waiting for RunSpaceJobs @ $StartTime - RunSpaceJobs Count : $($FilteredRunSpaces.count)"

		$Waiting = $true
		While ($waiting)
		{
			#Report Recently completed jobs in this Loop
			$RecentCompleted = $filteredRunspaces | 
				Where-Object {$_.Task.isCompleted -AND ($_.Elapsed -eq 0)} | 
				Foreach-Object {$_.End=(Get-Date); $_.Elapsed=$(New-TimeSpan -Start $_.Start).TotalSeconds; $_ }
			
			if ($ShowProgress)
			{
				$Remaining = ($FilteredRunspaces | Where-Object {-NOT $_.Task.IsCompleted}) | Measure-Object
				$RecentCompleted | Foreach-Object {Write-Host "$($_.JobName) Now Completed - Elapsed Time $($_.Elapsed) Seconds : $($Remaining.count) Jobs to complete"}
			}
			
			#Check all the jobs - If all are finished $isCompleted will contain only $true
			$isCompleted = [Array]$filteredRunspaces.Task.IsCompleted
			if ($isCompleted -Contains $false)
			{
				#Check timeout
				if (($Timeout -gt 0) -AND ($(New-Timespan -start $startTime).Totalseconds -gt $timeout))
				{
					$Active = @($isCompleted | Where-Object {-NOT $_})
					Write-Warning "Timeout expired - $($Active.count) RunSpaceJobs are still active"
					$waiting = $false
				}
				else
				{
					Start-Sleep -Seconds $SleepInterval
				}
			}
			else
			{
				#All Jobs Completed
				$waiting=$false
			}
		}

		$EndTime = Get-Date
		$TotalTime = $(new-timespan -Start $startTime -End $Endtime).TotalSeconds
		write-Host "RunSpaceJobs Complete @ $EndTime - RunSpaceJobs Count : $($FilteredRunSpaces.count) - Wait-Time $($TotalTime) Seconds"
		Return $FilteredRunSpaces
	}
}

<#
	.SYNOPSIS
		Display details RunSpaceJob
	.DESCRIPTION
		This function displays details of RunSpaceJob
	.PARAMETER RunSpaceJob
		List of RunSpace objects created by Start-RunSpaceJob
		
	.EXAMPLE
		Show-RunspaceJob -RunspaceJob $RS

		Description
		-----------
		Will show details of Runspace objects in list $RS

#>
Function Show-RunspaceJob
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true,Position=0,valuefrompipeline=$true)]$RunSpaceJob
	)
	
	begin {}
	
	Process
	{
		Foreach ($RS in $RunSpaceJob)
		{
			if ($RS.RunSpace -eq $Null)
			{
				Write-Host "Job $($RS.JobName) ID $($RS.ID) - RunSpace has been Disposed"
			}
			Else
			{
				Write-Host "Job $($RS.JobName) ID $($RS.ID) - Status $($RS.RunSpace.InvocationStateInfo.State) Reason $($RS.RunSpace.InvocationStateInfo.Reason)"
			}
		}
	}
	end {}
}

Function Remove-RunspaceJob
{
	[CmdletBinding()]
	param
	(
		[Parameter(mandatory=$true,Position=0,valuefrompipeline=$true)][Object[]]$RunSpaceJob
	)
	begin {}
	Process
	{
		Foreach ($RS in $RunSpaceJob)
		{
			if ($RS.RunSpace -ne $Null)
			{
				if ($RS.Task.isCompleted)
				{
					$RS.RunSpace.Stop()
					$RS.RunSpace.Dispose()
					$RS.RunSpace = $Null
					$RS.Task = $Null
					Write-verbose "RunSpace $($RS.ID) Has been disposed successfully" 
				}
				Else
				{
					Write-Warning "RunSpace $($RS.ID) Has not yet Completed!"
				}
			}
			Else
			{
				Write-Warning "RunSpace $($RS.ID) Has Already been disposed!"
			}
		}
	}
	end {}
}
 
Export-Modulemember -Function *