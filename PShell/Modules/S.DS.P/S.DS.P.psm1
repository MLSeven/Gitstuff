#Types used
Add-Type @'
public enum EncryptionType
{
	None=0,
	Kerberos,
	SSL
}
'@
<#
.SYNOPSIS
	Searches LDAP server in given search root and using given search filter
	

.OUTPUTS
	Search results as custom objects with requested properties as strings or byte stream

.EXAMPLE
Find-LdapObject -SearchFilter:"(&(sn=smith)(objectClass=user)(objectCategory=organizationalPerson))" -SearchBase:"cn=Users,dc=myDomain,dc=com"

Description
-----------
This command connects to local machine on port 389 and performs the search 

.EXAMPLE
Find-LdapObject -SearchFilter:"(&(cn=jsmith)(objectClass=user)(objectCategory=organizationalPerson))" -SearchBase:"ou=Users,dc=myDomain,dc=com" -PropertiesToLoad:@("sAMAccountName","objectSid") -BinaryProperties:@("objectSid")

Description
-----------
This command connects to local machine and performs the search, returning value of objectSid attribute as byte stream

.EXAMPLE
Find-LdapObject -LdapServer:mydc.mydomain.com -SearchFilter:"(&(sn=smith)(objectClass=user)(objectCategory=organizationalPerson))" -SearchBase:"ou=Users,dc=myDomain,dc=com" -UseSSL

Description
-----------
This command connects to given LDAP server and performs the search via SSL

.EXAMPLE
$MyConnection=new-object System.DirectoryServices.Protocols.LdapConnection(new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier("mydc.mydomain.com", 389))

Find-LdapObject -LdapConnection:$MyConnection -SearchFilter:"(&(sn=smith)(objectClass=user)(objectCategory=organizationalPerson))" -SearchBase:"cn=Users,dc=myDomain,dc=com"

Find-LdapObject -LdapConnection:$MyConnection -SearchFilter:"(&(cn=myComputer)(objectClass=computer)(objectCategory=organizationalPerson))" -SearchBase:"ou=Computers,dc=myDomain,dc=com" -PropertiesToLoad:@("cn","managedBy")

Description
-----------
This command creates the LDAP connection object and passes it as parameter. Connection remains open and ready for reuse in subsequent searches

.EXAMPLE
$MyConnection=new-object System.DirectoryServices.Protocols.LdapConnection(new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier("mydc.mydomain.com", 389))
Find-LdapObject -LdapConnection:$MyConnection -SearchFilter:"(&(cn=SEC_*)(objectClass=group)(objectCategory=group))" -SearchBase:"cn=Groups,dc=myDomain,dc=com" | 
Find-LdapObject -LdapConnection:$MyConnection -ASQ:"member" -SearchScope:"Base" -SearchFilter:"(&(objectClass=user)(objectCategory=organizationalPerson))" -propertiesToLoad:@("sAMAccountName","givenName","sn") |
Select-Object * -Unique

Description
-----------
This one-liner lists sAMAccountName, first and last name, and DN of all users who are members of at least one group whose name starts with "SEC_" string


.LINK
More about System.DirectoryServices.Protocols: http://msdn.microsoft.com/en-us/library/bb332056.aspx

#>
Function Find-LdapObject 

{

	Param 
	(
		[parameter(Mandatory = $true)]
		[String] 
		#Search filter in LDAP syntax
		$searchFilter,
		
		[parameter(Mandatory = $true, ValueFromPipeline=$true)]
		[Object] 
		#DN of container where to search
		$searchBase,
		
		[parameter(Mandatory = $false)]
		[String] 
		#LDAP server name
		#Default: Domain Controller
		$LdapServer=[String]::Empty,
		
		[parameter(Mandatory = $false)]
		[Int32] 
		#LDAP server port
		#Default: 389
		$Port=389,
		
		[parameter(Mandatory = $false)]
		[System.DirectoryServices.Protocols.LdapConnection]
			#existing LDAPConnection object.
			#When we perform many searches, it is more effective to use the same conbnection rather than create new connection for each search request.
			#Default: $null, which means that connection is created automatically using information in LdapServer and Port parameters
		$LdapConnection,
		
		[parameter(Mandatory = $false)]
		[System.DirectoryServices.Protocols.SearchScope]
			#Search scope
			#Default: Subtree
		$searchScope="Subtree",
		
		[parameter(Mandatory = $false)]
		[String[]]
		#List of properties we want to return for objects we find.
		#Default: empty array, meaning no properties are returned
		$PropertiesToLoad=@(),
		
		[parameter(Mandatory = $false)]
		[String]
		#Name of attribute for ASQ search. Note that searchScope must be set to Base for this type of seach
		#Default: empty string
		$ASQ,
		
		[parameter(Mandatory = $false)]
		[UInt32]
		#Page size for paged search. Zero means that paging is disabled
		#Default: 100
		$PageSize=100,
		
		[parameter(Mandatory = $false)]
		[String[]]
		#List of properties that we want to load as byte stream.
		#Note: Those properties must also be present in PropertiesToLoad parameter. Properties not listed here are loaded as strings
		#Default: empty list, which means that all properties are loaded as strings
		$BinaryProperties=@(),

		[parameter(Mandatory = $false)]
		[UInt32]
		#Number of seconds before connection times out.
		#Default: 120 seconds
		$TimeoutSeconds = 120,

		[parameter(Mandatory = $false)]
		[EncryptionType]
		#Type of encryption to use.
		#Applies only when existing connection is not passed
		$EncryptionType="None",

		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$UserName=$null,

		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$Domain=$null,

		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$Password=$null
	)

	Process 
	{
		#we want dispose LdapConnection we create
		[Boolean]$bDisposeConnection=$false
		#range size for ranged attribute retrieval
		#Note that default in query policy is 1500; we set to 1000
		$rangeSize=1000
		try 
		{
			if($LdapConnection -eq $null) 
			{
				[System.Net.NetworkCredential]$cred=$null
				if(-not [String]::IsNullOrEmpty($userName)) 
				{
					if([String]::IsNullOrEmpty($password)) 
					{
						$securePwd=Read-Host -AsSecureString -Prompt:"Enter password"
						$cred=new-object System.Net.NetworkCredential($userName,$securePwd, $Domain)
					} 
					else
					{
						$cred=new-object System.Net.NetworkCredential($userName,$Password, $Domain)
					}
					$LdapConnection=new-object System.DirectoryServices.Protocols.LdapConnection((new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($LdapServer, $Port)), $cred)
				} 
				else 
				{
					$LdapConnection=new-object System.DirectoryServices.Protocols.LdapConnection(new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($LdapServer, $Port))
				}
				$bDisposeConnection=$true
				switch($EncryptionType) 
				{
					"None" {break}
					"SSL" 
					{
						$options=$LdapConnection.SessionOptions
						$options.ProtocolVersion=3
						$options.StartTransportLayerSecurity($null)
						break			   
					}
					"Kerberos" 
					{
						$LdapConnection.SessionOptions.Sealing=$true
						$LdapConnection.SessionOptions.Signing=$true
						break
					}
				}
			}
			if($pageSize -gt 0) 
			{
				#paged search silently fails when chasing referrals
				$LdapConnection.SessionOptions.ReferralChasing="None"
			}

			#build request
			$rq=new-object System.DirectoryServices.Protocols.SearchRequest
			
			#search base
			switch($searchBase.GetType().Name) 
			{
			   "PSCustomObject" 
			   { 
					if($searchBase.distinguishedName -ne $null) 
					{
						$rq.DistinguishedName=$searchBase.distinguishedName
					}
				}
				"String" {$rq.DistinguishedName=$searchBase}
				default { return }
			}

			#search filter in LDAP syntax
			$rq.Filter=$searchFilter

			#search scope
			$rq.Scope=$searchScope

			#attributes we want to return - nothing now, and then use ranged retrieval for the propsToLoad
			$rq.Attributes.Add("1.1") | Out-Null

			#paged search control for paged search
			if($pageSize -gt 0) 
			{
				[System.DirectoryServices.Protocols.PageResultRequestControl]$pagedRqc = new-object System.DirectoryServices.Protocols.PageResultRequestControl($pageSize)
				$rq.Controls.Add($pagedRqc) | Out-Null
			}

			#server side timeout
			$rq.TimeLimit=(new-object System.Timespan(0,0,$TimeoutSeconds))

			if(-not [String]::IsNullOrEmpty($asq)) 
			{
				[System.DirectoryServices.Protocols.AsqRequestControl]$asqRqc=new-object System.DirectoryServices.Protocols.AsqRequestControl($ASQ)
				$rq.Controls.Add($asqRqc) | Out-Null
			}
			
			#initialize output objects via hashtable --> faster than add-member
			#create default initializer beforehand
			$propDef=[Ordered]@{}
			#we always return at least distinguishedName
			#so add it explicitly to object template and remove from propsToLoad if specified
			$propDef.Add("distinguishedName","")
			$PropertiesToLoad=@($propertiesToLoad | where-object {$_ -ne "distinguishedName"})
						
			#prepare template for output object
			foreach($prop in $PropertiesToLoad) 
			{
			   $propDef.Add($prop,@())
			}

			#process paged search in cycle or go through the processing at least once for non-paged search
			while ($true)
			{
				$rsp = $LdapConnection.SendRequest($rq, (new-object System.Timespan(0,0,$TimeoutSeconds))) -as [System.DirectoryServices.Protocols.SearchResponse];
				
				#for paged search, the response for paged search result control - we will need a cookie from result later
				if($pageSize -gt 0) 
				{
					[System.DirectoryServices.Protocols.PageResultResponseControl] $prrc=$null;
					if ($rsp.Controls.Length -gt 0)
					{
						foreach ($ctrl in $rsp.Controls)
						{
							if ($ctrl -is [System.DirectoryServices.Protocols.PageResultResponseControl])
							{
								$prrc = $ctrl;
								break;
							}
						}
					}
					if($prrc -eq $null) {
						#server was unable to process paged search
						throw "Find-LdapObject: Server failed to return paged response for request $SearchFilter"
					}
				}
				#now process the returned list of distinguishedNames and fetch required properties using ranged retrieval
				foreach ($sr in $rsp.Entries)
				{
					$dn=$sr.DistinguishedName
					#we return results as powershell custom objects to pipeline
					#initialize members of result object (server response does not contain empty attributes, so classes would not have the same layout
					#create empty custom object for result, including only distinguishedName as a default
					$data=new-object PSObject -Property $propDef
					$data.distinguishedName=$dn
					
					#load properties of custom object, if requested, using ranged retrieval
					foreach ($attrName in $PropertiesToLoad) 
					{
						$rqAttr=new-object System.DirectoryServices.Protocols.SearchRequest
						$rqAttr.DistinguishedName=$dn
						$rqAttr.Scope="Base"
						
						$start=-$rangeSize
						$lastRange=$false
						while ($lastRange -eq $false) 
						{
							$start += $rangeSize
							$rng = "$($attrName.ToLower());range=$start`-$($start+$rangeSize-1)"
							$rqAttr.Attributes.Clear() | Out-Null
							$rqAttr.Attributes.Add($rng) | Out-Null
							$rspAttr = $LdapConnection.SendRequest($rqAttr)
							foreach ($sr in $rspAttr.Entries) 
							{
								if($sr.Attributes.AttributeNames -ne $null) 
								{
									#LDAP server changes upper bound to * on last chunk
									$returnedAttrName=$($sr.Attributes.AttributeNames)
									#load binary properties as byte stream, other properties as strings
									if($BinaryProperties -contains $attrName) 
									{
										$vals=$sr.Attributes[$returnedAttrName].GetValues([byte[]])
									} 
									else 
									{
										$vals = $sr.Attributes[$returnedAttrName].GetValues(([string])) # -as [string[]];
									}
									$data.$attrName+=$vals
									if($returnedAttrName.EndsWith("-*") -or $returnedAttrName -eq $attrName) 
									{
										#last chunk arrived
										$lastRange = $true
									}
								} 
								else 
								{
									#nothing was found
									$lastRange = $true
								}
							}
						}

						#return single value as value, multiple values as array, empty value as null
						switch($data.$attrName.Count) 
						{
							0 
							{
								$data.$attrName=$null
								break;
							}
							1 
							{
								$data.$attrName = $data.$attrName[0]
								break;
							}
							default 
							{
								break;
							}
						}
					}
					#return result to pipeline
					$data
				}
				if($pageSize -gt 0) 
				{
					if ($prrc.Cookie.Length -eq 0) 
					{
						#last page --> we're done
						break;
					}
					#pass the search cookie back to server in next paged request
					$pagedRqc.Cookie = $prrc.Cookie;
				} 
				else 
				{
					#exit the processing for non-paged search
					break;
				}
			}
		}
		finally 
		{
			if($bDisposeConnection) 
			{
				#if we created the connection, dispose it here
				$LdapConnection.Dispose()
			}
		}
	}
}

Function Get-RootDSE 
{
	Param 
	(
		[parameter(Mandatory = $false)]
		[String] 
		#LDAP server name
		#Default: closest DC
		$LdapServer=[String]::Empty,
		
		[parameter(Mandatory = $false)]
		[Int32] 
		#LDAP server port
		#Default: 389
		$Port=389,
		[parameter(Mandatory = $false)]
		[System.DirectoryServices.Protocols.LdapConnection]
		#existing LDAPConnection object.
		#When we perform many searches, it is more effective to use the same connection rather than create new connection for each search request.
		#Default: $null, which means that connection is created automatically using information in LdapServer and Port parameters
		$LdapConnection,
		
		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$UserName=$null,

		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$Domain=$null,

		[parameter(Mandatory = $false)]
		[String]
		#Use different credentials when connecting
		$Password=$null
	)
	
	Process 
	{
		#we want dispose LdapConnection we create
		[Boolean]$bDisposeConnection=$false
		#range size for ranged attribute retrieval
		#Note that default in query policy is 1500; we set to 1000
		$rangeSize=1000

		try 
		{
			if($LdapConnection -eq $null) 
			{
				[System.Net.NetworkCredential]$cred=$null
				if(-not [String]::IsNullOrEmpty($userName)) 
				{
					if([String]::IsNullOrEmpty($password)) 
					{
						$securePwd=Read-Host -AsSecureString -Prompt:"Enter password"
						$cred=new-object System.Net.NetworkCredential($userName,$securePwd,$Domain)
					} 
					else 
					{
						$cred=new-object System.Net.NetworkCredential($userName,$Password,$Domain)
					}
					$LdapConnection=new-object System.DirectoryServices.Protocols.LdapConnection((new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($LdapServer, $Port)), $cred)
				} 
				else 
				{
					$LdapConnection=new-object System.DirectoryServices.Protocols.LdapConnection(new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($LdapServer, $Port))
				}
				$bDisposeConnection=$true
			}
			
			#initialize output objects via hashtable --> faster than add-member
			#create default initializer beforehand
			$PropertiesToLoad=@("rootDomainNamingContext", "configurationNamingContext", "schemaNamingContext","defaultNamingContext","dnsHostName")
			$propDef=@{}
			foreach($prop in $PropertiesToLoad) 
			{
				$propDef.Add($prop,@())
			}
			#build request
			$rq=new-object System.DirectoryServices.Protocols.SearchRequest
			$rq.Scope = "Base"
			$rq.Attributes.AddRange($PropertiesToLoad) | Out-Null
			[System.DirectoryServices.Protocols.ExtendedDNControl]$exRqc = new-object System.DirectoryServices.Protocols.ExtendedDNControl("StandardString")
			$rq.Controls.Add($exRqc) | Out-Null
			
			$rsp=$LdapConnection.SendRequest($rq)
			
			$data=new-object PSObject -Property $propDef
			
			$data.configurationNamingContext = (($rsp.Entries[0].Attributes["configurationNamingContext"].GetValues([string]))[0]).Split(';')[1];
			$data.schemaNamingContext = (($rsp.Entries[0].Attributes["schemaNamingContext"].GetValues([string]))[0]).Split(';')[1];
			$data.rootDomainNamingContext = (($rsp.Entries[0].Attributes["rootDomainNamingContext"].GetValues([string]))[0]).Split(';')[2];
			$data.defaultNamingContext = (($rsp.Entries[0].Attributes["rootDomainNamingContext"].GetValues([string]))[0]).Split(';')[2];
			$data.dnsHostName = ($rsp.Entries[0].Attributes["dnsHostName"].GetValues([string]))[0]
			$data
		}
		finally 
		{
			if($bDisposeConnection) 
			{
				#if we created the connection, dispose it here
				$LdapConnection.Dispose()
			}
		}
	}
}
  