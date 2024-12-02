# ==============================================================================================
# Microsoft PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2009
# NAME: 
# AUTHOR: Suddenlink User , Suddenlink Communications
# DATE  : 11/1/2011
# COMMENT: 
# 
# ==============================================================================================

# Test to verify that the Active Directory module is
# available.
$flag = $false
foreach ($module in $LoadedModules) {
    if ($module.Name -eq 'ActiveDirectory') {
        $ADflag = $true
    }
}

if (!$ADflag) {
    Import-Module ActiveDirectory
}

$ErrorActionPreference = "stop"
Try {
	$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sdldalpwemb01a.suddenlink.cequel3.com/PowerShell/ -Authentication Kerberos
	import-pssession $ExchSession
 } Catch {
	throw $_ 
}
$ErrorActionPreference = "Continue"


#$DebugPreference = "Continue"
$LogFilePath = (Get-Date -Format M-dd-yyyy-hh-mmtt)
$LogFilePath = "C:\Scripts\GroupAutomation\Logs\" + $LogFilePath + ".log"

$objRootDSE = [ADSI]"LDAP://rootDSE"
$ForestRoot = $objRootDSE.Get("rootDomainNamingContext")
$domainDN = $objRootDSE.Get("defaultNamingContext")
$domainDNS = "suddenlink.cequel3.com"
$objConnection = New-Object -ComObject "ADODB.Connection"
[void]$objConnection.Open("Provider=ADsDSOObject;")
$ADS_SCOPE_SUBTREE = 2
$ADS_PROPERTY_CLEAR = 1

$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.SearchScope = "Subtree"

$PeopleSoftOracleServer="HRPROD1.suddenlink.cequel3.com"
$PeopleSoftOracleInstance="hrprd"
$PeopleSoftOracleEmployeeView="PS_CQ_DIR_INTFCDAT"
$PeopleSoftOracleUser="ACT_USER"
$PeopleSoftOraclePassword="n3m4s"

$bRemoveExtraMembers = $true
$bNoRemoveExtraMembers = $false

####################################################################
#Debug Logging
# Writes debug information to Console and to a log file
####################################################################
Function DebugLog {
	Param ( [String] $LogText )
	
	Out-File -NoClobber:$true -Append:$True -FilePath:$script:LogFilePath -InputObject:$LogText
	Write-Host $LogText
}

####################################################################
#Group Processing
# Sets group membership based on search criteria
# Adds members that aren't currently
# optionally removes members not returned by the search
####################################################################
Function ProcessGroup {
	Param ( $SearchResults, [string] $GroupName, [boolean] $RemoveExtraMembers )

	$GroupMemberList = New-Object System.Collections.ArrayList
	
	DebugLog "Group Name: $GroupName"
	$objGroup = Get-ADGroup $GroupName
	
# 	if ($objGroup.GroupCategory -eq "Security") {
# 		$GroupMembers = Get-ADGroupMember $GroupName
# 	} elseif ($objGroup.GroupCategory -eq "Distribution") {
# 		$GroupMembers = Get-DistributionGroupMember $GroupName
# 	}
	
    $GroupDN = $objGroup.distinguishedName
    $filter = "(&(objectcategory=user)(memberof=$GroupDN))" 
    $ds = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"GC://$ForestRoot",$filter) 
    $ds.pagesize = 1000 
    $GroupMembers = $ds.Findall()
    if ($GroupMembers.Count -eq 0) {    
        $ds = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"",$filter)
        $ds.pagesize = 1000 
        $GroupMembers = $ds.Findall()
    }

	foreach ($GrpMember in $GroupMembers) {
		$tmpVal = $GrpMember.properties.distinguishedname[0]
		[void]$GroupMemberList.Add($tmpVal)
	}
# 	Write-Host "---------------"
# 	foreach ($tmp in $GroupMemberList) {Write-Host "{$tmp}"}
# 	Write-Host "---------------"
	Write-Host "Processing Additions"
#	Write-Host $SearchResults
	foreach ($objResult in $SearchResults) {
	    if ($objResult -and ($GroupMemberList -contains $objResult)) { 
	        DebugLog "* $objResult already a member"
	        $GroupMemberList.Remove($objResult)
	    } else {
	       DebugLog "+ Adding $objResult"
           $TmpUserObj = Get-ADUser -Identity $objResult -Server ldap-dal.suddenlink.cequel3.com:3268
	       if ($objGroup.GroupCategory -eq "Security") {
				Add-ADGroupMember -Identity $objGroup -Members $TmpUserObj -Confirm:$false
			} elseif ($objGroup.GroupCategory -eq "Distribution") {
				Add-DistributionGroupMember -Identity $GroupDN -Member $TmpUserObj.distinguishedName -Confirm:$false
			}
	       $GroupMemberList.Remove($objResult[0])
	    }
	}
	Write-Host "Processing removals"
	if ($RemoveExtraMembers -and $GroupMemberList.Count -gt 0) {
		foreach ($GrpMember in $GroupMemberList) {
			if ($GrpMember) {
				DebugLog "- Removing $GrpMember"
               $TmpUserObj = Get-ADUser -Identity $GrpMember -Server ldap-dal.suddenlink.cequel3.com:3268
		       if ($objGroup.GroupCategory -eq "Security") {
					Remove-ADGroupMember -Identity $objGroup -Members $TmpUserObj -Confirm:$false
				} elseif ($objGroup.GroupCategory -eq "Distribution") {
					Remove-DistributionGroupMember -Identity $GroupDN -Member $TmpUserObj.distinguishedName -Confirm:$false
				}
			}
		}
	}
	
	$GroupMemberList.Clear()
}

Function QueryPeopleSoft {
	Param ( [string]$QueryString )
	
	$sCheckActive = "(EMPL_status = 'A')"
	$QueryString = "SELECT EMPLID FROM " + $PeopleSoftOracleEmployeeView + " where " + $sCheckActive + " AND (" + $QueryString + ")"
	Write-Host $QueryString
	$colResults = New-Object System.Collections.ArrayList
	$PSDataSet = New-Object system.data.dataset	
	$PSAdapter = New-Object System.Data.OracleClient.OracleDataAdapter($QueryString, $PeopleOracleCN)
	[void]$PSAdapter.Fill($PSDataSet)
	Write-Host "--QPS PSDataSet: " $PSDataSet.Tables[0].Rows.Count
	foreach($tmpEntry in $PSDataSet.Tables[0]) {
		$strFilter = "employeeid -eq '$($tmpEntry.emplid)'"
		$tmpResults = Get-ADUser -Filter $strFilter
		if ($tmpResults -eq $null) {
			Write-Host "--QPS: $($tmpEntry.emplid) not found"
			$tmpResults = Get-ADUser -Server "cequel.cequel3.com" -Filter $strFilter
		}
		foreach ($tmpEntry in $tmpResults) {
			[void]$colResults.Add($tmpEntry.distinguishedname)
		}
	}
	Write-Host "--QPS colResults: " $colResults.Count
	return $colResults
}

####################################################################
#Main Body
# Define group membership criteria
# Call ProcessGroup with search results and target group name
####################################################################

[void][System.Reflection.Assembly]::LoadwithPartialName("System.Data.OracleClient")
$OracleConnectionString = "User Id=$PeopleSoftOracleUser;Password=$PeopleSoftOraclePassword;Data Source=" + $PeopleSoftOracleServer + ":1521/$PeopleSoftOracleInstance"
$PeopleOracleCN = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
[void]$PeopleOracleCN.Open()

#-------------------------------------------------------------------
#- SPS_KnowledgeLink_CareerLink; add all "Broadband Tech*"
#-------------------------------------------------------------------
  $ADGroupName = "SPS_KnowledgeLink_CareerLink"
  $OracleCommandText = "DESCR = 'Broadband Tech I' 
 	OR DESCR = 'Broadband Tech II' 
 	OR DESCR = 'Broadband Tech III'
 	OR DESCR = 'Broadband Tech IV'
 	OR DESCR = 'Broadband Tech V'
 	OR DESCR = 'Broadband Technician I'
 	OR DESCR = 'Broadband Technician II'
 	OR DESCR = 'Broadband Technician III'
 	OR DESCR = 'Broadband Technician IV'
 	OR DESCR = 'Broadband Technician V'"
 
 DebugLog "Working on $ADGroupName"
 $PSResults = QueryPeopleSoft $OracleCommandText
 ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
  
#-------------------------------------------------------------------
#- AP_BO_IT; all IT
#-------------------------------------------------------------------
 $ADGroupName = "AP_BO_IT"
 $OracleCommandText = "DESCR1 = 'IT'"
 DebugLog "Working on $ADGroupName"
 $PSResults = QueryPeopleSoft $OracleCommandText
 ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers
  
#-------------------------------------------------------------------
#- IT-FTE group; contains all Active Employees in the IT Department
#-------------------------------------------------------------------
$ADGroupName = "IT-FTE"
$OracleCommandText = "DESCR1 = 'IT' AND PER_ORG = 'EMP'"
DebugLog "Working on $ADGroupName"
$PSResults = QueryPeopleSoft $OracleCommandText
ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers

#-------------------------------------------------------------------
#- COR-FTE (location= "CORPORATE", Full time employees only)
#-------------------------------------------------------------------
$ADGroupName = "COR-FTE"
$OracleCommandText = "LOCATION = 'CORPORATE' AND PER_ORG = 'EMP'"
DebugLog "Working on $ADGroupName"
$PSResults = QueryPeopleSoft $OracleCommandText
ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers

#-------------------------------------------------------------------
#- SPS_KnowledgeLink_LeaderContent; copy members of "DL SDL - ALL Managers & Above - All Companies" - Prizm SSR 248
#-------------------------------------------------------------------
 $ADGroupName = "SPS_KnowledgeLink_LeaderContent"
 DebugLog "Working on $ADGroupName"
 $PSResults = New-Object System.Collections.ArrayList
 foreach ($TMPmember in $(Get-ADGroupMember "DL SDL - ALL Managers & Above - All Companies")) {
     [void]$PSResults.add($TMPMember.distinguishedName)
 }
 ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers
$TMPmember = $null

#-------------------------------------------------------------------
#- AP_Duo_Users; AP_Marquee_User; All active employees and contractors in Suddenlink
#-------------------------------------------------------------------
$ADGroupName = "AP_Duo_Users"

$colResults = New-Object System.Collections.ArrayList

$tmpResults = $null
$tmpResults =  Get-ADUser -LDAPFilter "(|(employeenumber=*)(&(employeeid=*)(extensionattribute5=A)))" | Select-Object distinguishedName
 foreach ($tmpEntry in $tmpResults) {
	[void]$colResults.Add($tmpEntry.distinguishedname)
}

DebugLog "Working on $ADGroupName"
$tmpResults = $null
$PSResults = $colResults
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers

$ADGroupName = "AP_Marquee_User"

DebugLog "Working on $ADGroupName"
$PSResults = $colResults
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers


#-------------------------------------------------------------------
#- AP_Duo_Nothing; All inactive employees and contractors in Suddenlink and Cequel
#-------------------------------------------------------------------
$ADGroupName = "AP_Duo_Nothing"

$colResults = New-Object System.Collections.ArrayList

$tmpResults = $null
$tmpResults =  Get-ADUser -Server "ldap-dal.suddenlink.cequel3.com:3268" -LDAPFilter "((employeeid=*)(|(extensionattribute5=S)(extensionattribute5=T)))" | Select-Object distinguishedName
 foreach ($tmpEntry in $tmpResults) {
	[void]$colResults.Add($tmpEntry.distinguishedname)
}

DebugLog "Working on $ADGroupName"
$tmpResults = $null
$PSResults = $colResults
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers

#-------------------------------------------------------------------
#- SPS_EmployeePortal_HealthFirst (Eligible for HealthFirst) - Prizm SSR 437 - May 19, 2014
#-------------------------------------------------------------------
$ADGroupName = "SPS_EmployeePortal_HealthFirst"
$OracleCommandText = "CQ_HF_IND = 'Y'"
DebugLog "Working on $ADGroupName"
$PSResults = QueryPeopleSoft $OracleCommandText
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers

#-------------------------------------------------------------------
#- AP_PaloAlto_AUPBlock ; add non-supervisor Call Center employees
#-------------------------------------------------------------------
  $ADGroupName = "AP_PaloAlto_AUPBlock"
  $OracleCommandText = "DESCR LIKE 'Quality Assurance Specialist%' 
 	OR DESCR LIKE 'Quality Assurance Inspector%'
 	OR DESCR LIKE 'CCR%'
 	OR DESCR LIKE 'Retention Specialist%'
 	OR DESCR LIKE 'Esco CCS%'
 	OR DESCR LIKE 'Voice Operations Spec%'
 	OR DESCR LIKE 'Customer Loyalty Spec%'
 	OR DESCR LIKE 'Collections Specialist%'
 	OR DESCR LIKE 'Monitoring Specialist%'
 	OR DESCR LIKE 'Customer Feedback Spec%'
	OR DESCR LIKE 'ICare Specialist%'"
 
 DebugLog "Working on $ADGroupName"
 $PSResults = QueryPeopleSoft $OracleCommandText
 ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers

#-------------------------------------------------------------------
#- DL SDL - ALL Credit Compliance; Prizm ticket 183171
#-------------------------------------------------------------------
$ADGroupName = "DL SDL - ALL Credit Compliance"

$colResults = New-Object System.Collections.ArrayList

$tmpResults = $null
<#$tmpResults =  get-aduser -Filter {(
(Department -eq 'CC Sales') -or
(Department -eq 'CC Account Services') -or
(Department -eq 'CC Tech Support') -or
(Department -eq 'CC Credit and Collections') -or
(Department -eq 'Customer Service') -or
(Department -eq 'Front Counter') -or
(Department -eq 'Ecommerce') -or
(Department -eq 'Direct and Retail Sales')
) -and (
(Title -like '*Supervisor*') -or
(Title -like '*Manager*') -or 
(Title -like '*Director*')
)} | Select-Object distinguishedName
#>

# Updated query, per Prizm 209522
$tmpResults =  get-aduser -Filter {
(
    (
        (Department -eq 'CC Sales') -or
        (Department -eq 'CC Account Services') -or
        (Department -eq 'CC Tech Support') -or
        (Department -eq 'CC Credit and Collections') -or
        (Department -eq 'Customer Service') -or
        (Department -eq 'Front Counter') -or
        (Department -eq 'Ecommerce') -or
        (Department -eq 'Direct and Retail Sales') -or
        (Department -eq 'Retail Shops')
    ) -and (
        (Title -like '*Supervisor*') -or
        (Title -like '*Manager*') -or 
        (Title -like '*Director*') -or
        (title -like '*VP*' -and title -notlike '*SVP*')
    ) -and -not (
        (Department -eq'Customer Service') -and
        (
            (title -like 'Manager IVR Operations') -or
            (title -like 'Sr Manager Business') -or
            (title -like 'Command Center Supervisor') -or
            (title -like 'Supervisor Commissions')
        )
    )
) -or (
    (Department -eq 'Installation & Service Leaders') -and
    (
        (title -like 'VP Operations') -or
        (title -like '*Director Operations') -or
        (title -like 'Manager System*')
    )
) -or (
    (Department -eq 'Marketing') -and
    (
        (title -like 'Manager Sales*') -or
        (title -like 'Director Sales*') -or
        (title -like 'VP*')
    )
)
} | Select-Object distinguishedName

 foreach ($tmpEntry in $tmpResults) {
	[void]$colResults.Add($tmpEntry.distinguishedname)
}

DebugLog "Working on $ADGroupName"
$PSResults = $colResults
ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers
$tmpResults = $null

#-------------------------------------------------------------------
#- DL SDL - WTX/OK Broadband Technicians; Prizm ticket 205602
#-------------------------------------------------------------------
$ADGroupName = "DL SDL - WTX OK Broadband Technicians"

$vTopUsers = 'Charles.Borthwick','jgebhart'
$vDLList = @()
DebugLog "Working on $ADGroupName"

ForEach($vTopUser in $vTopUsers){
    $vUserList = @()
    $vNextList = @()
    $vFinalList = @()
    $vUserList = Get-ADUser $vTopUser -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports
    $vFinalList = $vUserList

    DO{
        Foreach($vU in $vUserList){
            $vFinalList = $vFinalList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
            $vNextList = $vNextList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
        }
        $vCount = $vNextList.Count
        $vUserList = $vNextList
        $vNextList = @()
    }Until($vCount -eq 0)

    $vFinalList | Foreach{
        $vUser = Get-ADuser $_ -Properties Title
        If($vUser.Title -like 'Broadband Technician*'){$vDLList += ($vUser.distinguishedName).ToString()}
    }
}

$PSResults = $vDLList
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLList
Clear-Variable vFinalList
Clear-Variable vUserList
Clear-Variable vNextList

#-------------------------------------------------------------------
#- DL SDL - Care Center Tech Support Teams
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-CareCenterTechSupportTeams"


$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Tech Support'} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Tech Support Leaders' | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Tech Support Management' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}


ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - Care Center Billing Teams
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-CareCenterBillingTeams"

$vDLMembers = Get-ADUser -Filter {(Title -like 'CCR * - Account Services')} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Billing Leaders' | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - All Care Center Billing Management' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - Care Center Sales Teams
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-CareCenterSalesTeams"

$vDLMembers = Get-ADUser -Filter {(Title -like 'CCR * - Sales')} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Sales Leaders' | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - All Care Center Sales Management' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - Care Center Saves Teams
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-CareCenterSavesTeams"

$vDLMembers = Get-ADUser -Filter {(Title -like 'CSR * - Save')} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADUser -Filter {(Title -like 'Retention Specialist *')} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Saves Leaders' | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - All Care Center Saves Management' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - Care Center Escalations Teams
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-CareCenterEscalationsTeams"

$vDLMembers = Get-ADUser -Filter {(Title -like 'To be decided later')} -Properties Description | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Escalations Leaders' | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - Care Center Escalations Management' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - TYL Support Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-TYLSupportTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Tech Support'} -Properties Office,Department | Where-Object{($_.Office -like '16100')} | Where-Object {$_.Department -notlike 'Commercial Customer Service'} | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - TYL Support Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - TYL Billing Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-TYLBillingTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Account Services'} -Properties Description,Office | Where-Object{($_.Office -like '16100')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - TYL Billing Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - WTX Support Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-WTXSupportTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Tech Support'} -Properties Office,Department | Where-Object{($_.Office -like '20400')} | Where-Object {$_.Department -notlike 'Commercial Customer Service'} | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - WTX Support Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - WTX Billing Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-WTXBillingTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Account Services'} -Properties Description,Office | Where-Object{($_.Office -like '20400')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - WTX Billing Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - ARZ Support Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-ARZSupportTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Tech Support'} -Properties Description,Office | Where-Object{($_.Office -like '30006')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - ARZ Support Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - ARZ Billing Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-ARZBillingTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Account Services'} -Properties Description,Office | Where-Object{($_.Office -like '30006')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - ARZ Billing Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - ATL Support Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-ATLSupportTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Tech Support'} -Properties Description,Office | Where-Object{($_.Office -like '22300')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - ATL Support Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - ATL Billing Team
#-------------------------------------------------------------------
$ADGroupName = "DLSDL-ATLBillingTeam"

$vDLMembers = Get-ADUser -Filter {Title -like 'CCR * - Account Services'} -Properties Description,Office | Where-Object{($_.Office -like '22300')}  | Where-Object{($_.Description -notlike 'Contract*') -and ($_.Description -notlike 'Term*')} | Select distinguishedName
$vDLMembers += Get-ADGroupMember -Identity 'DL SDL - ATL Billing Leaders' | Select distinguishedName

$PSResults = New-Object System.Collections.ArrayList
foreach ($TMPmember in $vDLMembers) {
    [void]$PSResults.add($TMPMember.distinguishedName)
}

ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLMembers
Clear-Variable TMPmember

#-------------------------------------------------------------------
#- DL SDL - All Operations Field Employees
#-------------------------------------------------------------------
$ADGroupName = "DL SDL - All Operations Field Employees"

$vTopUsers = 'DGilles','Rodney.Cates'
$vDLList = @()
DebugLog "Working on $ADGroupName"

ForEach($vTopUser in $vTopUsers){
    $vUserList = @()
    $vNextList = @()
    $vFinalList = @()
    $vUserList = Get-ADUser $vTopUser -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports
    $vFinalList = $vUserList

    DO{
        Foreach($vU in $vUserList){
            $vFinalList = $vFinalList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
            $vNextList = $vNextList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
        }
        $vCount = $vNextList.Count
        $vUserList = $vNextList
        $vNextList = @()
    }Until($vCount -eq 0)

    $vFinalList | Foreach{
        $vUser = Get-ADuser $_ 
        $vDLList += ($vUser.distinguishedName).ToString()
    }
}

ForEach($vTopUser in $vTopUsers){
	$vDLList += ((Get-ADUser $vTopUser).DistinguishedName).ToString()
}

$PSResults = $vDLList
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLList
Clear-Variable vFinalList
Clear-Variable vUserList
Clear-Variable vNextList

#-------------------------------------------------------------------
#- DL SDL - All Operations Field Employees
#-------------------------------------------------------------------
$ADGroupName = "DL SDL - All Care Center Employees"

$vTopUsers = 'Gibbs.Jones'
$vDLList = @()
DebugLog "Working on $ADGroupName"

ForEach($vTopUser in $vTopUsers){
    $vUserList = @()
    $vNextList = @()
    $vFinalList = @()
    $vUserList = Get-ADUser $vTopUser -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports
    $vFinalList = $vUserList

    DO{
        Foreach($vU in $vUserList){
            $vFinalList = $vFinalList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
            $vNextList = $vNextList + (Get-ADUser $vU -Properties DirectReports | Select DirectReports -ExpandProperty DirectReports)
        }
        $vCount = $vNextList.Count
        $vUserList = $vNextList
        $vNextList = @()
    }Until($vCount -eq 0)

    $vFinalList | Foreach{
        $vUser = Get-ADuser $_ 
        $vDLList += ($vUser.distinguishedName).ToString()
    }
}

ForEach($vTopUser in $vTopUsers){
$vDLList += ((Get-ADUser $vTopUser).DistinguishedName).ToString()
}

$vUserList1 = Get-ADUser -Filter {((Department -like 'CC Account Services') -or (Department -like 'CC Credit and Collections') -or (Department -like 'CC CSR in Training') -or (Department -like 'CC Sales') -or (Department -like 'CC Tech Support') -or (Department -like 'Customer Care') -or (Department -like 'Customer Service') -or (Department -like 'iCare') -or (Department -like 'NOC') -or (Department -like 'HR') -or (Department -like 'Voice Operations Customer Care') -or (Department -like 'General Admin')) -and ((Title -like 'Facilities Assistant') -or (Title -like 'Facilities Coordinator II') -or (Title -like 'Facilitates Supervisor') -or (Title -like 'Facility Tech*') -or (Title -like 'Manager Facilities') -or (Title -like 'VP Customer Care') -or (Title -like 'Administrative Assistant*') -or (Title -like 'Administrative Coordinator*') -or (Title -like 'Manager Business Operations') -or (Title -like 'VP HR - Regional'))}
ForEach($vU in $vUserList1){
$vDLList += ((Get-ADUser $vU).DistinguishedName).ToString()
}

$PSResults = $vDLList
ProcessGroup $PSResults $ADGroupName $bRemoveExtraMembers
Clear-Variable vDLList
Clear-Variable vFinalList
Clear-Variable vUserList
Clear-Variable vUserList1
Clear-Variable vNextList

#-------------------------------------------------------------------
#- New group (1 of 2) requested by Angela Stricklin - Ticket 391965
#-------------------------------------------------------------------
# $ADGroupName = "DL SDL - ETX Mid-South Retail Sales Agents"
# $OracleCommandText = "((JOBCODE = 'RSA001') OR (JOBCODE = 'RSA002') OR (JOBCODE = 'RSA003')) AND ((BUSINESS_UNIT = '30700') OR (BUSINESS_UNIT = '30702'))"
# DebugLog "Working on $ADGroupName"
# $PSResults = QueryPeopleSoft $OracleCommandText
# ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers

#-------------------------------------------------------------------
#- New group (2 of 2) requested by Angela Stricklin - Ticket 391965
#-------------------------------------------------------------------
# $ADGroupName = "DL SDL - ETX MID-South Retail Sales Managers"
# $OracleCommandText = "((JOBCODE = 'CC0047') OR (JOBCODE = 'CC045') OR (JOBCODE = 'CC0048') OR (JOBCODE = 'CC0049')) AND ((BUSINESS_UNIT = '30700') OR (BUSINESS_UNIT = '30702'))"
# DebugLog "Working on $ADGroupName"
# $PSResults = QueryPeopleSoft $OracleCommandText
# ProcessGroup $PSResults $ADGroupName $bNoRemoveExtraMembers

#-------------------------------------------------------------------
#- AJA-Testing group; Any account with "AJA-Testing" in the extensionAttribute1, remove extra members
#-------------------------------------------------------------------
# $ADGroupName = "AJA-Testing"
# $colResults = New-Object System.Collections.ArrayList
# 
# $strFilter = "(&(objectCategory=User)(extensionAttribute1=$ADGroupName))"
# $objSearcher.Filter = $strFilter
# $tmpResults = $objSearcher.FindAll()
#  foreach ($tmpEntry in $tmpResults) {
#  	$tmpItems = $tmpEntry.Properties
#  	#Write-Host $tmpItems.distinguishedname
# 	[void]$colResults.Add($tmpItems.distinguishedname[0])
# }
# DebugLog "Working on $ADGroupName"
# ProcessGroup $colResults $bRemoveExtraMembers
# $colResults.clear()
# $tmpResults = $null
# $tmpItems = $null
# 
#-------------------------------------------------------------------
