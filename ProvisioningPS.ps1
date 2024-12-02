# ==============================================================================================
# NAME: ProvisioningPS.ps1
# AUTHOR: Andy Adrian , SuddenLink
# DATE  : 11/10/2008
# COMMENT: Account Provisioning / de-provisioning script
## 5/21/2013 - DONE - Added required functions for PeopleSoft Multi-factor Authentication
##			 - DONE - Added "Notice" check to Termination section - SSR 1061
##			 - DONE - Changed path for user creation to Windows7 OU in 'CreateAccount' function
##			 - DONE - Removed unused "AddtoGroup" function
## 			 - DONE - add MFA logic and function calls to Main
## 5/22/2013 - DONE - Removed TrackIT Self-Service group references (no longer needed)
# ==============================================================================================

Param ( 
    [switch]$SDLDebug,
    [switch]$FullSync,
    [switch]$DailySync,
    [switch]$SingleSync,
    [string]$SingleEmplID
)

$ErrorActionPreference = "stop"
Try {
	$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sdlstlpwemb01d.suddenlink.cequel3.com/PowerShell/ -Authentication Kerberos
	import-pssession $ExchSession
 } Catch {
	throw $_ 
}
$ErrorActionPreference = "Continue"
	
Import-Module -Name ActiveDirectory
Import-Module -Name Lync
Add-PSSnapin -Name SqlServerCmdletSnapin100

if ((Get-DomainController -DomainName suddenlink.cequel3.com | ?{$_.ADSite -like "*SDL-DAL"}).count -gt 1) {
	$CurDC = (Get-DomainController -DomainName suddenlink.cequel3.com | ?{$_.ADSite -like "*SDL-DAL"})[0].dnsHostName
} else {
	$CurDC = (Get-DomainController -DomainName suddenlink.cequel3.com | ?{$_.ADSite -like "*SDL-DAL"}).dnsHostName
}
Set-ADServerSettings -ViewEntireForest:$True -PreferredServer $CurDC
Remove-Variable CurDC

$DebugPreference = "Continue"
if ($SDLDebug) {$bDebugging = $True}
else {$bDebugging = $false}

$LogFilePath = (Get-Date -Format M-dd-yyyy-hh-mmtt)
if ($FullSync) {$LogFilePath += "-FULLSYNC"}
$LogFilePath = "C:\Scripts\ADAccountProvisioning\Logs\" + $LogFilePath + ".log"

$objRootDSE = [ADSI]"LDAP://rootDSE"
$ForestRoot = $objRootDSE.Get("rootDomainNamingContext")
$domainDN = $objRootDSE.Get("defaultNamingContext")
$domainDNS = "suddenlink.cequel3.com"
$objConnection = New-Object -ComObject "ADODB.Connection"
[void]$objConnection.Open("Provider=ADsDSOObject;")
$ADS_SCOPE_SUBTREE = 2
$ADS_PROPERTY_CLEAR = 1

$AccountProvisioningSQLServer="accountprovisioning-db"
$AccountProvisioningSQLDatabase="AccountProvisioning"

$PeopleSoftOraclePOIMappingDB="sysadm.ps_cq_poi_xref_ids"
$PeopleSoftOracleServer="HRPROD1.suddenlink.cequel3.com"
$PeopleSoftOracleInstance="hrprd"
$PeopleSoftOracleEmployeeView="PS_CQ_DIR_INTFCDAT"
$PeopleSoftOracleUser="ACT_USER"
$PeopleSoftOraclePassword="n3m4s"

$EmailNotificationFROMaddr = "ADProvisioning@Suddenlink.com"
$EmailNotificationDL = "DLSDL-ALLAccountProvisioning@Suddenlink.com"
$EmailNotificationBCCaddr = "DLSDL-ALLProvisioningBCC@suddenlink.com"
$EmailNotificationDebuggingTOaddr = "andy.adrian@suddenlink.com"

$CVCServiceNowEmailAddr = "cablevision@service-now.com"
$AUSEmailDomain = "@alticeusa.com"
$CVCDomainObj = New-Object System.DirectoryServices.DirectoryEntry("LDAP://ldapauth.cablevision.com:636/OU=NetIDUsers,DC=cvcauth,DC=com",'CVCAUTH\svcmailinteg','Monday123')
$CVCSearcherObj = New-Object System.DirectoryServices.DirectorySearcher
$CVCSearcherObj.SearchRoot = $CVCDomainObj
$CVCSearcherObj.PageSize = 1000
[void]$CVCSearcherObj.PropertiesToLoad.Add("displayname")
[void]$CVCSearcherObj.PropertiesToLoad.Add("proxyaddresses")
$CVCSearcherObj.SearchScope = "Subtree"

$AllowedLogonList = ""
foreach ($tempDC in (Get-DomainController | Sort-Object Name)) {
    $AllowedLogonList += $tempDC.Name + ","
}

foreach ($tempComp in (Get-ADComputer -Filter {Name -like '*lmfa*'} | Sort-Object Name)) {
    $AllowedLogonList += $tempComp.Name + ","
}

$AllowedLogonList = $AllowedLogonList.Substring(0,$AllowedLogonList.Length-1)
Remove-Variable tempDC,tempComp

$CurPW = ""
$CurAdminAccounts = $null
$CurTestAccounts = $null

######################################################################
#Execute Provisioning DB*
######################################################################
Function AccountProvisioningDBEx {
	Param ([string]$AccountProvisioningDBEXSQLStr)
	
	#DebugLog	$AccountProvisioningDBEXSQLStr
	$AccountProvisioningDBExCN = New-Object -ComObject "ADODB.Connection"
	if ($bDebugging -eq $false) {
	 	$AccountProvisioningDBExCN.ConnectionString = "Provider=sqloledb;Data Source='$AccountProvisioningSQLServer';Initial Catalog=$AccountProvisioningSQLDatabase;Integrated Security=SSPI;"
	 	[void]$AccountProvisioningDBExCN.Open()
	 	[void]$AccountProvisioningDBExCN.Execute($AccountProvisioningDBEXSQLStr)
	 	[void]$AccountProvisioningDBExCN.Close()
	}
	$AccountProvisioningDBExCN = $null
}

######################################################################
#Query Account - Return DN or FALSE*
# TESTED
######################################################################
Function QueryAccount {
	Param ([string]$EmployeeNum)
	DebugLog " -- Finding DN for $EmployeeNum"
	if ($EmployeeNum) {
		$QueryAccountResult = (Get-ADUser -Filter {Employeeid -eq $EmployeeNum}).distinguishedName
		if ($QueryAccountResult -eq $null) { 
    	    $OracleCommandText = "Select CQ_CVC_EMPLID FROM " + $PeopleSoftOraclePOIMappingDB + " where (EMPLID = '$EmployeeNum')"
		
		    $OraclePOILookupCN = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
		    [void]$OraclePOILookupCN.Open()
		    DebugLog "Oracle connection open"
		    $OraclePOILookupCMD = New-Object System.Data.OracleClient.OracleCommand($OracleCommandText, $OraclePOILookupCN)
		    $OraclePOILookupReader = $OraclePOILookupCMD.ExecuteReader()
            while ($OraclePOILookupReader.read()) {
        	    $CurCVCEMPLID = [string]$OraclePOILookupReader.GetValue($OraclePOILookupReader.GetOrdinal("CQ_CVC_EMPLID"))
                debuglog $CurCVCEMPLID
            }
            $OraclePOILookupCN.close()
            if ($CurCVCEMPLID -and $CurCVCEMPLID.count -gt 1) {
                DebugLog "Multiple entries returned, should not happen"
            } elseif (!$CurCVCEMPLID){
                DebugLog "Not found"
		        $QueryAccountResult = $false
            } else {
                $CurCVCContactObj = Get-MailContact -Filter "CustomAttribute1 -like $CurCVCEMPLID"
                if ($CurCVCContactObj) {
                    $QueryAccountResult = $CurCVCContactObj.distinguishedName
                } else {
		            $QueryAccountResult = $false
                }
                
            }
<#			$QueryAccountResult = (Get-ADUser -Server "cequel.cequel3.com" -Filter {Employeeid -eq $EmployeeNum}).distinguishedName
				if ($QueryAccountResult -eq $null) { 
					$QueryAccountResult = $False 
				}
#>
		}
	} else {
		$QueryAccountResult = $false
	}
	DebugLog " --- $QueryAccountResult"
	return $QueryAccountResult
}
######################################################################
#Query Sam Account - Return SamAccount or FALSE*
# TESTED
######################################################################
Function QuerySAMAccount {
	Param ([string]$CurEmplDN)
	
	DebugLog " -- Finding UserID for $CurEmplDN"
	if ($CurEmplDN.contains("DC=cequel,DC=cequel3,DC=com")) {
		$QuerySAMAccount = Get-ADUser -Server (Get-DomainController -DomainName "cequel.cequel3.com")[0].DnsHostName $CurEmplDN -Properties SamAccountName
	} else {
		$QuerySAMAccount = Get-ADUser $CurEmplDN -Properties SamAccountName
	}

	$QuerySAMAccountResult = $QuerySAMAccount.SamAccountName
	DebugLog " --- $QuerySAMAccountResult"
	return $QuerySAMAccountResult
}

######################################################################
#FindExtraAccounts - Find test and admin accounts
#
######################################################################

function FindExtraAccounts {
	Param (
		[string]$CurSamAccountName
	)
	
	$script:CurAdminAccounts = $null
	$script:CurTestAccounts =  $null
	
	DebugLog " -- Finding Admin and Test accounts for $CurSamAccountName"
	$TmpAdminAccountName = "a_" + $CurSamAccountName
	$script:CurAdminAccounts = Get-ADUser -Filter {samAccountName -eq $TmpAdminAccountName} -properties description,memberof
	$TmpTestAccountName = "t_" + $CurSamAccountName
	$script:CurTestAccounts = Get-ADUser -Filter {samAccountName -eq $TmpTestAccountName} -properties description,memberof
}

######################################################################
#DisableExtraAccounts - Find and disable test and admin accounts
#
######################################################################
Function DisableExtraAccounts {
	Param (
		[string]$CurSamAccountName,
		[string]$Region,
		[boolean]$isSuspended
	)
	
		
	if ($script:CurAdminAccounts) {
		foreach ($user in $script:CurAdminAccounts) {
			DebugLog " --- Found admin account ($($user.SamAccountName)), disabling..."
			if ($isSuspended) {
				$tmpDescription = $user.description
                if($tmpDescription -notlike "SUSPENDED - ") {
				    $tmpDescription = "SUSPENDED - " + $tmpDescription
                }
			} else {
				$tmpDescription = $user.description
                if ($tmpDescription -notlike "TERM - ") {
				    $tmpDescription = "TERM - " + $tmpDescription
                }
				DebugLog " ----- Removing from AD Groups..."
				$GroupList = $user.memberof
				foreach ($tmpGroup in $GroupList){
					DebugLog " ----- $tmpGroup"
					Remove-ADGroupMember -Identity $tmpGroup -Members $user -Confirm:$false
				}
			}
			DisableEmailAccess -UDN $user.distinguishedName
			set-ADUser $user -Enabled $false -description $tmpDescription
		}
	}
	
	if ($script:CurTestAccounts) {
		foreach ($user in $script:CurTestAccounts) {
			DebugLog " --- Found test account ($($user.SamAccountName)), disabling..."
			if ($isSuspended) {
				$tmpDescription = $user.description
				$tmpDescription = "SUSPENDED - " + $tmpDescription
			} else {
				$tmpDescription = $user.description
				$tmpDescription = "TERM - " + $tmpDescription
			}
			set-ADUser $user -Enabled $false -description $tmpDescription
			DisableEmailAccess -UDN $user.distinguishedName
		}
	}
}
######################################################################
#EmailManagerStatus
######################################################################
Function EmailManagerStatus {
	Param (
		[string]$EmailManagerStatusPass,
		[string]$EmailManagerStatusManagerID,
		[string]$EmailManagerStatusAddr,
		[string]$EmailManagerStatusOp,
		[string]$EmailManagerStatusSN,
		[string]$EmailManagerStatusMI,
		[string]$EmailManagerStatusGiven,
		[string]$EmailManagerStatusEMPID,
		[string]$EmailManagerStatussamaccount,
		[string]$EmailManagerStatusRegion,
		[string]$EmailManagerStatusTitle,
		[string]$EmailManagerStatusLocationID,
		[string]$EmailManagerStatusLocationName,
		[string]$EmailManagerStatusBUSUnit,
		[string]$EmailManagerStatusEffDate,
		[string]$EmailManagerStatusLastAction,
		[string]$EmailManagerStatusStatus,
		[string]$EmailManagerStatusConferenceCard,
		[string]$EmailManagerStatusCompanyPhone,
		[string]$EmailManagerStatusICOMSSalesID,
		[string]$EmailManagerStatusEmplType,
		[string]$EmailManagerStatusVendorName,
		[string]$EmailManagerStatusMobileNumber,
		[string]$EmailManagerStatusDeskNumber,
		[string]$EmailManagerStatusICOMSID,
		[string]$EmailManagerStatusDeptID
	)
	
	DebugLog " --- Sending update email"
	DebugLog " ---- $EmailManagerStatusAddr"
	DebugLog " ---- Message body follows ----"
	
	$EmailManagerStatusobjEmail = New-Object system.net.mail.smtpClient
	$EmailManagerStatusobjEmail.host = "smtp.suddenlink.com"
	$ADProvisionAddr = New-Object system.Net.Mail.MailAddress($EmailNotificationFROMaddr)
	if ($script:bDebugging -eq $false) {
		$ADProvisionDLAddr = New-Object system.Net.Mail.MailAddress($EmailNotificationDL)
        if ($EmailManagerStatusRegion -like "* *") {$EmailManagerStatusRegion = $EmailManagerStatusRegion.Replace(" ","")}
		$ArrTo = ("DLSDL-ALL$EmailManagerStatusRegion"+"AccountProvisioning@suddenlink.com")
	} else {
		$ADProvisionDLAddr = $ArrTo = $EmailNotificationDebuggingTOaddr
	}
	
	$EmailManagerStatusobjMsg = New-Object system.Net.Mail.MailMessage($ADProvisionAddr, $ADProvisionDLAddr)

	foreach ($Addr in $ArrTo) {
		$TempTo = New-Object system.Net.Mail.MailAddress($Addr)
		[void]$EmailManagerStatusobjMsg.To.Add($TempTo)
	}
	
	if ($script:bDebugging -eq $false) {
		if ($EmailManagerStatusAddr -ne $null) {
			$EmailCC = New-Object system.Net.Mail.MailAddress($EmailManagerStatusAddr)
			[void]$EmailManagerStatusobjMsg.CC.Add($EmailCC)
		}

	 	$TempBCC = New-Object system.Net.Mail.MailAddress($EmailNotificationBCCaddr)
	 	[void]$EmailManagerStatusobjMsg.BCC.Add($TempBCC)
	}
	$EmailManagerStatusEffDate = get-date($EmailManagerStatusEffDate) -format MM/dd/yyyy
	
	$EmailManagerStatusobjMsg.Subject = "Employee  $EmailManagerStatusOp : $EmailManagerStatusEMPID"
	$EmailManagerStatusobjMsg.Body = "Employee $EmailManagerStatusOp"
	$EmailManagerStatusobjMsg.Body += "`nLogin Account Name: $EmailManagerStatussamaccount"
	if ($EmailManagerStatusMI)  {
		$EmailManagerStatusobjMsg.Body += "`nUser: $EmailManagerStatusSN, $EmailManagerStatusGiven $EmailManagerStatusMI"
	} else {
		$EmailManagerStatusobjMsg.Body += "`nUser: $EmailManagerStatusSN, $EmailManagerStatusGiven"	
	}
	$EmailManagerStatusobjMsg.Body += "`nEmployee Number: $EmailManagerStatusEMPID"
	$EmailManagerStatusobjMsg.Body += "`nICOMS SalesRep ID/Tech ID: $EmailManagerStatusICOMSSalesID"
	$EmailManagerStatusobjMsg.Body += "`nICOMS Login: $EmailManagerStatusICOMSID"
	$EmailManagerStatusobjMsg.Body += "`nEmployment Type: $EmailManagerStatusEmplType"
	if ($EmailManagerStatusVendorName -ne "") {
		$EmailManagerStatusobjMsg.Body += "`nContractor Vendor: $EmailManagerStatusVendorName"
	}
	$EmailManagerStatusobjMsg.Body += "`nTitle: $EmailManagerStatusTitle"
	$EmailManagerStatusobjMsg.Body += "`nLocation Code: $EmailManagerStatusLocationID"
	$EmailManagerStatusobjMsg.Body += "`nLocation: $EmailManagerStatusLocationName"
	$EmailManagerStatusobjMsg.Body += "`nBusiness Unit: $EmailManagerStatusBUSUnit"
	$EmailManagerStatusobjMsg.Body += "`nDepartment Number: $EmailManagerStatusDeptID"
	$EmailManagerStatusobjMsg.Body += "`nManager EmployeeID: $EmailManagerStatusManagerID"
	$EmailManagerStatusobjMsg.Body += "`nConference Card: $EmailManagerStatusConferenceCard"
	$EmailManagerStatusobjMsg.Body += "`nCompany Cell Phone Number: $EmailManagerStatusMobileNumber"
	$EmailManagerStatusobjMsg.Body += "`nOffice Desk Phone Number: $EmailManagerStatusDeskNumber"
	$EmailManagerStatusobjMsg.Body += "`nLast Action Effective Date: $EmailManagerStatusEffDate"
	$EmailManagerStatusobjMsg.Body += "`nLast Action: $EmailManagerStatusLastAction"
	$EmailManagerStatusobjMsg.Body += "`nStatus: $EmailManagerStatusStatus"
	$EmailManagerStatusobjMsg.Body += "`nProvisioning Operation: $EmailManagerStatusOp"
	
	if (($EmailManagerStatusOp -eq "Creation") -or  ($EmailManagerStatusOp -eq "Re-Hire")) {
		$EmailManagerStatusobjMsg.Body += "`nPassword(Change Required At First Login): " + $EmailManagerStatusPass
		$EmailManagerStatusobjMsg.Body += "`n`nPlease note the new Suddenlink employee will be prompted to change their password after their initial sign on.  Please have the new hire login to the Suddenlink domain at the CTRL+ALT+DELETE screen.  At the time of the first login, the new hire will be prompted to change their initial password. This password needs to be changed prior to signing onto the Employee Portal.`nIf you have any questions, please contact the Helpdesk at (866) 244-3827."
	}
	DebugLog $EmailManagerStatusobjMsg.Body
	DebugLog " ---- End Message body ----"

	[void]$EmailManagerStatusobjEmail.Send($EmailManagerStatusobjMsg)
}

######################################################################
# EmailProvisioningCompleteStatus
######################################################################
Function EmailProvisioningCompleteStatus {

	DebugLog "Sending completion email"

	$EmailProvisioningCompleteobjEmail = New-Object system.Net.Mail.SmtpClient
	$EmailProvisioningCompleteobjEmail.host = "smtp.suddenlink.com"
	
	$EmailFrom = New-Object system.Net.Mail.MailAddress($EmailNotificationFROMaddr)
	if ($script:bDebugging -eq $false) {
	 	$EmailTo = New-Object system.Net.Mail.MailAddress($EmailNotificationBCCaddr)
	 } else {
		$EmailTo = New-Object system.Net.Mail.MailAddress($EmailNotificationDebuggingTOaddr)
	 }
	 $EmailProvisioningCompleteobjMsg = New-Object system.Net.Mail.MailMessage($EmailFrom,$EmailTo)	
	 $EmailProvisioningCompleteobjMsg.Subject = "Employee Provisioning Process has completed"

 	[void]$EmailProvisioningCompleteobjEmail.Send($EmailProvisioningCompleteobjMsg)
}

######################################################################
#Check for Existing Sam Account Name (During Creation) - Return TRUE or FALSE*
# TESTED
######################################################################
Function DuplicateAccount {
	Param ([string]$SAMName)
	
	DebugLog "Checking for duplicate account for $SAMName"
	
	$result = $False
	if (Get-ADUser -Filter {SamAccountName -eq $SAMName}) {$result = $True}
	return $result
}

#####################################################################
#Update Account
######################################################################
Function UpdateAccount {
	Param (
		[string]$UpdateAccountDN,
		[string]$UpdateAccountEMPLID,
		[string]$UpdateAccountFIRST_NAME,
		[string]$UpdateAccountLAST_NAME,
		[string]$UpdateAccountMIDDLE_INITIAL,
		[string]$UpdateAccountJobTitle,
		[string]$UpdateAccountLOCATION,
		[string]$UpdateAccountADDRESS1,
		[string]$UpdateAccountADDRESS2,
		[string]$UpdateAccountCITY,
		[string]$UpdateAccountSTATE,
		[string]$UpdateAccountPOSTAL,
		[string]$UpdateAccountManagerDN,
		[string]$UpdateAccountBUSINESS_UNIT,
		[string]$UpdateAccountDept,
		[boolean]$isConsultant,
		[string]$UpdateAccountVendorName,
		[boolean]$UpdateAccountisSuspended,
		[string]$UpdateAccountCompany,
		[string]$UpdateAccountRegion,
		[string]$UpdateAccountPSRegion,
		[string]$UpdateAccountDeptID,
		[string]$UpdateAccountICOMSID
	) 
	
	if ($isConsultant) {
		$updateAccountDescription = "Contractor - $UpdateAccountJobTitle"
	} else {
		$updateAccountDescription = $UpdateAccountJobTitle
	}
	
	if ($UpdateAccountisSuspended) {
		$updateAccountDescription = "SUSPENDED - $updateAccountDescription"
	}
	
	$UpdateAccountobjUser = Get-ADUser -Filter {employeeId -eq $UpdateAccountEMPLID} -Properties *
	if ($UpdateAccountobjUser -eq $null) { 
		$UpdateAccountobjUser = Get-ADUser -Server 'cequel.cequel3.com' -Filter {Employeeid -eq $UpdateAccountEMPLID} -Properties *
	}

	if ($UpdateAccountobjUser.ExtensionAttribute8 -eq 'DISABLED') {
		DebugLog " ---- Account not used, ignoring. (UpdateAccount)"
	} else {
	
		$UpdateAccountobjUserMailbox = Get-Mailbox $UpdateAccountDN -IgnoreDefaultScope
		
		if ($UpdateAccountFIRST_NAME -and ($UpdateAccountFIRST_NAME -ne $UpdateAccountobjUser.givenName)) {
			Set-User -Identity $UpdateAccountDN -FirstName $UpdateAccountFIRST_NAME -IgnoreDefaultScope
			DebugLog " ---- Updated FirstName"
		}
		if ($UpdateAccountLAST_NAME -and $UpdateAccountLAST_NAME -ne $UpdateAccountobjUser.sn) {
			Set-User -Identity $UpdateAccountDN -LastName $UpdateAccountLAST_NAME -IgnoreDefaultScope
			DebugLog " ---- Updated Last"
		}
		if ($UpdateAccountLAST_NAME -and $UpdateAccountFIRST_NAME -and ($UpdateAccountobjUser.displayName -ne "$UpdateAccountLAST_NAME, $UpdateAccountFIRST_NAME")) {
			Set-User -Identity $UpdateAccountDN -displayName "$($UpdateAccountLAST_NAME.trim()), $($UpdateAccountFIRST_NAME.trim())" -IgnoreDefaultScope
			DebugLog " ---- Updated displayName"
		}
		if ($UpdateAccountMIDDLE_INITIAL -and ($UpdateAccountMIDDLE_INITIAL -ne $UpdateAccountobjUser.initials)) {
			Set-User -Identity $UpdateAccountDN -Initials $UpdateAccountMIDDLE_INITIAL -IgnoreDefaultScope
			DebugLog " ---- Updated Initials"
		}
		if ($UpdateAccountJobTitle) { 
			if($UpdateAccountobjUser.title -ne $UpdateAccountJobTitle) {
				Set-User -Identity $UpdateAccountDN -Title $UpdateAccountJobTitle -IgnoreDefaultScope
				DebugLog " ---- Updated title"
				$tmpManagementLvl = GetManagementLevel($UpdateAccountJobTitle)
				SetMailboxLevel $UpdateAccountDN $tmpManagementLvl $false
			}
			$TmpObjUser = [ADSI]"LDAP://$UpdateAccountDN"
			if ((!$TmpObjUser.description) -or ($TmpObjUser.description -ne $updateAccountDescription)) {
				$TmpObjUser.description = $updateAccountDescription
				[void]$TmpObjUser.setinfo()
				DebugLog " ---- Updated description"
			} 
		}
		if ($UpdateAccountLOCATION) {
			if ($UpdateAccountLOCATION -ne $UpdateAccountobjUser.Office) {
				Set-User -Identity $UpdateAccountDN -Office $UpdateAccountLOCATION -IgnoreDefaultScope
				DebugLog " ---- Updated physicalDeliveryOfficeName"
			}
		}
		if ($UpdateAccountADDRESS1){
			if ($UpdateAccountobjUser.streetAddress -notmatch $UpdateAccountADDRESS1) {
				$tmpAddress = $UpdateAccountADDRESS1
				if ($UpdateAccountADDRESS2 -and ($UpdateAccountobjUser.streetAddress -notmatch $UpdateAccountADDRESS2)) {
					$tmpAddress += " $UpdateAccountAddress2"
				}
				$tmpAddress = $tmpAddress.trim()
				Set-User -Identity $UpdateAccountDN -streetAddress $tmpAddress -IgnoreDefaultScope
				DebugLog " ---- Updated streetAddress to $tmpAddress"
			}
		}
		if ($UpdateAccountCITY){
			if ($UpdateAccountCITY -ne $UpdateAccountobjUser.City) {
				Set-User -Identity $UpdateAccountDN -City $UpdateAccountCITY -IgnoreDefaultScope
				DebugLog " ---- Updated city"
			}
		}
		if ($UpdateAccountSTATE) {
			if ($UpdateAccountobjUser.State -ne $UpdateAccountSTATE) {
				Set-User -Identity $UpdateAccountDN -StateOrProvince $UpdateAccountSTATE.ToUpper() -IgnoreDefaultScope
				DebugLog " ---- Updated state"
			}
		}
		if ($UpdateAccountPOSTAL) {
			if ($UpdateAccountobjUser.postalCode -ne $UpdateAccountPOSTAL) {
				Set-User -Identity $UpdateAccountDN -postalCode $UpdateAccountPOSTAL.Substring(0,5) -IgnoreDefaultScope
				DebugLog " ---- Updated postalCode"
			}
		}

		if(($UpdateAccountobjUser.countryorRegion -ne "US")) {
			Set-User -Identity $UpdateAccountDN -CountryOrRegion "US" -IgnoreDefaultScope
			DebugLog " ---- Updated countryCode"
		}
		if($UpdateAccountobjUser.department -ne $UpdateAccountDept) {
			Set-User -Identity $UpdateAccountDN -department $UpdateAccountDept -IgnoreDefaultScope 
			DebugLog " ---- Updated department"
		}
		if($isConsultant) {
			if($UpdateAccountobjUser.company -ne $UpdateAccountVendorName) {
				Set-User -Identity $UpdateAccountDN -company  $UpdateAccountVendorName -IgnoreDefaultScope 
				DebugLog " ---- Updated company"
			}
		} else {
			if($UpdateAccountobjUser.company -ne $UpdateAccountCompany) {
				Set-User -Identity $UpdateAccountDN -company  $UpdateAccountCompany -IgnoreDefaultScope
				DebugLog " ---- Updated company"
			}
		}
		if (($UpdateAccountDeptID) -and ($UpdateAccountobjUser.departmentNumber -ne $UpdateAccountDeptID -or !($UpdateAccountobjUser.departmentNumber))) {
			Set-ADUser -Identity $UpdateAccountDN -replace @{departmentNumber=$UpdateAccountDeptID}
			DebugLog " ---- Updated Department Number"

		}

		if (($UpdateAccountManagerDN) -and ($UpdateAccountManagerDN -ne $false)-and ($UpdateAccountManagerDN -ne $UpdateAccountDN)) {
			if ($UpdateAccountobjUser.manager) {
				$TMPManager = (get-user $UpdateAccountobjUser.manager -IgnoreDefaultScope).distinguishedName
		
				if($TMPManager -ne $UpdateAccountManagerDN) {
					Set-User -Identity $UpdateAccountDN -manager $UpdateAccountManagerDN -IgnoreDefaultScope
					DebugLog " ---- Updated manager"
				}
			} else {
				Set-User -Identity $UpdateAccountDN -manager $UpdateAccountManagerDN -IgnoreDefaultScope
				DebugLog " ---- Updated manager"
			}
		}
		if ($UpdateAccountobjUserMailbox.CustomAttribute2 -ne $UpdateAccountRegion) {
			set-Mailbox -identity $UpdateAccountDN -CustomAttribute2 $UpdateAccountRegion
			DebugLog " ---- Updated IT Region"
		}
		if ($UpdateAccountobjUserMailbox.CustomAttribute4 -ne $UpdateAccountPSRegion) {
			set-Mailbox -identity $UpdateAccountDN -CustomAttribute4 $UpdateAccountPSRegion
			DebugLog " ---- Updated PS Region"
		}
		if ($UpdateAccountobjUserMailbox.CustomAttribute6 -ne $UpdateAccountICOMSID) {
			set-Mailbox -identity $UpdateAccountDN -CustomAttribute6 $UpdateAccountICOMSID
			DebugLog " ---- Updated ICOMS ID"
		}
		
		#Run Update-Recipient and Set-CSUser to repair potential account issues
		Update-Recipient -Identity $UpdateAccountDN
		Set-CsUser $UpdateAccountDN
	}
}

######################################################################
#Check Disabled Account - NOT USED (AJA - 6/10/2016)
# TESTED
######################################################################
Function CheckDisabledAccount {
	Param ([string]$CheckDisabledAccountUDN)
	DebugLog " -- Checking if account is disabled"
	if ($CheckDisabledAccountUDN.contains("DC=cequel,DC=cequel3,DC=com")) {
		$result = (Get-ADUser $CheckDisabledAccountUDN -Server "cequel.cequel3.com").Enabled
	} else {
		$result = (Get-ADUser $CheckDisabledAccountUDN).Enabled
	}
	DebugLog " --- Account.Enabled = $result"
	$result = !$result
	return $result
}

######################################################################
#Reenable Disabled Account
# Returns new password for account
# TESTED
######################################################################
Function ReenableDisabledAccount {
	param (
		[string]$UDN,
		[string]$Region
	)
	
	if ($UDN.contains("DC=cequel,DC=cequel3,DC=com")) {
		$ReenableUserobjUser = Get-ADUser $UDN -Server "cequel.cequel3.com" -Properties givenName,employeeID,sn,employeeType,extensionattribute5,LogonWorkstations,ExtensionAttribute8
	} else {
		$ReenableUserobjUser = Get-ADUser $UDN -Properties givenName,employeeID,sn,employeeType,extensionattribute5,LogonWorkstations,ExtensionAttribute8
	}
    
	if ($UpdateAccountobjUser.ExtensionAttribute8 -eq 'DISABLED') {
		DebugLog " ---- Account not used, ignoring. (ReenableDisabledAccount)"
	} else {

#   if ($ReenableUserobjUser.Enabled -eq $false) {
#		DebugLog " ---- Account disabled, Re-enabling (AD Disabled)"
#		Set-ADUser -Identity $ReenableUserobjUser -Enabled $true -Confirm:$false
#	}

		$UserGN = $ReenableUserobjUser.givenName
		$UserEmpID = $ReenableUserobjUser.employeeID
		$strTMPsn = $ReenableUserobjUser.sn
		$OpCode = 106
		
		if ($ReenableUserobjUser.extensionattribute5 -eq "S" -or $ReenableUserobjUser.extensionattribute5 -eq "T") {
			DebugLog " --- Account disabled, reenabling"
			RestoreGroups -UserDN $UDN -EmplID $UserEmpID
		}

		$CurDate = Get-Date
		
		$tmpCASMailbox = Get-CASMailbox $UDN
		$tmpMailbox = Get-Mailbox $UDN
		if ($tmpCASMailbox.OwaEnabled -eq $false) { 
			DebugLog " ---- Enabling access to OWA"
			Set-CASMailbox -Identity $UDN -OWAEnabled:$True
		}
		if ($tmpCASMailbox.MAPIEnabled -eq $false) { 
			DebugLog " ---- Enabling access to MAPI"
			Set-CASMailbox -Identity $UDN -MAPIEnabled:$True
		}
		if ($tmpCASMailbox.ActiveSyncEnabled -eq $false) { 
			DebugLog " ---- Enabling access to ActiveSync"
			Set-CASMailbox -Identity $UDN -ActiveSyncEnabled:$True 
		}
		if ($tmpMailbox.UMEnabled) {
			DebugLog " ---- Voicemail configured, enabling access"
			$tmpUMPolicy = (Get-UMMailbox $UDN).UMMailboxPolicy
			if ($tmpUMPolicy -like "*-SUSPENDED") {
				$tmpUMPolicy = $tmpUMPolicy.Replace("-SUSPENDED","")
				Set-UMMailbox -Identity $UDN -UMMailboxPolicy $tmpUMPolicy
			}
		}
		if ($tmpMailbox.HiddenFromAddressListsEnabled -eq $true -and $tmpmailbox.CustomAttribute7 -ne 'dualMailbox') {
			DebugLog " ---- Un-hiding from address lists"
			Set-Mailbox -Identity $UDN -HiddenFromAddressListsEnabled $false 
		}
		
		if ($tmpMailbox.Mailtip) {
			DebugLog " ---- Removing MailTip"
			Set-Mailbox -Identity $UDN -MailTip ""
		}
		
		if ((get-csuser -identity $UDN).Enabled -eq $false){
			DebugLog " ---- Enabling access to Lync"
			set-CSUser -identity $UDN -Enabled $true -Confirm:$false
		}
		
		if ($ReenableUserobjUser.LogonWorkstations) {
			DebugLog " ---- Clearing LogonWorkstations"
			if ($UDN.contains("DC=cequel,DC=cequel3,DC=com")) {
				Set-ADUser -Identity $UDN -replace @{extensionAttribute5 = "A"} -Server "cequel.cequel3.com" 
				Set-ADUser -Identity $UDN -LogonWorkstations $null -Server "cequel.cequel3.com" 
			} else {
				Set-ADUser -Identity $UDN -LogonWorkstations $null
			}
		}
		
		if ($ReenableUserobjUser.extensionattribute5 -ne "A") {
			DebugLog " ---- Setting extensionAttribute5"
					Set-ADUser -Identity $UDN -replace @{extensionAttribute5 = "A"}
		}
				
		$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$UserGN','$strTMPsn','$CurDate','$UserEmpID','$Region')"
		AccountProvisioningDBEx $SQLString
	}
}

######################################################################
#Disable Account*
# Returns False if account is already disabled
# Returns True if account was disabled successfully
# TESTED
######################################################################
Function DisableUser {
	param (
		[string]$UDN,
		[string]$Region,
		[boolean]$isSuspended
	)
	
	$result = $false

	if ($UDN.contains("DC=cequel,DC=cequel3,DC=com")) {
		$DisableUserobjUser = Get-ADUser $UDN -Server (Get-DomainController -DomainName "cequel.cequel3.com")[0].DnsHostName -Properties employeeID,sn,description,employeeType,manager,extensionAttribute5
	} else {		
		$DisableUserobjUser = Get-ADUser $UDN -Properties employeeID,sn,description,employeeType,manager,extensionAttribute5
	}

	$UserGN = $DisableUserobjUser.givenName
	$UserEmpID = $DisableUserobjUser.employeeID
	$strTMPsn = $DisableUserobjUser.sn

	if ($isSuspended) {
		if ($DisableUserobjUser.extensionAttribute5 -eq 'A') {
			DebugLog " ---- Account enabled, disabling"
			SaveGroups -UserDN $UDN -EmplID $UserEmpID
			set-aduser -identity $DisableUserobjUser -Replace @{extensionAttribute5 = "S"} -LogonWorkstations $AllowedLogonList
			DebugLog " ---- Setting MailTip"
			if ($DisableUserobjUser.Manager) {
				$tmpDisableUserObjManager = Get-ADUser $($DisableUserObjUser.manager)
				set-mailbox -identity $UDN -MailTip "This employee is out of the office until further notice.  Please contact $($tmpDisableUserObjManager.GivenName) $($tmpDisableUserObjManager.SurName) for immediate assistance."
			} else {
				set-mailbox -identity $UDN -MailTip "This employee is out of the office until further notice.  Please contact their supervisor for immediate assistance."
			}
			DisableEmailAccess -UDN $UDN
			$OpCode = 102
			$result = $True
			$CurDate = Get-Date
			$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$UserGN','$strTMPsn','$CurDate','$UserEmpID','$Region')"
			AccountProvisioningDBEx $SQLString
		}
	} else {
		if ($DisableUserobjUser.extensionAttribute5 -eq 'A') {
			DebugLog " ---- Account enabled, disabling"
			if ($isSuspended -eq $false) { #UpdateUser is called for Suspension, but not for Termination, so must set this here.
				if ($DisableUserobjUser.description -notlike "TERM - ") {
					$tmpDescription = $DisableUserobjUser.description
					$tmpDescription = "TERM - " + $tmpDescription
					Set-ADUser $UDN -Description $tmpDescription
				}
			}
			DebugLog " ---- Setting MailTip"
			if ($DisableUserobjUser.Manager) {
				$tmpDisableUserObjManager = Get-ADUser $($DisableUserObjUser.manager)
				set-mailbox -identity $UDN -HiddenFromAddressListsEnabled $true -MailTip "The employee you are trying to reach is no longer with the company, please contact $($tmpDisableUserObjManager.GivenName) $($tmpDisableUserObjManager.SurName) for assistance."
			} else {
				set-mailbox -identity $UDN -HiddenFromAddressListsEnabled $true -MailTip "The employee you are trying to reach is no longer with the company, please contact their supervisor for assistance."
			}
			SaveGroups -UserDN $UDN -EmplID $UserEmpID
			Set-ADUser -identity $UDN -Replace @{extensionAttribute5 = "T"}
            Set-ADUser -identity $UDN -LogonWorkstations $AllowedLogonList
			set-CSUser -identity $UDN -Enabled $false -Confirm:$false
			
			DisableEmailAccess -UDN $UDN
			
			$OpCode = 102
			$result = $True
			$CurDate = Get-Date
			$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$UserGN','$strTMPsn','$CurDate','$UserEmpID','$Region')"
			AccountProvisioningDBEx $SQLString
		} elseif ($DisableUserobjUser.extensionAttribute5 -eq 'S') {
			DebugLog " ----- Account was previously suspended, terminating"
			Set-ADUser -identity $UDN -Replace @{extensionAttribute5 = "T"} 
            Set-ADUser -identity $UDN -LogonWorkstations $AllowedLogonList
			$tmpDescription = $DisableUserobjUser.description
			$tmpDescription = $tmpDescription.Replace("SUSPENDED - ","TERM - ")
			Set-ADUser $UDN -Description $tmpDescription

			DisableEmailAccess -UDN $UDN

			$OpCode = 102
			$result = $True
			$CurDate = Get-Date
			$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$UserGN','$strTMPsn','$CurDate','$UserEmpID','$Region')"
			AccountProvisioningDBEx $SQLString
		}

		if ((get-aduser -identity $UDN -properties extensionAttribute5).extensionAttribute5 -ne "T") {
			set-aduser -identity $UDN -Replace @{extensionAttribute5 = "T"}
		}
	}

	if ((get-csuser -identity $UDN).Enabled) {
		DebugLog " ---- Removing Lync access"
		set-CSUser -identity $UDN -Enabled $false -Confirm:$false
	}
	
	if (((Get-ADUser -identity $UDN -properties LogonWorkstations).LogonWorkstations) -and (Get-ADUser -identity $UDN -properties LogonWorkstations).LogonWorkstations.trim() -ne $AllowedLogonList) {
		DebugLog " ---- Setting LogonWorkstations"
		Set-ADUser -identity $UDN -LogonWorkstations $AllowedLogonList
	}
	$DisableUserobjUser = $null
	return $result
}
######################################################################
#Disable email access*
######################################################################
function DisableEmailAccess{
	Param (
		[string] $UDN
	)
	
	if ((get-user -identity $UDN).RecipientType -eq 'UserMailbox') {
		DebugLog " ---- Removing OWA, MAPI, and ActiveSync access"
		$tmpCASMailbox = Get-CASMailbox $UDN
		if ($tmpCASMailbox.OwaEnabled) { Set-CASMailbox -Identity $UDN -OWAEnabled:$False }
		if ($tmpCASMailbox.MAPIEnabled) { Set-CASMailbox -Identity $UDN -MAPIEnabled:$False}
		if ($tmpCASMailbox.ActiveSyncEnabled) { Set-CASMailbox -Identity $UDN -ActiveSyncEnabled:$False }
		if ((Get-Mailbox $UDN).UMEnabled) {
			DebugLog " ---- Voicemail configured, removing access"
			$tmpUMPolicy = (Get-UMMailbox $UDN).UMMailboxPolicy
			if ($tmpUMPolicy -notlike "*-SUSPENDED") {
				$tmpUMPolicy = $tmpUMPolicy + "-SUSPENDED"
				Set-UMMailbox -Identity $UDN -UMMailboxPolicy $tmpUMPolicy
			}
		}
	}
}


######################################################################
#Create Account*
######################################################################
Function CreateAccount {
	Param (
		[string]$First,
		[string]$Last,
		[string]$Region,
		[string]$EmpID
	)
			
	$TMPFirst = $First
	$TMPFirst = $TMPFirst.replace("'",$null)
    $TMPFirst = $TMPFirst.replace("’",$null)
    $TMPFirst = $TMPFirst.replace(".",$null)
    $TMPFirst = $TMPFirst.replace(" ",$null)
    $TMPFirst = $TMPFirst.replace(",",$null)
	
	$TMPLast = $Last
	$TMPLast = $TMPLast.replace("'",$null)
    $TMPLast = $TMPLast.replace("’",$null)
    $TMPLast = $TMPLast.replace(".",$null)
    $TMPLast = $TMPLast.replace(" ",$null)
    $TMPLast = $TMPLast.replace(",",$null)
	
	$AccountName = $TMPFirst + "." + $TMPLast
	$AccountNameINC = 1
	$dupAccountName = $TMPFirst + "." + $TMPLast

	#crop to correct length for 20 char max
	if ($AccountName.length -gt 20) {$AccountName = $AccountName.substring(0,20)}

	#crop to correct length for 19 char max - for duplicate cases
	if ($dupAccountName.length -ge 20) {$dupAccountName = $dupAccountName.substring(0,19)}

	while (DuplicateAccount $AccountName -eq $false) {
		$AccountNameINC += 1
		$AccountName = $dupAccountName + $AccountNameINC
	}

	DebugLog " ---- Account to be created will be $AccountName"
	$CreateAccountobjParent = [ADSI]"LDAP://OU=Windows7,OU=General,OU=Accounts,OU=SDL-$Region,$domainDN"
	DebugLog " ---- OU will be OU=General,OU=Accounts,OU=SDL-$Region,$domainDN"
	$CreateAccountobjUser = $CreateAccountobjParent.Create("user", "CN=$AccountName")
	[void]$CreateAccountobjUser.Put("sAMAccountName", $AccountName)
	[void]$CreateAccountobjUser.SetInfo()
	[void]$CreateAccountobjUser.Put("employeeID", $EmpID)
	[void]$CreateAccountobjUser.Put("givenName", $First)
	[void]$CreateAccountobjUser.Put("sn", $Last)
	[void]$CreateAccountobjUser.Put("displayName", "$Last, $First")
	[void]$CreateAccountobjUser.Put("userPrincipalName", "$AccountName@suddenlink.com")
# 	[void]$CreateAccountobjUser.Put("mail", "$AccountName@suddenlink.com")
	[void]$CreateAccountobjUser.SetInfo()
	
	$script:CurPW = New-Password
	[void]$CreateAccountobjUser.SetPassword($script:CurPW.toString())
	[void]$CreateAccountobjUser.Put("userAccountControl", 512)
	[void]$CreateAccountobjUser.Put("pwdLastSet", 0)
	$CurDate = Get-Date
	[void]$CreateAccountobjUser.SetInfo()

	$UDN = $CreateAccountobjUser.Get("distinguishedName")
	$UserGN = $First
	$UserEmpID = $EmpID
	$strTMPsn = $Last
	$OpCode = 101
	
	$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$UserGN','$strTMPsn','$CurDate','$UserEmpID','$Region')"
	AccountProvisioningDBEx $SQLString

	DebugLog " ---- Creation Password: $script:CurPW"
# 	return $CreateAccountpassword
}

######################################################################
# Generate email address
######################################################################
Function GenerateEmailAddress {
	Param( $MailboxObj )
    
    $DuplicateDigit = 1
    $tmpEmailAddr = $MailboxObj.alias + $AUSEmailDomain

    DebugLog " ---- Checking SDL for $tmpEmailAddr"

    $tmpSDLSearchResult = (Get-Recipient $tmpEmailAddr -ErrorAction SilentlyContinue)

    while ($tmpSDLSearchResult) {
        $tmpEmailAddr = $MailboxObj.alias + $DuplicateDigit.ToString() + $AUSEmailDomain
        DebugLog "Checking SDL for $tmpEmailAddr"
        $tmpSDLSearchResult = (Get-Recipient $tmpEmailAddr -ErrorAction SilentlyContinue)
        $DuplicateDigit ++
    }

    DebugLog " ----- $tmpEmailAddr ok in SDL"
    DebugLog " ---- Checking CVC for $tmpEmailAddr"

    $CVCSearcherObj.Filter = "(proxyaddresses=*$tmpEmailAddr*)"
    $tmpSearchResults = $CVCSearcherObj.FindAll()
    $tmpResults = @()
    foreach ($tmpSearchResult in $tmpSearchResults){$tmpResults += $tmpSearchResult}

    while ($tmpResults.count -gt 0) {
        $tmpEmailAddr = $MailboxObj.alias + $DuplicateDigit.ToString() + $AUSEmailDomain
        DebugLog " ---- Checking CVC for $tmpEmailAddr"
        $DuplicateDigit ++

        $CVCSearcherObj.Filter = "(proxyaddresses=*$tmpEmailAddr*)"
        $tmpSearchResults = $CVCSearcherObj.FindAll()
        $tmpResults = @()
        foreach ($tmpSearchResult in $tmpSearchResults){$tmpResults += $tmpSearchResult}

    }

    DebugLog " ----- $tmpEmailAddr ok in CVC"

    return $tmpEmailAddr    
}

######################################################################
#EnableMailbox	
######################################################################
Function SetMailboxLevel {
	Param( [string]$SetMailboxLevelUser, [string]$SetMailboxLevel, [boolean]$bNewUser )
	
	if ($bNewUser) {
        DebugLog " --- Enabling mailbox"
		$tmpMailbox = Enable-Mailbox $SetMailboxLevelUser
        $AUSEmailAddress = GenerateEmailAddress $tmpMailbox

        $tmpStartTime = Get-Date
		do {
            $tmpMailboxCheck = $false
			$tmpCurTime = Get-Date
			$tmpMailboxCheck = get-mailbox -identity $SetMailboxLevelUser -errorAction SilentlyContinue
		} until ($tmpMailboxCheck -ne $false)
		$tmpWaitSpan = New-TimeSpan $tmpStartTime $tmpCurTime
		DebugLog " --- took $tmpWaitSpan seconds to complete."
        
        DebugLog " --- Setting email address to $AUSEmailAddress"
        Set-Mailbox $SetMailboxLevelUser -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $AUSEmailAddress 
	} else {
	    $tmpMailbox = Get-Mailbox $SetMailboxLevelUser
    }

	$tmpWarningQuota = $tmpMailbox.IssueWarningQuota
	$tmpProhibitSendQuota = $tmpMailbox.ProhibitSendQuota
	$tmpManagedFolderMailboxPolicy = $tmpMailbox.ManagedFolderMailboxPolicy
	$tmpRetentionPolicy = $tmpMailbox.RetentionPolicy
	
	$bQuotaChanged = $False
	
	if ($tmpManagedFolderMailboxPolicy -eq $null -and $tmpRetentionPolicy -eq $null) {
		set-mailbox $SetMailboxLevelUser -RetentionPolicy SDL-Default-Delete-180Days -Confirm:$False
	}
	
	Set-CASMailbox $SetMailboxLevelUser -ImapEnabled $false
	
	switch ($SetMailboxLevel) {
		"User" {}
		"Manager" {}
		"Director" {
			if (($bNewUser) -or ($tmpWarningQuota -lt "1.8GB") -or ($tmpMailbox.useDatabaseQuotaDefaults)) {
				Set-Mailbox $SetMailboxLevelUser -IssueWarningQuota 1.8GB
				$bQuotaChanged = $true
			}
			if (($bNewUser) -or ($tmpProhibitSendQuota -lt "2GB") -or ($tmpMailbox.useDatabaseQuotaDefaults)) {
				Set-Mailbox $SetMailboxLevelUser -ProhibitSendQuota 2GB
				$bQuotaChanged = $True
			}
		}
		"VP" {
			if (($bNewUser) -or ($tmpWarningQuota -lt "1.8GB") -or ($tmpMailbox.useDatabaseQuotaDefaults)) {
				Set-Mailbox $SetMailboxLevelUser -IssueWarningQuota 1.8GB
				$bQuotaChanged = $True
			}
			if (($bNewUser) -or ($tmpProhibitSendQuota -lt "2GB") -or ($tmpMailbox.useDatabaseQuotaDefaults)) {
				Set-Mailbox $SetMailboxLevelUser -ProhibitSendQuota 2GB
				$bQuotaChanged = $True
			}
		}
	}
	if ($bQuotaChanged) { 
		Set-Mailbox	$SetMailboxLevelUser -ProhibitSendReceiveQuota unlimited -UseDatabaseQuotaDefaults $False
		DebugLog " ---- Mailbox Quota adjusted" 
	}
}

######################################################################
#Delete Account	
######################################################################
Function DeleteAccount {
	Param(
		[string]$UDN,
		[string]$Region
	)
	
	if ($UDN.contains("DC=cequel,DC=cequel3,DC=com")) {
		$DeleteAccountobjUser = Get-ADUser $UDN -Server "cequel.cequel3.com" -Properties sn,employeeID
	} else {
		$DeleteAccountobjUser = Get-ADUser $UDN -Properties sn,employeeID
	}

    $DeleteAccountTMPSN = $DeleteAccountobjUser.SN
    $DeleteAccountTMPGivenName = $DeleteAccountobjUser.givenName
    $DeleteAccountTMPEMPID = $DeleteAccountobjUser.EmployeeID

	$CurDate = Get-Date
	if ((get-user -identity $UDN).RecipientType -eq 'UserMailbox') {
		Remove-Mailbox -Identity $UDN -Confirm:$False > $null
	} else {
		$tmpADUserChildren = Get-AdObject -Filter * -SearchScope oneLevel -SearchBase $UDN
		if ($tmpADUserChildren) { $tmpADUserChildren | Remove-AdObject -Recursive -Confirm:$false }
		remove-aduser -identity $UDN -Confirm:$false
	}
	Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "DELETE FROM GroupTemp Where EmployeeID = '$DeleteAccountTMPEMPID'"

	$OpCode = 100
	
	$SQLString = "INSERT INTO OperationLog(DN,Operation,FirstName,LastName,OperationDate,EmployeeID,Region) Values('$UDN',$OpCode,'$DeleteAccountTMPGivenName','$DeleteAccountTMPSN','$CurDate','$DeleteAccountTMPEMPID','$Region')"
	AccountProvisioningDBEx $SQLString
}

####################################################################
# FUNCTION NAME: New-Password
# TESTED
# See USAGE() function for docs.
# WRITTEN BY: Derek Mangrum
# REVISION HISTORY:
#     2008-10-23 : Initial version
####################################################################
function New-Password {
    param (
        [int]$length = 9,
        [switch]$lowerCase = $true,
        [switch]$upperCase = $true,
        [switch]$numbers = $true,
        [switch]$specialChars = $true
    )

   $lCase = 'abcdefghijklmnopqrstuvwxyz'
   $uCase = $lCase.ToUpper()
   $nums = '1234567890'
   $specChars = '!@#$&*'
	
   if ($lowerCase) { 
       $charsToUse += $lCase
       $regexExp += "(?=.*[$lCase])"
   }
   if ($upperCase) { 
       $charsToUse += $uCase 
       $regexExp += "(?=.*[$uCase])"
   }
   if ($numbers) { 
       $charsToUse += $nums 
       $regexExp += "(?=.*[$nums])"
   }
   if ($specialChars) { 
       $charsToUse += $specChars
       $regexExp += "(?=.*[\W])"
   }
   
   $test = [regex]$regexExp
   $seed = ([system.Guid]::NewGuid().GetHashCode())
   $rnd = New-Object System.Random($seed)
   
   do {
       $pw = $null
       for ($i = 0 ; $i -lt $length ; $i++) {
           $pw += $charsToUse[($rnd.Next(0,$charsToUse.Length))]
       }
   } until ($pw -match $test)
   return [string]$pw
}

######################################################################
#Get Managment Level User, Manager, Director and Above*
######################################################################
Function GetManagementLevel {
	param ([string]$LevelTitle)
	if ($LevelTitle -like "*Manager*") {$result = "Manager"}
	elseif ($LevelTitle -like "*Mgr*") {$result = "Manager"}
	elseif ($LevelTitle -like "Supervisor*") {$result = "Manager"}
	elseif ($LevelTitle -like "Director*") {$result = "Director"}
	elseif ($LevelTitle -like "*VP*") {$result = "VP"}
	elseif ($LevelTitle -like "*Chairman*") {$result = "VP"}
	else {$result = "User"}

	DebugLog " -- Management Level: $result"
	return $result
}

######################################################################
#Get Region Name*
######################################################################
Function GetRegionName {
	Param ([string]$GetRegionNameDESC)

	$AccountProvisioningDBRFSQLStr = "SELECT * from PSLocationLookup where PeopleSoftLocationCode='$GetRegionNameDESC'"

	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection "Server = $AccountProvisioningSQLServer;Database=$AccountProvisioningSQLDatabase;Integrated Security=True"
  	[void]$sqlConnection.Open()

  	$sqlCommand = New-Object system.data.sqlclient.SqlCommand($AccountProvisioningDBRFSQLStr,$sqlConnection)
  	$sqlDataAdapter = new-object System.Data.SqlClient.SQLDataAdapter($sqlCommand) 
  	$LookupReader = $sqlCommand.ExecuteReader()

	while ($LookupReader.Read()) {
		if ($LookupReader.IsDBNull($LookupReader.GetOrdinal("ITLocationCode"))) {
			$ITSiteCode = "COR"
		} else {
			$ITSiteCode = $LookupReader.GetValue($LookupReader.GetOrdinal("ITLocationCode"))
  		}
   	}

  	[void]$sqlconnection.close() # close connection

	
	if ($ITSiteCode -eq "" -or $ITSiteCode -eq $Null) {
		$ITSiteCode = "COR"
	} else {
		$ITSiteCode = $ITSiteCode.Trim()	
	}
	
# 	DebugLog " -- Region: $ITSiteCode"
	return $ITSiteCode
}

######################################################################
#QueryEmail*
# Returns email address of UDN submitted as UDN
# TESTED
######################################################################
Function QueryEmail {
	Param ([string]$UDN)
    
    $QueryEmailRecipObj = Get-Recipient -Identity $UDN -ResultSize 10 -ErrorAction silentlycontinue
	if ($QueryEmailRecipObj -and $QueryEmailRecipObj.count -lt 2) {
		$result = $QueryEmailRecipObj.PrimarySMTPAddress
	} else {
		$QueryEmailRecipObj = Get-ADObject -Identity $UDN -Properties mail
        if ($QueryEmailRecipObj.mail) {
            $result = $QueryEmailRecipObj.mail
        } else {
            $result = $false
        }
	}
	return $result
}


####################################################################
#Debug Logging
# Writes debug information to Console and to a log file
####################################################################
Function DebugLog {
	Param ( [String] $LogText )
	
	Out-File -NoClobber:$true -Append:$True -FilePath:$script:LogFilePath -InputObject:$LogText
	Write-Host $LogText
}

######################################################################
#Save Group Membership
#  Reads current Group Membership and saves to the database
######################################################################
Function SaveGroups {
    Param ([string]$UserDN,
        [string]$Emplid    
    )
    $DUGroupRID = [int]513
    $DuoNothingGroupUID = (Get-ADGroup	-identity "AP_DUO_Nothing").ObjectGUID
    
    $UserObj = Get-SDLADUser -EmployeeNum $Emplid
    $CurPrimaryGroupID = $UserObj.primaryGroupID
    if ($CurPrimaryGroupID -eq $DUGroupRID) {
		DebugLog " --- Saving group membership"
        $GroupList = $UserObj.memberOf
        pushd
        foreach ($tmpGroup in $GroupList) {
            $GrpGUID = (Get-SDLADGroup $tmpGroup).ObjectGUID
            if ($GrpGUID -ne $DuoNothingGroupUID) {
	            if (-not $(Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "SELECT GroupUID FROM GroupTemp WHERE EmployeeID = '$Emplid' AND GroupUID = '$GrpGUID'")) {
	   	 	        Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "INSERT INTO GroupTemp Values('$Emplid','$GrpGUID')"
	            }
            }
        }
        popd
		pushd
		foreach ($tmpGroup in $GroupList) {
			$GrpGUID = (Get-SDLADGroup $tmpGroup).ObjectGUID
            if ($GrpGUID -ne $DuoNothingGroupUID) {
				if (-not $(Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "SELECT GroupUID FROM GroupTemp WHERE GroupUID = '$GrpGUID'")) {
					Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "INSERT INTO GroupTemp Values('$Emplid','$GrpGUID')"
					$bDBCheck = $false
				}
				Remove-SDLADGroupMember -GroupDN $tmpGroup -UserObj $UserObj
			}
		}
		popd
		Set-SDLDefaultADGroup -UserObj $UserObj -Restrict:$true
    }
}

######################################################################
#Restore Saved Group Membership
#  Reads saved Groups from the database and adds User as member again
######################################################################
Function RestoreGroups {
    Param([string]$UserDN,
        [string]$Emplid   
    )
    
    $RGEmplid = $Emplid
    DebugLog " --- Restoring Group membership"
    $UserObj = Get-SDLADUser -EmployeeNum $RGEmplid
    $DBGroups = Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "SELECT GroupUID FROM GroupTemp WHERE EmployeeID = '$RGEmplid'"
    $ExistingGroups = @()
    foreach ($tmpGroupGUID in $DBGroups) {
        $tmpGroup = Get-SDLADGroup $tmpGroupGUID.GroupUID
        if ($tmpGroup) {
            $ExistingGroups += $tmpGroupGUID
            Add-SDLADGroupMember -GroupDN $tmpGroup.distinguishedName -UserObj $UserObj
        }
    }

    Set-SDLDefaultADGroup -UserObj $UserObj -Restrict:$false

    $bGroupCheck = $false
    while ($bGroupCheck -ne $True) {
        $UserObj = Get-SDLADUser -EmployeeNum $RGEmplid
        $bGroupCheck = $true
        if ($ExistingGroups.Count -ne $UserObj.memberOf.Count -and $ExistingGroups.Count -gt 0) {
            foreach ($tmpGroupGUID in $ExistingGroups) {
                $tmpGroup = Get-SDLADGroup $tmpGroupGUID.GroupUID
                if ($UserObj.memberOf -notcontains $tmpGroup.DistinguishedName) {
                    $bGroupCheck = $false
                    write-host $tmpGroup.Name
                    Add-SDLADGroupMember -GroupDN $tmpGroup -UserObj $UserObj
                }
            }
        }
#			else {
#             $bGroupCheck = $true
#         }
    }
    pushd
    Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "DELETE FROM GroupTemp Where EmployeeID = '$RGEmplid'"
    popd
}

######################################################################
#Get AD User 
#  Custom function to get AD User object from the employeeID,  
#   accomodates the multiple domains at Suddenlink
######################################################################
Function Get-SDLADUser {
	Param ([string]$EmployeeNum)
	DebugLog " -- Getting AD Object for $EmployeeNum"
	$Result = Get-ADUser -Filter {employeeId -eq $EmployeeNum} -Properties memberOf,primaryGroupID,LogonWorkstations,employeeType,mobile,telephoneNumber,extensionAttribute5
	if ($Result -eq $null) { 
		$Result = Get-ADUser -Server 'cequel.cequel3.com' -Filter {Employeeid -eq $EmployeeNum} -Properties memberOf,primaryGroupID,LogonWorkstations,employeeType,mobile,telephoneNumber,extensionAttribute5
			if ($Result -eq $null) { 
				$Result = $False 
			}
	}
	return $Result
}

######################################################################
#Get AD Group 
#  Custom function to get AD group object from the Group GUID,  
#   accomodates the multiple domains at Suddenlink
######################################################################
Function Get-SDLADGroup {
	Param ([string]$GroupUID)
	DebugLog " -- Getting AD Object for $GroupUID"
    $error.clear()
    try {
        $Result = Get-ADGroup -Identity $GroupUID
    } catch {
        $error.clear()
        try {
            $Result = Get-ADGroup -Identity $GroupUID -Server cequel.cequel3.com
        } catch {
            $error.clear()
            $Result = Get-ADGroupMember -Identity $GroupUID -Server cequel3.com
        }
    }
	return $Result
}

######################################################################
#Remove Group Member
#  Custom function to remove group members, accomodates the multiple 
#   domains at Suddenlink
######################################################################

Function Remove-SDLADGroupMember {
    Param([string]$GroupDN,
        [System.Object]$UserObj
    )

    if ($GroupDN -like '*,DC=cequel,DC=cequel3,DC=com') {
        Remove-ADGroupMember -Identity $GroupDN -Members $UserObj -Server cequel.cequel3.com -confirm:$false
    } elseif ($GroupDN -like '*,DC=suddenlink,DC=cequel3,DC=com') {
        Remove-ADGroupMember -Identity $GroupDN -Members $UserObj -confirm:$false
    }else {
        Remove-ADGroupMember -Identity $GroupDN -Members $UserObj -Server cequel3.com -confirm:$false
    }
}

######################################################################
#Add Group Member
#  Custom function to add group members, accomodates the multiple 
#   domains at Suddenlink
######################################################################
Function Add-SDLADGroupMember {
    Param([string]$GroupDN,
        [System.Object]$UserObj
    )

    if ($GroupDN -like '*,DC=cequel,DC=cequel3,DC=com') {
        Add-ADGroupMember -Identity $GroupDN -Members $UserObj -Server cequel.cequel3.com
    } elseif ($GroupDN -like '*,DC=suddenlink,DC=cequel3,DC=com') {
        Add-ADGroupMember -Identity $GroupDN -Members $UserObj
    } else {
        Add-ADGroupMember -Identity $GroupDN -Members $UserObj -Server cequel3.com
    }

}

######################################################################
#Set Default Group
#  Sets the Default AD Group for the indicated user.
#  Sets to AP_Provisioning_Nothing when $Restrict = $true
#  Sets to Domain Users when $Restrict = $false
######################################################################
Function Set-SDLDefaultADGroup {
    Param([System.Object]$UserObj,
        [boolean]$Restrict
    )

    $DUGroupRID = [int]513

    if ($Restrict) {
    	DebugLog " ---- Setting default group to restricted group"
        if ($UserObj.distinguishedName -like '*,DC=cequel,DC=cequel3,DC=com') {
            $NothingGroup = Get-ADGroup -Identity "AP_Provisioning_Nothing_CQL" -server cequel.cequel3.com
            Add-SDLADGroupMember -GroupDN $NothingGroup.distinguishedName -UserObj $UserObj

            $RestrictedGroupRID = [int]$NothingGroup.SID.toString().Split('-')[7]
            $GroupRID = $RestrictedGroupRID

            $UserObj | Set-ADObject -Replace @{primaryGroupID = "$GroupRID"} -Server cequel.cequel3.com
            Remove-ADGroupMember -Identity 'Domain Users' -Members $UserObj -Server cequel.cequel3.com -confirm:$false
        } elseif ($UserObj.DistinguishedName -like '*,DC=suddenlink,DC=cequel3,DC=com') {
            $NothingGroup = Get-ADGroup -Identity "AP_Provisioning_Nothing"
            Add-SDLADGroupMember -GroupDN $NothingGroup.distinguishedName -UserObj $UserObj
            $RestrictedGroupRID = [int]$NothingGroup.SID.toString().Split('-')[7]
            $GroupRID = $RestrictedGroupRID

            $UserObj | Set-ADObject -Replace @{primaryGroupID = "$GroupRID"}
            Remove-ADGroupMember -Identity 'Domain Users' -Members $UserObj -confirm:$false
        }

    } else {
    	DebugLog " ---- Setting default group to Domain Users"
        $GroupRID = $DUGroupRID
        if ($UserObj.DistinguishedName -like '*,DC=cequel,DC=cequel3,DC=com') {
            $NothingGroup = Get-ADGroup -Identity "AP_Provisioning_Nothing_CQL" -server cequel.cequel3.com
            Add-ADGroupMember -Identity (Get-ADGroup -Identity 'Domain Users' -server cequel.cequel3.com) -Members $UserObj -Server cequel.cequel3.com -confirm:$false
            $UserObj | Set-ADObject -Replace @{primaryGroupID = "$GroupRID"} -Server cequel.cequel3.com
        } elseif ($UserObj.DistinguishedName -like '*,DC=suddenlink,DC=cequel3,DC=com') {
            $NothingGroup = Get-ADGroup -Identity "AP_Provisioning_Nothing"
            Add-ADGroupMember -Identity (Get-ADGroup -Identity 'Domain Users') -Members $UserObj -confirm:$false
            $UserObj | Set-ADObject -Replace @{primaryGroupID = "$GroupRID"}
        }

        Remove-SDLADGroupMember -GroupDN $NothingGroup.distinguishedName -UserObj $UserObj
    }
}

######################################################################
#Get Account Flags*
######################################################################
Function GetAccountFlags {
	Param ([string]$GetAccountFlagsEmplID)

	$GetAccountFlagsDBRFSQLStr = "SELECT * from AccountFlags where EmployeeID='$GetAccountFlagsEmplID'"

    $QueryResults = Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query $GetAccountFlagsDBRFSQLStr
	
	return $QueryResults
}

######################################################################
#ServiceNowTermEmail
######################################################################
Function ServiceNowTermEmail {
	Param (
		[string]$ServiceNowTermEmailManagerID,
		[string]$ServiceNowTermEmailSN,
		[string]$ServiceNowTermEmailGiven,
		[string]$ServiceNowTermEmailEMPID,
		[string]$ServiceNowTermEmailTitle,
		[string]$ServiceNowTermEmailLocationID,
		[string]$ServiceNowTermEmailLocationName,
		[string]$ServiceNowTermEmailBUSUnit,
		[string]$ServiceNowTermEmailEffDate,
		[string]$ServiceNowTermEmailEmplType,
		[string]$ServiceNowTermEmailVendorName,
		[string]$ServiceNowTermEmailDeskNumber,
		[string]$ServiceNowTermEmailDeptID,
		[string]$ServiceNowTermEmailDepartment,
		[string]$ServiceNowTermEmailCompanyCode,
		[string]$ServiceNowTermEmailCompany
	)

    $tmpManagerObj = get-aduser -Filter {employeeID -eq $ServiceNowTermEmailManagerID}
	DebugLog " --- Creating ServiceNow ticket email"
	DebugLog " ---- Message body follows ----"
	
	$ServiceNowTermEmailobjEmail = New-Object system.net.mail.smtpClient
	$ServiceNowTermEmailobjEmail.host = "smtp.suddenlink.com"
	$ADProvisionAddr = New-Object system.Net.Mail.MailAddress($EmailNotificationFROMaddr)
	if ($script:bDebugging -eq $false) {
		$ADProvisionDLAddr = New-Object system.Net.Mail.MailAddress($CVCServiceNowEmailAddr)
	} else {
		$ADProvisionDLAddr = $ArrTo = $EmailNotificationDebuggingTOaddr
	}
	
	$ServiceNowTermEmailobjMsg = New-Object system.Net.Mail.MailMessage($ADProvisionAddr, $ADProvisionDLAddr)

	$ServiceNowTermEmailEffDate = get-date($ServiceNowTermEmailEffDate) -format MM/dd/yyyy

    $ServiceNowTermEmailEMPID = "S" + $ServiceNowTermEmailEMPID

	$ServiceNowTermEmailobjMsg.Subject = "Employee Separation Request"
	$ServiceNowTermEmailobjMsg.Body = "request_type: Employee Separation Request"
    $ServiceNowTermEmailobjMsg.Body += "`nu_resource:$ServiceNowTermEmailEMPID;SDL"
	$ServiceNowTermEmailobjMsg.Body += "`ncomments:"
	$ServiceNowTermEmailobjMsg.Body += "`nLast Day Worked - $ServiceNowTermEmailEffDate"
	$ServiceNowTermEmailobjMsg.Body += "`nTermination Date - $ServiceNowTermEmailEffDate"
	$ServiceNowTermEmailobjMsg.Body += "`nEmployee ID - $ServiceNowTermEmailEMPID"
	$ServiceNowTermEmailobjMsg.Body += "`nEmployee Name - $ServiceNowTermEmailGiven $ServiceNowTermEmailSN"
	$ServiceNowTermEmailobjMsg.Body += "`nJob Title - $ServiceNowTermEmailTitle"
	$ServiceNowTermEmailobjMsg.Body += "`nWork Phone - $ServiceNowTermEmailDeskNumber"
	$ServiceNowTermEmailobjMsg.Body += "`nBusiness Unit - $ServiceNowTermEmailBUSUnit"
	$ServiceNowTermEmailobjMsg.Body += "`nCompany Code - $ServiceNowTermEmailCompanyCode"
    $ServiceNowTermEmailobjMsg.Body += "`nCompany Description - $ServiceNowTermEmailCompany"
	$ServiceNowTermEmailobjMsg.Body += "`nDepartment ID - $ServiceNowTermEmailDeptID"
	$ServiceNowTermEmailobjMsg.Body += "`nDepartment Description - $ServiceNowTermEmailDepartment"
	$ServiceNowTermEmailobjMsg.Body += "`nLocation Code - $ServiceNowTermEmailLocationID"
	$ServiceNowTermEmailobjMsg.Body += "`nLocation Description - $ServiceNowTermEmailLocationName"
	$ServiceNowTermEmailobjMsg.Body += "`nReports To Name - $($tmpManagerObj.GivenName) $($tmpManagerObj.surname)"
	
	DebugLog $ServiceNowTermEmailobjMsg.Body
	DebugLog " ---- End Message body ----"

	[void]$ServiceNowTermEmailobjEmail.Send($ServiceNowTermEmailobjMsg)
}

####################################################################
# Main
####################################################################
function Main {
	
	if ($script:bDebugging -eq $false) {
		#------------------------------------
		#	Connection to PRODUCTION ORACLE
		#====================================
		DebugLog "Connecting to Production database..."
		[void][System.Reflection.Assembly]::LoadwithPartialName("System.Data.OracleClient")
		$OracleConnectionString = "User Id=$PeopleSoftOracleUser;Password=$PeopleSoftOraclePassword;Data Source=" + $PeopleSoftOracleServer + ":1521/$PeopleSoftOracleInstance"
	# 	DebugLog $OracleConnectionString
		if ($FullSync) {
			$OracleCommandText = "Select * FROM " + $PeopleSoftOracleEmployeeView + " where (Company='CLA'or Company='EXC')"#or Company = 'CEQ' )"
        } elseif ($DailySync) {
			$OracleCommandText = "Select * FROM " + $PeopleSoftOracleEmployeeView + " where (Company='CLA'or Company='EXC') AND (EFFDT > CURRENT_DATE - interval '95' day) order by EFFDT desc"
		} elseif ($SingleSync) {
			$OracleCommandText = "Select * FROM " + $PeopleSoftOracleEmployeeView + " where EMPLID = '$SingleEmplID'"
		} else {
			$OracleCommandText = "Select * FROM " + $PeopleSoftOracleEmployeeView + " where (Company='CLA' or Company='EXC') AND (EFFDT > CURRENT_DATE - interval '8' day) order by EFFDT desc"
		}
		#$OracleCommandText = "Select * FROM " + $PeopleSoftOracleEmployeeView + " where (EMPLID = '010733')"
		
		$PeopleOracleCN = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
		[void]$PeopleOracleCN.Open()
		DebugLog "Oracle connection open"
		$PeopleOracleCMD = New-Object System.Data.OracleClient.OracleCommand($OracleCommandText, $PeopleOracleCN)
		$PSReader = $PeopleOracleCMD.ExecuteReader()
		#------------------------------------
		
	} else {
		#---------------------------------------------------------------
		# 	Connection to AccountProvisioning SQL database for test data
		#===============================================================
		#Change the data in the TestData table as needed
		DebugLog "Connecting to SQL database for test data..."
		$OracleConnectionString = "Provider=sqloledb;Data Source='$AccountProvisioningSQLServer';Initial Catalog=$AccountProvisioningSQLDatabase;Integrated Security=SSPI;"
		$OracleCommandText='select * from dbo.TestData'
		if ($SingleSync) {
			$OracleCommandText = "select * from dbo.TestData where EMPLID = '$SingleEmplID'"
		}
		$PeopleOracleCN = New-Object System.Data.OleDb.OleDbConnection($OracleConnectionString)
		[void]$PeopleOracleCN.open()
		$PeopleOracleCMD = new-object System.Data.OleDb.OleDbCommand($OracleCommandText,$PeopleOracleCN) 
		$PSReader = $PeopleOracleCMD.ExecuteReader() 
		#------------------------------------
	}
		
	DebugLog "Getting data from Oracle"
	while ($PSReader.read()) {
		$CurEMPLID = [string]$PSReader.GetValue($PSReader.GetOrdinal("EMPLID"))
		$CurEmplAction = $PSReader.GetString($PSReader.GetOrdinal("XLATLONGNAME"))
		$CurEmplStatus = $PSReader.GetString($PSReader.GetOrdinal("EMPL_STATUS"))
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("EXEMPT_JOB_LBR"))) {
			$CurEmplStatusException = $PSReader.GetString($PSReader.GetOrdinal("EXEMPT_JOB_LBR"))
		} else {$CurEmplStatusException = "N"}
		#Field 4 (ACTION_DT) is not used
		$CurEffectiveDate = $PSReader.GetDateTime($PSReader.GetOrdinal("EFFDT"))
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("PREFERRED_NAME"))) {
			$CurPreferredName = $PSReader.GetString($PSReader.GetOrdinal("PREFERRED_NAME"))
		} else {$CurPreferredName = $false}
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("CQ_PREF_LAST_NAME"))) {
			$CurPreferredLastName = $PSReader.GetString($PSReader.GetOrdinal("CQ_PREF_LAST_NAME"))
		} else {$CurPreferredLastName = $false}
		$CurFirstName = $PSReader.GetString($PSReader.GetOrdinal("FIRST_NAME"))
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("MIDDLE_INITIAL"))) {
			$CurMiddleInitial = $PSReader.GetString($PSReader.GetOrdinal("MIDDLE_INITIAL"))
		} else {$CurMiddleInitial = ""}
		$CurLastName = $PSReader.GetString($PSReader.GetOrdinal("LAST_NAME"))
		#Field 10 (JOBCODE) is not used
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("DESCR"))) {
			$CurJobTitle = $PSReader.GetString($PSReader.GetOrdinal("DESCR"))
		} else {$CurJobTitle = ""}

		$CurLocation = $PSReader.GetString($PSReader.GetOrdinal("LOCATION"))
		$CurITRegion = GetRegionName $CurLocation
		$CurAddress1 = $PSReader.GetString($PSReader.GetOrdinal("ADDRESS1"))
		$CurAddress2 = $PSReader.GetString($PSReader.GetOrdinal("ADDRESS2"))
		$CurCity = $PSReader.GetString($PSReader.GetOrdinal("CITY"))
		$CurState = $PSReader.GetString($PSReader.GetOrdinal("STATE"))
		$CurZipCode = [string]$PSReader.GetValue($PSReader.GetOrdinal("POSTAL"))
	 	#Field 17 (PHONE) is not used in the old script
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("DEPTID"))) {
			$CurDeptID = $PSReader.GetString($PSReader.GetOrdinal("DEPTID"))
		} else {$CurDeptID = ""}
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("DESCR1"))) {
			$CurDepartment = $PSReader.GetString($PSReader.GetOrdinal("DESCR1"))
		} else {$CurDepartment = ""}
		$CurBusinessUnit = [string]$PSReader.GetValue($PSReader.GetOrdinal("BUSINESS_UNIT"))
		#Field 21 (DESCR2) is not used in the old script
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("SUPERVISOR_ID"))) {
			$CurManagerID = $PSReader.GetString($PSReader.GetOrdinal("SUPERVISOR_ID"))
		} else {$CurManagerID = ""}
		$CurCompany = $PSReader.GetString($PSReader.GetOrdinal("COMPANY"))
		if ($CurCompany -eq "CEQ") {
			$CurCompany = "AlticeUSA" 
		} elseif ($CurCompany -eq "CLA") {
			$CurCompany = "Suddenlink"
		} elseif ($CurCompany -eq "TPT") {
			$CurCompany = "Cablevision"
		} else {
			$CurCompany = "Suddenlink"
		}
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("CQ_REGIONS"))) {
			$CurPSRegion = $PSReader.GetString($PSReader.GetOrdinal("CQ_REGIONS"))
		} else {$CurPSRegion = ""}
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("CQ_CONFERENCE_CARD"))) {
			$CurConferenceCard = $PSReader.GetString($PSReader.GetOrdinal("CQ_CONFERENCE_CARD"))
		} else {$CurConferenceCard = ""}
		if (!$PSReader.isDBNull($PSReader.GetOrdinal("CQ_PHON_XTF_STATUS"))) {
			$CurPhoneNumber = $PSReader.GetString($PSReader.GetOrdinal("CQ_PHON_XTF_STATUS"))
		} else {$CurPhoneNumber = ""}
		$CurSalesID = [string]$PSReader.Getvalue($PSReader.GetOrdinal("ALTER_EMPLID"))
		$CurEmplType = [string]$PSReader.Getvalue($PSReader.GetOrdinal("PER_ORG"))
		#Field (VENDOR_ID) is not used
		if ($CurEmplType -eq "CWR" -and !$PSReader.isDBNull($PSReader.GetOrdinal("CQ_VENDOR_NAME"))) {
			$CurVendorName = [string]$PSReader.Getvalue($PSReader.GetOrdinal("CQ_VENDOR_NAME"))
		} else {
			$CurVendorName = ""
		}
		$CurICOMSID = [string]$PSReader.Getvalue($PSReader.GetOrdinal("CQ_ICOMS_ID"))
		$CurSupervisorSalesID = [string]$PSReader.Getvalue($PSReader.GetOrdinal("CQ_SUPV_SALES_ID"))

        if ($PSReader.GetString($PSReader.GetOrdinal("CQ_AD_ACCT_REQ_FLG")) -eq 'N') {
            $isADAccountReqd = $false
        } else {
            $isADAccountReqd = $true
        }

		$TmpCityState = "$CurCity, $CurState"
		
		$isConsultant = $False
		if ($CurEmplType -eq "EMP" -or $CurEmplType -eq "POI") { $CurEmplType = "Employee" }
		elseif ($CurEmplType -eq "CWR") { $isConsultant = $True; $CurEmplType = "Contractor" }
		#elseif ($CurEmplType -eq "POI") { $CurJobTitle = "$CurCompany User" }
		
		$TmpTime = (Get-Date -Format hh:mm:ss)
		DebugLog ""
		DebugLog $TmpTime
		DebugLog "Working on $CurEMPLID : $CurLastName, $CurFirstName in $CurITRegion"	
		
		$CurManagerDN = $Null
		$CurManagerDN = QueryAccount $CurManagerID
		$CurEmplDN = QueryAccount $CurEMPLID
        if ($CurEmplDN -like 'CN=o365-*') {$CurEmplDN = $false} #If a Contact is found for the employeeID of the current record, treat as nothing found.
		$CurManagerEmail = $null
		
		If (($CurManagerDN -ne $False) -and ($CurManagerDN -ne $CurEmplDN)) {
		    $CurManagerEmail = QueryEmail $CurManagerDN
		    DebugLog " -- Manager Email Address: $CurManagerEmail"
		}
		
		$isSuspended = $False
		if ($script:bDebugging) { DebugLog "CurEmplStatus = '$CurEmplStatus'"}
		
		if ($CurPreferredName -and ($CurPreferredName -ne $CurFirstName)) {$TMPFirstName = $CurPreferredName} 
		else {$TMPFirstName = $CurFirstName}
 		$TMPFirstName = $TMPFirstName.replace("'",$null)
        $TMPFirstName = $TMPFirstName.replace("’",$null)
#        $TMPFirstName = $TMPFirstName.replace(".",$null)
# 		 $TMPFirstName = $TMPFirstName.replace(" ",$null)

		if ($CurPreferredLastName -and ($CurPreferredLastName -ne $CurLastName)) {$TMPLastName = $CurPreferredLastName}
		else {$TMPLastName = $CurLastName}
  		$TMPLastName = $TMPLastName.replace("'",$null)
        $TMPLastName = $TMPLastName.replace("’",$null)
#         $TMPLastName = $TMPLastName.replace(".",$null)
#         $TMPLastName = $TMPLastName.replace(" ",$null)

    #PAR indicates person should have an AD account
        if ($isADAccountReqd) {
            if ($CurEmplStatus -eq "L" -and $CurEmplStatusException -eq "Y") {  # Set Employment Status to Active if LOA with Exception.
                $CurEmplStatus = "A"
                DebugLog "- Employee is on LOA with exception, treating as Active."
            } 

		    if ($CurEmplStatus -eq "S" -or ($CurEmplStatus -eq "L" -and $CurEmplStatusException -eq "N")) {
     #
     # Employee is SUSPENDED
     #
			    DebugLog " -- Suspended in PS, Effective Date $CurEffectiveDate"
			    $isSuspended = $True
			    if ($CurEmplDN -ne $False) {
				    DebugLog " --- Account exists"

		 		    $ignore = DisableUser $CurEmplDN $CurITRegion $isSuspended
		 		
		 		    $CurSAMaccount = QuerySAMAccount($CurEmplDN)
				    FindExtraAccounts $CurSAMaccount
		 		    DisableExtraAccounts $CurSAMaccount $CurITRegion $isSuspended
		 		
				    DebugLog " --- Updating account with latest data from PS"
				    UpdateAccount $CurEmplDN $CurEMPLID $TMPFirstName $TMPLastName $CurMiddleInitial $CurJobTitle $CurLocation $CurAddress1 $CurAddress2 $CurCity $CurState $CurZipCode $CurManagerDN $CurBusinessUnit $CurDepartment $isConsultant $CurVendorName $isSuspended $CurCompany $CurITRegion $CurPSRegion $CurDeptID $CurICOMSID
			    } else {DebugLog " --- no account - THIS SHOULDN'T HAPPEN"}
		    } elseif (($CurEmplStatus -eq "R") -or ($CurEmplStatus -eq "Q") -or ($CurEmplStatus -eq "D") -or ($CurEmplStatus -eq "U") -or ($CurEmplStatus -eq "T") -or ($CurEmplStatus -eq "V") -or ($CurEmplAction -eq "Notice")) {
     #
     # Employee is TERMINATED
     #
			    DebugLog " -- Terminated in PS, Effective Date $CurEffectiveDate"
		 	    if ($CurEmplDN -ne $false) { #if terminated in PS and account exists, disable it
		 		    DebugLog " --- Account exists"
		 		    $CurSAMaccount = QuerySAMAccount($CurEmplDN)
				    FindExtraAccounts $CurSAMaccount
		 		    DisableExtraAccounts $CurSAMaccount $CurITRegion $isSuspended
		 		    if ((DisableUser $CurEmplDN $CurITRegion $isSuspended) -eq $true) { #if account not already disabled, do so
		 			    $TmpUsrObj = $null
		 			    $TmpUsrObj = Get-SDLADUser -EmployeeNum $CurEMPLID
		 			
					    $TmpMobileNumber = $TmpUsrObj.mobile
					    if ($TmpMobileNumber) {$TmpMobileNumber = $TmpMobileNumber.Trim()}
					    $TmpDeskNumber = $TmpUsrObj.telephoneNumber
					    if ($TmpDeskNumber) {$TmpDeskNumber = $TmpDeskNumber.Trim()}
		 			    EmailManagerStatus $null $CurManagerID $CurManagerEmail "Termination/Disabled" $CurLastName $CurMiddleInitial $CurFirstName $CurEMPLID $CurSAMaccount $CurPSRegion $CurJobTitle $CurLocation $TmpCityState $CurBusinessUnit $CurEffectiveDate $CurEmplAction $CurEmplStatus $CurConferenceCard $CurPhoneNumber $CurSalesID $CurEmplType $CurVendorName $TmpMobileNumber $TmpDeskNumber $CurICOMSID $CurDeptID
                        $TmpCityStateAddr = "$CurCity-$CurState-$CurAddress1"		
                        ServiceNowTermEmail $CurManagerID $CurLastName $CurFirstName $CurEMPLID $CurJobTitle $CurLocation $TmpCityStateAddr $CurBusinessUnit $CurEffectiveDate $CurEmplType $CurVendorName $TmpDeskNumber $CurDeptID $CurDepartment $CurCompanyCode $CurCompany
		 			    $CurSAMaccount = $null
		 			    $TmpCityState = $null
		 		    } else { #if account was already restricted, delete after 90 days
		 			    $dtmPwdLastSet = [datetime]$CurEffectiveDate
					    $spanPwdLastSet = New-TimeSpan $dtmPwdLastSet $(Get-Date)
					    DebugLog " --- Effective date was $($spanPwdLastSet.Days) days ago."
					    If ($spanPwdLastSet.Days -gt 90) {
						    DebugLog " --- Deleting account"
						    DeleteAccount $CurEmplDN $CurITRegion
							$CurAccountFlags = GetAccountFlags $CurEmplID
							if ($CurAccountFlags) {
								Invoke-Sqlcmd -ServerInstance $AccountProvisioningSQLServer -Database $AccountProvisioningSQLDatabase -Query "DELETE FROM AccountFlags where EmployeeID = $CurEmplID"
							}
						    if ($script:CurAdminAccounts) {
							    foreach ($tmpUser in $script:CurAdminAccounts) {
								    DebugLog " --- Deleting Admin Account ($($tmpUser.SamAccountName))"
								    Remove-ADUser -Identity $tmpUser.distinguishedName -Confirm:$false
							    }
						    }
						    if ($script:CurTestAccounts) {
							    foreach ($tmpUser in $script:CurTestAccounts) {
								    DebugLog " --- Deleting Test Account ($($tmpUser.SamAccountName))"
								    Remove-ADUser -Identity $tmpUser.distinguishedName -Confirm:$false
							    }
						    }
					    }
		 		    }

				    if ((Get-User -Identity $CurEmplDN -IgnoreDefaultScope).RecipientType -eq "UserMailbox") {
					    #Remove ActiveSync device relationships
					    $TMPASDeviceList = $false
					    $TMPASDeviceList = get-MobileDevice -mailbox $CurEmplDN
					    if ($TMPASDeviceList) {
						    DebugLog " ---- Found ActiveSync devices, removing them"
						    foreach ($TMPASDevice in $TMPASDeviceList) {
							    DebugLog " ----- Removing $($tmpASDevice.FriendlyName)"
							    Remove-MobileDevice -identity $TMPASDevice.identity -confirm:$false
						    } 
					    }
				
					    #delete mailbox 14 days after Term
					    $dtmEffectiveDate = [datetime]$CurEffectiveDate
					    $spanEffective = New-TimeSpan -Start $dtmEffectiveDate -End $(Get-Date)
					    If ($spanEffective.Days -gt 14 -and $spanEffective.Days -le 90) {
						    DebugLog " --- Effective date was $($spanEffective.Days) days ago."
						    DebugLog " --- Deleting mailbox"
						    Disable-Mailbox -Identity $CurEmplDN -Confirm:$false
					    }
				    }
		 	    } else {
				    DebugLog " --- no account"
				    $dtmPwdLastSet = [datetime]$CurEffectiveDate
				    $spanPwdLastSet = New-TimeSpan $dtmPwdLastSet $(Get-Date)
				    DebugLog " --- Effective date was $spanPwdLastSet days ago."
				    If ($spanPwdLastSet.Days -lt 90) {
					    DebugLog " --- Re-creating account"
					    CreateAccount $CurFirstName $CurLastName $CurITRegion $CurEMPLID
				    } 
			    }
		    } elseif ($CurEmplStatus -eq "A") {
     #
     # Employee is ACTIVE
     #
			    DebugLog " -- Active in PS, Effective Date $CurEffectiveDate"
			
			    if ($CurEmplDN -eq $false) { #if no account already exists, create one
				    DebugLog " --- No account, creating one"
				    CreateAccount $TMPFirstName $TMPLastName $CurITRegion $CurEMPLID
				    $CurEmplDN = QueryAccount($CurEMPLID)
				    $tmpStartTime = Get-Date
				    do {
					    $tmpCurTime = Get-Date
					    $CurEmplDN = QueryAccount($CurEMPLID)
				    } until ($CurEmplDN -ne $false)
				    $tmpWaitSpan = New-TimeSpan $tmpStartTime $tmpCurTime
				    DebugLog " --- took $tmpWaitSpan seconds to complete."
				    $CurSAMaccount = QuerySAMAccount($CurEmplDN)
				    $TmpMobileNumber = $null
				    $TmpDeskNumber = $null
				    EmailManagerStatus $script:CurPW $CurManagerID $CurManagerEmail "Creation" $CurLastName $CurMiddleInitial $CurFirstName $CurEMPLID $CurSAMaccount $CurPSRegion $CurJobTitle $CurLocation $TmpCityState $CurBusinessUnit $CurEffectiveDate $CurEmplAction $CurEmplStatus $CurConferenceCard $CurPhoneNumber $CurSalesID $CurEmplType $CurVendorName $TmpMobileNumber $TmpDeskNumber $CurICOMSID $CurDeptID
	 			    $CurSAMaccount = $null
	 			    $TmpCityState = $null
			    }

                if (($CurCity -ne "Plano" -and $CurCity -ne "Addison") -or $CurState -ne "TX") {
			#
            # Do not create mailbox or enable Lync for users in Plano or Addison, TX
            #

			        if ((Get-User -Identity $CurEmplDN -IgnoreDefaultScope).RecipientType -eq "User") {
				        DebugLog " --- no mailbox, creating one"
					    $tmpManLvl = GetManagementLevel($CurJobTitle)
					    SetMailboxLevel $CurEmplDN $tmpManLvl $True
			        }
			
			        $ErrorActionPreference = "stop" #the following error is non-terminating.  we have to set this so the error can be caught
			        Try {
				        $bLyncEnabled = (Get-CsUser -Identity $CurEmplDN).enabled
			        } Catch [system.exception] {
				        $bLyncEnabled = $False
			        }
			        $ErrorActionPreference = "Continue" #set the error action back to default
			
			        if ($bLyncEnabled -eq $False) {
				        Enable-CsUser -Identity $CurEmplDN -SipAddressType UserPrincipalName -RegistrarPool 'lync-pool-dal.suddenlink.cequel3.com'
			        }
			    } else {
					DebugLog " --- Plano/Addison, no mailbox or Lync"
				}
                
			    DebugLog " --- Updating account with latest data from PS"
			    UpdateAccount $CurEmplDN $CurEMPLID $TMPFirstName $TMPLastName $CurMiddleInitial $CurJobTitle $CurLocation $CurAddress1 $CurAddress2 $CurCity $CurState $CurZipCode $CurManagerDN $CurBusinessUnit $CurDepartment $isConsultant $CurVendorName $isSuspended $CurCompany $CurITRegion $CurPSRegion $CurDeptID $CurICOMSID
			    $TmpUsrObj = Get-SDLADUser -EmployeeNum $CurEMPLID

			    ReenableDisabledAccount $CurEmplDN $CurITRegion

			    if ($TmpUsrObj.extensionAttribute5 -eq "S" -or $TmpUsrObj.extensionAttribute5 -eq "T") {
                    if ($TmpUsrObj.extensionAttribute5 -eq "T") {
    				    $CurSAMaccount = QuerySAMAccount($CurEmplDN)
				        EmailManagerStatus $script:CurPW $CurManagerID $CurManagerEmail "Re-hire" $CurLastName $CurMiddleInitial $CurFirstName $CurEMPLID $CurSAMaccount $CurPSRegion $CurJobTitle $CurLocation $TmpCityState $CurBusinessUnit $CurEffectiveDate $CurEmplAction $CurEmplStatus $CurConferenceCard $CurPhoneNumber $CurSalesID $CurEmplType $CurVendorName $TmpMobileNumber $TmpDeskNumber $CurICOMSID $CurDeptID
			        }
                } else {
            	    DebugLog " --- Account is already enabled."
				    if ($TmpUsrObj.extensionAttribute5 -ne 'A') {set-aduser $CurEmplDN -replace @{extensionAttribute5 = "A"}}
			    }
			    $TmpUsrObj = $null	 			
 			    $CurSAMaccount = $null
		    } else {
     #
     # Employment status is not recognized
     #
                DebugLog " --- Not a valid Employment status ($CurEmplStatus)"
		    }
        } else { #PAR indicates person should NOT have an AD account
			
			DebugLog " ---- No AD Account requested on PAR"
			$CurPW = "No AD Account requested on PAR"
			$CurSAMaccount = "No AD Account requested on PAR"
			$TmpMobileNumber = $null
			$TmpDeskNumber = $null
            $CurEmailManagerFlag = $false

            $SQLString = $null
			$CurAccountFlags = GetAccountFlags $CurEmplID
			if ($CurAccountFlags) {
                if ($CurEmplStatus -eq 'A') {
					$CurAction = "Creation"
                    if ($CurAccountFlags.HireEmailSent -ne 1) {
                        $CurHireEmailSentFlag = 1
                        $CurEmailManagerFlag = $true
                        $SQLString = "UPDATE AccountFlags SET HireEmailSent = 1 WHERE EmployeeID = $CurEmplid"
                    }
				} elseif ($CurEmplStatus -eq 'T') {
					$CurAction = "Termination/Disabled"
                    if ($CurAccountFlags.TermEmailSent -ne 1) {
                        $CurTermEmailSentFlag = 1
                        $CurEmailManagerFlag = $true
                        $SQLString = "UPDATE AccountFlags SET TermEmailSent = 1 WHERE EmployeeID = $CurEmplid"
                    }
				}
            } else {
                $CurEmailManagerFlag = $true

				if ($CurEmplStatus -eq 'A') {
					$CurAction = "Creation"
                    $CurHireEmailSentFlag = 1
                    $CurTermEmailSentFlag = 0
				} elseif ($CurEmplStatus -eq 'T') {
					$CurAction = "Termination/Disabled"
                    $CurHireEmailSentFlag = 1
                    $CurTermEmailSentFlag = 1
				}

				$SQLString = "INSERT INTO AccountFlags(EmployeeID,NoAccountFlag,HireEmailSent,TermEmailSent) Values('$CurEMPLID',1,$CurHireEmailSentFlag,$CurTermEmailSentFlag)"
            }				
            
            if ($SQLString) {AccountProvisioningDBEx $SQLString}

            If ($CurEmailManagerFlag) {
				EmailManagerStatus $CurPW $CurManagerID $CurManagerEmail $CurAction $CurLastName $CurMiddleInitial $CurFirstName $CurEMPLID $CurSAMaccount $CurPSRegion $CurJobTitle $CurLocation $TmpCityState $CurBusinessUnit $CurEffectiveDate $CurEmplAction $CurEmplStatus $CurConferenceCard $CurPhoneNumber $CurSalesID $CurEmplType $CurVendorName $TmpMobileNumber $TmpDeskNumber $CurICOMSID $CurDeptID
		   } else {
				DebugLog " ---- $CurAction Email already sent."
		   }
        }
		$CurEMPLID = $null
		$CurEmplAction = $null
		$CurEmplStatus = $null
		$CurEffectiveDate = $null
		$CurPreferredName = $null
		$CurFirstName = $null
 		$CurFirstName = $null
		$CurMiddleInitial = $null
		$CurLastName = $null
 		$CurLastName = $null
		$CurJobTitle = $null
		$CurLocation = $null
		$CurITRegion = $null
		$CurAddress1 = $null
		$CurAddress2 = $null
		$CurCity = $null
		$CurState = $null
		$CurZipCode = $null
		$CurDepartment = $null
		$CurBusinessUnit = $null
		$CurManagerID = $null
		$CurConferenceCard = $null
		$CurPhoneNumber = $null
		$CurSalesID = $null
		$CurEmplType = $null
		$CurVendorName = $null
		$TmpCityState = $null
		$TMPFirstName = $null
		$TMPLastName = $null
		$TmpMobileNumber = $null
		$TmpDeskNumber = $null
		$CurPSRegion = $null
		$CurICOMSID = $null
		$script:CurAdminAccounts = $null
		$script:CurTestAccounts = $null
	}
	[void]$PSReader.close()
	[void]$PeopleOracleCN.close()
	EmailProvisioningCompleteStatus
}
Main
