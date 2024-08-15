# Clean up/ Server decommisioned
$vServerName = read-host -Prompt "Server to decommission?"
$vDisableAccount = read-host -Prompt "Any service accounts to disable? (Y/N)"
if ($vDisableAccount -match "[yY]") {
    $vAcctNames = read-host -Prompt "List all service accounts, comma separated (account1, account2)"
    $vAcctNames = $vAcctNames.split(",")
    $vDisableAccount = $true
}
$vDC = read-host "Nearest domain controller? (to disable accounts and delete DNS entries)"
$vDisDate = (Get-Date).ToShortDateString()
$vDelDate = (Get-Date).AddDays(14).ToShortDateString()
$vPW = Read-Host "What do you want the password for the local Administrator to be?"
$vLPW = ConvertTo-SecureString -String $vPW -AsPlainText -Force
$vLocalCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$($vServerName)\Administrator", $vLPW

## Unjoin server from domain
# Change local Administrator password so we can get back in
([adsi]“WinNT://$($vServerName)/Administrator”).SetPassword($vPW)
# Unjoin server remotely
Remove-Computer -ComputerName $vServerName -UnJoinDomainCredential $(Get-Credential) -WorkGroup DECOMMISSIONED -Force -Restart
# Shutdown server
Stop-Computer -ComputerName $vServerName -Credential $vLocalCred -Force

# Disable service account(s) in AD
#Get-ADUser -Filter {SamAccountName -like $vAcctName} -Server $vDC -Properties Description | Select-Object Name, Enabled, Description
if ($vDisableAccount) {
    foreach ($vAcctName in $vAcctNames) {
        $vAcctName = $vAcctName.Trim()
        $vSerAcct = Get-ADUser -Filter { SamAccountName -like $vAcctName } -Server $vDC -Properties Description
        ForEach ($vAcct in $vSerAcct) {
            if ((read-host -Prompt "Found $($vAcct.samaccountname), disable? (Y/N)") -match "[yY]") {
                Set-ADUser -Server $vDC -Identity $vAcct -Enabled $False -Description "Disabled $($vDisDate) - Can be deleted after $($vDelDate) - $($vAcct.Description)" -Confirm:$false
            }
        }
    }
}

# Disable computer object in AD
$vSTD = Get-ADComputer -Identity $vServerName -Properties Description -Server $vDC
Set-ADComputer -Identity $vSTD.Name -Server $vDC -Enabled $false -Description "Decommissioned on $($vDisDate) - Can be deleted after $($vDelDate) - $($vSTD.Description)"

# Remove DNS entries from forward lookup zone
$vDNSzoneName = read-host -Prompt "What DNS domain to remove from?"
$vForDNS = Get-DnsServerResourceRecord -ZoneName $vDNSzoneName -ComputerName $vDC | Where-Object { $_.HostName -like $($vServerName) }
ForEach ($vEntry2 in $vForDNS) { Remove-DnsServerResourceRecord -ZoneName $vDNSzoneName -ComputerName $vDC -Name $vEntry2.HostName -RRType $vEntry2.RecordType -Force }

# Remove DNS entries from reverse lookup zone
$vDNSRZoneName = read-host -Prompt "What reverse DNS zone to remove from? (x.y.z.in-addr.arpa)"
$vRevDNS = Get-DnsServerResourceRecord -ZoneName $vDNSRZoneName -ComputerName $vDC | Where-Object { $_.RecordData.ptrDomainName -like "$($vServerName)*" }
ForEach ($vEntry in $vRevDNS) { Remove-DnsServerResourceRecord -ZoneName $vDNSRZoneName -ComputerName $vDC -Name $vEntry.HostName -RRType Ptr -Force }

## "disable" groups by removing members. The list of members is copied into the Notes field of the group incase group is still in use
## Be extremely sure ALL the groups you are filtering are correct!!!
<#

## -----------------------------------------------------------------------------------------------------------------
## Commented out because here lies dragons and danger. 
## Modify so the search can only find groups associated with the server to be decomissioned (if such things exist)
## Maybe also add a validation inside the foreach to double-check and approve the change in real-time
## -----------------------------------------------------------------------------------------------------------------

$vGrpName = "*$vServerName*"
$vGrpList = Get-ADGroup -Filter { Name -like $vGrpName } -Properties Description
#Get-ADGroup -Filter {Name -like $vGrpName} -Properties Description -Server $vDC | Select-Object Name,Description

ForEach ($vGrp in $vGrpList) {
    $vGrpMembers = Get-ADGroupMember $vGrp.Name | Select-Object SamAccountName
    $vInfo = $vGrpMembers.SamAccountName
    Set-ADGroup -Identity $vGrp.Name -Server $vDC -Replace @{info = "Members used to be: $vInfo" } -Description "Group members removed $($vDisDate) to decommission group - Can be deleted after $($vDelDate) - $($vGrp.Description)"
    ForEach ($vMem in $vGrpMembers) {
        Remove-ADGroupMember -Identity $vGrp.Name -Server $vDC -Members $vMem.SamAccountName -Confirm:$false
    }
}
#>