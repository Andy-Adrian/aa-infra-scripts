# Search DNS for text string and return entries

Param(
    [switch]$OutputCSV = $false,
    [Parameter(mandatory)][string]$DNSServer = "",
    [string]$DNSdomain = "",
    [Parameter(mandatory)][string[]]$Hostname = ""
)

if ($DNSDomain -eq "") {
    if ($Hostname -like "*.*") {
        $tHostnameArr = $Hostname.split(".")
        $tdomainarr = $tHostnameArr[1..$tHostnameArr.Length]
        $vDNSdomain = $tdomainarr -join "."
        $vLookup = $tHostnameArr[0]
    }
}

if ($vLookup -eq "") {
    $vLookup = '*' + $Hostname + '*'
} else {
    $vLookup = '*' + $vLookup + '*'
}
$vRecords4 = @()
$vOutput = @()

# Get CNAME records from zone
$vRecords1 = Get-DnsServerResourceRecord -ZoneName $vDNSDomain -ComputerName $DNSServer -RRType CNAME
$tFound = $false
ForEach ($vRecord1 in $vRecords1) {
    # Find Find CNAME entries that match lookup string 
    If ($vRecord1.RecordData.HostNameAlias -like $vLookup) {
        $vOutput += $vRecord1
        $tFound = $true
        Write-Host "The DNS CNAME Entry " -NoNewline
        Write-Host $vRecord1.HostName  -NoNewline -ForeGroundColor Green
        Write-Host " is pointing to the name "  -NoNewline
        Write-Host $vRecord1.RecordData.HostNameAlias -ForeGroundColor Green
    }
    # Find Host Names that match lookup string
    If ($vRecord1.HostName -like $vLookup) {
        $vOutput += $vRecord1
        $tFound = $true
        Write-Host "The DNS CNAME Entry " -NoNewline
        Write-Host $vRecord1.HostName  -NoNewline -ForeGroundColor Green
        Write-Host " is pointing to the name "  -NoNewline
        Write-Host $vRecord1.RecordData.HostNameAlias -ForeGroundColor Green
    }
}
if ($tFound -eq $false) {Write-Host "No CNAME records found matching $vLookup"}

# Get A records from zone
$vRecords2 = Get-DnsServerResourceRecord -ZoneName $vDNSDomain -ComputerName $DNSServer -RRType A
$tFound = $false
ForEach ($vRecord2 in $vRecords2) { 
    # Find Host Names that match lookup string
    If ($vRecord2.HostName -like $vLookup) {
        $vOutput += $vRecord2
        $tFound = $true
        Write-Host "The DNS A Entry " -NoNewline
        Write-Host $vRecord2.HostName -NoNewline -ForegroundColor Green
        Write-Host " is pointing to the IP address " -NoNewline
        Write-Host $vRecord2.RecordData.IPv4Address -ForegroundColor Green
    }
}
if ($tFound -eq $false) {Write-Host "No A records found matching $vLookup"}
   
$vReverseZones = Get-DnsServerZone | Where-Object{$_.IsReverseLookupZone -and -not $_.IsAutoCreated}
$tFound = $false

foreach ($vRZ in $vReverseZones) {
    #Get PTR entries from zone
    $vRecords3 = Get-DnsServerResourceRecord -ZoneName $vRZ.ZoneName -ComputerName $DNSServer

    ForEach ($vRecord3 in $vRecords3) { 
        # Find Reverse entries that match lookup string
        If ($vRecord3.RecordData.ptrDomainName -like $vLookup) {
            $vOutput += $vRecord3
            $tFound = $true
            Write-Host "The DNS PTR entry " -NoNewline
            Write-Host $vRecord3.HostName -NoNewline -ForegroundColor Green
            Write-Host " in the " -NoNewline
            Write-Host $vRZ.ZoneName -NoNewline -ForegroundColor Green
            Write-Host " is pointing to the name " -NoNewline
            Write-Host $vRecord3.RecordData.ptrDomainName -ForegroundColor Green
            $vRecords4 = $vRecords4 + $vRecord3.HostName
        } 
    }
}
if ($tFound -eq $false) {write-host "No PTR records found for $vLookup"}

ForEach ($vRecord4 in $vRecords4) {
    $tFound = $false
    ForEach ($vRZ in $vReverseZones) {
        $vRecords5 = Get-DnsServerResourceRecord -ZoneName $vRZ.ZoneName -ComputerName $DNSServer -Name $vRecord4

        ForEach ($vRecord5 in $vRecords5) {
            #$vOutput += $vRecord5
            $tFound = $true
            $vIP = @()
            $vOctSplit = $vRZ.ZoneName.Split(".")
            [array]::Reverse($vOctSplit)
            foreach ($tOct in $vOctSplit) {
                if ($tOct -match "^\d+$") {$vIP += $tOct}
            }
            $vIP += $vRecord5.Hostname
            Write-Host "There is also a PTR entry for IP " -NoNewline
            Write-Host "$($vIP -join ".")" -NoNewline -ForegroundColor Green
            Write-Host " pointing to the name " -NoNewline
            Write-Host $vRecord5.RecordData.PtrDomainName -ForegroundColor Green
        }
    }
    if ($tFound -eq $false) {write-host "No PTR records found for $vRecord5"}
}

if ($OutputCSV) {$vOutput | export-csv -Path c:\temp\DNS_Search_Results.csv -Append -NoTypeInformation}