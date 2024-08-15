
$DNSServer = read-host -Prompt "DNS Server to query?"
$AllInternalZones = Get-DnsServerZone -ComputerName $DNSServer | Where-Object {$_.zonetype -eq 'Primary' -and $_.IsReverseLookupZone -eq $false}

$AllRecords = @()

foreach ($zone in $AllInternalZones) {
    $AllRecords += Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ComputerName $DNSServer -RRType A
}

$vLookup = read-host -Prompt "IP Address to find records for?"

    ForEach ($vRecord2 in $AllRecords) { 
        # Find Host Names that match lookup string
        If ($vRecord2.RecordData.IPv4Address -eq $vLookup) {
            Write-Host "The DNS A Entry " -NoNewline
#            if ($vRecord2.HostName -eq '@') { 
#                write-host $vRecord2.DistinguishedName.split(',')[1].split('=')[1] -ForegroundColor DarkYellow -NoNewline
#            } else {
               #Write-Host $vRecord2.HostName -NoNewline -ForegroundColor Green
               Write-Host $vRecord2.DistinguishedName.split(',')[0].split('=')[1] $vRecord2.DistinguishedName.split(',')[1].split('=')[1] -NoNewline -ForegroundColor Green
#            }
            Write-Host " is pointing to the IP address " -NoNewline
            Write-Host $vRecord2.RecordData.IPv4Address -ForegroundColor Green
        }
    }

