Start-Transcript c:\scripts\cert-renewal.log -Append

$hostname ="vcl01.oitpm.se"

$oldCert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.Subject -eq "CN=$hostname" } |
    Sort-Object -Descending NotAfter |
    Select-Object -First 1

if (Submit-Renewal -NoSkipManualDns) {
    $oldCert | Remove-Item

    $certname = "vcc-$(get-date -format yyyyMMdd)"
    $pfxfile = "C:\Posh-ACME\$certname.pfx"

    $pfxpass = 'poshacme'

    $cert = Get-PACertificate -MainDomain $domain
    Copy-Item -Path $cert.PfxFile -Destination C:\Posh-ACME\$certname.pfx
    $securepfxpass = $pfxpass | ConvertTo-SecureString -AsPlainText -Force
    Import-PfxCertificate -FilePath $pfxfile -Password $securepfxpass -CertStoreLocation cert:\localMachine\my -Exportable

    Connect-VBRServer -Server localhost
    # Write-Host "Connected to VBRServer"

    $thumbprint = Get-PACertificate -MainDomain $domain
    $thumbprint = $thumbprint.Thumbprint
    $certificate = Get-VBRCloudGatewayCertificate -FromStore | Where {$_.Thumbprint -eq $thumbprint}

    Add-VBRCloudGatewayCertificate -Certificate $certificate
    # Write-Host "Added certificate to VBRServer"

    Disconnect-VBRServer
    # Write-Host "Disconnected from VBRServer"
}
else {
    # Write-Host "Ingen uppdatering gjord"
}
Stop-Transcript