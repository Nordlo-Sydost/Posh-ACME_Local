$savepfx = True
$destination_folder = "C:\Certificate"
$domain = ""
$name = $domain
$contact = ""
$pfxfile = "C:\Posh-ACME\$domain-$(get-date -format yyyyMMdd).pfx"
$PfxPassSecure = ConvertTo-SecureString 'xxxxxxx' -AsPlainText -Force
$pluginargs = @{
    LoopiaUser = 'nordlo_sydost@loopiaapi'; 
    LoopiaPass = ConvertTo-SecureString 'd9yDEj^895BQ@YKUN7baq2Qo' -AsPlainText -Force
}

if(-Not Get-PAAccount) {
    #New-PAAccount -Contact 'me@example.com' -AcceptTOS
}
#$cert = New-PACertificate $domain -AcceptTOS -Contact $contact -Plugin Loopia -PluginArgs $pluginargs -Install -Verbose
if($savepfx) {
    if (-not (Test-Path $destination_folder)) {
        New-Item -ItemType Directory -Path $destination_folder
    }
    Copy-Item -Path $cert.PfxFile -Destination "$destination_folder\$pfxfile"
}
