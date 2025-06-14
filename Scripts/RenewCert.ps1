  <#
.SYNOPSIS
Uppdaterar certifikat för IIS, Exchange och RDS med Posh-ACME och Loopia API.

.DESCRIPTION
Uppdaterar certifikat för IIS, Exchange och RDS med Posh-ACME och Loopia API. 
Skriptet skapar en schemalagd uppgift för automatisk certifikatförnyelse om det anges.

.PARAMETER CreateSchedule
Anges till $true för att skapa en schemalagd uppgift för automatisk certifikatförnyelse.

.EXAMPLE 

RenewCert.ps1 -CreateSchedule $true

.NOTES
Filen skall läggas i mappen C:\Scripts\.

#>
param (
    [bool]$CreateSchedule = $false  # Välj om schemaläggning ska skapas
)

# ==== Konfigurationsinställningar ====
$Domän = "example.com"
$Epost = "admin@example.com"
$PFXPath = "C:\Cert\RDS_Cert.pfx"
$MainPath = "C:\Cert"
$BackupPath = "C:\Cert\Backup\Cert_$(Get-Date -Format yyyyMMdd).pfx"
$IISBinding = "0.0.0.0:443"
$LogFile = "C:\Cert\CertUpdateLog.txt"
$CertPassword = ConvertTo-SecureString "SuperSäkertLösenord" -AsPlainText -Force
$UseSMTP = $true
$SMTPServer = "smtp.example.com"
$EmailFrom = "admin@example.com"
$EmailTo = "alerts@example.com"
$RDSServices = @("TermService", "SessionEnv", "UmRdpService", "W3SVC")
$RDSRoles = @("RDGateway", "RDWebAccess", "RDRedirector", "RDPublishing")
$ExchangeServices = @("MSExchangeTransport", "MSExchangeFrontEndTransport", "MSExchangeIS", "MSExchangePOP3", "MSExchangeIMAP4")

# ==== Roller att uppdatera certifikat för ====
$IIS = $false
$Exchange = $false
$RDS = $false

# ==== Loggfunktion ====
function Write-Log {
    param($Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Out-File -Append -FilePath $LogFile
}

# ==== Installera och importera Posh-ACME ====
function Install-PoshACME {
    try {
        if (-not (Get-Module -ListAvailable -Name Posh-ACME)) {
            Install-Module -Name Posh-ACME -Force -ErrorAction Stop
            Write-Log "Posh-ACME installerad."
        }
        Import-Module Posh-ACME -ErrorAction Stop
        Write-Log "Posh-ACME modulen laddad."
    } catch {
        Write-Log "Fel: Kunde inte installera/importera Posh-ACME - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
        exit
    }
}

# ==== Backup av befintligt certifikat ====
function Backup-Certificate {
    try {
        if (Test-Path $PFXPath) {
            Copy-Item -Path $PFXPath -Destination $BackupPath -Force
            Write-Log "Backup av certifikat skapad: $BackupPath"
        }
    } catch {
        Write-Log "Fel: Kunde inte skapa certifikat-backup - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
    }
}

# ==== Certifikatförnyelse ====
function Renew-Certificate {
    try {
        if ($Cert = Submit-Renewal $Domän) {
            Write-Log "Certifikat erhållet för $Domän."
            return $Cert
        }
        else {
            Write-Log "Fel: Certifikat kunde inte hämtas."
            exit
        }
    } catch {
        Write-Log "Fel: Problem med certifikatskapning - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
        exit
    }
}

# ==== Exportera certifikatet till PFX ====
function Export-Certificate {
    param ($Cert)
    try {
        Export-PfxCertificate -Cert $Cert.PfxFile -FilePath $PFXPath -Password $CertPassword -ErrorAction Stop
        Write-Log "Certifikat exporterat till $PFXPath."
    } catch {
        Write-Log "Fel: Problem med certifikatexport - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
        exit
    }
}

# ==== Uppdatera RDS-certifikaten ====
function Update-RDSCertificates {
    try {
        foreach ($Role in $RDSRoles) {
            Set-RDCertificate -Role $Role -ImportPath $PFXPath -Password $CertPassword -ErrorAction Stop
            Write-Log "Certifikat uppdaterat för $Role."
        }
    } catch {
        Write-Log "Fel: Problem med RDS-certifikatuppdatering - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
        exit
    }
}

# ==== Uppdatera IIS-certifikatet ====
function Update-IISCertificate {
    try {
        $CertThumbprint = (Get-PfxCertificate -FilePath $PFXPath).Thumbprint
        Import-PfxCertificate -FilePath $PFXPath -CertStoreLocation Cert:\LocalMachine\My -Password $CertPassword -Exportable
        Write-Log "Certifikat importerat till Windows certifikatlager."

        # Hitta IIS-webbplatsens inställningar
        $Binding = Get-WebBinding -Protocol "https" | Where-Object { $_.BindingInformation -eq $IISBinding }
        
        if ($Binding -ne $null) {
            # Uppdatera IIS-bindningen med det nya certifikatet
            $Binding.AddSslCertificate($CertThumbprint, "My")
            Write-Log "IIS-certifikat uppdaterat!"
        } else {
            Write-Log "Fel: Kunde inte hitta IIS-bindning!"
            if($UseSMTP) {
                Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
            }
        }
    } catch {
        Write-Log "Fel: Problem med IIS-certifikatuppdatering - $_"
    }
}

# ==== Uppdatera Exchange-certifikatet ====
function Update-ExchangeCertificate {
    try {
        Import-PfxCertificate -FilePath $PFXPath -CertStoreLocation Cert:\LocalMachine\My -Password $CertPassword -Exportable
        $CertThumbprint = (Get-PfxCertificate -FilePath $PFXPath).Thumbprint
        Write-Log "Certifikat importerat till Windows certifikatlager för Exchange."

        # Uppdatera Exchange-certifikat för IIS, SMTP, IMAP och POP
        Enable-ExchangeCertificate -Thumbprint $CertThumbprint -Services "IIS,SMTP,IMAP,POP" -ErrorAction Stop
        Write-Log "Exchange-certifikat uppdaterat!"
    } catch {
        Write-Log "Fel: Problem med Exchange-certifikatuppdatering - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
    }
}

# ==== Starta om tjänster ====
function Restart-Services {
    if($RDS) {
        foreach ($Service in $RDSServices) {
            try {
                Restart-Service -Name $Service -Force
                Write-Log "Tjänsten $Service har startats om."
            } catch {
                Write-Log "Fel: Kunde inte starta om tjänsten $Service - $_"
                if($UseSMTP) {
                    Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
                }
            }
        }
    }
    if($IIS) {
        try {
            Restart-Service -Name "W3SVC" -Force
            Write-Log "IIS-tjänsten har startats om."
        } catch {
            Write-Log "Fel: Kunde inte starta om IIS-tjänsten - $_"
            if($UseSMTP) {
                Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
            }
        }
    }
    # ==== Starta om Exchange-tjänster ====
    if($Exchange) {
        foreach ($Service in $ExchangeServices) {
            try {
                Restart-Service -Name $Service -Force
                Write-Log "Exchange-tjänsten $Service har startats om."
            } catch {
                Write-Log "Fel: Kunde inte starta om Exchange-tjänsten $Service - $_"
                if($UseSMTP) {
                    Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
                }
            }
        }
    }
}

# ==== Schemaläggning ====
function ScheduleCertificateRenewal {
    try {
        $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File C:\Scripts\RenewCert.ps1"
        $Trigger = New-ScheduledTaskTrigger -Daily -At 3am
        Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "Automatisk Certifikatförnyelse" -User "System"
        Write-Log "Schemalagd uppgift skapad för automatisk certifikatförnyelse."
    } catch {
        Write-Log "Fel: Problem med schemaläggning - $_"
        if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatfel" -Body "Ett fel uppstod vid certifikatuppdatering. Kontrollera loggen."
        }
    }
}

# ==== Kontrollera och skapa nödvändiga mappar ====
function Folder-Check {
    if (-not(Test-Path $MainPath -PathType Container)) {
        try {
            New-Item -path $MainPath -ItemType Directory -ErrorAction stop 
        }
        catch {
            Write-Log  "Kund inte skapa huvudmappen: $MainPath"
            Write-Log $_.Exception.Message
        }
    }
    if (-not(Test-Path $BackupPath -PathType Container)) {
        try {
            New-Item -path $BackupPath -ItemType Directory -ErrorAction stop 
        }
        catch {
            Write-Log  "Kund inte skapa backup-mappen: $BackupPath"
            Write-Log $_.Exception.Message
        }
    }
}

# ==== Huvudskript - Kör alla funktioner ====
Write-Log "Startar certifikatsuppdatering..."
Install-PoshACME
Folder-Check
if($CreateSchedule -eq $true) {
    ScheduleCertificateRenewal
    Write-Log "Schemaläggning av certifikatförnyelse skapad."
    Write-Host "Schemaläggning av certifikatförnyelse har skapats. Den kommer att köras dagligen kl. 03:00."
} else {
    Backup-Certificate
    $Cert = Renew-Certificate
    Export-Certificate -Cert $Cert
    if($IIS -eq $true) {
        Update-IISCertificate
    }
    if($Exchange -eq $true) {
        Update-ExchangeCertificate
    }
    if($RDS -eq $true) {
        Update-RDSCertificates
    }
    Restart-Services
    Write-Log "Certifikatsuppdatering slutförd!"
    Write-Host "Certifikatsuppdateringen har lyckats! Kontrollera loggfilen för detaljer: $LogFile"
    if($UseSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject "Certifikatuppdatering $Domän" -Body "Certifikatuppdatering för domänen $Domän lyckades." -Attachments $LogFile
    }
}