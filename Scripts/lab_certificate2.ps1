# === Inställningar ===
$domain = "example.com"
$email = "admin@example.com"
$siteName = "Default Web Site"
$certPath = "C:\Certs\$domain"
$logFile = "C:\Certs\ssl_install_log.txt"

# === Loggningsfunktion ===
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$level] $message"
    Write-Output $logEntry
    Add-Content -Path $logFile -Value $logEntry
}

# === Starta loggning ===
Write-Log "Startar SSL-certifikatinstallation för $domain"

try {
    # Installera Posh-ACME om det inte finns
    if (-not (Get-Module -ListAvailable -Name Posh-ACME)) {
        Write-Log "Installerar Posh-ACME-modulen"
        Install-Module -Name Posh-ACME -Scope CurrentUser -Force
    }

    Import-Module Posh-ACME

    # Skapa katalog för certifikat
    New-Item -ItemType Directory -Path $certPath -Force | Out-Null

    # Skapa konto om det inte finns
    if (-not (Get-PAAccount)) {
        Write-Log "Skapar nytt ACME-konto"
        New-PAAccount -Contact "mailto:$email" -AcceptTOS
    }

    # Begär certifikat
    Write-Log "Begär certifikat för $domain"
    New-PACertificate -Domain $domain -AcceptTOS -Install -UseIIS

    # Hämta certifikat
    $cert = Get-PACertificate $domain
    $thumbprint = $cert.Certificate.Thumbprint
    Write-Log "Certifikat erhållet med thumbprint: $thumbprint"

    # Installera i certifikatbutiken
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","LocalMachine")
    $store.Open("ReadWrite")
    $store.Add($cert.Certificate)
    $store.Close()
    Write-Log "Certifikat installerat i certifikatbutiken"

    # Binda till IIS
    Import-Module WebAdministration
    $binding = Get-WebBinding -Name $siteName -Protocol "https"
    if ($binding) {
        Remove-WebBinding -Name $siteName -Protocol "https"
        Write-Log "Tidigare HTTPS-bindning borttagen"
    }

    New-WebBinding -Name $siteName -Protocol https -Port 443 -HostHeader $domain
    Push-Location IIS:\SslBindings
    Get-Item "0.0.0.0!443" | Remove-Item -Force -ErrorAction SilentlyContinue
    New-Item "0.0.0.0!443" -Thumbprint $thumbprint -SSLFlags 1
    Pop-Location
    Write-Log "Certifikat bundet till IIS-webbplatsen $siteName"

    Write-Log "SSL-certifikatinstallation slutförd för $domain"
}
catch {
    Write-Log "Fel inträffade: $_" "ERROR"
}
