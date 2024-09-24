# Requires -Modules Microsoft.Graph.Authentication
# Requires -Modules ExchangeOnlineManagement
# Requires -Modules PNP.Powershell

# Function to connect to Microsoft Graph
function Connect-MicrosoftGraph {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -ClientId $env:DigitalSupportAppID `
                    -TenantId $env:DigitalSupportTenantID `
                    -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
                    -NoWelcome | Out-Null
    Write-Host "Connected to Microsoft Graph."
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineService {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -AppId $env:DigitalSupportAppID `
                           -Organization "mhud.onmicrosoft.com" `
                           -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
                           -ShowBanner:$false
    Write-Host "Connected to Exchange Online."
}

# Function to connect to PnP PowerShell
function Connect-PnPPowerShell {
    Write-Host "Connecting to PnP PowerShell..."
    Connect-PnPOnline -Url "https://mhud.sharepoint.com" `
                      -ClientId $env:DigitalSupportAppID `
                      -Tenant 'mhud.onmicrosoft.com' `
                      -Thumbprint $env:DigitalSupportCertificateThumbprint
    Write-Host "Connected to PnP PowerShell."
}