#______________________________________________________________________________________________________________________________________________________
# Requires -Modules ExchangeOnlineManagement
# Requires -Modules Microsoft.Graph.Authentication
# Requires -Modules PNP.Powershell
# Requires -Modules MicrosoftTeams
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Microsoft Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected to Graph" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token        = $CollectToken.RequestMessage.Headers.Authorization.Parameter
    $Token | Out-Null
    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Exchange Online
try{
    Connect-ExchangeOnline `
        -AppId $env:DigitalSupportAppID `
        -Organization "mhud.onmicrosoft.com" `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -ShowBanner:$false
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Exchange Online. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

#______________________________________________________________________________________________________________________________________________________

# Function to connect to SharePoint Online
if ($PSVersionTable.PSVersion -gt [Version]"7.0") {
    try {
        # Disable PnP PowerShell update check
        $env:PNPPOWERSHELL_UPDATECHECK = "Off"
        
        Connect-PnPOnline `
            -Url "https://mhud.sharepoint.com" `
            -ClientId $env:DigitalSupportAppID `
            -Tenant 'mhud.onmicrosoft.com' `
            -Thumbprint $env:DigitalSupportCertificateThumbprint
        Write-Host "Connected to PnP PowerShell."  -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to PNP SharePoint Online. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "PowerShell version is below 7.2. Skipping PnP PowerShell connection." -ForegroundColor Yellow
}
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Microsoft Teams
try {
    Connect-MicrosoftTeams `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -ApplicationId $env:DigitalSupportAppID | Out-Null
    Write-Host "Connected to Microsoft Teams." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Microsoft Teams. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}