Clear-Host
Write-Host '## EntraID User Export ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication

# Connect to Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token = $CollectToken.RequestMessage.Headers.Authorization.Parameter

        
    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

# Selected Values
$select = @(
    'id'
    'givenName'
    'surname'
    'displayName'
    'userPrincipalName'
    'mail'
    'userType'
    'accountEnabled'
    'jobTitle'
    'department'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup'
    'company'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'
    'manager'
    'OfficeLocation'
    'streetAddress'
    'City'
    'postalCode'
    'state'
    'country'
    'businessPhones'
    'mobilePhone'
    'createdDateTime'
    'signInActivity'
    'signInActivity'
    'usageLocation'
    'passwordPolicies'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime'
    'assignedLicenses'
    'SecurityIdentifier'
    'extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox'
    'extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox'
) -join ','

# Graph API Call
$uri = "https://graph.microsoft.com/v1.0/users?`$select=$select&`$expand=manager"
$headers = @{
        "Authorization" = $Token
        "Content-Type"  = "application/json"
}

# Results
$output = @()

do {
    $req = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $uri = $req.'@odata.nextLink'

    foreach ($user in $req.value) {
        $output += [PSCustomObject]@{
            # Identity
            'ID'                                     = $user.id
            'FirstName'                              = $user.givenName
            'LastName'                               = $user.surname
            'Display Name'                           = $user.displayName
            'User Principal Name'                    = $user.userPrincipalName
            'Email'                                  = $user.mail
            'User Type'                              = $user.userType
            'Account Enabled'                        = $user.accountEnabled

            # Organisational Structure
            'Job Title'                              = $user.jobTitle
            'Department'                             = $user.department
            'Organisational Group'                   = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
            'Organisation'                           = $user.company
            'Employee Type'                          = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType
            'Employee Category'                      = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory
            'Manager Display Name'                   = $user.manager.displayName
            'Manager UPN'                            = $user.manager.userPrincipalName
            'Manager Job Title'                      = $user.manager.jobTitle

            # Contact and Location
            'Office'                                 = $user.OfficeLocation
            'Address'                                = $user.streetAddress
            'City'                                   = $user.City
            'Postal Code'                            = $user.postalCode
            'State'                                  = $user.state
            'Country'                                = $user.country
            'Phone'                                  = $user.businessPhones -join ',' 
            'Mobile'                                 = $user.mobilePhone 
             
             # Account
            'Created Date Time'                      = $user.createdDateTime
            'Last Interactive Sign In Date Time'     = $user.signInActivity.lastSignInDateTime
            'Last Non-Interactive Sign In Date Time' = $user.signInActivity.lastNonInteractiveSignInDateTime
            'Usage Location'                         = $user.usageLocation
            'Password Policies'                      = $user.passwordPolicies
            'Start Date'                             = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate
            'Leave Date'                             = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime
            'E5 License'                             = if ($user.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06") { $true } else { $false }
            'No Licenses'                            = if ($user.assignedLicenses.count -eq 0) { $true } else { $false }
                        
             # Other
            'Security Identifier'                    = $user.SecurityIdentifier
            'Room Mailbox'                           = $User.extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox
            'Shared Mailbox'                         = $User.extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox
        }
    }
} while ($uri)

# Output the user details
#$output

Write-Host "Open Save Dialog"

$Date     = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Entra All Users Export"

# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog  
$SaveFileDialog.Filter   = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title    = "Save as"
$SaveFileDialog.FileName = $FileName

# Show the SaveFileDialog and get the selected file path  
$SaveFileResult = $SaveFileDialog.ShowDialog()

if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
				$SelectedFilePath = $SaveFileDialog.FileName
    $output | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
    
    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet    = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow    = 1
    $endRow      = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn   = $worksheet.Dimension.End.Column

    # Set horizontal alignment to left for all cells in the used range
    for ($col = $startColumn; $col -le $endColumn; $col++) {
        for ($row = $startRow; $row -le $endRow; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        }
    }

    # Autosize columns if needed
    foreach ($column in $worksheet.Dimension.Start.Column.$worksheet.Dimension.End.Column) {
        $worksheet.Column($column).AutoFit()
        }
    
    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

    Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green

} else {
Write-Host "Save cancelled" -ForegroundColor Yellow
}
