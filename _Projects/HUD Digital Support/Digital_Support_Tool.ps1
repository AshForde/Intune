function Show-Menu {
    # Define the menu items for each column
    $column1Items = @(
        "ENTRA ID"
        ""
        "  1.  Activate Entra PIM Role(s)"
        "  2.  Entra All Users report"
        "  3.  Entra Nested Security Groups report"
        "  4.  Check User Aho Assignment Status"
        "  5.  Change Username and Email Address of User"
        "  *6.  Add/Remove domain from external access (TBC)*"
        ""
        "EXCHANGE ONLINE"
        ""
        "  7.  Create a new shared mailbox" 
        "  8.  Get mailbox access report for user" 
        "  9.  Modify calendar permission/delegate access"
        "  10. Update approved senders on distribution lists"
        "  11. Remove calendar Events for user"
        "  12. Distribution list members report"
        ""
        "SHAREPOINT ONLINE"
        ""
        "  13. Get List Item Report"
        "  14. Generate Basic Site Report"
        "  15. Move between sites/libraries"
    )
    
    $column2Items = @(
        "INTUNE"
        ""
        "  17. Get All Apps and Group Assignments"
        "  18. Generate All Discovered Apps Report"
        ""
        "TEAMS"
        ""
        "  19. Get All Teams Owner and Members Report"
        "  20. Get Users Teams Access Report"
        ""
        "COMPLIANCE (PURVIEW)"
        ""
        "  21. Audit Report - User Activity"
        "  22. Audit Report - SPO Activity"
    
    
    
    
    
    )
    
    # Define a fixed width for the first column, enough to accommodate the longest line
    $column1Width = 60
    
    # Write the header
    Write-Host "## Digital Support Tool ##" -ForegroundColor Green
    Write-Host ""
    Write-Warning "Ensure you have the right permissions to run these commands"
    Write-Host ""
    
    # Print the menu items side by side
    for ($i = 0; $i -lt [Math]::Max($column1Items.Length, $column2Items.Length); $i++) {
        $column1Text = $column1Items[$i] -replace "`t","    " # replace tabs with spaces if needed
        $column2Text = $column2Items[$i] -replace "`t","    " # replace tabs with spaces if needed
    
        # Check if we have an item for the current index in each column
        if ($null -eq $column1Text) { $column1Text = "" }
        if ($null -eq $column2Text) { $column2Text = "" }
    
        # Determine if the item is a heading
        $isColumn1Heading = $column1Text -match "^\D+$" # Matches text that doesn't contain numbers
        $isColumn2Heading = $column2Text -match "^\D+$" # Matches text that doesn't contain numbers
    
        # Print the items with padding to align the columns
        $formattedColumn1Text = $column1Text.PadRight($column1Width)
        
        # Apply color to headings
        if ($isColumn1Heading) {
            Write-Host $formattedColumn1Text -NoNewline -ForegroundColor Cyan
        } else {
            Write-Host $formattedColumn1Text -NoNewline
        }
        
        if ($isColumn2Heading) {
            Write-Host $column2Text -ForegroundColor Cyan
        } else {
            Write-Host $column2Text
        }
    }
        Write-Host ""
        $option = Read-Host "Enter your number choice (or Q to exit)"
        return $option
    }
    
    function Get-TeamAccessReportForUser {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Teams Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Teams/Get-TeamAccessReportForUser.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    
    function Get-AllTeamMembersAndOwners {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Teams Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Teams/Get-AllTeamMembersAndOwners.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    
    
    function Get-SPOAuditUser {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Get-SPOUserAuditLog.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-DelegateAccess {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Get-CalendarDelegates.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-AppAssignments {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Intune Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Intune/Reporting/Report_App_Assignments.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-DiscoveredApps {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Intune Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Intune/Reporting/Report_Discovered_Apps.ps1"
    
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $scriptContent | Out-File -FilePath "C:\HUD\00_Staging\Report_Discovered_Apps.ps1" -Encoding UTF8
    
        $Platform = Read-Host "Enter the Platform value (Windows, AndroidWorkProfile, iOS)" 
    
        & C:\HUD\00_Staging\Report_Discovered_Apps.ps1 -Platform $Platform
    }
    function New-PIMSession {
        Clear-Host
        Write-Warning "Please ensure you have been granted access to your selected role before attempting to activate"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Enable_PIM_Assignments.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Export-AllUserReport {
        Clear-Host
      
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Reports/Entra_All_Users.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Export-NestedGroupReport {
        Clear-Host
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Reports/Entra_Nested_Security_Group_Report.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-EmployeeAssignment {
        Clear-Host
      
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Get-AhoEmployeeAssignment.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Update-AllowedDomains {
        Write-Host ""
        Write-host "This function is a work in progress" -ForegroundColor Red
        Write-Host ""
    }
    function Update-UserNameAndEmail {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: User Administrator & Exchange Recipient Administrator"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Update-Username.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    
    }
    function New-SharedMailbox {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: User Administrator & Exchange Recipient Administrator"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/New-SharedMailbox.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-UserMailboxAccess {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Check_Mailbox_Access.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Edit-DLApprovedSenders {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"
    
        start-sleep 3
    
        $DL = Read-Host "Please Enter Distribution List Name"
    
        $Action = Read-Host "Please specify 'Add', 'Remove', 'Review' [Default is 'Review']"
        
        if (-not $Action) {
            $Action = 'Review'
        }   
       
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Update_DL_Approved_Senders.ps1"
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $scriptBlock = [scriptblock]::Create($scriptContent)
        Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $DL, $Action
    
    }
    function Remove-CalendarEventsForUser {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Remove-CalenderEvents.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)   
    }
    function Get-DLGroupMember {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Reports/Distribution_Group_Member_Export.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
    }
    
    function Get-SPOListItemReport {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: SharePoint Administrator"
        Write-Warning "Alternatively you must have full permissions to the respective site to run this"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Get-SPListItemReport.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
    
    }
    
    function Get-SPOBasicSiteReport {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: SharePoint Administrator"
        Write-Warning "Alternatively you must have full permissions to the respective site to run this"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Get-BasicSiteReport.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
        
    }
    function Move-SPOFolders {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: SharePoint Administrator"
        Write-Warning "Alternatively you must have full permissions to the respective site to run this"
    
        start-sleep 3
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Move-Folders.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
        
    }
    function Get-UserActivityAuditReport {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Purview/Audit_User_Report.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    function Get-SPOActivityAuditReport {
        Clear-Host
        Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    
        start-sleep 3
    
        $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Purview/SPO_Activity_Report.ps1"
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    }
    #Select task based on Show-Menu function
    do
    {
        Clear-Host
        $selection = Show-Menu
        switch ($selection) {
                     '1'  {New-PIMSession}
                     '2'  {Export-AllUserReport}
                     '3'  {Export-NestedGroupReport}
                     '4'  {Get-EmployeeAssignment}
                     '5'  {Update-UserNameAndEmail}
                     '6'  {Update-AllowedDomains}
                     '7'  {New-SharedMailbox}
                     '8'  {Get-UserMailboxAccess}
                     '9'  {Get-DelegateAccess}
                     '10' {Edit-DLApprovedSenders}
                     '11' {Remove-CalendarEventsForUser}
                     '12' {Get-DLGroupMember}
                     '13' {Get-SPOListItemReport}
                     '14' {Get-SPOBasicSiteReport}
                     '15' {Move-SPOFolders}
                     '16' {"Please use option 21"}
                     '17' {Get-AppAssignments}
                     '18' {Get-DiscoveredApps}
                     '19' {Get-AllTeamMembersAndOwners}
                     '20' {Get-TeamAccessReportForUser}
                     '21' {Get-UserActivityAuditReport}
                     '22' {Get-SPOActivityAuditReport}
                     'q'  {return}
                     }
            pause
            }
    until ($selection -eq 'q')
    
    
    <#
     #New-HUDUser
     #New-SharedMailbox
     Search-MailboxAccess
     Add-PhoneNumber
    
     #>