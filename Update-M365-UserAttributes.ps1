<#
.SYNOPSIS
    Microsoft 365 User Attribute Management Script

.DESCRIPTION
    Interactive script to manage user attributes in Microsoft 365 (UPN, Email, etc.)
    
.NOTES
    Version: 1.8
#>

# Load Windows Forms for file dialogs
Add-Type -AssemblyName System.Windows.Forms

# Global Variables
$Global:ConnectionStatus = $false
$Global:TenantDomain = ""

# Function to display banner
function Show-Banner {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "   Microsoft 365 User Attribute Management Tool" -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Global:ConnectionStatus) {
        Write-Host "Connection Status: " -NoNewline -ForegroundColor White
        Write-Host "CONNECTED" -ForegroundColor Green
        Write-Host "Tenant Domain: " -NoNewline -ForegroundColor White
        Write-Host "$Global:TenantDomain" -ForegroundColor Yellow
    } else {
        Write-Host "Connection Status: " -NoNewline -ForegroundColor White
        Write-Host "DISCONNECTED" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Function to display menu
function Show-Menu {
    Write-Host "Please select an option:" -ForegroundColor White
    Write-Host ""
    Write-Host "  1" -NoNewline -ForegroundColor Yellow
    Write-Host " - Connect to Exchange Online / Microsoft 365" -ForegroundColor White
    Write-Host ""
    Write-Host "  2" -NoNewline -ForegroundColor Yellow
    Write-Host " - Export current mailbox users to CSV" -ForegroundColor White
    Write-Host ""
    Write-Host "  3" -NoNewline -ForegroundColor Yellow
    Write-Host " - Run in Test Mode (Preview changes only)" -ForegroundColor White
    Write-Host ""
    Write-Host "  4" -NoNewline -ForegroundColor Yellow
    Write-Host " - Run in Change Mode (Apply changes)" -ForegroundColor White
    Write-Host ""
    Write-Host "  5" -NoNewline -ForegroundColor Yellow
    Write-Host " - Disconnect from Exchange Online / Microsoft 365" -ForegroundColor White
    Write-Host ""
    Write-Host "  Q" -NoNewline -ForegroundColor Yellow
    Write-Host " - Quit" -ForegroundColor White
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Function to check and install Exchange Online Module
function Install-ExchangeOnlineModule {
    Write-Host "Checking for Exchange Online Management Module..." -ForegroundColor White
    
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "Exchange Online Management Module is already installed." -ForegroundColor Green
        return $true
    } else {
        Write-Host "Exchange Online Management Module is not installed." -ForegroundColor Yellow
        Write-Host "Installing Exchange Online Management Module..." -ForegroundColor White
        
        try {
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host "Exchange Online Management Module installed successfully." -ForegroundColor Green
            return $true
        } catch {
            Write-Host "ERROR: Failed to install Exchange Online Management Module." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            return $false
        }
    }
}

# Function to check if already connected
function Test-ExistingConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineSession {
    Write-Host ""
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Connecting to Exchange Online" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    # Check if already connected
    if (Test-ExistingConnection) {
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            $Global:TenantDomain = $orgConfig.Name
            $Global:ConnectionStatus = $true
            
            Write-Host "Already connected to Exchange Online!" -ForegroundColor Green
            Write-Host "Tenant: $Global:TenantDomain" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press any key to return to menu..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        } catch {
            # Continue with new connection if check fails
        }
    }
    
    # Check and install module
    if (-not (Install-ExchangeOnlineModule)) {
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }
    
    Write-Host "Initiating authentication..." -ForegroundColor White
    Write-Host "Please sign in with your Microsoft 365 administrator credentials." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        
        # Get tenant domain
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        $Global:TenantDomain = $orgConfig.Name
        $Global:ConnectionStatus = $true
        
        Write-Host ""
        Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
        Write-Host "Tenant: $Global:TenantDomain" -ForegroundColor Yellow
        
    } catch {
        Write-Host ""
        Write-Host "ERROR: Failed to connect to Exchange Online." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        $Global:ConnectionStatus = $false
    }
    
    Write-Host ""
    Write-Host "Press any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to validate connection
function Test-ExchangeConnection {
    if (-not $Global:ConnectionStatus) {
        Write-Host ""
        Write-Host "ERROR: Not connected to Exchange Online." -ForegroundColor Red
        Write-Host "Please select Option 1 to connect first." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $false
    }
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    } catch {
        Write-Host ""
        Write-Host "ERROR: Connection lost. Please reconnect." -ForegroundColor Red
        $Global:ConnectionStatus = $false
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $false
    }
}

# Function to show Save File Dialog
function Get-SaveFilePath {
    param(
        [string]$DefaultFileName = "users.csv",
        [string]$Title = "Save CSV File"
    )
    
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $SaveFileDialog.Title = $Title
    $SaveFileDialog.FileName = $DefaultFileName
    $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $SaveFileDialog.FileName
    }
    return $null
}

# Function to show Open File Dialog
function Get-OpenFilePath {
    param(
        [string]$Title = "Select CSV File"
    )
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    }
    return $null
}

# Function to export mailboxes to CSV
function Export-MailboxUsers {
    Write-Host ""
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Export Mailbox Users to CSV" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not (Test-ExchangeConnection)) {
        return
    }
    
    Write-Host "Please select the location to save the CSV file..." -ForegroundColor White
    Write-Host ""
    
    $csvPath = Get-SaveFilePath -DefaultFileName "M365_Users_Export.csv" -Title "Save Mailbox Users Export"
    
    if ([string]::IsNullOrWhiteSpace($csvPath)) {
        Write-Host "Export cancelled." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }
    
    Write-Host ""
    Write-Host "Retrieving mailboxes from Exchange Online..." -ForegroundColor White
    
    try {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | 
            Where-Object { $_.DisplayName -notlike "*Discovery Search Mailbox*" -and $_.RecipientTypeDetails -ne "DiscoveryMailbox" }
        
        Write-Host "Found $($mailboxes.Count) mailboxes (excluding Discovery mailboxes)." -ForegroundColor Green
        Write-Host "Retrieving user details..." -ForegroundColor White
        
        $exportData = @()
        $counter = 0
        
        foreach ($mailbox in $mailboxes) {
            $counter++
            Write-Progress -Activity "Processing mailboxes" -Status "Processing $counter of $($mailboxes.Count)" -PercentComplete (($counter / $mailboxes.Count) * 100)
            
            # Get user details for FirstName and LastName
            try {
                $user = Get-User -Identity $mailbox.UserPrincipalName -ErrorAction Stop
                $firstName = if ($user.FirstName) { $user.FirstName } else { "" }
                $lastName = if ($user.LastName) { $user.LastName } else { "" }
            } catch {
                $firstName = ""
                $lastName = ""
            }
            
            $exportData += [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                FirstName = $firstName
                LastName = $lastName
                MailNickname = $mailbox.Alias
                UserPrincipalName = $mailbox.UserPrincipalName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress.ToString()
            }
        }
        
        Write-Progress -Activity "Processing mailboxes" -Completed
        
        Write-Host "Exporting to CSV..." -ForegroundColor White
        # Use semicolon delimiter for European locales
        $exportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        
        Write-Host ""
        Write-Host "Export completed successfully!" -ForegroundColor Green
        Write-Host "File saved to: $csvPath" -ForegroundColor Yellow
        Write-Host "Total users exported: $($exportData.Count)" -ForegroundColor Green
        Write-Host "CSV Format: Semicolon-delimited (;)" -ForegroundColor Cyan
        
    } catch {
        Write-Host ""
        Write-Host "ERROR: Failed to export mailboxes." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Press any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to process users in Test or Change mode
function Process-Users {
    param(
        [bool]$TestMode
    )
    
    $modeText = if ($TestMode) { "Test Mode" } else { "Change Mode" }
    
    Write-Host ""
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Run in $modeText" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not (Test-ExchangeConnection)) {
        return
    }
    
    # Ask about domain change
    Write-Host "Do you want to change the domain for users?" -ForegroundColor White
    Write-Host "Enter new domain (e.g., newdomain.com) or press Enter to skip:" -ForegroundColor Gray
    Write-Host ""
    $newDomain = Read-Host "New Domain"
    
    $changeDomain = -not [string]::IsNullOrWhiteSpace($newDomain)
    
    if ($changeDomain) {
        Write-Host ""
        Write-Host "Domain will be changed to: $newDomain" -ForegroundColor Yellow
    } else {
        Write-Host ""
        Write-Host "Domain will remain unchanged." -ForegroundColor Yellow
    }
    
    # Ask for CSV path using file dialog
    Write-Host ""
    Write-Host "Please select the CSV file to import..." -ForegroundColor White
    Write-Host ""
    
    $csvPath = Get-OpenFilePath -Title "Select Mailbox Users CSV File"
    
    if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path $csvPath)) {
        Write-Host "Import cancelled or file not found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }
    
    Write-Host ""
    Write-Host "Importing CSV file..." -ForegroundColor White
    
    try {
        # Try semicolon delimiter first (European format)
        $users = Import-Csv -Path $csvPath -Delimiter ";" -ErrorAction Stop
        Write-Host "Imported $($users.Count) users from CSV (semicolon-delimited)." -ForegroundColor Green
        
        # Verify we have the required columns
        if ($users.Count -gt 0) {
            $firstUser = $users[0]
            $hasUPN = $firstUser.PSObject.Properties.Name -contains "UserPrincipalName"
            
            if (-not $hasUPN) {
                # Try comma delimiter as fallback
                Write-Host "Semicolon format failed, trying comma-delimited format..." -ForegroundColor Yellow
                $users = Import-Csv -Path $csvPath -Delimiter "," -ErrorAction Stop
                Write-Host "Imported $($users.Count) users from CSV (comma-delimited)." -ForegroundColor Green
            }
        }
        
        # Show detected columns
        if ($users.Count -gt 0) {
            $firstUser = $users[0]
            Write-Host "CSV Columns: $($firstUser.PSObject.Properties.Name -join ', ')" -ForegroundColor Cyan
        }
        
    } catch {
        Write-Host ""
        Write-Host "ERROR: Failed to import CSV file." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host ""
        Write-Host "Press any key to return to menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }
    
    Write-Host ""
    Write-Host "Processing users..." -ForegroundColor White
    Write-Host ""
    
    $results = @()
    $errorList = @()
    $counter = 0
    
    foreach ($csvUser in $users) {
        $counter++
        Write-Progress -Activity "Processing users in $modeText" -Status "Processing $counter of $($users.Count)" -PercentComplete (($counter / $users.Count) * 100)
        
        try {
            # Get current mailbox
            $mailbox = Get-Mailbox -Identity $csvUser.UserPrincipalName -ErrorAction Stop
            
            # Get FirstName and LastName from CSV
            $firstName = ""
            $lastName = ""
            
            if ($csvUser.FirstName -and ![string]::IsNullOrWhiteSpace($csvUser.FirstName)) {
                $firstName = $csvUser.FirstName.ToString().Trim()
            }
            
            if ($csvUser.LastName -and ![string]::IsNullOrWhiteSpace($csvUser.LastName)) {
                $lastName = $csvUser.LastName.ToString().Trim()
            }
            
            # Check if we have valid FirstName and LastName
            if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) {
                $errorList += [PSCustomObject]@{
                    DisplayName = $mailbox.DisplayName
                    MailNickname = $mailbox.Alias
                    UserPrincipalName = $mailbox.UserPrincipalName.ToString()
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress.ToString()
                    Error = "Missing FirstName ('$firstName') or LastName ('$lastName') in CSV"
                }
                continue
            }
            
            $newMailNickname = "$($firstName).$($lastName)".ToLower()
            
            # Determine domain
            $domain = if ($changeDomain) {
                $newDomain
            } else {
                $mailbox.PrimarySmtpAddress.ToString().Split('@')[1]
            }
            
            $newPrimarySmtp = "$newMailNickname@$domain"
            $newUPN = "$newMailNickname@$domain"
            
            # Get old email addresses - only SMTP addresses
            $oldEmailAddresses = @()
            foreach ($email in $mailbox.EmailAddresses) {
                $emailString = $email.ToString()
                if ($emailString -like "smtp:*" -or $emailString -like "SMTP:*") {
                    $oldEmailAddresses += $emailString
                }
            }
            
            $newEmailAddresses = @("SMTP:$newPrimarySmtp")
            
            foreach ($oldEmail in $oldEmailAddresses) {
                $emailAddress = $oldEmail -replace "^smtp:", "" -replace "^SMTP:", ""
                if ($emailAddress -ne $newPrimarySmtp) {
                    $newEmailAddresses += "smtp:$emailAddress"
                }
            }
            
            # Apply changes if not in test mode
            if (-not $TestMode) {
                try {
                    # Update email addresses (suppress SPO warnings)
                    Set-Mailbox -Identity $mailbox.UserPrincipalName `
                        -EmailAddresses $newEmailAddresses `
                        -WarningAction SilentlyContinue `
                        -ErrorAction Stop
                    
                    # Update alias
                    Set-Mailbox -Identity $mailbox.UserPrincipalName `
                        -Alias $newMailNickname `
                        -WarningAction SilentlyContinue `
                        -ErrorAction Stop
                    
                    # Update primary SMTP
                    Set-Mailbox -Identity $mailbox.UserPrincipalName `
                        -WindowsEmailAddress $newPrimarySmtp `
                        -WarningAction SilentlyContinue `
                        -ErrorAction Stop
                    
                    # Update UPN using Set-User (better for UPN changes)
                    Set-User -Identity $mailbox.UserPrincipalName `
                        -UserPrincipalName $newUPN `
                        -WarningAction SilentlyContinue `
                        -ErrorAction Stop
                        
                } catch {
                    $errorList += [PSCustomObject]@{
                        DisplayName = $mailbox.DisplayName
                        MailNickname = $mailbox.Alias
                        UserPrincipalName = $mailbox.UserPrincipalName.ToString()
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress.ToString()
                        Error = $_.Exception.Message
                    }
                    continue
                }
            }
            
            # Add to results
            $results += [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                MailNickname = $newMailNickname
                UserPrincipalName = $newUPN
                PrimarySmtpAddress = $newPrimarySmtp
            }
            
        } catch {
            $errorList += [PSCustomObject]@{
                DisplayName = if ($csvUser.DisplayName) { $csvUser.DisplayName } else { "Unknown" }
                MailNickname = if ($csvUser.MailNickname) { $csvUser.MailNickname } else { "Unknown" }
                UserPrincipalName = if ($csvUser.UserPrincipalName) { $csvUser.UserPrincipalName } else { "Unknown" }
                PrimarySmtpAddress = if ($csvUser.PrimarySmtpAddress) { $csvUser.PrimarySmtpAddress } else { "Unknown" }
                Error = $_.Exception.Message
            }
        }
    }
    
    Write-Progress -Activity "Processing users in $modeText" -Completed
    
    # Display results
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Results Overview" -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($results.Count -gt 0) {
        Write-Host "Successfully processed users:" -ForegroundColor Green
        Write-Host ""
        $results | Format-Table -Property DisplayName, MailNickname, UserPrincipalName, PrimarySmtpAddress -AutoSize
    }
    
    # Display errors if any
    if ($errorList.Count -gt 0) {
        Write-Host ""
        Write-Host "------------------------------------------------------------" -ForegroundColor Red
        Write-Host "  Errors Encountered ($($errorList.Count) users)" -ForegroundColor Red
        Write-Host "------------------------------------------------------------" -ForegroundColor Red
        Write-Host ""
        
        foreach ($err in $errorList) {
            Write-Host "FAILED: $($err.DisplayName)" -ForegroundColor Red
            Write-Host "  MailNickname: $($err.MailNickname)" -ForegroundColor Yellow
            Write-Host "  UserPrincipalName: $($err.UserPrincipalName)" -ForegroundColor Yellow
            Write-Host "  PrimarySmtpAddress: $($err.PrimarySmtpAddress)" -ForegroundColor Yellow
            Write-Host "  Error: $($err.Error)" -ForegroundColor Red
            Write-Host ""
        }
    }
    
    Write-Host ""
    Write-Host "Processing completed." -ForegroundColor Green
    Write-Host "Successful: $($results.Count) | Failed: $($errorList.Count)" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to disconnect from Exchange Online
function Disconnect-ExchangeOnlineSession {
    Write-Host ""
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Disconnecting from Exchange Online" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    if (-not $Global:ConnectionStatus) {
        Write-Host "Not currently connected." -ForegroundColor Yellow
    } else {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            $Global:ConnectionStatus = $false
            $Global:TenantDomain = ""
            Write-Host "Successfully disconnected from Exchange Online." -ForegroundColor Green
        } catch {
            Write-Host "ERROR: Failed to disconnect." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Press any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main script loop
do {
    Show-Banner
    Show-Menu
    
    $choice = Read-Host "Enter your choice"
    
    switch ($choice.ToUpper()) {
        "1" { Connect-ExchangeOnlineSession }
        "2" { Export-MailboxUsers }
        "3" { Process-Users -TestMode $true }
        "4" { Process-Users -TestMode $false }
        "5" { Disconnect-ExchangeOnlineSession }
        "Q" { 
            Write-Host ""
            Write-Host "Exiting script..." -ForegroundColor Yellow
            if ($Global:ConnectionStatus) {
                Write-Host "Disconnecting from Exchange Online..." -ForegroundColor White
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            Write-Host "Goodbye!" -ForegroundColor Green
            Write-Host ""
            break
        }
        default {
            Write-Host ""
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
    
} while ($choice.ToUpper() -ne "Q")