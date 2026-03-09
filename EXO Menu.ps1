# 1. Environment Preparation (Exchange Online Only)
Write-Host "Configuring environment..." -ForegroundColor Cyan
Write-Host "Version 1.5" -ForegroundColor Yellow


if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -Force -Confirm:$false -Scope CurrentUser
}
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

# 2. Connect
Connect-ExchangeOnline

# Helper to find User Identity
function Get-ExoUserIdentity {
    param($InputString)
    $user = Get-EXOMailbox -Identity $InputString -ErrorAction SilentlyContinue
    if ($null -eq $user) {
        $user = Get-EXOMailbox -Filter "DisplayName -eq '$InputString' -or Alias -eq '$InputString'" -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    return $user.UserPrincipalName
}

# 3. Interactive Menu
do {
    Clear-Host
    Write-Host "--- Exchange Online Management Menu ---" -ForegroundColor Cyan
    Write-Host "1. List all Mailboxes"
    Write-Host "2. Get Mailbox Statistics"
    Write-Host "3. List Shared Mailboxes"
    Write-Host "4. Check Mailbox Permissions"
    Write-Host "5. Display ACTIVE Inbox Rules"
    Write-Host "6. Display HIDDEN Inbox Rules"
    Write-Host "7. Display EXCHANGE LICENSES (via SKUAssigned)"
    Write-Host "8. Manual Change Shell (Set-Mailbox)"
    Write-Host "9. Enable Archiving (Single or ALL)"
    Write-Host "Q. Disconnect and Quit"
    
    $choice = Read-Host "`nSelect an option"
    
    if ($choice -eq 'q') { 
        Disconnect-ExchangeOnline -Confirm:$false
        Clear-Host
        break 
    }

    # Options that don't require a specific user input first
    if ($choice -eq '1' -or $choice -eq '3') {
        switch ($choice) {
            '1' { Get-EXOMailbox | Select-Object DisplayName, UserPrincipalName | Out-GridView }
            '3' { Get-EXOMailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName, PrimarySmtpAddress | FT }
        }
        Read-Host "`nPress Enter to return to menu..."
        continue
    }

    # Special handling for Archiving (Option 9) to allow "All" selection
    if ($choice -eq '9') {
        Write-Host "`nArchiving Options:" -ForegroundColor Yellow
        Write-Host "1. Enable for Single Mailbox"
        Write-Host "2. Enable for ALL User Mailboxes (where not already enabled)"
        $archChoice = Read-Host "Select sub-option"

        if ($archChoice -eq '1') {
            $rawInput = Read-Host "Enter User (UPN, Display Name, or Alias)"
            $upn = Get-ExoUserIdentity -InputString $rawInput
            if ($upn) {
                Enable-Mailbox -Identity $upn -Archive
                Write-Host "Archive enabled for $upn" -ForegroundColor Green
            } else {
                Write-Host "User not found!" -ForegroundColor Red
            }
        }
        elseif ($archChoice -eq '2') {
            $confirm = Read-Host "Are you sure you want to enable archiving for ALL users? (y/n)"
            if ($confirm -eq 'y') {
                Write-Host "Processing... this may take a moment." -ForegroundColor Cyan
                # Filters for users who do not already have an ArchiveGuid assigned
                Get-EXOMailbox -RecipientTypeDetails UserMailbox -Filter 'ArchiveGuid -eq "00000000-0000-0000-0000-000000000000"' | Enable-Mailbox -Archive
                Write-Host "Archiving enabled for all applicable user mailboxes." -ForegroundColor Green
            }
        }
        Read-Host "`nPress Enter to return to menu..."
        continue
    }

    # Standard User-specific options
    $rawInput = Read-Host "Enter User (UPN, Display Name, or Alias)"
    $upn = Get-ExoUserIdentity -InputString $rawInput

    if (-not $upn) { 
        Write-Host "User not found!" -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue 
    }

    switch ($choice) {
        '2' { Get-EXOMailboxStatistics -Identity $upn | Select-Object DisplayName, TotalItemSize, LastLogonTime | FT }
        '4' { Get-EXOMailboxPermission -Identity $upn | Where-Object { $_.User -notlike "NT AUTHORITY\SELF" } | FT }
        '5' { Get-InboxRule -Mailbox $upn | Select-Object Name, Enabled, Priority | FT }
        '6' { Get-InboxRule -Mailbox $upn -IncludeHidden | Select-Object Name, Enabled, Priority | FT }
        '7' { 
            $box = Get-EXOMailbox -Identity $upn -Property SKUAssigned
            if ($box.SKUAssigned) {
                Write-Host "License Status: Active" -ForegroundColor Green
                $box | Select-Object DisplayName, SKUAssigned | FT
            } else {
                Write-Host "No Exchange license detected for this user." -ForegroundColor Yellow
            }
        }
        '8' { 
            Write-Host "Manual Mode for $upn. Type 'exit' to return." -ForegroundColor Magenta
            while ($true) {
                $cmd = Read-Host "EXO-Shell ($upn)"
                if ($cmd -eq "exit") { break }
                try { Invoke-Expression $cmd } catch { Write-Host $_ -ForegroundColor Red }
            }
        }
    }
    
    Read-Host "`nPress Enter to return to menu..."
} while ($true)
