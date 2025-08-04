<#
    New-OutlookEmail.ps1 - v2.0
    • Now checks for an active, open email window in Outlook.
    • If an email is open, attaches the file to it.
    • If no email is open, creates a new email with the file attached (original behavior).
#>

param (
    [string]$Article,
    [string]$TypeKey,
    [string[]]$Attachments
)

try {
    $Outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $MailItem = $null

    # --- CORE IMPROVEMENT ---
    # Check for a currently active (open) email window first.
    if ($Outlook.ActiveInspector()) {
        $MailItem = $Outlook.ActiveInspector().CurrentItem
    } 
    else {
        # If no email is open, create a new one.
        $MailItem = $Outlook.CreateItem(0) 
        $MailItem.Subject = "$Article $TypeKey"
    }
    # --- END IMPROVEMENT ---

    foreach ($file in $Attachments) {
        if (Test-Path $file) {
            $MailItem.Attachments.Add($file) | Out-Null
            Write-Host "✓ File attached to Outlook email" -ForegroundColor Green
        }
    }

    $MailItem.Display()
    Start-Sleep -Milliseconds 800
}
catch {
    Write-Warning "Could not create Outlook email: $($_.Exception.Message)"
}
