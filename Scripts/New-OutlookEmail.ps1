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
    $Inspector = $Outlook.ActiveInspector() # Get the inspector once

    # --- FINAL FIX ---
    # Check if there is an active window AND if that window contains a MailItem.
    # The '.Class' property of a MailItem is 43 (olMail).
    if ($Inspector -and $Inspector.CurrentItem.Class -eq 43) {
        # Use the existing, open email because it's a valid mail item.
        $MailItem = $Inspector.CurrentItem
    }
    else {
        # If no window is active, or the active item is NOT an email (e.g., a calendar item),
        # create a new email from scratch to avoid errors.
        $MailItem = $Outlook.CreateItem(0) # 0 is olMailItem
    }
    
    # Set subject line if we have article info and email doesn't have a subject
    if ($Article -and (-not $MailItem.Subject -or $MailItem.Subject.Trim() -eq '')) {
        if ($TypeKey -and $TypeKey.Trim() -ne '') {
            $MailItem.Subject = "$Article - $TypeKey"
        } else {
            $MailItem.Subject = $Article
        }
        # Set email subject (silent)
    }
    # --- END IMPROVEMENT ---

    foreach ($file in $Attachments) {
        if (Test-Path $file) {
            $MailItem.Attachments.Add($file) | Out-Null
            # File attached to Outlook email (silent)
        }
    }

    $MailItem.Display()
    Start-Sleep -Milliseconds 800
}
catch {
    Write-Host "ERROR: Could not create Outlook email: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure Outlook is installed and running." -ForegroundColor Yellow
    # IMPROVEMENT: Re-throw the exception as a terminating error
    # This allows the calling script (navigator.ps1) to catch the failure.
    throw $_
}
finally {
    # Release all COM objects meticulously to prevent memory leaks and orphaned processes.
    # This is the key to fixing the intermittent "Array index out of bounds" error.
    if ($MailItem)  { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MailItem) | Out-Null }
    if ($Inspector) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Inspector) | Out-Null }
    if ($Outlook)   { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }

    # Force immediate garbage collection to ensure all references are cleared.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
