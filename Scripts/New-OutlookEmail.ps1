# New-OutlookEmail.ps1
#
# DESCRIPTION: Attaches a file to a new Outlook email using a pre-defined subject.

param(
    # The Article Number, passed from the navigator script
    [string]$Article,

    # The Type Key, passed from the navigator script (optional)
    [string]$TypeKey,

    # The full path to the file to attach
    [Parameter(Mandatory=$true, ValueFromRemainingArguments=$true)]
    [string]$FilePath
)

try {
    # Validate the file path exists
    if (-not (Test-Path $FilePath -PathType Leaf)) {
        throw "File not found: $FilePath"
    }

    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    # --- NEW SUBJECT BUILDER ---
    # Uses the accurate parts passed in from the navigator script.
    if ($Article -and $TypeKey) {
        $mail.Subject = "$Article - $TypeKey"
    }
    elseif ($Article) {
        $mail.Subject = $Article
    }
    else {
        # Fallback if for some reason no article was passed
        $mail.Subject = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    }

    # Attach the file.
    $mail.Attachments.Add($FilePath)

    $mail.Display()
}
catch {
    Write-Error "An error occurred in New-OutlookEmail.ps1: $($_.Exception.Message)"
    Write-Host "This window will close in 20 seconds..."
    Start-Sleep -Seconds 20
}