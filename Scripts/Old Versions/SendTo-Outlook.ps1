# SendTo-Outlook.ps1 - HARDENED VERSION
#
# DESCRIPTION: Attaches selected files to a new Outlook email and intelligently sets the subject.
#              This version is hardened to ignore invalid arguments passed by some Windows shell configurations.

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$files
)

# Helper function to parse filename (no changes needed here)
function Get-SubjectParts([string]$fileName) {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fileName)
    $cleanName = ($baseName -replace '(?i)_EN.*$').Trim()
    $tokens = $cleanName -split '\s+'
    $article = $tokens | Where-Object { $_ -match '^\d{6}[A-Z]?$' } | Select-Object -First 1
    $typeKey = $tokens | Where-Object { $_ -match '-' } | Select-Object -First 1
    return @($article, $typeKey)
}

try {
    # --- HARDENED FILE VALIDATION ---
    # Filter the incoming arguments to get a list of only REAL files.
    # This ignores any junk arguments like a literal '%*'.
    $validFiles = $files | Where-Object { Test-Path $_ -PathType Leaf }

    if (-not $validFiles) {
        Write-Error "CRITICAL: No valid files were found in the arguments provided."
        Start-Sleep -Seconds 15
        exit 1
    }

    # Use the first valid file for all subject logic.
    $firstFile = $validFiles[0]

    # --- MAIN SCRIPT ---
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    # --- SUBJECT BUILDER ---
    # Parse the FIRST VALID file to generate the email subject.
    $article, $typeKey = Get-SubjectParts -fileName $firstFile

    if ($article -and $typeKey) { $mail.Subject = "$article - $typeKey" }
    elseif ($article) { $mail.Subject = $article }
    elseif ($typeKey) { $mail.Subject = $typeKey }
    else { $mail.Subject = [IO.Path]::GetFileNameWithoutExtension($firstFile) }

    # Add a note if there are multiple files.
    if ($validFiles.Count -gt 1) {
        $otherFilesCount = $validFiles.Count - 1
        $mail.Subject += " (+ $otherFilesCount more)"
    }

    # Attach every VALID file to the email.
    foreach ($file in $validFiles) {
        $mail.Attachments.Add($file)
    }

    $mail.Display()
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Host "This window will close in 20 seconds..."
    Start-Sleep -Seconds 20
}