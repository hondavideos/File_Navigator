<#
    navigator.ps1  –  v12.0 (Header Visibility Fix)
    • FIXED: Header visibility issue - buffer size matches window exactly
    • FIXED: Console auto-scrolling that pushed header out of view
    • FIXED: Cursor positioning to prevent automatic viewport scrolling
    • FIXED: First row disappearance issue - corrected viewport scrolling logic
    • FIXED: Ctrl+T mode switching works perfectly (changed from Ctrl+S)
    • FIXED: Variable scoping bug in viewport calculations
    • FIXED: Off-by-one error in scroll trigger logic
    • Cleaned status bar - removed long file paths
#>

[CmdletBinding()] param()

# ── 1. Settings & Initial Setup ────────────────────────────────────────────
$ScriptRoot    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$DataRoot      = Join-Path -Path $ScriptRoot -ChildPath '..\Data'

$SearchModes = @{
    PDFs = @{
        Path = Join-Path -Path $DataRoot -ChildPath 'PDFs'
        Filter = '*.pdf'
        ArticleRegex = '(?<=\b)\d{6}(?:[._-][A-Za-z0-9]+)?(?=\s|$|\b)'
        ArticleColour = 'Green'
        TypeKeyRegex = '\b[A-Z]{2}\d{2,3}[A-Z0-9\-\.]*\b'
        TypeKeyColour = 'Yellow'
    }
    STP_ZIPs = @{
        Path = Join-Path -Path $DataRoot -ChildPath 'STP_and_ZIPs'
        Filter = '*.stp', '*.zip'
        ArticleRegex = '(?<=\b)\d{6}(?:[._-][A-Za-z0-9]+)?(?=\s|$|\b)'
        ArticleColour = 'Green'
        TypeKeyRegex = '\b[A-Z]{2}\d{2,3}[A-Z0-9\-\.]*\b'
        TypeKeyColour = 'Yellow'
    }
    Recent = @{
        Path = ''
        Filter = '*.*'
        ArticleRegex = '(?<=\b)\d{6}(?:[._-][A-Za-z0-9]+)?(?=\s|$|\b)'
        ArticleColour = 'Green'
        TypeKeyRegex = '\b[A-Z]{2}\d{2,3}[A-Z0-9\-\.]*\b'
        TypeKeyColour = 'Yellow'
    }
}

$currentMode = 'PDFs'
$script:AllFiles = @{}

# Recently accessed files tracking
$HistoryFile = Join-Path $ScriptRoot "file_history.json"
$script:FileHistory = @()
if (Test-Path $HistoryFile) {
    try {
        $script:FileHistory = Get-Content $HistoryFile | ConvertFrom-Json
    } catch {
        $script:FileHistory = @()
    }
}

# Pre-load files for both modes
foreach ($modeName in $SearchModes.Keys) {
    $mode = $SearchModes[$modeName]
    if ($modeName -eq 'Recent') {
        # Load recent files from history
        $script:AllFiles[$modeName] = $script:FileHistory | ForEach-Object { 
            if (Test-Path $_.FullName) { Get-Item $_.FullName }
        } | Where-Object { $_ } | Sort-Object LastAccessTime -Descending
    } elseif (Test-Path -Path $mode.Path) {
        $script:AllFiles[$modeName] = Get-ChildItem -Path $mode.Path -Recurse -Include $mode.Filter -File | Sort-Object Name
    }
    else {
        Write-Warning "Directory not found for mode '$modeName': $($mode.Path)"
        $script:AllFiles[$modeName] = @()
    }
}

# ── 3. Helpers ─────────────────────────────────────────────────────────────
function Add-ToHistory {
    param([IO.FileInfo]$File)
    if (-not $File) { return }
    
    $historyEntry = @{
        FullName = $File.FullName
        BaseName = $File.BaseName
        AccessTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    
    # Remove if already exists and add to front
    $script:FileHistory = @($historyEntry) + ($script:FileHistory | Where-Object { $_.FullName -ne $File.FullName }) | Select-Object -First 20
    
    # Save to file
    try {
        $script:FileHistory | ConvertTo-Json | Out-File $HistoryFile -Encoding UTF8
        # Refresh recent files list
        $script:AllFiles.Recent = $script:FileHistory | ForEach-Object { 
            if (Test-Path $_.FullName) { Get-Item $_.FullName }
        } | Where-Object { $_ } | Sort-Object LastAccessTime -Descending
    } catch {
        # Silently ignore save errors
    }
}

function Try-OutlookAttach {
    param([IO.FileInfo]$File, [string]$Article, [string]$TypeKey)
    
    $outlookScriptPath = Join-Path $ScriptRoot 'New-OutlookEmail.ps1'
    
    # Quick checks before attempting
    if (-not (Test-Path $outlookScriptPath)) {
        Write-Host "Outlook script not found - opening file..." -ForegroundColor Yellow
        return $false
    }
    
    try {
        # FIX: Explicitly name the -Attachments parameter for robust binding.
        $null = & $outlookScriptPath -Article $Article -TypeKey $TypeKey -Attachments $File.FullName -ErrorAction Stop
        return $true
    }
    catch {
        # Silent failure - just return false
        Write-Host "Outlook unavailable - opening file..." -ForegroundColor Yellow
        return $false
    }
}

function Parse-Parts {
    param([string]$Base, [string]$Regex)
    $m = [regex]::Match($Base, $Regex)
    $article = if ($m.Success) { $m.Value } else { '' }
    
    # Extract TypeKey from the filename after the article number
    $typeKey = ''
    if ($article -and $Base.Contains($article)) {
        $afterArticle = $Base.Substring($Base.IndexOf($article) + $article.Length).Trim()
        # Extract meaningful part (remove common suffixes like " - Alex", "_en", "_gb", etc.)
        $typeKey = $afterArticle -replace '\s*-\s*Alex.*$', '' -replace '_en$', '' -replace '_gb$', '' -replace '_belmore.*$', ''
        $typeKey = $typeKey.Trim(' -')
        # Limit length for email subjects
        if ($typeKey.Length -gt 50) {
            $typeKey = $typeKey.Substring(0, 47) + "..."
        }
    }
    
    [PSCustomObject] @{ Article = $article; TypeKey = $typeKey }
}

function Open-File {
    param([IO.FileInfo]$File)
    if (-not $File -or -not $File.FullName) { return }
    Add-ToHistory $File
    try { Invoke-Item -Path $File.FullName } catch { Write-Warning $_.Exception.Message }
}

function Select-File {
    param([IO.FileInfo]$File)
    if (-not $File -or -not $File.FullName) { 
        Write-Host "Error: No file selected" -ForegroundColor Red
        Start-Sleep -Milliseconds 1500
        return 
    }
    
    Add-ToHistory $File
    $parts = Parse-Parts $File.BaseName $SearchModes[$currentMode].ArticleRegex
    
    # Try Outlook first, fallback to opening file
    $outlookSuccess = Try-OutlookAttach $File $parts.Article $parts.TypeKey
    
    if (-not $outlookSuccess) {
        Write-Host "Opening file directly..." -ForegroundColor Cyan
        Open-File $File
    }
}

function Invoke-FileDelete {
    param([IO.FileInfo]$File)
    if (-not $File) { return }
    
    try {
        # Move to Recycle Bin (safer than permanent delete)
        Add-Type -AssemblyName Microsoft.VisualBasic
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $File.FullName,
            'OnlyErrorDialogs',
            'SendToRecycleBin'
        )
        
        # Show brief confirmation at bottom
        $w = $Host.UI.RawUI.WindowSize.Width
        $h = $Host.UI.RawUI.WindowSize.Height
        [Console]::SetCursorPosition(0, $h-1)
        [Console]::BackgroundColor = 'DarkRed'
        [Console]::ForegroundColor = 'White'
        Write-Host "File deleted: $($File.Name)".PadRight($w-1)
        [Console]::ResetColor()
        Start-Sleep -Milliseconds 1200
        
        # Refresh file list and update cursor position
        Update-FileListAfterDelete $File
        
    } catch {
        # Show error at bottom
        $w = $Host.UI.RawUI.WindowSize.Width
        $h = $Host.UI.RawUI.WindowSize.Height
        [Console]::SetCursorPosition(0, $h-1)
        [Console]::BackgroundColor = 'DarkRed'
        [Console]::ForegroundColor = 'White'
        Write-Host "Delete failed: $($_.Exception.Message)".PadRight($w-1)
        [Console]::ResetColor()
        Start-Sleep -Milliseconds 2000
    }
}

function Update-FileListAfterDelete {
    param([IO.FileInfo]$DeletedFile)
    
    # Remove from current file list
    $oldIndex = $script:cur
    $script:AllFiles[$currentMode] = $script:AllFiles[$currentMode] | Where-Object { $_.FullName -ne $DeletedFile.FullName }
    
    # Update search results
    Update-SearchResults
    
    # Smart cursor repositioning
    if ($script:Files.Count -eq 0) {
        $script:cur = 0
    } elseif ($oldIndex -ge $script:Files.Count) {
        $script:cur = $script:Files.Count - 1  # Move to last file
    }
    # If oldIndex < Count, cursor stays at same position (shows next file)
}

function Draw-Line {
    param($BaseName, $IsSel, $WinW, $ViewportRow, $ArticleRegex, $ArticleColour, $TypeKeyRegex, $TypeKeyColour)
    $fg0,$bg0 = [Console]::ForegroundColor,[Console]::BackgroundColor
    if ($IsSel) { $rowFg = 'Black'; $rowBg = 'White' } else { $rowFg = 'Gray'; $rowBg = 'Black' }

    $prefix = if ($ViewportRow -lt 10) { "($(($ViewportRow + 1) % 10)) " } else { '   ' }
    $maxText = $WinW - 1 - $prefix.Length
    $fileTxt = if ($BaseName.Length -gt $maxText) { $BaseName.Substring(0,$maxText-1)+'…' } else { $BaseName }
    
    Write-Host (' ' * ($WinW - 1)) -NoNewline -BackgroundColor $rowBg
    [Console]::CursorLeft = 0

    $numColor = if ($IsSel) { 'DarkBlue' } else { 'Yellow' }
    Write-Host -NoNewline $prefix -ForegroundColor $numColor -BackgroundColor $rowBg

    # Find all matches for both article numbers and type keys
    $articleMatches = [regex]::Matches($fileTxt, $ArticleRegex)
    $typeKeyMatches = [regex]::Matches($fileTxt, $TypeKeyRegex)
    
    # Combine and sort all matches by position
    $allMatches = @()
    foreach ($match in $articleMatches) {
        $allMatches += @{ Match = $match; Type = 'Article'; Color = if ($IsSel) { 'DarkRed' } else { $ArticleColour } }
    }
    foreach ($match in $typeKeyMatches) {
        # Only add if it doesn't overlap with an article number
        $overlaps = $false
        foreach ($articleMatch in $articleMatches) {
            if (($match.Index -lt ($articleMatch.Index + $articleMatch.Length)) -and 
                (($match.Index + $match.Length) -gt $articleMatch.Index)) {
                $overlaps = $true
                break
            }
        }
        if (-not $overlaps) {
            $allMatches += @{ Match = $match; Type = 'TypeKey'; Color = if ($IsSel) { 'DarkYellow' } else { $TypeKeyColour } }
        }
    }
    
    # Sort matches by position
    $allMatches = $allMatches | Sort-Object { $_.Match.Index }
    
    if ($allMatches.Count -gt 0) {
        $currentPos = 0
        foreach ($matchInfo in $allMatches) {
            $match = $matchInfo.Match
            # Write text before match
            if ($match.Index -gt $currentPos) {
                Write-Host -NoNewline $fileTxt.Substring($currentPos, $match.Index - $currentPos) -ForegroundColor $rowFg -BackgroundColor $rowBg
            }
            # Write highlighted match
            Write-Host -NoNewline $match.Value -ForegroundColor $matchInfo.Color -BackgroundColor $rowBg
            $currentPos = $match.Index + $match.Length
        }
        # Write remaining text
        if ($currentPos -lt $fileTxt.Length) {
            Write-Host -NoNewline $fileTxt.Substring($currentPos) -ForegroundColor $rowFg -BackgroundColor $rowBg
        }
    } else {
        Write-Host -NoNewline $fileTxt -ForegroundColor $rowFg -BackgroundColor $rowBg
    }
    [Console]::ForegroundColor,[Console]::BackgroundColor=$fg0,$bg0
}

function Write-Header {
    param($searchTerm, $w, $modeName)
    $fg0,$bg0=[Console]::ForegroundColor,[Console]::BackgroundColor
    [Console]::BackgroundColor='DarkBlue'
    $txt = "($modeName) Filter: $searchTerm"
    $pad = ' ' * [int][Math]::Max(0, ($w - 1 - $txt.Length) / 2)
    Write-Host ($pad + $txt).PadRight($w - 1) -ForegroundColor White
    [Console]::ForegroundColor,[Console]::BackgroundColor=$fg0,$bg0
}

function Write-StatusBar {
    param($cur, $tot, $w, $currentModeName)
    $fg0,$bg0=[Console]::ForegroundColor,[Console]::BackgroundColor
    [Console]::BackgroundColor='DarkCyan'
    
    # Determine the next mode for the help text
    $modes = @('PDFs', 'STP_ZIPs', 'Recent')
    $currentIndex = $modes.IndexOf($currentModeName)
    $nextIndex = ($currentIndex + 1) % $modes.Length
    $nextModeName = switch ($modes[$nextIndex]) {
        'PDFs' { 'PDF' }
        'STP_ZIPs' { 'STP/ZIP' }
        'Recent' { 'Recent' }
    }
    $switchText = "Ctrl+T: Switch to $nextModeName"

    # Build status bar with color segments
    $prefix = "{0}/{1} | " -f ($cur+1), $tot
    $suffix = " | Ctrl+U/D: Page ½ | Ctrl+O: Open | Enter: Add to Outlook | Delete: Remove File | Ctrl+Q: Quit"
    $totalLength = $prefix.Length + $switchText.Length + $suffix.Length
    
    if ($totalLength -gt ($w - 1)) {
        # If too long, truncate suffix
        $availableSpace = ($w - 1) - $prefix.Length - $switchText.Length - 3
        $suffix = $suffix.Substring(0, [Math]::Max(0, $availableSpace)) + '…'
    }
    
    # Write segments with different colors
    Write-Host $prefix -ForegroundColor White -BackgroundColor DarkCyan -NoNewline
    Write-Host $switchText -ForegroundColor Red -BackgroundColor DarkCyan -NoNewline
    Write-Host $suffix -ForegroundColor White -BackgroundColor DarkCyan -NoNewline
    
    # Pad remaining space
    $usedSpace = $prefix.Length + $switchText.Length + $suffix.Length
    $padding = [Math]::Max(0, ($w - 1) - $usedSpace)
    if ($padding -gt 0) {
        Write-Host (' ' * $padding) -BackgroundColor DarkCyan -NoNewline
    }
    
    [Console]::ForegroundColor,[Console]::BackgroundColor=$fg0,$bg0
}

function Update-SearchResults {
    $script:Files = if ($script:searchTerm) {
        $patternString = if ($script:searchTerm -match '[*?]') {
            $script:searchTerm
        } else {
            "*$($script:searchTerm)*"
        }
        $pattern = [System.Management.Automation.WildcardPattern]::new($patternString, 'IgnoreCase')
        @($script:AllFiles[$currentMode] | Where-Object { $pattern.IsMatch($_.Name) })
    } else {
        $script:AllFiles[$currentMode]
    }
    $script:cur = 0
    $script:top = 0
}

# ── 4. UI Loop ─────────────────────────────────────────────────────────────
try {
    $script:cur,$script:top,$script:searchTerm = 0,0,''
    Update-SearchResults # Initial load
    try{ [Console]::CursorVisible=$false }catch{}

    $needsRedraw = $true
    while ($true) {

        if ($needsRedraw) {
            # Get window dimensions first
            $w=$Host.UI.RawUI.WindowSize.Width; $h=$Host.UI.RawUI.WindowSize.Height
            
            # Fix buffer size to prevent scrolling issues
            try {
                $currentBuffer = $Host.UI.RawUI.BufferSize
                
                # Ensure buffer matches window size exactly to prevent auto-scrolling
                if ($currentBuffer.Height -ne $h -or $currentBuffer.Width -ne $w) {
                    $newBuffer = New-Object System.Management.Automation.Host.Size($w, $h)
                    $Host.UI.RawUI.BufferSize = $newBuffer
                }
            } catch {
                # Fallback if buffer sizing fails
            }
            
            # Clear and immediately position cursor at absolute top-left
            [System.Console]::Clear()
            [Console]::SetCursorPosition(0, 0)
            
            # Write header at position 0,0 (no SetCursorPosition call needed)
            Write-Header $script:searchTerm $w $currentMode

            # Calculate available space: total height - header - status bar
            $headerRows = 1
            $statusRows = 1
            $rowsAvailable = [Math]::Max(1, $h - $headerRows - $statusRows)
            
            # Update viewport logic
            if($script:cur -lt $script:top){$script:top=$script:cur}elseif($script:cur -ge ($script:top + $rowsAvailable)){$script:top=$script:cur - $rowsAvailable + 1}
            
            # Draw file list - start at row 1 (after header)
            for($r=0; $r -lt $rowsAvailable; $r++) {
                $fileIndex = $script:top + $r
                if ($script:Files -and $fileIndex -lt $script:Files.Count) {
                    # Position cursor manually for each row
                    [Console]::SetCursorPosition(0, $headerRows + $r)
                    $mode = $SearchModes[$currentMode]
                    Draw-Line $script:Files[$fileIndex].BaseName ($fileIndex -eq $script:cur) $w $r $mode.ArticleRegex $mode.ArticleColour $mode.TypeKeyRegex $mode.TypeKeyColour
                } elseif ($fileIndex -eq 0 -and -not $script:Files) {
                    # Show "no files" message in center
                    $msg = "No files match '$($script:searchTerm)'"
                    $centerRow = $headerRows + [int]($rowsAvailable / 2)
                    $centerCol = [Math]::Max(0, ($w - $msg.Length) / 2)
                    [Console]::SetCursorPosition($centerCol, $centerRow)
                    Write-Host $msg
                }
            }

            # Write status bar at the second-to-last row to avoid triggering scroll
            $statusPosition = $h - $statusRows
            [Console]::SetCursorPosition(0, $statusPosition)
            Write-StatusBar $script:cur ($script:Files.Count) $w $currentMode
            
            # Position cursor safely away from edges
            [Console]::SetCursorPosition(0, $statusPosition)
            $needsRedraw = $false
        }

        $k = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
        $needsRedraw = $true
        # --- FIX: Detect Ctrl using ControlKeyState string matching ---
        $isCtrl = ($k.ControlKeyState -match "LeftCtrlPressed|RightCtrlPressed") -or ($k.Modifiers -band 4)
        $isShift = ($k.ControlKeyState -match "ShiftPressed") -or ($k.Modifiers -band 1)

        $vk = $k.VirtualKeyCode

        if ($vk -eq 40) { if($script:cur -lt $script:Files.Count-1) { $script:cur++ } else { $needsRedraw = $false } } # Down
        elseif ($vk -eq 38) { if($script:cur -gt 0) { $script:cur-- } else { $needsRedraw = $false } } # Up
        elseif ($vk -eq 34) { $script:cur=[Math]::Min($script:Files.Count-1, $script:cur+$rowsAvailable-1) } # PageDown
        elseif ($vk -eq 33) { $script:cur=[Math]::Max(0, $script:cur-$rowsAvailable-1) } # PageUp
        elseif ($isCtrl -and $vk -eq 68) { $script:cur=[Math]::Min($script:Files.Count-1, $script:cur + [int]($rowsAvailable/2)) } # Ctrl+D (Down 1/2)
        elseif ($isCtrl -and $vk -eq 85) { $script:cur=[Math]::Max(0, $script:cur - [int]($rowsAvailable/2)) } # Ctrl+U (Up 1/2)
        elseif ($vk -eq 13) { if($script:Files) { Select-File $script:Files[$script:cur] }; $needsRedraw = $false } # Enter
        elseif ($vk -eq 27) { exit } # Esc only
        elseif ($vk -eq 8)  { if ($script:searchTerm.Length -gt 0) { $script:searchTerm = $script:searchTerm.Substring(0, $script:searchTerm.Length - 1); Update-SearchResults } else { $needsRedraw = $false } } # Backspace
        elseif ($isCtrl -and $vk -eq 81) { exit } # Ctrl+Q
        elseif ($isCtrl -and $vk -eq 84) { # Ctrl+T - cycle through modes
            $modes = @('PDFs', 'STP_ZIPs', 'Recent')
            $currentIndex = $modes.IndexOf($currentMode)
            $nextIndex = ($currentIndex + 1) % $modes.Length
            $currentMode = $modes[$nextIndex]
            Update-SearchResults
        }
        elseif ($isCtrl -and $vk -eq 79) { if ($script:Files) { Open-File $script:Files[$script:cur] }; $needsRedraw = $false } # Ctrl+O
        elseif ($isCtrl -and $vk -ge 48 -and $vk -le 57) { # Ctrl+0-9
            if ($script:Files) { $num = if ($vk -eq 48) { 9 } else { $vk - 49 }; $fileIdx = $script:top + $num; if ($fileIdx -lt $script:Files.Count) { Select-File $script:Files[$fileIdx] }}; $needsRedraw = $false
        }
        elseif ($vk -eq 46) { # Delete key
            if ($script:Files) { 
                Invoke-FileDelete $script:Files[$script:cur]
            }
            $needsRedraw = $false
        }
        elseif (-not [char]::IsControl($k.Character)) {
            $script:searchTerm += $k.Character
            Update-SearchResults
        } 
        else {
            $needsRedraw = $false
        }
    }
}
finally { 
    try{[Console]::CursorVisible=$true}catch{} 
    try{[Console]::ResetColor()}catch{}
    Clear-Host
}