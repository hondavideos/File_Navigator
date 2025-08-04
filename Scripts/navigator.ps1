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
        ArticleRegex = '(?<!\d)\d{6}(?:[A-Z]{1,2}|[/_][A-Z]\d{2})?(?!\d)'
        ArticleColour = 'Green'
    }
    STP_ZIPs = @{
        Path = Join-Path -Path $DataRoot -ChildPath 'STP_and_ZIPs'
        Filter = '*.stp', '*.zip'
        ArticleRegex = '.*' # Match everything
        ArticleColour = 'Yellow'
    }
}

$currentMode = 'PDFs'
$script:AllFiles = @{}

# Pre-load files for both modes
foreach ($modeName in $SearchModes.Keys) {
    $mode = $SearchModes[$modeName]
    if (Test-Path -Path $mode.Path) {
        $script:AllFiles[$modeName] = Get-ChildItem -Path $mode.Path -Recurse -Include $mode.Filter -File | Sort-Object Name
    }
    else {
        Write-Warning "Directory not found for mode '$modeName': $($mode.Path)"
        $script:AllFiles[$modeName] = @()
    }
}

# ── 3. Helpers ─────────────────────────────────────────────────────────────
function Parse-Parts {
    param([string]$Base, [string]$Regex)
    $m = [regex]::Match($Base, $Regex)
    $article = if ($m.Success) { $m.Value } else { '' }
    [PSCustomObject] @{ Article = $article; TypeKey = '' }
}

function Open-File {
    param([IO.FileInfo]$File)
    if (-not $File -or -not $File.FullName) { return }
    try { Invoke-Item -Path $File.FullName } catch { Write-Warning $_.Exception.Message }
}

function Select-File {
    param([IO.FileInfo]$File)
    if (-not $File -or -not $File.FullName) { return }
    $parts = Parse-Parts $File.BaseName $SearchModes[$currentMode].ArticleRegex
    $outlookScriptPath = Join-Path $ScriptRoot 'New-OutlookEmail.ps1'
    if (Test-Path $outlookScriptPath) {
        & $outlookScriptPath -Article $parts.Article -TypeKey $parts.TypeKey $File.FullName
    }
    else { Write-Warning "Could not find '$outlookScriptPath'." }
}

function Draw-Line {
    param($BaseName, $IsSel, $WinW, $ViewportRow, $Regex, $Colour)
    $fg0,$bg0 = [Console]::ForegroundColor,[Console]::BackgroundColor
    if ($IsSel) { $rowFg = 'Black'; $rowBg = 'White' } else { $rowFg = 'Gray'; $rowBg = 'Black' }

    $prefix = if ($ViewportRow -lt 10) { "($(($ViewportRow + 1) % 10)) " } else { '   ' }
    $maxText = $WinW - 1 - $prefix.Length
    $fileTxt = if ($BaseName.Length -gt $maxText) { $BaseName.Substring(0,$maxText-1)+'…' } else { $BaseName }
    
    Write-Host (' ' * ($WinW - 1)) -NoNewline -BackgroundColor $rowBg
    [Console]::CursorLeft = 0

    $numColor = if ($IsSel) { 'DarkBlue' } else { 'Yellow' }
    Write-Host -NoNewline $prefix -ForegroundColor $numColor -BackgroundColor $rowBg

    $m=[regex]::Match($fileTxt,$Regex)
    if ($m.Success) {
        Write-Host -NoNewline $fileTxt.Substring(0,$m.Index) -ForegroundColor $rowFg -BackgroundColor $rowBg
        $fgArt = if ($IsSel) { 'DarkRed' } else { $Colour }
        Write-Host -NoNewline $m.Value -ForegroundColor $fgArt -BackgroundColor $rowBg
        Write-Host -NoNewline $fileTxt.Substring($m.Index+$m.Length) -ForegroundColor $rowFg -BackgroundColor $rowBg
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
    $nextModeName = if ($currentModeName -eq 'PDFs') { 'STP/ZIP' } else { 'PDF' }
    $switchText = "Ctrl+T: Switch to $nextModeName"

    # Build status bar with color segments
    $prefix = "{0}/{1} | " -f ($cur+1), $tot
    $suffix = " | Ctrl+U/D: Page ½ | Ctrl+O: Open | Enter: Attach to Outlook Msg | Ctrl+Q: Quit"
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
        @($script:AllFiles[$currentMode].Where({$_.Name -like "*$script:searchTerm*"})) 
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
                    Draw-Line $script:Files[$fileIndex].BaseName ($fileIndex -eq $script:cur) $w $r $mode.ArticleRegex $mode.ArticleColour
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
        elseif ($vk -eq 27 -or ($isCtrl -and $vk -eq 67)) { exit } # Esc or Ctrl+C
        elseif ($vk -eq 8)  { if ($script:searchTerm.Length -gt 0) { $script:searchTerm = $script:searchTerm.Substring(0, $script:searchTerm.Length - 1); Update-SearchResults } else { $needsRedraw = $false } } # Backspace
        elseif ($isCtrl -and $vk -eq 81) { exit } # Ctrl+Q
        elseif ($isCtrl -and $vk -eq 84) { $currentMode = if ($currentMode -eq 'PDFs') { 'STP_ZIPs' } else { 'PDFs' }; Update-SearchResults } # Ctrl+T
        elseif ($isCtrl -and $vk -eq 79) { if ($script:Files) { Open-File $script:Files[$script:cur] }; $needsRedraw = $false } # Ctrl+O
        elseif ($isCtrl -and $vk -ge 48 -and $vk -le 57) { # Ctrl+0-9
            if ($script:Files) { $num = if ($vk -eq 48) { 9 } else { $vk - 49 }; $fileIdx = $script:top + $num; if ($fileIdx -lt $script:Files.Count) { Select-File $script:Files[$fileIdx] }}; $needsRedraw = $false
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
