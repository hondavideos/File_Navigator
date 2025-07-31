<#
    navigator.ps1  –  v8.3 (Definitive Edition)
    • Implements a robust, flicker-free screen drawing method to fix all UI glitches.
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
    $maxText = $WinW - $prefix.Length
    $fileTxt = if ($BaseName.Length -gt $maxText) { $BaseName.Substring(0,$maxText-1)+'…' } else { $BaseName }
    
    Write-Host (' ' * $WinW) -NoNewline -BackgroundColor $rowBg
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
    $pad = ' ' * [int][Math]::Max(0, ($w - $txt.Length) / 2)
    Write-Host ($pad + $txt).PadRight($w) -ForegroundColor White
    [Console]::ForegroundColor,$bg0=[Console]::ForegroundColor,$bg0
}

function Write-StatusBar {
    param($cur, $tot, $w, $fullPath)
    $fg0,$bg0=[Console]::ForegroundColor,[Console]::BackgroundColor
    [Console]::BackgroundColor='DarkCyan'
    $pathText = if ($fullPath) { " | $fullPath"} else { "" }
    $txt="{0}/{1}$pathText | Ctrl+S: Switch | Ctrl+O: Open | Enter: Attach | q: Quit" -f ($cur+1),$tot
    if ($txt.Length -gt $w) { $txt = $txt.Substring(0, $w-1) + '…' }
    Write-Host $txt.PadRight($w) -ForegroundColor White -BackgroundColor DarkCyan
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
            $w=$Host.UI.RawUI.WindowSize.Width; $h=$Host.UI.RawUI.WindowSize.Height
            
            # --- FIX: Robust, flicker-free screen clearing and drawing ---
            [Console]::SetCursorPosition(0, 0)
            $blankLine = ' ' * $w

            $headerRows = 0
            if ($script:searchTerm) {
                Write-Header $script:searchTerm $w $currentMode
                $headerRows = 1
            } else {
                # Manually clear header area if no search term
                Write-Host $blankLine
            }

            $rowsAvailable = $h - $headerRows - 1
            if($script:cur -lt $script:top){$script:top=$script:cur}elseif($script:cur -ge $top+$rowsAvailable){$script:top=$cur-$rowsAvailable+1}
            
            $linesDrawn = 0
            if ($script:Files) {
                for($r=0;$r -lt $rowsAvailable;$r++){
                    $i=$script:top+$r; if($i -ge $script:Files.Count){break}
                    $mode = $SearchModes[$currentMode]
                    [Console]::SetCursorPosition(0, $headerRows + $r)
                    Draw-Line $script:Files[$i].BaseName ($i -eq $script:cur) $w $r $mode.ArticleRegex $mode.ArticleColour
                    $linesDrawn++
                }
            }
            
            for ($r = $linesDrawn; $r -lt $rowsAvailable; $r++) {
                [Console]::SetCursorPosition(0, $headerRows + $r)
                Write-Host $blankLine
            }

            $selectedFullPath = if ($script:Files) { $script:Files[$script:cur].FullName } else { '' }
            [Console]::SetCursorPosition(0, $h - 1)
            Write-StatusBar $script:cur ($script:Files.Count) $w $selectedFullPath
            $needsRedraw = $false
        }

        $k = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        if ($k.VirtualKeyCode -in 16, 17, 18) { continue } # Ignore modifier-only keypresses
        
        $needsRedraw = $true
        $isCtrl = ($k.ControlKeyState -band 8)

        switch($k.VirtualKeyCode){
            40 { if($script:cur -lt $script:Files.Count-1) { $script:cur++ } else { $needsRedraw = $false } } # Down
            38 { if($script:cur -gt 0) { $script:cur-- } else { $needsRedraw = $false } } # Up
            34 { $script:cur=[Math]::Min($script:Files.Count-1, $script:cur+$rowsAvailable-1) } # PageDown
            33 { $script:cur=[Math]::Max(0, $script:cur-$rowsAvailable-1) } # PageUp
            13 { if($script:Files) { Select-File $script:Files[$script:cur] }; $needsRedraw = $false } # Enter
            79 { if ($isCtrl -and $script:Files) { Open-File $script:Files[$script:cur] }; $needsRedraw = $false } # Ctrl+O
            83 { if ($isCtrl) { $currentMode = if ($currentMode -eq 'PDFs') { 'STP_ZIPs' } else { 'PDFs' }; Update-SearchResults } } # Ctrl+S
            81 { exit }
            27 { exit }
            3  { exit }
            8  { if ($script:searchTerm.Length -gt 0) { $script:searchTerm = $script:searchTerm.Substring(0, $script:searchTerm.Length - 1); Update-SearchResults } else { $needsRedraw = $false } } # Backspace
            {$_ -ge 48 -and $_ -le 57} { # 0-9 keys
                if ($isCtrl) { if ($script:Files) { $num = if ($_.VirtualKeyCode -eq 48) { 9 } else { $_.VirtualKeyCode - 49 }; $fileIdx = $script:top + $num; if ($fileIdx -lt $script:Files.Count) { Select-File $script:Files[$fileIdx] }}; $needsRedraw = $false } 
                else { $script:searchTerm += $k.Character; Update-SearchResults }
            }
            default {
                if (-not [char]::IsControl($k.Character)) { $script:searchTerm += $k.Character; Update-SearchResults } 
                else { $needsRedraw = $false }
            }
        }
    }
}
finally { 
    try{[Console]::CursorVisible=$true}catch{} 
    try{[Console]::ResetColor()}catch{}
    Clear-Host
}
