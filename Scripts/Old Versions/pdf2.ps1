<#
    pdf_nav.ps1 – v3.0g  (centred list + centred status bar)
#>

[CmdletBinding()] param(
    [Parameter(Position=0)] [string]$SearchTerm
)

# ── 1. Settings ────────────────────────────────────────────────────────────
$Root          = 'C:\Users\Jack Belmore\Documents\Ziehl\Crossbase Data Sheets\PDF'
$SearchPattern = '*.pdf'
$ArticleColour = 'Green'
$NumberColour  = 'Yellow'

# article highlight: 6 digits + (1–2 letters  OR  /_H01 style)
$ArticleRegex  = '(?<!\d)\d{6}(?:[A-Z]{1,2}|[/_][A-Z]\d{2})?(?!\d)'

# ── 2. Load PDF list ───────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
    Write-Error "Root path '$Root' not found."; exit 1
}
$AllFiles = Get-ChildItem -Path $Root -Recurse -Filter $SearchPattern -File |
            Sort-Object Name
if (-not $AllFiles) { Write-Host "No PDFs found."; exit }

# ── 3. Helpers ─────────────────────────────────────────────────────────────
function Parse-Parts {
    param([string]$Base)

    # remove localisation tails
    $clean = $Base -replace '(?i)(_belmore_en|_en1|_en|_gb)$'

    # capture 6-digit article + optional direct / separated suffix
    $m = [regex]::Match(
        $clean,
        '(?<!\d)(?<art>\d{6})(?<dir>[A-Z]{1,2})?(?:[_/-](?<suf>[A-Z]\d{2}))?'
    )

    $article = ''
    if ($m.Success) {
        $article = $m.Groups['art'].Value
        if     ($m.Groups['suf'].Success) { $article += "/$($m.Groups['suf'].Value)" }
        elseif ($m.Groups['dir'].Success) { $article += $m.Groups['dir'].Value }
        $clean = $clean.Remove($m.Index, $m.Length)
    }

    $tm = [regex]::Match($clean, '(?<!\d)(?<tk>[A-Z]{2,}[A-Z0-9.\-]*)')
    $typeKey = if ($tm.Success) { $tm.Groups['tk'].Value } else { '' }

    [PSCustomObject]@{ Article = $article; TypeKey = $typeKey }
}

function Select-File {
    param([IO.FileInfo]$File)

    # ---- open Explorer and pre-select the file ---------------------------
    # (quotes are doubled so paths with spaces work)
    $quoted = '"' + $File.FullName + '"'
    Start-Process explorer.exe "/select, $quoted"

    # ---- existing parse/display logic ------------------------------------
    $parts = Parse-Parts $File.BaseName
    Clear-Host
    Write-Host "`nChosen file:`n`t$($File.FullName)`n"
    Write-Host "--- Parts Captured ---" -ForegroundColor Yellow
    Write-Host "Article: $($parts.Article)"
    Write-Host "Type Key: $($parts.TypeKey)"
    Write-Host "----------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "`n(File Explorer is open. Drag the file into Outlook or elsewhere."
    Write-Host "Press any key to return to the list.)"

    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Draw-Line {
    param($BaseName,$IsSel,$WinW,$RowIdx)

    # ---------- choose row colours ----------
    $fg0,$bg0 = [Console]::ForegroundColor,[Console]::BackgroundColor
    switch ($true) {
        { $IsSel }          { $rowFg='Black'; $rowBg='White' }
        { $RowIdx % 2 -eq 1}{ $rowFg='Gray' ; $rowBg='DarkGray' }
        default             { $rowFg=$fg0   ; $rowBg=$bg0 }
    }

    # ---------- build display string ----------
    $prefix = if ($RowIdx -lt 10){ "$(($RowIdx+1)%10) - " } else { '   ' }
    $maxText = $WinW - $prefix.Length
    $text    = if ($BaseName.Length -gt $maxText) { $BaseName.Substring(0,$maxText-1)+'…' } else { $BaseName }
    $display = $prefix + $text

    # ---------- centring ----------
    $leftPadLen = [Math]::Max(0, ($WinW - $display.Length) / 2)
    $leftPad    = ' ' * [int]$leftPadLen

    # wipe row
    Write-Host (' ' * $WinW) -NoNewline -BackgroundColor $rowBg
    [Console]::CursorLeft = $leftPadLen

    # write prefix (always plain colour)
    Write-Host -NoNewline $prefix -ForegroundColor $NumberColour -BackgroundColor $rowBg

    # highlight article substring inside $text
    $m=[regex]::Match($text,$script:ArticleRegex)
    if ($m.Success) {
        Write-Host -NoNewline $text.Substring(0,$m.Index) -ForegroundColor $rowFg -BackgroundColor $rowBg
        $fgArt = if ($IsSel) { 'DarkRed' } else { $script:ArticleColour }
        Write-Host -NoNewline $m.Value                     -ForegroundColor $fgArt -BackgroundColor $rowBg
        Write-Host -NoNewline $text.Substring($m.Index+$m.Length) -ForegroundColor $rowFg -BackgroundColor $rowBg
    } else {
        Write-Host -NoNewline $text -ForegroundColor $rowFg -BackgroundColor $rowBg
    }

    Write-Host
    [Console]::ForegroundColor,[Console]::BackgroundColor=$fg0,$bg0
}

function Write-StatusBar {
    param($cur,$tot,$winW)
    $fg0,$bg0=[Console]::ForegroundColor,[Console]::BackgroundColor
    [Console]::BackgroundColor='DarkCyan'
    $txt="{0}/{1} – 1-0 select • s search • u/d page • j/k ↑/↓ • Enter • q" -f ($cur+1),$tot
    $pad=' ' * [int][Math]::Max(0,($winW - $txt.Length)/2)
    Write-Host ($pad+$txt).PadRight($winW) -ForegroundColor White -BackgroundColor DarkCyan
    [Console]::ForegroundColor,[Console]::BackgroundColor=$fg0,$bg0
}

# ── 4. Search / UI loop ────────────────────────────────────────────────────
$firstRun = $true
:Outer while ($true) {
    if ($firstRun) { $firstRun = $false }
    else           { $SearchTerm = Read-Host 'Enter new search term (blank = all)' }

    $Files = if ($SearchTerm) { @($AllFiles.Where({ $_.Name -like "*$SearchTerm*" })) } else { $AllFiles }
    if (-not $Files) {
        Write-Host 'No hits. Press any key…'; $null=$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); continue
    }

    $cur,$top=0,0
    try{ [Console]::CursorVisible=$false }catch{}
    :UI while ($true) {
        $w=$Host.UI.RawUI.WindowSize.Width; $h=$Host.UI.RawUI.WindowSize.Height; $rows=$h-1
        if     ($cur -lt $top)       {$top=$cur}
        elseif ($cur -ge $top+$rows) {$top=$cur-$rows+1}

        Clear-Host
        for ($r=0; $r -lt $rows; $r++) {
            $i=$top+$r; if ($i -ge $Files.Count) { break }
            Draw-Line $Files[$i].BaseName ($i -eq $cur) $w $r
        }
        Write-StatusBar $cur $Files.Count $w

        $k=$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        $num=$null
        if     ($k.VirtualKeyCode -ge 49 -and $k.VirtualKeyCode -le 57){$num=$k.VirtualKeyCode-49}
        elseif ($k.VirtualKeyCode -eq 48){$num=9}
        elseif ($k.VirtualKeyCode -ge 97 -and $k.VirtualKeyCode -le 105){$num=$k.VirtualKeyCode-97}
        elseif ($k.VirtualKeyCode -eq 96){$num=9}

        if ($num -ne $null) {
            $idx=$top+$num; if ($idx -lt $Files.Count) { Select-File $Files[$idx] }
        } else {
            switch ($k.VirtualKeyCode) {
                74{if($cur -lt $Files.Count-1){$cur++}}  40{if($cur -lt $Files.Count-1){$cur++}}  # j / ↓
                75{if($cur -gt 0){$cur--}}              38{if($cur -gt 0){$cur--}}               # k / ↑
                68{$cur=[Math]::Min($Files.Count-1,$cur+($h-1))}   # d page-down
                85{$cur=[Math]::Max(0,$cur-($h-1))}               # u page-up
                13{Select-File $Files[$cur]}                      # Enter
                83{break UI}                                      # s re-search
                81{exit} 27{exit} 3{exit}                         # q / Esc / Ctrl-C
            }
        }
    }
} finally { try{[Console]::CursorVisible=$true}catch{} }

