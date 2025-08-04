# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

File Navigator is a PowerShell-based file management tool that provides a fast, keyboard-driven interface for searching and managing files in the `Data` directory. The tool allows users to search for PDF and STP/ZIP files, attach them to Outlook emails, or open them directly.

## Architecture

### Core Components

- **navigator.ps1**: Main application script that provides the interactive file browser interface
- **New-OutlookEmail.ps1**: Utility script that creates or attaches files to Outlook emails
- **Data/**: Directory containing organized files:
  - `PDFs/`: Contains PDF technical documents and datasheets
  - `STP_and_ZIPs/`: Contains STP and ZIP files (directory may need to be created)

### Key Features

- **Dual Search Modes**: Toggle between PDF search and STP/ZIP search using `Ctrl+S`
- **Live Search**: Real-time filtering as you type
- **Article Number Extraction**: Uses regex patterns to identify and highlight article numbers in filenames
- **Outlook Integration**: Automatically creates emails with file attachments or adds to existing open emails
- **Keyboard Navigation**: Complete keyboard-driven interface with hotkeys

### Technical Details

#### Search Modes Configuration
```powershell
$SearchModes = @{
    PDFs = @{
        Path = 'Data/PDFs'
        Filter = '*.pdf'
        ArticleRegex = '(?<!\d)\d{6}(?:[A-Z]{1,2}|[/_][A-Z]\d{2})?(?!\d)'
        ArticleColour = 'Green'
    }
    STP_ZIPs = @{
        Path = 'Data/STP_and_ZIPs'  
        Filter = '*.stp', '*.zip'
        ArticleRegex = '.*'
        ArticleColour = 'Yellow'
    }
}
```

#### Article Number Pattern
The PDF mode uses a specific regex pattern to identify article numbers: 6 digits optionally followed by 1-2 letters or underscore/slash + letter + 2 digits (e.g., "123456", "123456AB", "123456_H01").

## Commands

### Running the Application
```powershell
# Navigate to Scripts directory and run:
.\navigator.ps1

# If execution policy error occurs, run as administrator:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Recent Fixes (v10.0)
- **FIXED**: First row disappearance issue that was causing display problems
- **FIXED**: Unwanted scrolling when typing search terms
- **Implemented**: "Total screen paint" approach using `[System.Console]::Clear()` for reliable rendering
- **Improved**: Sequential top-to-bottom rendering eliminates cursor position artifacts
- **Simplified**: Draw-Line function removes complex cursor manipulation

### Keyboard Controls
- **Live Search**: Start typing to filter files
- **Ctrl+S**: Toggle between PDF and STP/ZIP search modes
- **Up/Down Arrows**: Navigate file list
- **Enter**: Attach selected file to Outlook email
- **Ctrl+O**: Open selected file in default application
- **Ctrl+1-0**: Quick select files 1-10 from visible list
- **Shift+Q**: Quit application
- **Backspace**: Remove characters from search term

### Development Commands
```powershell
# Test the navigator script
.\Scripts\navigator.ps1

# Test email functionality directly
.\Scripts\New-OutlookEmail.ps1 -Article "123456" -TypeKey "ABC" "path\to\file.pdf"

# Check if data directories exist
Test-Path "Data\PDFs"
Test-Path "Data\STP_and_ZIPs"
```

## File Organization

### Required Directory Structure
```
File_Navigator/
├── Data/
│   ├── PDFs/           # PDF technical documents
│   └── STP_and_ZIPs/   # STP and ZIP files
├── Scripts/
│   ├── navigator.ps1           # Main application
│   ├── New-OutlookEmail.ps1   # Email integration
│   └── Old Versions/          # Legacy scripts
└── README.md
```

### File Naming Conventions
- PDF files often contain article numbers in format: `XXXXXX[A-Z]{1,2}` or `XXXXXX_H01` style
- Files may have language suffixes like `_en`, `_belmore_en`, `_gb` 
- The system automatically extracts and highlights article numbers for easy identification

## Integration Points

### Outlook COM Integration
The `New-OutlookEmail.ps1` script uses Outlook COM automation to:
1. Check for existing open email in Outlook
2. Attach files to open email OR create new email with attachment
3. Auto-populate subject line with extracted article number and type key

### Error Handling
- Graceful handling of missing directories
- COM object error handling for Outlook integration
- Console state restoration on exit (cursor visibility, colors)

## Development Notes

- The application uses PowerShell's `Host.UI.RawUI.ReadKey()` for keyboard input handling
- Console manipulation includes cursor positioning, color management, and viewport scrolling
- File pre-loading occurs at startup for better performance during search
- All paths are constructed relative to the script location for portability