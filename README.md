# STPigator

  ![STPigator Logo](https://github.com/user-attachments/assets/7144dda6-cdde-4f36-b6d6-7dc8b4771de9)

  A fast, keyboard-driven file navigator for searching and managing PDF, STP, and ZIP files with
  intelligent highlighting and Outlook integration.

  ![STPigator
  Interface](https://github.com/user-attachments/assets/cc013616-3d67-41a8-856e-6d9302966bc1)

  ## 📋 Table of Contents

  - [🚀 Quick Start](#-quick-start)
  - [⚙️ First-Time Setup](#️-first-time-setup)
  - [🎮 How to Use](#-how-to-use)
  - [⌨️ Keyboard Shortcuts](#️-keyboard-shortcuts)
  - [🔍 Smart Features](#-smart-features)
  - [📁 File Organization](#-file-organization)
  - [🔧 Troubleshooting](#-troubleshooting)
  - [💡 Tips & Tricks](#-tips--tricks)

  ## 🚀 Quick Start

  1. **Populate your data folders** with PDF and STP/ZIP files
  2. **Double-click `STPigator.bat`** to launch
  3. **Start typing** to search files instantly
  4. **Press `Enter`** to attach to Outlook or **`Ctrl+O`** to open directly

  ## ⚙️ First-Time Setup

  ### 1. Organize Your Files
  STPigator/
  ├── Data/
  │   ├── PDFs/           # Place all PDF files here
  │   └── STP_and_ZIPs/   # Place all STP and ZIP files here
  └── STPigator.bat       # Double-click to run

  ### 2. Launch the Application
  - **Double-click** `STPigator.bat`

  ### 3. Fix Execution Policy (If Needed)
  If you see a PowerShell execution policy error:

  1. **Right-click PowerShell** → **"Run as administrator"**
  2. **Run this command:**
     ```powershell
     Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  3. Close PowerShell and run STPigator normally

  🎮 How to Use

  Basic Navigation

  - Type to search - Instant file filtering as you type
  - Arrow keys - Navigate up/down through results
  - Ctrl+T - Cycle between PDFs → STP/ZIPs → Recent Files
  - Esc - Exit the application

  File Actions

  - Enter - Attach file to Outlook email
  - Ctrl+O - Open file in default application
  - Delete - Move file to Recycle Bin
  - Ctrl+1-0 - Quick select files 1-10 from visible list

  ⌨️ Keyboard Shortcuts

  | Shortcut | Action                               |
  |----------|--------------------------------------|
  | Type     | Live search filtering                |
  | ↑/↓      | Navigate file list                   |
  | Ctrl+T   | Switch between PDFs/STP/Recent modes |
  | Enter    | Attach to Outlook email              |
  | Ctrl+O   | Open file directly                   |
  | Delete   | Move file to Recycle Bin             |
  | Ctrl+1-0 | Quick select numbered files          |
  | Ctrl+U/D | Page up/down (half page)             |
  | Ctrl+Q   | Quit application                     |
  | Esc      | Exit                                 |

  🔍 Smart Features

  Intelligent Highlighting

  - 🟢 Green - Article numbers (e.g., 116890.A01, 100120E)
  - 🟡 Yellow - Type keys (e.g., GR28C-2DN.D5.CR, FE050-VDA.4I.V7)

  Advanced Search

  - Wildcard patterns - Use ? and * (e.g., FN?80*V finds FN080ZIQGLV7P3)
  - Partial matching - Type any part of filename
  - Real-time results - Instant filtering as you type

  File Modes

  - 📄 PDFs - Technical documents and datasheets
  - 📦 STP/ZIPs - CAD files and archives
  - 🕒 Recent - Last 20 accessed files (auto-tracked)

  Outlook Integration

  - Smart attachment - Adds files to open draft emails or creates new ones
  - Automatic fallback - Opens files directly if Outlook unavailable
  - Silent operation - No interrupting success messages

  📁 File Organization

  Supported Formats

  - PDFs: .pdf files in Data/PDFs/
  - STP Files: .stp files in Data/STP_and_ZIPs/
  - Archives: .zip files in Data/STP_and_ZIPs/

  Article Number Recognition

  Automatically detects and highlights these patterns:
  - 123456 - 6-digit base numbers
  - 123456E - With letter suffix
  - 123456.A01 - With dot separator
  - 123456-A01 - With dash separator
  - 123456_H01 - With underscore separator

  🔧 Troubleshooting

  Common Issues

  "No files found"
  - Verify files are in correct Data/PDFs/ or Data/STP_and_ZIPs/ folders
  - Check file extensions are .pdf, .stp, or .zip

  "Outlook attachment failed"
  - Application automatically opens file instead
  - Ensure Outlook is running for email attachment

  Search not working
  - Clear search with Backspace
  - Try different search terms or wildcards

  Advanced Troubleshooting

  - PowerShell errors: Ensure execution policy is set (see setup)
  - File access issues: Run as administrator if needed
  - Performance: Tool pre-loads files at startup for fast searching

  💡 Tips & Tricks

  Efficient Workflows

  1. Use Recent Files - Ctrl+T to access your last 20 used files
  2. Wildcard Search - FN?80* finds all FN080 variants quickly
  3. Quick Select - Ctrl+3 jumps to 3rd file in current view
  4. Safe Deletion - Delete key moves to Recycle Bin (recoverable)

  Power User Features

  - Article extraction - Automatically populates email subjects
  - Type key detection - Identifies technical specifications
  - Memory efficiency - Files cached for instant search results
  - Cross-platform paths - Works with network drives and cloud storage

  ---
  Need help? Open an issue on GitHub or check the #-troubleshooting.
