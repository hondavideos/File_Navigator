![20250821_2112_STPigator Logo Reveal_simple_compose_01k3746c61ffbs72yfxqdjr017](https://github.com/user-attachments/assets/7144dda6-cdde-4f36-b6d6-7dc8b4771de9)


![STPigator-output](https://github.com/user-attachments/assets/cc013616-3d67-41a8-856e-6d9302966bc1)

This tool provides a fast, keyboard-driven interface to search for files in the `Data` directory and perform actions on them, like attaching them to an Outlook email or opening them directly.

## First-Time Setup

1.  **Populate Data Folders:**
    *   Place all your PDF files into the `/Data/PDFs/` folder.
    *   Place all your STP and ZIP files into the `/Data/STP_and_ZIPs/` folder.

2.  **Run the Script:**
    *   Double Click on STPigator.bat 

3.  **Execution Policy (If you see an error):**
    *   If you get an error message about the execution policy, you will need to run a one-time command to allow scripts to run on your machine.
    *   Open PowerShell as an **Administrator** (search for PowerShell in the Start Menu, right-click it, and select "Run as administrator").
    *   In the blue administrator window, type the following command and press Enter:
        ```powershell
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
        ```
    *   This is a safe, one-time change that allows signed scripts (like this one) to run. You can now close the administrator window and run the navigator script normally.
    
<img width="1116" height="630" alt="image" src="https://github.com/user-attachments/assets/6ed3d7ed-168a-468f-bc13-f07150890a49" />

## How to Use

*   **Live Search:** Start typing to instantly filter the file list.
*   **Switch Modes (`Ctrl+T`):** Press `Ctrl+T` to toggle between searching for PDFs and searching for STP/ZIP files.
*   **Navigate:** Use the `Up/Down` arrow keys to move the selection.
*   **Attach to Email (`Enter`):** Press `Enter` on a selected file to create a new Outlook email with that file as an attachment.
*   **Open File (`Ctrl+O`):** Press `Ctrl+O` to open the selected file in its default application.
*   **Quick Select (`Ctrl` + `1-0`):** Press `Ctrl` and a number key to instantly select the corresponding file from the visible list.
*   **Quit (`Esc`):** Press `Esc` to exit the navigator.
