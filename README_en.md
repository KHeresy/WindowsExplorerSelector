# Windows Explorer Selector

> [!CAUTION]
> **⚠️ Caution: The code, documentation, and icons in this project were generated entirely by Gemini-CLI.**  
> Users should assess the safety risks themselves; the author does not guarantee the absolute correctness or security of the code.

This is a Windows Explorer enhancement tool that provides real-time search and file selection directly within the native Explorer window.

![app](images/ExplorerSelector-en.png)

## Features

*   **Resident Mode**:
    *   The program minimizes to the system tray after launch.
    *   Supports single instance handling via a named pipe for efficient processing.
*   **Global Hotkey**:
    *   Default: `Ctrl + F3`. Press in an Explorer window to open the search interface immediately.
    *   The hotkey can be customized in the settings.
*   **Shell Extension Integration**:
    *   Natively integrated into the Explorer context menu as "Find File".
    *   Implemented using pure Win32 API for the DLL plugin (IExplorerCommand), lightweight and low overhead.
*   **Advanced Search and Selection**:
    *   **Bidirectional Linking**: Double-click a search result to select it in Explorer; click "Select in Explorer" to batch select all highlighted items.
    *   **Multi-select Support**: Supports Select All, Invert Selection, and Clear All.
    *   **Auto Select All (Default)**: Can be configured to automatically select all results after a search.
*   **Personalization**:
    *   **Multi-language**: Supports switching between English and Traditional Chinese (requires restart).
    *   **Auto-start**: Supports running automatically at Windows logon.
    *   **Stay on Top**: Keep the search window always in front.
    *   **Auto-close After Selection**: Automatically hide the window after performing a selection action.

## System Requirements

*   Windows 11 (64-bit)
*   Visual Studio 2026 (for compilation)
*   Qt 6.10.1 (MSVC 2022 64-bit)

## Project Structure

*   **SelectorExplorerPlugin** (DLL): Windows Shell Extension responsible for the context menu logic (supports IContextMenu and IExplorerCommand).
*   **ExplorerSelector** (EXE): Qt 6 application responsible for the search interface, global hotkey, and Explorer control.
*   **AppxManifest.xml** / **package.ps1**: Configuration and scripts for Windows 11 Sparse Package registration and packaging.

## Build Instructions

1.  Ensure Visual Studio 2026 and Qt 6.10.1 (with the Network module) are installed.
2.  Open `ExplorerSelector.slnx`.
3.  Select the `Release` configuration and `x64` platform.
4.  Build the solution.

## Manual Installation

If you do not need support for the Windows 11 first-level context menu, you can register only the DLL:

1.  **Register the context menu**:
    Run as **Administrator**:
    ```powershell
    regsvr32 "SelectorExplorerPlugin.dll"
    ```
2.  **Run the main program**:
    Launch `ExplorerSelector.exe`; you will see the icon in the system tray.

You may need to install Visual C++ v14 Redistributable Package first: https://aka.ms/vc14/vc_redist.x64.exe

## Notes

*   **Searching Status**: While searching, the list will display "Searching..." and be locked; it automatically recovers when the search is complete.
*   **Wildcards**: Supports `*` and `?` wildcard character searches.
*   **Single Instance**: This program uses `QLocalServer` to ensure only a single instance is running.

## License and Disclaimer

*   This software uses the **Qt Framework** (Qt 6), dynamically linked under the **LGPL v3** license.
*   Qt is a registered trademark of The Qt Company Ltd.
*   The source code and modifications of this project should comply with LGPL requirements, ensuring users have the right to replace the version of the Qt library used.
*   For detailed license information, please refer to "About Qt" in the application settings or the LICENSE files in the installation directory.
