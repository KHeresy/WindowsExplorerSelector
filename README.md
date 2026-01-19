# Windows Explorer Selector

> [!CAUTION]
> **⚠️ 警語：本專案之程式碼、文件與圖示完全由 Gemini-CLI 自動產生。**  
> 使用者應自行評估安全風險，作者不保證程式碼的絕對正確性與安全性。

這是一個強大的 Windows 檔案總管增強工具，提供即時搜尋並直接在原生檔案總管視窗中選取檔案的功能。

## 功能特色

*   **常駐執行 (Resident Mode)**：
    *   程式啟動後會縮小至系統匣 (System Tray)。
    *   支援單一執行實體 (Single Instance)，透過具名管道 (Named Pipe) 進行高效處理。
*   **全域快速鍵 (Global Hotkey)**：
    *   預設 `Ctrl + F3`。在檔案總管視窗中按下，可立即開啟搜尋介面。
    *   快速鍵可於設定中自定義。
*   **右鍵選單整合 (Shell Extension)**：
    *   **完整支援 Windows 11**：可註冊至第一層右鍵選單（Modern Context Menu）。
    *   原生整合於檔案總管右鍵選單「尋找檔案」。
    *   使用純 Win32 API 實作 DLL 插件 (IExplorerCommand)，輕量且無負擔。
*   **進階搜尋與選取**：
    *   **雙向連動**：點兩下搜尋結果即可在檔案總管中選取；點擊「在檔案總管選取」可批次選取所有選中項。
    *   **多選支持**：支援選取全部、反向選取、全部取消。
    *   **預設自動全選**：可設定搜尋後自動選取所有結果。
*   **個人化設定**：
    *   **多國語言**：支援英文、繁體中文介面切換（需重啟生效）。
    *   **自動啟動**：支援隨 Windows 登入時自動執行。
    *   **最上層顯示**：保持搜尋視窗永遠在最前方。
    *   **選取後自動關閉**：執行選取動作後自動隱藏視窗。

## 系統需求

*   Windows 10 / 11 (64-bit)
*   Visual Studio 2026 (用於編譯)
*   Qt 6.10.1 (MSVC 2022 64-bit)

## 專案結構

*   **SelectorExplorerPlugin** (DLL): Windows Shell Extension，負責右鍵選單邏輯 (支援 IContextMenu 與 IExplorerCommand)。
*   **ExplorerSelector** (EXE): Qt 6 應用程式，負責搜尋介面、全域快速鍵與檔案總管控制。
*   **AppxManifest.xml** / **Install.ps1**: 用於 Windows 11 Sparse Package 註冊的設定與腳本。

## 編譯說明

1.  確保已安裝 Visual Studio 2026 與 Qt 6.10.1 (並包含 Network 模組)。
2.  開啟 `ExplorerSelector.slnx`。
3.  選擇 `Release` 設定，平台選擇 `x64`。
4.  建置方案 (Build Solution)。

## 打包與安裝 (推薦)

為了方便部署與完整支援 Windows 11 右鍵選單，本專案提供了一鍵打包腳本。

1.  **打包專案**：
    以系統管理員身分執行 `package.ps1`：
    ```powershell
    .\package.ps1
    ```
    此腳本會將執行檔、DLL、Qt 依賴庫以及安裝腳本整理至專案根目錄下的 `App` 資料夾。

2.  **安裝 (Windows 11 Modern Context Menu)**：
    進入 `App` 資料夾，以**系統管理員身分**執行 `Install.ps1` :
    ```powershell
    cd App
    .\Install.ps1
    ```
    此腳本會：
    *   註冊 COM 元件 (SelectorExplorerPlugin.dll)。
    *   建立並信任開發用憑證。
    *   註冊 Sparse Package 以啟用 Windows 11 第一層右鍵選單。

3.  **解除安裝**：
    進入 `App` 資料夾，以**系統管理員身分**執行 `Uninstall.ps1`：
    ```powershell
    .\Uninstall.ps1
    ```

## 手動安裝 (僅舊版選單)

若不需要 Windows 11 第一層選單支援，可僅註冊 DLL：

1.  **註冊右鍵選單**：
    以**系統管理員身分**執行：
    ```powershell
    regsvr32 "SelectorExplorerPlugin.dll"
    ```
2.  **執行主程式**：
    執行 `ExplorerSelector.exe` 後，會在系統匣看到圖示。

## 注意事項

*   **搜尋中的狀態**：搜尋時會在列表顯示 "Searching..." 並鎖定列表，搜尋完成後自動恢復。
*   **萬用字元**：支援 `*` 與 `?` 萬用字元搜尋。
*   **單一實體**：本程式使用 `QLocalServer` 確保單一執行實體。

## 授權與聲明 (License)

*   本軟體使用 **Qt Framework** (Qt 6)，依據 **LGPL v3** 授權協議進行動態連結 (Dynamic Linking)。
*   Qt 是 The Qt Company Ltd. 的註冊商標。
*   本專案的原始碼與修改均應符合 LGPL 規範，確保使用者有權更換所使用的 Qt 函式庫版本。
*   詳細授權資訊請參閱程式設定頁面中的 "About Qt" 或安裝目錄下的 LICENSE 文件。
