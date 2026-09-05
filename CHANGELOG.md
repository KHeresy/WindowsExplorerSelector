# Changelog

## [0.2.0] - 2026-09-05

### Added

- 新增 `F2` 視窗快捷鍵，可快速將焦點移至搜尋欄並全選目前搜尋文字。
- 新增 GitHub Actions 自動化建置流程。
- 新增 LGPL 授權文件 `LICENSE`。
- 新增應用程式與 Shell Extension 的版本資訊。

### Changed

- Qt 版本由 6.10.1 升級至 6.11.2。
- 更新 README 與 README_en.md，補充常駐搜尋、選用插件、安裝方式及建置需求等說明。
- `package.ps1` 改由環境變數 `QT_ROOT_DIR` 指定 Qt 安裝路徑，不再寫死 Qt 路徑。
- `package.ps1` 的必要檔案與 `windeployqt.exe` 檢查改為直接拋出錯誤，避免封裝流程在檔案缺少時繼續執行。
- 啟用 Qt deprecated API 警告控制，將警告檢查範圍設為 Qt 6.9 以前的 API。
- 新增 Visual Studio 資源專案檔與 Windows x64 發行檔案至專案內容。

## [0.1.0]

- 基準版本。詳細功能請參閱當時的專案文件與原始碼。
