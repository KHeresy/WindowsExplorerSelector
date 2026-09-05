# Changelog

## [0.2.0] - 2026-09-05

### Added

- Added the `F2` window shortcut to quickly focus the search field and select the current search text.
- Added a GitHub Actions automated build workflow.
- Added the LGPL license document `LICENSE`.
- Added application and Shell Extension version information.

### Changed

- Upgraded Qt from 6.10.1 to 6.11.2.
- Updated README and README_en.md with additional information about resident search, the optional plugin, installation, and build requirements.
- Updated `package.ps1` to use the `QT_ROOT_DIR` environment variable for the Qt installation path instead of a hard-coded path.
- Changed the required-file and `windeployqt.exe` checks in `package.ps1` to throw errors directly, preventing the packaging process from continuing when files are missing.
- Enabled Qt deprecated API warning control, with the warning check range set to APIs deprecated before Qt 6.9.
- Added Visual Studio resource project files and Windows x64 release files to the project.

## [0.1.0]

- Baseline version. See the project documentation and source code at that time for detailed functionality.
