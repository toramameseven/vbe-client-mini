# Change Log

All notable changes to the "vbecm" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]
### Added
- When vbecm exports modules, vbecm exports modules to a .vbe folder. 
  Then vbecm copes to a src folder and a .base folder form the .vbe folder.


## [0.0.4] - 2022-09-04
### Changed
* Some titles of context menu are changed.

### Fixed
- Verify logic is fail. When VBA engine import modules, the case conversion and whitespace conversions occur.
  So, sometimes the verify errors occur. Now vbecm does not check modules when you import modules.
- Context menus is displayed when no vba modules is selected. Fix this.

## [0.0.3]
### Fixed
- When you commit a frm module, you fail to commit at 0.0.2. Fix this.

## [0.0.2]
### Added
- Some features add.

### Fixed
- When you run a sub functions that are in some modules, vbecm can not detect the module the function includes.
  This version can detect the module you select.

## [0.0.1]


<!-- 
### Added
 for new features.
### Changed
 for changes in existing functionality.
### Deprecated
 for soon-to-be removed features.
### Removed
 for now removed features.
### Fixed
 for any bug fixes.
### Security
 in case of vulnerabilities.
 -->
