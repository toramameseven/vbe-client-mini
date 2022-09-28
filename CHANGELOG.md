# Change Log

All notable changes to the "vbecm" extension will be documented in this file.

## [Unreleased]


## [0.0.6] - 2022-09-28

### Added
- Options: Automatically update modification. default false.
  For Automatically update modification, and feel slow operation.
- Editor context menu: Goto Vbe module
  Goto the module on the vbe form the editor you select.
- Source folder context menu:
  - pull all modules
  - compile vba

### Changed
- The output tab name is changed to vbecm from myOutput.

### Fixed
- when a vba is running, vbecm stops to access to an excel book. 
  vbecm tests if a continue command and a pause command are enabled in the VBE.
- Detect modification for no modification module. Fixed. #5
- Sometime, vbecm can not detect the target for compile. Fixed.


## [0.0.5] - 2022-09-23
### Added
- Some features for checking modification.

### Fixed
- Some fixes


## [0.0.4] - 2022-09-04
### Changed
- Some titles of context menu are changed.

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
Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.
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
