# vbecm README

vbe-client-mini

This is vbe client mini.
Vba module export, import extension for vs code.
This extension thinks Excel as Excel Server.

## Requirement

* Windows 10
* Excel

## Features

before use this extension, you must back up your xlsm or xlam file.

### Explorer context menu for an xlsm and an xlam

* Export all VBA modules
  * Export vba modules form a xlsm file to a folder.
  * select a xlsm file in the explorer, and right click, and select export.
  you can find a Src_[xxx.xlsm] folder at the same folder.
* Import all VBA modules
  * Import vba modules to a xlsm file.
  * you can import modules from a src_file.xlsm folder.
* compile VBA
  * Do compile.
  * vbecm can not detect complete to compile, so please check on the VBE.
* Export only frx modules
  * Export frx modules
  * You can not know frx is modified or not.  If you modified a frm on the vscode,
    you can export only frx files.

### Explorer context menu for a Src_[xxx.xlsm] folder

* Push All Vba modules
  * Push all modules form a Src_[xxx.xlsm] folder
  * you can push all modules in the source folder to the excel book.
* Check modified
  * Check modification among srcXXX, .vbe and .base.
  * The result is shown at Explore Diff

### Explorer context menu for DIFF:

Modified modules are shown in the DIFF.

#### src(base)

If there is some modification in the src folder, the modified modules are shown.
Vbecm checks the modification between the src and base folder.

You can click to compare the module in base and the module in the src.

You can right click to push the module to a Excel. 

#### vbe(base)

If there is some modification in the VBE, the modified modules are shown.
Vbecm checks the modification between the src and vbe folder.

You can click to compare the module in base and the module in the vbe.

At first you right click to diff between vbe and src.
And Edit the module in the src.
And right click to resolve the conflicting. 
The The resolve command is to copy vbe to base.

The resolve command is very confused. Sorry.


### editor context menu

* Push VBA module
  * Push vba module to a xlsm file on an editor.
  * Selected module is imported to a xlsm file.
* Run Sub function
  * run Sub() function on a editor.
  * select a sub XXX() line, and right click, and select run.
* Pull VBA module
  * Pull selected module form xlsm file.
  * If you modified the module, vbe overwrites the modules in the src with vbe
* Goto Vbe module
  * Goto the code of module on the vbe form the editor you select.
    It does not work for a workbook module.


### Folders

#### src_[excelBookName] folder

When you export modules, the modules are in the src_[excelBookName] folder.


#### .base folder

In the src_[excelBookName] folder, you can see a .base folder.
When you commit or import modules, this folder is updated.
And when you export modules, vbecm tests if there are difference between the src_[excelBookName]  and .base folders.

#### .vbe folder

In the src_[excelBookName] folder, you can see a .vbe folder.
When you commit or import modules or check modification, this folder is updated.
And when you import modules, vbecm tests if there are difference between the .vbe  and .base folders.

### exclude settings

You should better exclude the .base and .vbe folder at the explorer and the search.

```
  "files.exclude": {
    "**/.base": true,
    "**/.vbe": true,
  },
  "search.exclude": {
      "**/.base": true,
      "**/.vbe": true,
  },
```
### settings

* At cjk language area, vbs message not work good beside Japanese. Set encoding option
  * vbecm.vbsEncode, for japanese 'windows-31j'

## Recommendation

if you need, it is better to install VSCode VBA below.
* https://marketplace.visualstudio.com/items?itemName=spences10.VBA


## How to build

https://code.visualstudio.com/api/working-with-extensions/publishing-extension

```
npm install -g vsce
vsce package --target win32-x64
vsce publish
```

## Known problems

* If the encoding of your vba modules is not convertible to utf8, do not save the modules on utf8.
  Sometimes I do this, and messages in modules become unreadable.
* At cjk language area, vbs message not work good beside Japanese. Please customize your encoding on the settings.
* Sometimes, Excel remain on background. You should kill the process on a task manager. Sorry.
* Sometimes, you meet export or import error. So you recover from a backup excel book.
* When vbecm is working, [[vbecm]] is displayed on the status bar. If no vbecm menus displayed with out [[vbecm]] , maybe bugs, please reload your vscode.
* When you import a sheet module, sometimes you find new empty line at end of module.
* Sometime, you may see the folder src_GUID. It stays when some errors occur. Please delete it.

## Shallow dive

Not deep dive.

### Sheet modules and Workbook modules

Sheet modules and Workbook modules are exported to [ModuleName].sht.cls.
For vbecm distinguishes normal class modules from sheet(book) modules.
Thanks for the [VbaDeveloper](https://github.com/hilkoc/vbaDeveloper "VbaDeveloper")


### Opened excel file

While vbecm is working, Excel dose not close. Please close the Excel when you end using it.

### For xlam file

When you use an xlam file, vbecm add a book to the excel instance.

### Import frm

When you import a frm module, you find a extra line added.
vbecm will delete the line.

### Diff algorism
we are using the dir-compare and we are very grateful to be able to use it.
https://github.com/gliviu/dir-compare
And we want the option IGNORE CONTENT CASE, we fork and add some modification.
We think there's probably a way to easily extend it without forking.

## Release Notes

0.0.5 test release.

* Add Some features for checking modification.
* Some fixes

0.0.4 test release.

* Verify logic is fail. When VBA engine import modules, the case conversion and whitespace conversions occur.
  So, sometimes the verify errors occur. Now vbecm does not check modules when you import modules.
* Context menus is displayed when no vba modules is selected. Fix this.
* Some titles of context menu are changed.

0.0.3 test release.

* When you commit a frm module, you fail to commit at 0.0.2. Fixed.

0.0.2 test release.

* When you run a sub functions that are in some modules, vbecm can not detect the module the function includes.
  This version can detect the module you select.
* Some features add.

0.0.1 test release.


