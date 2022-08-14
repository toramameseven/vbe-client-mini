# vbecm README

vbe-client-mini

This is vbe client mini.
Vba module export, import extension for vs code.
This extension thinks Excel as Excel Server.

## Features

before use this extension, you must back up your xlsm or xlam file.

### explorer context menu

* export vba module form a xlsm file to a folder.
  * select a xlsm file in the explorer, and right click, and select export.
  you can find a src_file.xlsm folder at the same folder.
* import vba module to a xlsm file.
  * you can import modules from a src_file.xlsm folder.

### editor context menu

* commit vba module to a xlsm file on an editor.
  * Selected module is imported to a xlsm file.
* run Sub() function on a editor.
  * select a sub XXX() line, and right click, and select run.
* checkout a module form excel.

### settings

* form modules and sheet modules are not imported on default settings. If you want, you can enable the settings
  * Enable importing form modules(frm).
  * Enable importing sheet Modules(cls). 
* At cjk language area, vbs message not work good beside Japanese. Set encode option
  * vbecm.vbsEncode, for japanese 'windows-31j'

let's fun vba life.


## Recommendation

if you need, it is better to install VSCode VBA below.
* https://marketplace.visualstudio.com/items?itemName=spences10.VBA


## How to build

https://code.visualstudio.com/api/working-with-extensions/publishing-extension

```
npm install -g vsce
$ vsce package
$ vsce publish
```

## Known problems

* At cjk language area, vbs message not work good beside Japanese. Please customize your encode.
* Sometimes, Excel remain on background. You should kill the process on a task manager.
* When you export modules, vbecm asks "Do you want to export" in notification window.
  Unless you click Yes or No, you can not select next command.
* Sometimes, you meet export or import error. So you recover from a backup file.
* When vbecm is working, [[vbecm Working]] is displayed on the status bar. Check the notification window, if some confirm dialog exists. Or some bug includes, please reload your vscode.

## Shallow dive

Not deep dive.

### Sheet modules and Workbook modules

Sheet modules and Workbook modules are exported to [ModuleName].sht.cls.
For vbecm distinguishes normal class modules from sheet(book) modules.
Thanks for the VbaDeveloper.

### Opened excel file

While vbecm is working, Excel dose not close. Please close the file when you end using it.

### For xlam file

When you use an xlam file, vbecm add a book to the excel instance.

### Import frm

When you import a frm module, you find a extra line added.
vbecm will delete the line.


## Release Notes

