# vbecm README

vbe-client-mini

This is vbe client mini.
Vba module export, import extension.
This extension thinks Excel as Excel Server.

## Features

before use this extension, you must back up your xlsm file.

* export vba module form a xlsm file to a folder.
  * select a xlsm file in the explorer, and right click, and select export.
  you can find a src_file.xlsm folder at the same folder.
* import vba module to a xlsm file.
  * you can import modules from a src_file.xlsm folder.
* commit vba module to a xlsm file on an editor.
  * Although from an editor window, all modules are imported to a xlsm file.
* run Sub() function on a editor.
  * select a sub XXX() line, and right click, and select run.
* form modules and sheet modules are not imported on default settings. If you want, you can enable the settings
  * Enable importing form modules(frm).
  * Enable importing sheet Modules(cls). 
* At cjk language area, vbs message not work good beside Japanese. Set encode option
  * vbecm.vbsEncode, for japanese 'windows-31j'

let's fun vba life.


## Requirements

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

## Release Notes

