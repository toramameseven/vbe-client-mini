# vbecm README

vbe-client-mini

This is vbe client mini.
Vba module export, import extension.

## Features

before use this extension, you must back up your xlsm file.

* export vba module form a xlsm file to a folder.
  * select a xlsm file in the explorer, and right click, and select export.
  you can find a src_file.xlsm folder at the same folder.
* import vba module to a xlsm file.
  * you can import modules from a src_file.xlsm folder.
* commit vba module to a xlsm file on an editor.
  * Although form an editor window, all modules are imported to a xlsm file.
* run Sub() function on a editor.
  * select a sub XXX() line, and right click, and select run.
* form modules and sheet modules are not imported on default settings. If you want, you can enable the settings
  * Enable importing form modules(frm).
  * Enable importing sheet Modules(cls). 

let's fun vba life.


## Requirements

if you need, it is better to install VSCode VBA below.
https://marketplace.visualstudio.com/items?itemName=spences10.VBA


## How to build

https://code.visualstudio.com/api/working-with-extensions/publishing-extension

```
npm install -g vsce
$ vsce package
$ vsce publish
```

## Release Notes

