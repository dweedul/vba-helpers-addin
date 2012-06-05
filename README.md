# VBE Helper Addin v2.0.0
An addin for Excel 2007+ that does some nifty stuff in the code editor.

## In the VBE
A new toolbar is added to the VB IDE.

### Exporting and Reloading Code

Two menus appear on the toolbar, one for exporting, one for reloading.

The export and reload options are stored within each module. These options start with a `'!`.  This can be changed to your liking in the config module.

Current options:
* `no-export`            - file is not exported. I use this in a lot of quick testing code.
* `no-reload`            - file will not be reloaded from the file path, even if the command if given
* `absolute-path <path>` - absolute path to export to/reload from.
* `relative-path <path>` - path to save to/load from. This is relative to the current workbook's file location.

__Import From Folder__ will import all files in the folder into the selected project.


### Copy path to clipboard
Copies the currently selected projects path to the clipboard.

### Command-code dropdown
Lists the command codes for use in a project.

### 0X menu
USE THIS STUFF WITH CAUTION!!!

__Clear all code__ will clear the selected project of all code.

## In the Main Excel Window

A new ribbon tab appears!

### Type name of selection button.
This button will reflect the typename of the selected object.  Clicking the button refreshes the target.

## ideas
* multiline option parsing: for documentation, etc
* array options: for references and requires
