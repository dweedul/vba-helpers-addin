# VBE Helper Addin
An addin for Excel 2007+ that does some nifty stuff in the code editor.

## In the VBE

A new toolbar shows up with the following features.

### Exporting and Importing Code

Four buttons here:
* Export active module.
* Export active project.
* Reload active module.
* Reload active project.

The export and import are controlled by the use of `#commands` in the comments at the top of a code module.

* `#NoExport` - file is not exported. I use this in a lot of quick testing code.
* `#NoRefresh` - file will not be refreshed, even if the command if given
* `#RelativePath` - path to save to/load from. This is relative to the current workbook's file location.

### Copy path the clipboard button
I forget to open my console into the correct folder so often that copying the path to the clipboard seemed like a good idea.

## In the Main Excel Window

The following buttons are located in a new group on the Developer tab:

### Class name of current selection.
This button will change its name to the class type of the selected object. (e.g. if a chart axis is selected, it will say "Axis")

## Usage
Goto the [downloads](https://github.com/dweedul/VBEHelpersAddin/downloads) and use it!

If you want to peek at the code inside, the password is `qwerty`.  I do this so that it won't show me all that code when I'm working on other code.
