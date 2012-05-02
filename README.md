VBE Helper Addin
================
An addin for Excel 2007+ that does some nifty stuff in the code editor.

Features
--------
__Exporting code__
Export all modules and classes or just one.

__Importing code__
Imports all modules and classes or just one.

The export and import are controlled by the use of `#commands` in the comments at the top of a code moduel.

* `#NoExport` - file is not exported. I use this in a lot of quick testing code.
* `#NoRefresh` - file will not be refreshed, even if the command if given
* `#RelativePath` - path to save to/load from. This is relative to the current workbook's file location.
* `#AbsolutePath` - path to save to/load from. __Not yet implemented__

__Copy path the clipboard button__
I forget to open my console into the correct folder so often that copying the path to the clipboard seemed like a good idea.

__What the hell am I selecting button__
On the developer tab, there is a new button. When pressed, this button will display the name of the class of the object that is currently selected.
For example, select the axis on a graph and press the button. The button will say 'Axis'