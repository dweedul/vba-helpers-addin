

__Issues__

  * [bug] refresh selected module - this does not add the file back, may need to close the workbook first...
    - use `Application.OnTime Now(), ThisWorkbook.Name & "!ImportFiles"`? (see SourceTools.xla!ImportFiles for example)
  * [feature] reloading should also set project references and project properties (e.g. name, compile switches)
  * [feature] inject timer code at #timer comments (#timer name/comment)
  * [feature] reloads should follow paths
  * [feature] implement absolute path export/import
  * [feature] dropdown for #codes
  * [feature] auto add of watched vars for a given project: #watch variable context etc
  * [feature] find grouped code
    - use the {{#section}} {{/section}} start and end tags (like {{moustache}}) (maybe get rid of the moustaches?)
  * [feature] build PowerPoint version
  * [feature] comment/uncomment all debug.* lines
  * [feature] build a template dropdown that injects snippets and classes.

__Closed Issues__

  * [complete] export all in this project
  * [complete] refresh all in this project
  * [complete] export selected module
  * [complete] Add a button to copy to the clipboard the path the current project/file.
  * [removed]  VBScript to refresh all in this project
  * [removed]  aggregate @TODOs in a tree like on the MZTools summarize functions -> can use MZTools Find function for this