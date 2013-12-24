# Confluence Export table to Excel Plugin

This plugin enables you to download tables on Confluence pages as Office Open XML / OOXML sheets.  

The plugin consists of the following components:
* A servlet which creates an Office Open XML (.xslx) Workbook from JSON
* A user-macro which to decorate tables with a button to initiate the export

## Restrictions
* Currently the macro must be added as user macro. It will be included in the plugin somewhere in the future.
* Currently nested tables are not supported. They are simply ommitted in the output.
* Currently only PNG and JPEG images are exported.

## Installing the plugin
1. Clone and build the repository and run the package command to create the jar file.
2. Open the Confluence administration and go to "Manage add-ons".
3. Use "Upload add-on" select the packaged jar file and confirm to upload and install the plugin.

## Adding the macro
Open the Confluence administration and go to the user macros section and create a new user macro.
Enter:
* Macro Name: tabletoexcel
* Macro Title: Table to Excel
* Macro body processing: Rendered
* Paste the contents of table-to-excel-user-macro.txt to the Template section (empty it before)
* Save

## Usage Example
1. Create a new page
2. Add the table-to-excel macro
3. Add a table to the body of the macro
4. Save
5. Now you see your table decorated with a button that says "Export to Excel"
6. Click the button and download your Excel sheet

## How to build
1. Install the Atlassian Plugin SDK from https://marketplace.atlassian.com/download/plugins/atlassian-plugin-sdk-windows
2. Clone the repository
3. Open a command prompt in the repository folder and execute:
  * atlas-compile
  * atlas-package
  * atlas-debug

## How to import to Eclipse IDE
1. Open a command prompt in the repository folder
2. Execute atlas-mvn eclipse:eclipse which generates the Eclipse project
3. Import project in Eclipse

Or follow the instructions on https://developer.atlassian.com/display/DOCS/Set+Up+the+Eclipse+IDE+for+Windows  
Alternatively use IntelliJ: https://developer.atlassian.com/display/DOCS/Configure+IDEA+to+use+the+SDK  

## Other hints
Shell commands:  
* atlas-clean   -- removes the target directory  
* atlas-compile -- compiles the plugin  
* atlas-package -- creates the jar file (compiles if necessary)  
* atlas-run     -- installs this plugin into the product and starts it on localhost  
* atlas-debug   -- same as atlas-run, but allows a debugger to attach at port 5005  
* atlas-cli     -- after atlas-run or atlas-debug, opens a Maven command line window:  
                   - 'pi' reinstalls the plugin into the running product instance  
* atlas-help    -- prints description for all commands in the SDK  

Full documentation is always available at:  
https://developer.atlassian.com/display/DOCS/Introduction+to+the+Atlassian+Plugin+SDK
