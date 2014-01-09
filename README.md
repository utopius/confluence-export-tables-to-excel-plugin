# Confluence Export table to Excel Plugin
This plugin allows to download tables on Confluence pages as Office Open XML / OOXML sheets.

The plugin consists of the following components:

* A servlet which creates an Office Open XML (.xslx) Workbook from JSON
* A macro to decorate tables with a button to initiate the export
* JavaScript functions which convert html tables to the JSON format the servlet expects

## Restrictions

* Other image types than PNG and JPEG images are not yet supported.
* Multiple images in one cell are not supported, yet. The first image wins.

## Installing the plugin

1. Clone and build the repository and run the package command to create the jar file.
2. Open the Confluence administration and go to "Manage add-ons".
3. Use "Upload add-on" select the packaged jar file and confirm to upload and install the plugin.

## Adding the macro

1. Open the Confluence administration and go to the user macros section and create a new user macro.
	* Macro Name: tabletoexcel
	* Macro Title: Table to Excel
	* Macro body processing: Rendered
2. Paste the contents of table-to-excel-user-macro.txt to the Template section (empty it before)
3. Save

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

To run the plugin in a developer instance run:

* atlas-run
* atlas-debug to run in debug mode (remote debugger available at port 5005)

## Setting up your IDE
As the repository contains no project files you have to generate them using Maven. See the following guides on how to do that:

* [Guide for Eclipse](https://developer.atlassian.com/display/DOCS/Set+Up+the+Eclipse+IDE+for+Windows)
* [Guide for IntelliJ](https://developer.atlassian.com/display/DOCS/Configure+IDEA+to+use+the+SDK )

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

## License

Copyright 2013 Florian Herbel, Holger Steffan

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
