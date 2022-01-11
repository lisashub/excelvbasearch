# Excel VBA Search
A VBA application to find one or more search terms within a single file's Excel VBA project code. It can be used to help with reverse engineering existing 
projects to better understand where code elements appear and which elements should be updated when making changes. Instead of cycling through single 
instances of a term using the in-built VBA search feature, this app reveals all instances of a given search string in one place.

## Prerequisites
(1) The following VBA IDE libraries should be active in the app's Excel file:
  * Visual Basic for Applications
  * Microsoft Excel 16.0 Object Library (older versions may also work)
  * Microsoft Scripting Runtime
  * Microsoft Visual Basic for Application Extensibility 5.3
  * OLE Automation
NOTE: The above libraries are already active in the downloadable .xlsm version.
  
 (2) The “Trust access to the VBA project object model” option must be enabled in your Trust Center's macro settings. This options is
  found along the navigation path: File --> Options --> Trust Center --> Trust Center Settings --> Macro Settings --> Developer Macro Settings

## Usage

To use the application as-is, click on the .xlsm file from the repository and click the "Download" button. Once downloaded, open the file and enable macros. Click the "Begin Search" button and follow the directions as they appear on screen.

Some use tips:
* This app is designed to examine local files only; it will not work with URL links
* Make sure to enter the full path with the filename - for example, "C:\User\Desktop\filename.xlsm" (without quotes)
* Searches are not case sensitive
* Searches are full-word only with the exception of delimiters like "."
  * This means that a query of "name" will return results for "VBCodeMod.Name" but not "FileName"

The three source files of the downloadable .xlsm version are also available in this repository for download and tweaking.

## Future Feature Ideas
* Search multiple files in one session
* Option to save results on seperate sheets
* Option to work with URL links

## Acknowledgements
Much gratitude for the work of Zaid in London for laying out the foundations for programmatically sifting through VBA project objects; 
much of his template for backend structure was adopted into this application (with the exclusion of non-Excel related elements). Variable names 
were changed to guide me building a deeper understanding of how each program element functions. You can check out Zaid's open source work 
here: https://datapluscode.com/general/programmatically-search-vba-code/

Additional thanks to the many Microsoft and StackExchange contributors who have helped me fortify a foundational understanding
of the Microsoft Visual Basic for Applications Extensibility library and more.

## Feedback
Bugs? Questions? Other Feedback? Feel free to reach out to me at gf184grmu@mozmail.com (a Firefox Relay address).
