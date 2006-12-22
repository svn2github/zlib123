ODF Add-in for Word2003
================================

Thank you for downloading ODF Add-in for Word2003.
The purpose of the software is to enable users to open ODF Specification based documents in Word 2003 and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Word 2003 - SP2 with .NET Programmability installed, on Windows XP – SP2 or Windows Server 2003 with SP1
* .NET Framework 2.0
* Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats.

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Run the downloaded executable to install the add-in for Word 2003.


Test the add-in 
---------------
If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Word 2003. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Word as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).


Troubleshooting Guide
---------------------
If you don't see the ODF entry, it could be due to missing features in Word (see the Software Requirements listed above). After completing the ODF Add-in installation, you need to activate the Add-in by doing the following actions:
* Go to Tools ? Customize, and then click the Commands tab. 
* In the Categories box, click Tools. 
* Drag COM Add-Ins from the Commands box over the Tools menu. When the Tools menu displays the menu commands, point to the location where you want the COM Add-Ins command to appear on the menu, and then release the mouse button. 
* Click Close. 
* On the Tools menu, click COM Add-Ins and do any of the following: 
* To load an add-in, select the check box next to the add-in name in the Add-Ins available list. If the add-in you want isn't in the Add-Ins available list, click Add, locate the add-in, and then click OK.


ODF to OpenXml conversion currently supports :
----------------------------------------------

Characters formatting
Paragraph formatting
Page formatting
Columns
Tables
Numbering
Table of content
Index of tables
Alphanumerical index
Pictures
OLE Objects
Headers and footers
Footnotes and endnotes
Bookmarks and cross-references
Change tracking
Line numbering
Fields (common and documents fields)


Features/options lost in the ODF to OpenXML translation
-------------------------------------------------------

Table of content protection
Text background color outside the 16 basic colors
Page number offset
Fields (chapter, description, printed by)
Nested frames
Frame absolute position
Annotations in text-boxes
Annotations in headers or footers
Hidden sections
Page break before endnotes
Notes in lists
SVM images
Cropped images
Embedded objects
Shape top and bottom wrapping
Distance between numbering and text
Spacing at top of a page/table
Next page style if no page break
Alignment of last line in paragraphs
Paragraph's background image
Automatic page breaks
Widow and orphan user defined line numbers
Text blinking
Text font weight
Basic text rotation
Capitalized text
Lowercase text
Text scale greater than 600%
Tab stop leader text
Individual page background color
Header or footer dynamic text adaptation
Shadow borders style
Subtables borders and padding
Tables repeat header option
Tables keep with next paragraph option
Unsplittable table option
Table background image
Tables with more than 64 columns
Table cell protection
Table cell shadow


OpenXML to ODF conversion currently supports :
----------------------------------------------

Document structure
Basic text and paragraph formatting
Basic tables
Basic document properties






