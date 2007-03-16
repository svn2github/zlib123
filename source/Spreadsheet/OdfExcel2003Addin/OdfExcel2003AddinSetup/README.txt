ODF Add-in for Excel 2003
================================

Thank you for downloading ODF Add-in for Excel 2003.
The purpose of the software is to enable users to open ODF Specification based documents in Excel 2003 and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Excel 2003 - SP2 with .NET Programmability installed, on Windows XP – SP2 or Windows Server 2003 with SP1
* .NET Framework 2.0
* Compatibility Pack  for Word, Excel, and PowerPoint 2007 File Formats (beta 2 or higher).

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Run the downloaded executable to install the add-in for Excel 2003.



Test the add-in 
-------------------------------
If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Excel 2003. It allows you to either import an ODF spreadsheet file or export your current working document as an ODF spreadsheet file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Excel as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).

Troubleshooting Guide:
-------------------------------
If you don't see the ODF entry, it could be due to missing features in Excel (see the Software Requirements listed above). After completing the ODF Add-in installation, you need to activate the Add-in by doing the following actions:
* Go to Tools ? Customize, and then click the Commands tab. 
* In the Categories box, click Tools. 
* Drag COM Add-Ins from the Commands box over the Tools menu. When the Tools menu displays the menu commands, point to the location where you want the COM Add-Ins command to appear on the menu, and then release the mouse button. 
* Click Close. 
* On the Tools menu, click COM Add-Ins and do any of the following: 
* To load an add-in, select the check box next to the add-in name in the Add-Ins available list. If the add-in you want isn't in the Add-Ins available list, click Add, locate the add-in, and then click OK

Conversion currently supports :
----------------------------------------------	

Basic and Advanced Table Model
Basic Text and Paragraph Formatting
Document Metadata
Document structure


Features/options lost in the ODF to OpenXML translation
-------------------------------------------------------

Angle orientation between (90,270) degrees
Automatic height of rows with merged cells
Font language
Custom optimal column width
Custom optimal row height
Non-Standard underline styles


Features/options lost in the OpenXML to ODF translation
-------------------------------------------------------

Columns above 256 ("IV")
Rows above 65536
Sheets above 256
Multiline vertically stacked text
'Justify' and 'Distributed' cell vertical text alignment
'Center Across Selection' and 'Distributed' cell horizontal
Accounting underline styles
