ODF Add-in for Excel XP
================================

Thank you for downloading ODF Add-in for Excel XP.
The purpose of the software is to enable users to open ODF Specification based documents in Excel XP and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Excel XP with SP3 on Windows XP with SP2 or Windows Server 2003 with SP1
* .NET Framework 2.0
* Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats.

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Run the downloaded executable to install the add-in for Excel XP.



Test the add-in 
-------------------------------
If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Excel XP. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Excel as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).

Troubleshooting Guide:
-------------------------------
If you don’t see the ODF entry, it could be due to the Macro Security Settings. To set proper security settings do the following:
1.	Go to Tools ? Options and click the Security tab.
2.	Under Macro Security Click “Macro Security…”
3.	Click the Trusted Sources tab in the Security Dialog.
4.	Check the Option “Trust all installed add-ins and templates”.
5.	Click “Ok” of the Security Dialog and also the Options Dialog.
6.	Close and re-open Excel to see the Open ODF… entry.

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

