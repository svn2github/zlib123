ODF Add-in for Excel2007
================================

Thank you for downloading ODF Add-in for Excel2007.
The purpose of the software is to enable users to open ODF Specification based documents in Excel 2007 and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Excel 2007 RTM
* .NET Framework 2.0
Warning: if you install Excel 2007 before the .NET Framework, you will need to update your Excel installation after installing the  .NET Framework, by running the following these steps below.
* Open "Add or remove Programs" from the Control Panel
* Find the Excel 2007 entry and choose "Modify"
* Within the installation process, choose "Add or remove features" then add ".NET Programmability Support" in the "Microsoft Office Excel" section. 

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Double click the MSI file to install the add-in for Excel 2007.



Test the add-in 
-------------------------------
If installation is successful, you should see a new "ODF" entry in the “File” menu in Excel 2007. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Excel as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).

Troubleshooting Guide:
-------------------------------
If you don't see the ODF entry, it could be due to missing features in Excel (see the Software Requirements listed above). After completing the ODF Add-in installation,  you need to activate the Add-in by doing the following actions:

* Click "Excel Options" in the File menu
* In the "Add-ins" section, select "COM Add-ins" in the list box and press the "Go" button
* Make sure to select the checkbox "ODF Excel 2007 Add-in" and validate with the "OK" button.

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

