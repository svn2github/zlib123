ODF Add-in for Word XP
================================

Thank you for downloading ODF Add-in for Word XP.
The purpose of the software is to enable users to open ODF Specification based documents in Word XP and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).


Software Requirements
--------------------------------
To test the add-in, you will need the following:

* Microsoft .NET framework 2.0 (http://www.microsoft.com/downloads/details.aspx?familyid=0856eacb-4362-4b0d-8edd-aab15c5e04f5&displaylang=en)
* Microsoft Word XP with Compatibility Pack for Office 2007 (http://www.microsoft.com/office/preview/beta/converter.mspx)
Note that the Compatibility Pack for Office 2007 requires the Service Pack 3 for Office XP to be installed.


Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Double click the the MSI file to install the add-in for Word XP.


Test the add-in 
-------------------------------
If installation is successful, you should see two new entries related to ODF (Open and Save) in the "File" menu in Word XP. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Word as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).


Troubleshooting Guide:
-------------------------------
If an error occurs when you try to open an ODF file, it might be that the Service Pack 3 for Office XP or the Compatibility Pack for Word 2007 is not installed (see Software Requirements).


