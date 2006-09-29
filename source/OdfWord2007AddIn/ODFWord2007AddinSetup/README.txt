ODF Add-in for Word2007
================================

Thank you for downloading ODF Add-in for Word2007.
The purpose of the software is to enable users to open ODF Specification based documents in Word 2007 and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Word 2007 beta 2 Technical Refresh
* .NET Framework 2.0
Warning: if you install Word 2007 before the .NET Framework, you will need to update your Word installation after installing the  .NET Framework, by running the following these steps below.
* Open "Add or remove Programs" from the Control Panel
* Find the Word 2007 entry and choose "Modify"
* Within the installation process, choose "Add or remove features" then add ".NET Programmability Support" in the "Microsoft Office Word" section. 

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Double click the MSI file to install the add-in for Word 2007.



Test the add-in 
-------------------------------
If installation is successful, you should see a new "ODF" entry in the “File” menu in Word 2007. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Word as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).

Troubleshooting Guide:
-------------------------------
If you don't see the ODF entry, it could be due to missing features in Word (see the Software Requirements listed above). After completing the ODF Add-in installation,  you need to activate the Add-in by doing the following actions:

* Click "Word Options" in the File menu
* In the "Add-ins" section, select "COM Add-ins" in the list box and press the "Go" button
* Make sure to select the checkbox "ODF Word 2007 Add-in" and validate with the "OK" button.
