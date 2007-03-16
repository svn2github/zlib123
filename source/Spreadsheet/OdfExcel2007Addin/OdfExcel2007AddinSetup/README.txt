ODF Add-in for Excel2007
========================

Thank you for downloading ODF Add-in for Excel XP. The purpose of the software is to enable users to open ODF Specification based documents in Excel XP and demonstrate the interoperability between OpenXML & ODF specifications. Since this is still a beta release, you may use and test it... at your own risk! Any feedback is welcome though! ;)

Software Requirements
---------------------
To test the add-in, you will need the following:
 1. Excel 2007
 2. .NET Framework 2.0*

IMPORTANT NOTES:

(*) Warning: if you install Microsoft Excel before the .NET Framework, you will need to update your Excel installation, by running the following these steps below.

   1. Open "Add or remove Programs" from the Control Panel
   2. Find the Microsoft Excel entry and choose "Modify"
   3. Within the installation process, choose "Add or remove features" then add ".NET Programmability Support" in the "Microsoft Office Excel" section.

Installing the add-in
---------------------
Run the downloaded install program and follow the steps...

If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Excel. It allows you to either import an ODF spreadsheet file or export your current working document as an ODF spreadsheet file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is imported in Microsoft Excel as a read-only DOCX file. If you want to save it back as ODF, you will have to first make a copy of the document ("Menu>Save As") and then use the "Export as ODF" menu.

Troubleshooting Guide
---------------------
If you don't see the ODF entry, it could be due to missing features in Excel (see the Software Requirements listed above). After completing the ODF Add-in installation,  you need to activate the Add-in by doing the following actions:

* Click "Excel Options" in the File menu
* In the "Add-ins" section, select "COM Add-ins" in the list box and press the "Go" button
* Make sure to select the checkbox "ODF Excel 2007 Add-in" and validate with the "OK" button.

Conversion currently supports
-----------------------------

For a complete list of features, please refer to our roadmap on sourceForge.net (https://sourceforge.net/project/showfiles.php?group_id=169337).


Features/options lost in the ODF to OpenXML translation
-------------------------------------------------------

Please to refer to sourceForge (https://sourceforge.net/tracker/?group_id=169337&atid=932582) for a detailed list of unsupported features.
