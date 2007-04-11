ODF Add-in for Excel 2003
================================

Thank you for downloading ODF Add-in for Excel 2003. The purpose of the software is to enable users to open ODF Specification based documents in Excel 2003 and demonstrate the interoperability between OpenXML & ODF specifications. Since this is still a beta release, you may use and test it... at your own risk! Any feedback is welcome though! ;)


Software Requirements
---------------------
To test the add-in, you will need the following:
 1. Excel 2003
 2. .NET Framework 2.0*
 3. Compatibility Pack** for Word, Excel, and PowerPoint 2007 File Formats (beta 2 or higher).

IMPORTANT NOTES:

(*) Warning: if you install Microsoft Excel before the .NET Framework, you will need to update your Excel installation, by running the following these steps below.

   1. Open "Add or remove Programs" from the Control Panel
   2. Find the Microsoft Excel entry and choose "Modify"
   3. Within the installation process, choose "Add or remove features" then add ".NET Programmability Support" in the "Microsoft Office Excel" section.

(**) Please make sure you run the complete Office update before installing the compatibility pack. Office update is available from Office Online (http://office.microsoft.com/en-us/default.aspx)


Installing the add-in
---------------------

Run the downloaded setup.exe program and follow the steps...

If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Excel. It allows you to either import an ODF spreadsheet file or export your current working document as an ODF spreadsheet file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is imported in Microsoft Excel as a read-only XLSX file. If you want to save it back as ODF, you will have to first make a copy of the document ("Menu>Save As") and then use the "Export as ODF" menu.


Troubleshooting Guide
---------------------

If you don't see the ODF entry, it could be due to missing features in Excel (see the Software Requirements listed above). After completing the ODF Add-in installation, you need to activate the Add-in by doing the following actions:
* Go to "Tools > Customize", and then click the "Commands" tab.
* In the "Categories" box, click "Tools".
* Drag COM Add-Ins from the "Commands" box over the "Tools" menu. When the "Tools" menu displays the menu commands, point to the location where you want the COM Add-Ins command to appear on the menu, and then release the mouse button.
* Click Close.
* In the Tools menu, click "COM Add-Ins" and check the box next to the ODF add-in in the available list. If the add-in you want isn't in the list, click Add, locate the add-in, and then click OK.


Conversion currently supports
-----------------------------

For a complete list of features, please refer to our roadmap on sourceForge.net (https://sourceforge.net/project/showfiles.php?group_id=169337).


Features/options lost in the ODF to OpenXML translation
-------------------------------------------------------

Please to refer to sourceForge (https://sourceforge.net/tracker/?group_id=169337&atid=932582) for a detailed list of unsupported features.

Release notes
---------------

(1) The installation routine for M1 is in French only
(2) The installation works for the current user only.
(3) The AddIn does not work when Office 2007 and 2003 are installed in
parallel

Unresolved bugs in M1
---------------------

#1686766 Roundtrip Conversion Crash: Diagrammtypen.xlsx
#1673978 Columns not wide enough(OOX -> ODT)

unresolved bugs in M2
---------------------

#1698244 Add-in crash - corrupt ods file
#1697444 Command line tool cannot convert xls files
#1697386 Column widths
#1696803 Converter not responding
#1693493 Comments lost after conversion(large text content)
#1693462 File converter failed to open the file
#1693458 Total editing time not retained
#1692861 Thousand separators (conversion not proper for 16000 as 16)
#1692680 Add-in crashes when header contains large text
#1692669 Progress status in progress bar is not shown for direct and
#1692663 Text in cells lost after roundtrip conversion
#1686746 Format percentage becomes number
#1679589 Readme/help document
#1677196 Content lost(fixed lines)
#1677193 Content lost(first line)
#1674178 Cell content is lost(2jahre_onpsx.ods)