ODF Add-in for Word2003
================================

Thank you for downloading ODF Add-in for WordXP.
The purpose of the software is to enable users to open ODF Specification based documents in Word XP and demonstrate the interoperability between OpenXML & ODF specifications. 
This add-in is under development (see the roadmap at http://sourceforge.net/projects/odf-converter for more details about planned releases).
We will provide stable nightly builds. You can install and test them at your own risk!

Software Requirements
--------------------------------
To test the add-in, you will need the following:
* Word XP – SP3 on Windows XP – SP2 or Windows 2003
* .NET Framework 2.0
* Office 2007 File Format Awareness Update for Microsoft Office XP
* Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats (beta 2 or higher).
* Office XP PIA

Download and install the add-in
-------------------------------
Visit the Download Page and download the latest binary distribution of the add-in.  
Run the setup to install the add-in for Word XP.



Test the add-in 
-------------------------------
If installation is successful, you should see a new "Open ODF" entry in the “File” menu in Word XP. It allows you to either import an ODF text file or export your current working document as an ODF text file (note that during development process, those functionalities might be temporary unavailable).

Important note: The ODF file opened by the add-in is converted into Office OpenXML (Office 2007 new file format) and imported into Word as a read-only file. If you want to save it as ODF, you have to use the "Export as ODF" button and provide a new file name (that can be the same as the current file name).

Troubleshooting Guide:
-------------------------------
If you don’t see the ODF entry, it could be due to the Macro Security Settings. To set proper security settings do the following:
1.	Go to Tools ? Options and click the Security tab.
2.	Under Macro Security Click “Macro Security…”
3.	Click the Trusted Sources tab in the Security Dialog.
4.	Check the Option “Trust all installed add-ins and templates”.
5.	Click “Ok” of the Security Dialog and also the Options Dialog.
6.	Close and re-open Word to see the Open ODF… entry.
Still if you don't see the ODF entry, it could be due to missing features in Word (see the Software Requirements listed above). After completing the ODF Add-in installation, you need to activate the Add-in by doing the following actions:
1.	Go to Tools ? Customize, and then click the Commands tab. 
2.	In the Categories box, click Tools. 
3.	Drag COM Add-Ins from the Commands box over the Tools menu. When the Tools menu displays the menu commands, point to the location where you want the COM Add-Ins command to appear on the menu, and then release the mouse button. 
4.	Click Close. 
5.	On the Tools menu, click COM Add-Ins and do any of the following: 
5.1.	To load an add-in, select the check box next to the add-in name in the Add-Ins available list. If the add-in you want isn't in the Add-Ins available list, click Add, locate the add-in, and then click OK.
