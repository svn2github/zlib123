ODF Command Line Translator
================

Thank you for downloading ODF Command Line Translator(OdfConverterTest)

The purpose of the software is to enable users to open ODF Specification based documents in Word 2007 and demonstrate the interoperability between OpenXML & ODF specifications. 

The ODF Command Line Translator is a tool for translating ODF documents from the command line to Office Open XML format.

This software is still under development (see the roadmap on http://sourceforge.net/projects/odf-converter for more details about planned releases).

We will provide stable nightly builds. You can install and test them at your own risk!


Software Requirements
---------------------


To test the program, you need :

* .NET framework 2.0 (http://www.microsoft.com/downloads/details.aspx?familyid=0856eacb-4362-4b0d-8edd-aab15c5e04f5&displaylang=en)


OdfConverterTest Usage
----------------------

	OdfConverterTest.exe /I Filename [/O PathOrFilename] [/V] [/XSLT Path]
  Where:
	/I Filename        Name of the ODF file to transform
	/O PathOrFilename  Path of the folder where to put the output file (must exist) or name of the output file
	/V                 Validate the result of the transformation against OpenXML schemas
	/DEBUG             Debug mode
	/XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)


Known Issues
------------

	The OdfConverterTest command line tool only works on local drives.
