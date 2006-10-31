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

OdfConverter.exe /I PathOrFilename [/O PathOrFilename] [BATCH-ODT] [BATCH-DOCX] [/V] [/OPEN] [/XSLT Path] [/NOPACKAGING] [/SKIP name] [/REPORT Filename] [/LEVEL Level]
    Where options are:
       /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)
       /O PathOrFilename  Name of the output file (or output folder)
       /BATCH-ODT         Do a batch conversion over every ODT file in the input folder (note: existing files will be replaced)
       /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (note: existing files will be replaced)
       /V                 Validate the result of the transformation against the schemas
       /OPEN              Try to open the converted files (works only for ODF->OOX, Microsoft Word required)
       /XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)
       /NOPACKAGING       Don't package the result of the transformation into a ZIP archive (produce raw XML)
       /SKIP name         Skip a post-processing (provide the post-processor's name)
       /REPORT Filename   Name of the report file that must be generated (existing files will be replaced)
       /LEVEL Level       Level of reporting: 1=DEBUG, 2=INFO, 3=WARNING, 4=ERROR


Known Issues
------------

	The OdfConverterTest command line tool only works on local drives.
