This is the source distribution of the "ODF Add-in for Microsoft Word" project.


Content
-------

The project is made of several subprojects:
- OdfZipUtils: a library used to zip/unzip files
- OdfConverterLib: the main converter library
- OdfConverterTest: a command line tool
- OdfWordAddinLib: a library used by all the Add-in projects
- OdfWord2007Addin: the ODF Add-in for Word 2007
- OdfWord2007AddinSetup: the setup program for the ODF Add-in for Word 2007
- OdfWord2003Addin: the ODF Add-in for Word 2003
- OdfWord2003AddinSetup: the setup program for the ODF Add-in for Word 2003
- OdfWordXPAddin: the ODF Add-in for Word XP
- OdfWordXPAddinSetup: the setup program for the ODF Add-in for Word XP
- OdfAddInForWordSetup : a unified ODF Add-in setup program embracing WordXP, 2003 and 2007.

There are also some additional tools that are used by the main projects (they are located in the AdditionalTools subfolder):
- OdfWordXPAddinDriver: a COM façade for managed Add-ins (used to activate/deactivate the ODF Add-in in Word XP)
Note: it is not required to compile the additional tools in order to compile the main projects.


Requirements
------------

To compile the sources, you will need .NET Framework 2.0 SDK.
The Add-in setup programs (Word 2007, Word 2003 and Word XP) also require Visual Studio 2005.
To compile Word 2003 & Word XP Add-in setup programs, you will also need to patch Visual Studio, by following the instructions given at the following address: http://support.microsoft.com/kb/908002/en-us.
To generate the OdfAddInForWordSetup project, you will have to copy the content of KB908002 and DotNetFX folders (from the OdfForWordAddInSetup) folders to <.NET 2.0 Framework SDK Folder>\Bootstrapper\Packages folder
- OdfAddInForWordSetup/DotNetFX only contains localized strings , add these folders to the existing dotnetfx package folder
- OdfAddInForWordSetup/KB908002 only contains an updated product.xml and new tool which detects if the office patch has been installed