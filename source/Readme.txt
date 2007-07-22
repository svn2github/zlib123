This is the source distribution of the "ODF Add-in for Microsoft Word" project.


Content
-------

The project is made of several subprojects and has the following struture:

+ lib                                Shared Dlls
+ packages                           Visual Studio Packages (patches for building setup executable)
+ source
  + Additional Tools                 External tools/source used by the converter
  | + OdfInstallHelper               Custom actions for the VS Windows Installer
  | + OdfWordXPAddInDriver           A COM façade for managed Add-ins (used to activate/deactivate the ODF Add-in in Word XP)
  | + Tenuto                         Validator RelaxNG
  | + zlib123                        ZLib source
  + Common                           Common projects
  | + OdfConverterLib                The main converter library
  | + OdfZipUtils                    A library used to zip/unzip files
  + Presentation                     Presentation converter
  + Spreadsheet			     Spreadsheet converter
  + Shell                            Shell and standalone converter
  | + OdfConverterLauncher           The project in charge of windows explorer contextual menu
  | + OdfConverterTest               A command line tool
  + Word                             Word Converters
    + Converter			     Word converter engine
    + Setup                          Setup projects
    | + OdfAddInForWordSetup         The ODF Add-in setup programs in english
    | + OdfAddInForWordSetup-xx      The ODF Add-in setup programs in different languages
    + OdfWordAddInLib                A library used by all the Word Add-in projects
    + OdfWordXXXAddIn                The ODF Add-in for Word 2007/2003/XP



There are also some additional tools that are used by the main projects (they are located in the AdditionalTools subfolder):
- OdfWordXPAddinDriver:
Note: it is not required to compile the additional tools in order to compile the main projects.


Requirements
------------

To compile the sources, you will need .NET Framework 2.0 SDK.
The Add-in setup programs (Word 2007, Word 2003 and Word XP) also require Visual Studio 2005.
To compile Word 2003 & Word XP Add-in setup programs, you will also need to patch Visual Studio, by following the instructions given at the following address: http://support.microsoft.com/kb/908002/en-us

Notes
-----

To generate the OdfAddInForWordSetup project, you will have to copy the KB908002, CompatibilityPack and DotNetFX folders (from the packages directory) to "<.NET 2.0 Framework SDK Folder>\Bootstrapper\Packages" folder
- OdfAddInForWordSetup/DotNetFX only contains localized strings, add these folders to the existing dotnetfx package folder
- OdfAddInForWordSetup/KB908002 only contains an updated product.xml and new tool which detects if the office patch has been installed
