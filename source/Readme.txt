This is the source distribution of the "ODF Add-in for Microsoft Office" project.


Content
-------

The project is made of several subprojects and has the following struture:

+ build
|
+ docs
|
+ lib                                Shared Dlls
|
+ packages                           Visual Studio Packages (patches for building setup executable)
|
+ signing                            A sample configuration for creating a digitally signed add-in
|
+ source
  + Additional Tools                 External tools/source used by the converter
  | + OdfInstallHelper               Custom actions for the VS Windows Installer
  | + ProcessHelper                  A helper executable fixing a problem with IExpress installations
  | + Tenuto                         Validator RelaxNG
  | + zlib123                        ZLib source
  |
  + Common                           Common projects
  | + OdfAddinLib                    A library containing the basic add-in logic and UI
  | + OdfConverterLib                The main converter library
  | + OdfZipUtils                    A library used to zip/unzip files
  |
  + Presentation                     Presentation converter and add-in
  | + Converter                      The XSLT-based converter engine
  | + OdfPowerPointAddin             The ODF Add-in for PowerPoint
  | + OdfPowerPointAddinShim         A custom add-in loader providing support for digital signed add-ins
  |
  + Setup                            A WiX 3.0.4624 project to create an MSI installer for the add-in
  |
  + Shell                            Shell and standalone converter
  | + OdfConverter                   A command line tool
  | + OdfConverterLauncher           The project in charge of windows explorer contextual menu
  |
  + Spreadsheet			     Spreadsheet converter and add-in
  | + Converter                      The XSLT-based converter engine
  | + OdfExcelAddin                  The ODF Add-in for Excel
  | + OdfExcelAddinShim              A custom add-in loader providing support for digital signed add-ins
  |
  + Word                             Word converter and add-in
  | + Converter			     The XSLT-based converter engine
  | + OdfWordAddin                   The ODF Add-in for Word 2007/2003/XP/2000
  | + OdfWordAddinShim               A custom add-in loader providing support for digital signed add-ins
  | + Templates                      A set of Word templates designed to be compatible with the ODF format



There are also some additional tools that are used by the main projects (they are located in the AdditionalTools subfolder):
These tools may be build separately and the latest version should be placed in the respective folder (e.g. \lib)


Requirements
------------

To compile the sources, you will need the .NET Framework 2.0 SDK.

The Add-in programs (Word 2007, Word 2003 and Word XP) also require Visual Studio 2008.

The setup project requires the Windows Installer XML Toolset 3.0.4624.

Notes
-----

To generate the setup project, you will have to copy the KB908002, CompatibilityPack and DotNetFX folders (from the packages directory) to "<.NET 2.0 Framework SDK Folder>\Bootstrapper\Packages" folder.
