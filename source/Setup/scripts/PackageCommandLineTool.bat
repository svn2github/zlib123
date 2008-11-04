rem Creates a ZIP package for the CommandLineTool
rem Requires the free command line zip tools for WinZip

rem %1 - $(TargetDir)
rem %2 - $(ConfigurationName)

echo Packaging CommandLineTool

setlocal
cd %1..\..\..\Shell\OdfConverter\bin\%2\

%1\..\..\scripts\minizip.exe -o CommandLineTool.zip OdfConverter.exe
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip OdfConverterLib.dll        
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip OdfZipUtils.dll            
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip PresentationConverter.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip PresentationConverter2Odf.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip PresentationConverter2Oox.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip SpreadsheetConverter2Odf.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip SpreadsheetConverter2Oox.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip SpreadsheetConverter.dll       
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip WordprocessingConverter.dll   
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip WordprocessingConverter2Odf.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip WordprocessingConverter2Oox.dll  
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip relaxngDatatype.dll
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip tenuto.core.dll
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip zlibwapi.dll
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip Readme.txt
%1\..\..\scripts\minizip.exe -a CommandLineTool.zip License.txt

move CommandLineTool.zip %1CommandLineTool.zip
endlocal
