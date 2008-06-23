:: called with the following parameters:
::     %1 - path to project folder
::     %2 - BuiltOutputPath, e.g. .\Release\OdfAddInForWordSetup.msi
::     %3 - Configuration, e.g. "Release", "Release to Manufacturing (signed)"

:: add custom install actions
ECHO "Adding custom actions"
ECHO %1..\Scripts\AddCustomActions.vbs %2 ..\..\..\..\..\lib\OdfInstallHelper.dll
cscript.exe %1..\Scripts\AddCustomActions.vbs %2 ..\..\..\..\..\lib\OdfInstallHelper.dll


cscript.exe %1..\Scripts\CorrectConditions.vbs %2 %1..\Scripts\Conditions.xml
cscript.exe %1..\Scripts\patchForVista.vbs %2

::Bug #1789989  enable back button on wizard page 4
::removed bugfix 20080606/divo
::cscript.exe %1..\Scripts\AddProperty.vbs %2

::build self-extracting installer
%1..\Scripts\MakeSetupExe.bat %1%3 ..\OdfAddInForWordSetup.sed




