%1AddCustomActions.vbs %2 ..\..\..\..\..\lib\OdfInstallHelper.dll
cscript.exe %1CorrectConditions.vbs %2 %1Conditions.xml
cscript.exe %1patchForVista.vbs %2
%1MakeSetupExe.bat %1%3

::ECHO.%1AddCustomActions.vbs %2 ..\..\..\..\..\lib\OdfInstallHelper.dll>"postBuild.log"
::ECHO.%1MakeSetupExe.bat %1%3>>"postBuild.log"

