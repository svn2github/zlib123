pushd %1
AddExcelCustomActions.vbs %2 ..\..\..\..\lib\OdfInstallHelper.dll
popd
::ECHO.%1AddExcelCustomActions.vbs %2 ..\..\..\..\lib\OdfInstallHelper.dll>"postBuild.log"

