if not exist "$(TargetDir)de-DE" md "$(TargetDir)de-DE"   
copy "$(TargetDir)$(TargetName).de-de.msi" "$(TargetDir)de-DE\$(TargetName).de-de.msi"
cd
cd "$(TargetDir)de-DE"
cd
echo IExpress /N "..\..\..\OdfAddInForOfficeSetup.de-DE.sed"
rem IExpress /N ..\..\..\OdfAddInForOfficeSetup.de-DE.sed
cd "$(TargetDir)"
    
copy "$(TargetDir)$(TargetName).en-us.msi" "$(TargetPath)"

"C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin\msitran.exe" -g "$(TargetPath)" "$(TargetDir)$(TargetName).de-de.msi" "$(TargetDir)de.mst"
cscript "C:\Program Files\Microsoft Platform SDK\Samples\SysMgmt\Msi\Scripts\wisubstg.vbs" "$(TargetName).msi" de.mst 1031

cscript "C:\Program Files\Microsoft Platform SDK\Samples\SysMgmt\Msi\Scripts\WiSumInf.vbs" OdfAddinForOfficeSetup.msi Template = Intel;1033,1031