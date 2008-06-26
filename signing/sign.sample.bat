@echo off

:: configure project here
set PROJECT_NAME=OpenXML/ODF Translator for Office
set PROJECT_URL=http://odf-converter.sourceforge.net

:: configure certificate here
:: CERT_FILE must contain the full path to the certificate, 
:: %~p0 expands to the path this script is in
set CERT_FILE=%~p0OdfConverter.sample.pfx
set CERT_PASSWORD=odfconverter

:: optional parameter for timestamping
::set TIMESTAMP_SERVER=/t http://timestamp.verisign.com/scripts/timstamp.dll

:: path to signtool.exe
set SIGNTOOL=C:\Program Files\Microsoft Visual Studio 8\SDK\v2.0\bin\signtool.exe

IF NOT EXIST "%SIGNTOOL%" goto :error

@echo on 
"%SIGNTOOL%" sign /d "%PROJECT_NAME%" /du "%PROJECT_URL%" %TIMESTAMP_SERVER% /f "%CERT_FILE%" /p "%CERT_PASSWORD%" /v "%~f1"
@echo off

exit /b 0

:error
echo ERROR: signtool.exe not found (Please configure the path correctly in file %~f0).
