IF "%~1" == "" GOTO ERROR
IF "%~2" == "" GOTO ERROR

mkdir %~2
for %%F in ("%~1"\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~2\"%%~nF"%%~xF.html

GOTO END

:ERROR
@echo off
echo Usage: "build.bat <input folder containing xsl files> <output folder>"

:END