IF "%~1" == "" GOTO ERROR

mkdir %~1\Word\odf2oox
mkdir %~1\Word\oox2odf

for %%F in (..\..\OdfConverterLib\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Word\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\OdfConverterLib\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Word\oox2odf\"%%~nF"%%~xF.html

for %%F in (..\..\..\Word\Converter\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Word\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\..\Word\Converter\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Word\oox2odf\"%%~nF"%%~xF.html


mkdir %~1\Presentation\odf2oox
mkdir %~1\Presentation\oox2odf

for %%F in (..\..\OdfConverterLib\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Presentation\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\OdfConverterLib\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Presentation\oox2odf\"%%~nF"%%~xF.html

for %%F in (..\..\..\Presentation\Converter\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Presentation\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\..\Presentation\Converter\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Presentation\oox2odf\"%%~nF"%%~xF.html


mkdir %~1\Spreadsheet\odf2oox
mkdir %~1\Spreadsheet\oox2odf

for %%F in (..\..\OdfConverterLib\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Spreadsheet\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\OdfConverterLib\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Spreadsheet\oox2odf\"%%~nF"%%~xF.html

for %%F in (..\..\..\Spreadsheet\Converter\resources\odf2oox\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Spreadsheet\odf2oox\"%%~nF"%%~xF.html
for %%F in (..\..\..\Spreadsheet\Converter\resources\oox2odf\*.xsl) do msxsl.exe "%%F" xsl2html.xsl > %~1\Spreadsheet\oox2odf\"%%~nF"%%~xF.html

GOTO END

:ERROR
@echo off
echo Usage: "build.bat <output folder>"

:END