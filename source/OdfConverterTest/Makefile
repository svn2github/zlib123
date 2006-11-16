top_srcdir=../..
include $(top_srcdir)/build/common.mk

SOURCE = \
    OdfCommandLineException.cs \
    OdfConverterTest.cs \
    OoxValidator.cs \
    OdfValidator.cs

all-local: OdfConverterTest.exe OdfConverter
OdfConverterTest.exe : $(SOURCE) ../OdfZipUtils/OdfZipUtils.dll ../OdfConverterLib/OdfConverterLib.dll
	$(CSC) $(CSC_DEBUG) -d:MONO -target:exe -out:OdfConverterTest.exe \
		-reference:../OdfZipUtils/OdfZipUtils.dll \
		-reference:../OdfConverterLib/OdfConverterLib.dll \
		$(SOURCE)

OdfConverter : OdfConverterTest.exe Makefile
	rm -f OdfConverter
	MONO_PATH='../OdfZipUtils:../OdfConverterLib' mkbundle2 --static --deps OdfConverterTest.exe -o OdfConverter

check-local: all
	rm -f /tmp/foo.odt
	MONO_TRACE_LISTENER=Console.Error mono --debug ./OdfConverterTest.exe /LEVEL 1 /I /tmp/tst.docx /O /tmp/foo.odt 
	MONO_TRACE_LISTENER=Console.Error mono --debug ./OdfConverterTest.exe /LEVEL 1 /I /tmp/foo.odt /O /tmp/baa.docx 

clean-local:
	rm -f OdfConverterTest.exe

install-local: OdfConverter
	install -D OdfConverter $(DESTDIR)/usr/lib/ooo-2.0/program/OdfConverter
