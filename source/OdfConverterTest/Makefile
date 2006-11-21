top_srcdir=../..
include $(top_srcdir)/build/common.mk

SOURCE = \
    OdfCommandLineException.cs \
    OdfConverterTest.cs \
    OoxValidator.cs \
    OdfValidator.cs
DISTFILES=$(SOURCE)

all-local: OdfConverterTest.exe OdfConverter
OdfConverterTest.exe : $(SOURCE) ../OdfZipUtils/OdfZipUtils.dll ../OdfConverterLib/OdfConverterLib.dll
	$(CSC) $(CSC_DEBUG) -d:MONO -target:exe -out:OdfConverterTest.exe \
		-reference:../OdfZipUtils/OdfZipUtils.dll \
		-reference:../OdfConverterLib/OdfConverterLib.dll \
		$(SOURCE)

OdfConverterTestStatic.exe : $(SOURCE) ../OdfZipUtils/OdfZipUtilsStatic.dll ../OdfConverterLib/OdfConverterLib.dll
	$(CSC) $(CSC_DEBUG) -d:MONO -target:exe -out:OdfConverterTestStatic.exe \
		-reference:../OdfZipUtils/OdfZipUtilsStatic.dll \
		-reference:../OdfConverterLib/OdfConverterLib.dll \
		$(SOURCE)

main.c bundle.o : OdfConverterTestStatic.exe
	MONO_PATH='../OdfZipUtils:../OdfConverterLib' mkbundle2 -c -o main.c -oo bundle.o --deps OdfConverterTestStatic.exe

OdfConverter : main.c bundle.o Makefile
	rm -f OdfConverter
	$(CC) -o OdfConverter -Wall `pkg-config --cflags mono` main.c \
	    `pkg-config --libs-only-L mono` -L../AdditionalTools/zlib123/contrib/minizip \
	    -Wl,--export-dynamic -lz -Wl,--undefined=zipOpen -Wl,--undefined=unzOpen \
	    -Wl,-Bstatic -lmono -lzlibwapi -Wl,-Bdynamic  \
	    `pkg-config --libs-only-l mono | sed -e "s/\-lmono //"` bundle.o

check-local: all
	rm -f /tmp/foo.odt
	MONO_TRACE_LISTENER=Console.Error ./OdfConverter /LEVEL 1 /I /tmp/tst.docx /O /tmp/foo.odt 
	MONO_TRACE_LISTENER=Console.Error ./OdfConverter /LEVEL 1 /I /tmp/foo.odt /O /tmp/baa.docx 

clean-local:
	rm -f OdfConverterTest.exe OdfConverterTest *.mdb

install-local: OdfConverter
	install -D OdfConverter $(DESTDIR)/usr/lib/ooo-2.0/program/OdfConverter
dist-local: dist-common
