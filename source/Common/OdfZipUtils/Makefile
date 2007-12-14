top_srcdir=../../..

include $(top_srcdir)/build/common.mk

SOURCE = $(wildcard *.cs)
DISTFILES = $(SOURCE)

all-local: OdfZipUtils.dll OdfZipUtilsStatic.dll
dist-local: dist-common

OdfZipUtils.dll : $(SOURCE)
	$(CSC) $(CSC_DEBUG) -target:library -r:System.Drawing \
	     -out:$@ $(SOURCE)

OdfZipUtilsStatic.dll : $(SOURCE)
	$(CSC) $(CSC_DEBUG) -d:STATIC -target:library -r:System.Drawing \
	     -out:$@ $(SOURCE)

clean-local:
	rm -f *.dll *.mdb *.pdb
