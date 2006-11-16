top_srcdir=../..

include $(top_srcdir)/build/common.mk

SOURCE = $(wildcard *.cs)

all-local: OdfZipUtils.dll


OdfZipUtils.dll : $(SOURCE)
	$(CSC) $(CSC_DEBUG) -debug+ -target:library \
	     -out:OdfZipUtils.dll $(SOURCE)

