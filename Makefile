package := odf-converter-0.0.6

top_srcdir=.

SUBDIRS= \
	source/AdditionalTools/zlib123/contrib/minizip \
	source/Common/OdfZipUtils \
	source/Common/OdfConverterLib \
	source/Word/Converter \
	source/Presentation/Converter \
	source/Spreadsheet/Converter \
	source/Shell/OdfConverterTest
DISTFILES= \
	build/common.mk \
	source/AdditionalTools/zlib123/zconf.h \
	source/AdditionalTools/zlib123/zlib.h

include $(top_srcdir)/build/common.mk

dist-pre:
	rm -rf $(package)
	mkdir $(package)

dist-tarball: dist-pre
	$(MAKE) distdir='$(package)' dist-recursive
	rm -f $(package).tar.gz
	tar czf $(package).tar.gz $(package)

dist: dist-tarball

dist-local : dist-common

.PHONY: all-local $(RECURSIVE_TARGETS:=-local)
