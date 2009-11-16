package := odf-converter-3.0.5270

top_srcdir=.

SUBDIRS= \
	source/AdditionalTools/zlib123/contrib/minizip \
	source/Common/OdfZipUtils \
	source/Common/OdfConverterLib \
	source/Word/Converter \
	source/Presentation/Converter \
	source/Spreadsheet/Converter \
	source/Shell/OdfConverter
DISTFILES= \
	LICENSE.TXT \
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
	find $(package) -name '*.cs' -o -name '*.xsl' | while read F; do dos2unix $$F; done
	tar czf $(package).tar.gz $(package)

dist: dist-tarball

dist-local : dist-common

.PHONY: all-local $(RECURSIVE_TARGETS:=-local)
