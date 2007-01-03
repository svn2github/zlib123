package := odf-converter-0.0.6

top_srcdir=.

SUBDIRS= \
	source/AdditionalTools/zlib123/contrib/minizip \
	source/OdfZipUtils \
	source/OdfConverterLib \
	source/OdfConverterTest
DISTFILES=build/common.mk

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
