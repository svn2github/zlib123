top_srcdir=.

SUBDIRS= \
	source/AdditionalTools/zlib123/contrib/minizip \
	source/OdfZipUtils \
	source/OdfConverterLib \
	source/OdfConverterTest

include $(top_srcdir)/build/common.mk

.PHONY: all-local $(RECURSIVE_TARGETS:=-local)
