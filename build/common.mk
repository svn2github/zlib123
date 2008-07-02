all: do-all

OS:=$(shell uname -o 2> /dev/null)

# uname doesn't have a -o option on all platforms
ifndef OS
OS:=$(shell uname)
endif

CC=gcc
LD=ld
CSC=gmcs
CSC_DEBUG=-d:DEBUG -debug+
RECURSIVE_TARGETS = check clean install uninstall

$(RECURSIVE_TARGETS): %: do-%

do-% : %-recursive
	$(MAKE) $*-local

%-recursive:
	@list='$(SUBDIRS)'; for dir in $$list ; do \
	    (cd $$dir && $(MAKE) $*) || { error_terminate="exit 1"; } ; \
	done; \
	$$error_terminate

dist-recursive: dist-local
	list='$(SUBDIRS)'; for dir in $$list ; do \
	    (cd $$dir && $(MAKE) distdir=$(distdir)/$$dir $@) || exit 1 ; \
	done

dist-common:
	-mkdir -p $(top_srcdir)/$(distdir)
	for file in Makefile $(DISTFILES) ; do \
	    install -D $$file $(top_srcdir)/$(distdir)/$$file || exit 1 ; \
	done
