top_srcdir=../../../../..

LD=ld
CC=cc
DEBUG=-g -O0
CFLAGS=-Wall

SOURCE = zip.c unzip.c ioapi.c
HEADER = $(wildcard *.h)
DISTFILES=$(SOURCE) $(HEADER)

include $(top_srcdir)/build/common.mk

all-local: libzlibwapi.so libzlibwapi.a

libzlibwapi.so : $(SOURCE) Makefile
	$(CC) -lz -shared -fPIC $(DEBUG) $(CFLAGS) \
	    -o libzlibwapi.so $(SOURCE)

libzlibwapi.a : $(SOURCE) Makefile
	$(CC) -c $(DEBUG) $(CFLAGS) $(SOURCE)
	ar rc libzlibwapi.a zip.o unzip.o ioapi.o # FIXME - do better

clean-local:
	rm -f *.o libzlibwapi.so libzlibwapi.a

dist-local: dist-common
