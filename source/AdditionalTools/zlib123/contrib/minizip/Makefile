top_srcdir=../../../../..

LD=ld
CC=cc
DEBUG=-g -O0
CFLAGS=-Wall

SOURCES = zip.c unzip.c ioapi.c

include $(top_srcdir)/build/common.mk

all-local: libzlibwapi.so

libzlibwapi.so : $(SOURCES)
	$(CC) -lz -shared $(DEBUG) $(CFLAGS) \
	    -o libzlibwapi.so $(SOURCES)

clean-local:
	rm -f *.o libzlibwapi.so


