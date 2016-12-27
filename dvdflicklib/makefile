CFLAGS= -Wall -O3 -c
LIBS= -lgdi32

all:

	mkdir -p obj

	gcc $(CFLAGS) -o obj/util.o source/util.c
	gcc $(CFLAGS) -o obj/main.o source/main.c

	gcc -shared -o dvdflick.dll obj/*.o $(LIBS) -Wl,--add-stdcall-alias
	strip dvdflick.dll
	
	cp dvdflick.dll ../dvdflick.dll
	cp dvdflick.dll c:/windows/system32/dvdflick.dll


clean:

	rm -rf obj
	rm -f dvdflick.a