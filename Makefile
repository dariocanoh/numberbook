
.PHONY: docs test

all:
	@echo "\tTo install run: sudo make install"
	@echo "\tRequires luarocks."

install:
	@luarocks make

docs:
	@make -C dev/docs html

html: docs

pdf:
	@make -C dev/docs latexpdf

cleandocs:
	@make -C dev/docs clean

alldocs: cleandocs docs pdf
	@cp -r dev/docs/build/html docs
	@cp -r dev/docs/build/latex/xlsxwriter_lua.pdf docs

pdf_release: pdf
	@cp -r dev/docs/build/latex/xlsxwriter_lua.pdf docs


test:
	@lua -v
	@prove --exec=lua --ext=lua -r test/unit

testall: test5.1 test5.2 testluajit

test5.1:
	@lua5.1 -v
	@prove --exec=lua5.1 --ext=lua -r test/unit

test5.2:
	@lua5.2 -v
	@prove --exec=lua5.2 --ext=lua -r test/unit

testluajit:
	@luajit -v
	@prove --exec=luajit --ext=lua -r test/unit

readthedocs:
	@curl -X POST http://readthedocs.org/build/xlsxwriterlua
