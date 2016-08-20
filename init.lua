-- /////////////////////////////////////////////////////////////////////////////////////////////////
-- // Name:        numberbook/init.lua
-- // Purpose:     Numberbook class definition: is a interface of xlsxwriter.lua
-- // Author:      John McNamara <jmcnamara@cpan.org>
-- // Modified:    Hernán D. Cano <hdcano@protonmail.com>
-- // Created:     2016/08/14
-- // Copyright:   (c) 2016 Hernán D. Cano
-- // License:     lide license
-- /////////////////////////////////////////////////////////////////////////////////////////////////
-- 
--- - CREDITS:
--
-- 	Numberbook class is a fork of xlsxwriter.lua project, this fork is created because we need more
--  integration with Lide framework and John's library, thank you very much to the original author.
--
--	* xlsxwriter.lua 0.0.6-1
--
--	 The MIT/X11 License.
-- 	 Copyright 2014-2015, John McNamara <jmcnamara@cpan.org>
--   https://github.com/jmcnamara/xlsxwriter.lua
--  

-- load required libraries
local check 	 = lide.core.base.check

-- load required funcitons
local isString = lide.core.base.isstring
local isNumber = lide.core.base.isnumber
local isObject = lide.core.base.isobject
local isTable  = lide.core.base.istable

-- load required classes
local Object       = lide.classes.object

--- class definition:

local Numbersheet = class 'Numbersheet' : subclassof 'Object'
		: global ( false )

function Numbersheet:Numbersheet ( objBook, strName, nID )
	isTable(objBook) isString(strName)
	
	-- object Object:new ( string sObjectName, number nObjectID )
	self.super:init ( strName, nID or lide.core.base.newid() )

	-- add the current worksheet
	self.Cobj = objBook:add_worksheet( self.Name )
end

function Numbersheet:writeText ( strCell, strText, objFormat )
	isString(strCell) isString(strText)	
	
	if objFormat and isObject(objFormat) then 
		objFormat = objFormat:getCobj()
	end

	self.Cobj:write_string(strCell, strText, objFormat)
end

-- set_column(first_col, last_col, width, format, options)
function Numbersheet:setColWidth( first_col, last_col, width, format, options )
	if format and isObject(format) then
		format = format:getCobj()
	end

	self.Cobj:set_column(first_col , last_col , width, format, options)
end

function Numbersheet:setRowHeight( nRow, nHeight )
	self.Cobj:set_row(nRow -1, nHeight, nil, nil)
end

-- merge_range(first_row, first_col, last_row, last_col, format)
--- ("B3:D4", "Merged Cells", merge_format)
function Numbersheet:mergeCells ( strCells, strText, oFormat )
	isString(strCells); isString(strText); isObject(oFormat);
	
	self.Cobj:merge_range(strCells, strText, oFormat:getCobj())
end

function Numbersheet:getCobj( )
	return self.Cobj
end


--- class Definition
--
--	 	Numberbook : new { 
--			string Name, string Filename
--	 	}
--	
--- public fields
--	
--

local Numberbook = class 'Numberbook' : subclassof 'Object'
	: global(false)

function Numberbook:Numberbook ( fields )

	check.fields {
		'string Name', 'string Filename'
	}

	private {
		Filename = fields.Filename,
		Cobj     = {}
	}

	-- object Object:new ( string sObjectName, number nObjectID )
	self.super:init ( fields.Name, fields.ID  or lide.core.base.newid() )
	
	-- load required libraries locally 	
	local Workbook = require 'numberbook.xlsxwriter.workbook'	

	--- create the current object:
	-- table Workbook:new( filename[,options] )
	self.Cobj = Workbook:new ( self.Filename, nil )
end

--[table *format] workbook:add_format(table [properties])
function Numberbook:newFormat ( fields )
	return require 'numberbook.cellformat' : new ( self, fields )
end

-- worksheet workbook:add_worksheet([sheetname])
function Numberbook:addSheet( strName )
	isString(strName)

	return Numbersheet:new ( self.Cobj, strName )
end

function Numberbook:close( )
	self.Cobj:close()
end

function Numberbook:getCobj ( ... )
	return self.Cobj
end

return Numberbook