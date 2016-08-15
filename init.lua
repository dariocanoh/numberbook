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
--- - DEPENDENCIES : 
---		
---		zipwriter 0.1.5-1  (in luarocks)
---		lua-zlib  0.4-1	   (in luarocks)
---		struct 1.4-1 
--- 

-- load required libraries
local check = lide.core.base.check

-- load required funcitons
local isString = lide.core.base.isstring
local isNumber = lide.core.base.isnumber
local isObject = lide.core.base.isobject
local isTable  = lide.core.base.istable

-- load required classes
local Object       = lide.classes.object
local Numbersheet  = require 'numberbook.numbersheet'


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
	
	local format = self.Cobj:add_format {
		bold = fields.Bold or nil,
		font_size = fields.FontSize or nil,
		font_name = fields.FontName or nil,
		align     = fields.Align or nil,
	}

	format:set_font_name ( fields.FontName or nil )
	
	return format
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