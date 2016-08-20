-- /////////////////////////////////////////////////////////////////////////////////////////////////
-- // Name:        numberbook/init.lua
-- // Purpose:     CellFormat class
-- // Author:      John McNamara <jmcnamara@cpan.org>
-- // Modified:    Hernán D. Cano <hdcano@protonmail.com>
-- // Created:     2016/08/16
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
local check = lide.core.base.check

--[[ load required funcitons
local isObject = lide.core.base.isobject
]]

local isString = lide.core.base.isstring
local isNumber = lide.core.base.isnumber
local isTable  = lide.core.base.istable

-- load required classes
local Object       = lide.classes.object

--- class definition:
local CellFormat

CellFormat = class 'CellFormat' -- : subclassof 'Object'
		
		: global ( false )

function CellFormat:CellFormat ( objBook, fields )
	--isTable(objBook); isTable(fields);
	
	--check.fields { 'string FontName' } -- , 'string FontColor', 'bool Bold', 'string Align'}

	--- object Object:new ( string sObjectName, number nObjectID )
	--self.super:init ( fields.Name or 'cell_format', nID or lide.core.base.newid() )

	-- add the current worksheet
	local props = {}
	props.font_color = fields.FontColor
	props.bold       = fields.Bold
	props.font_name  = fields.FontFace
	props.align      = fields.Align
	props.font_size  = fields.FontSize
	
	self.Cobj = objBook:getCobj() : add_format ( props )
end

function CellFormat:getCobj ( ... )
	return self.Cobj
end

--- set_font_name(string fontname)
function CellFormat:setFontName ( sFontName )
	isString(sFontName);
	self.Cobj:set_font_name (sFontName )
end

--- "#FFFFFF", "green"
--- 
--- set_bg_color(color)
function CellFormat:setBackgroundColor( sColor )
	isString(sColor);
	self.Cobj:set_bg_color(sColor)
end

---		possible borders:
---
---		0 	None 	0 	 
---		1 	Continuous 	1 	-----------
---		2 	Continuous 	2 	-----------
---		3 	Dash 	1 	- - - - - -
---		4 	Dot 	1 	. . . . . .
---		5 	Continuous 	3 	-----------
---		6 	Double 	3 	===========
---		7 	Continuous 	0 	-----------
---		8 	Dash 	2 	- - - - - -
---		9 	Dash Dot 	1 	- . - . - .
---		10 	Dash Dot 	2 	- . - . - .
---		11 	Dash Dot Dot 	1 	- . . - . .
---		12 	Dash Dot Dot 	2 	- . . - . .
---		13 	SlantDash Dot 	2 	/ - . / - .
---
--- set_border(style)
function CellFormat:setBorder( nBorderStyle )
	isNumber ( nBorderStyle );
	self.Cobj:set_border(nBorderStyle)
end

-----------------------------------------------------------
---		Horizontal			---		Vertical alignment	---
---		center				---			top				---
---		right				---			vcenter			---
---		fill				---			bottom			---
---		justify				---			vjustify		---
---		center_across		---							---
-----------------------------------------------------------
---
--- set_align(alignment)
function CellFormat:setAlign( sAlignment )
	isString ( sAlignment );
	self.Cobj:set_align(sAlignment)

	return self
end

--- set_italic()
function CellFormat:setItalic ( )
	self.Cobj:set_italic()
end

--- set_text_wrap()
function CellFormat:setTextWrap ( )
	self.Cobj:set_text_wrap()
end

function CellFormat:setTop( nBorderStyle)
	self.Cobj:set_top(nBorderStyle)
end
--- set_bottom()
function CellFormat:setBottom( nBorderStyle)
	self.Cobj:set_bottom(nBorderStyle)
end

function CellFormat:setLeft( nBorderStyle)
	self.Cobj:set_left(nBorderStyle)
end

function CellFormat:setRight( nBorderStyle)
	self.Cobj:set_right(nBorderStyle)
end

--[[local Arial_10 = CellFormat:new { FontName = 'Arial', FontSize = 10 }

Arial_10:setFontName 'Arial'
Arial_10:setBorder(9) --- {1,2,3,...9}
Arial_10:setTop(9)
Arial_10:setBottom(9)
Arial_10:setLeft(9)
Arial_10:setRight(9)
Arial_10:setBackgroundColor '#FFDDFF'

]]

return CellFormat

--[[
function CellFormat:writeText ( strCell, strText, objFormat )
	isString(strCell) isString(strText)	

	self.Cobj:write_string(strCell, strText, objFormat)
end

-- set_column(first_col, last_col, width, format, options)
function CellFormat:setColWidth( first_col, last_col, width, format, options )
	self.Cobj:set_column(first_col , last_col , width, format, options)
end

function CellFormat:setRowHeight( nRow, nHeight )
	self.Cobj:set_row(nRow -1, nHeight, nil, nil)
end

-- merge_range(first_row, first_col, last_row, last_col, format)
--- ("B3:D4", "Merged Cells", merge_format)
function CellFormat:mergeCells ( strCells, strText, oFormat )
	self.Cobj:merge_range(strCells, strText, oFormat)
end

function CellFormat:getCobj( )
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

	return CellFormat:new ( self.Cobj, strName )
end

function Numberbook:close( )
	self.Cobj:close()
end

function Numberbook:getCobj ( ... )
	return self.Cobj
end
]]
