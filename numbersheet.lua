-- load required libraries:
local check = lide.core.base.check

-- load required funcitons:
local isString = lide.core.base.isstring
local isNumber = lide.core.base.isnumber
local isObject = lide.core.base.isobject
local isTable  = lide.core.base.istable

-- load required classes:
local Object     = lide.classes.object

--- class definition:
local NumberSheet = class 'NumberSheet' : subclassof 'Object'
		: global ( false )

function NumberSheet:NumberSheet ( objBook, strName, nID )
	isTable(objBook) isString(strName)
	
	-- object Object:new ( string sObjectName, number nObjectID )
	self.super:init ( strName, nID or lide.core.base.newid() )

	-- add the current worksheet
	self.Cobj = objBook:add_worksheet( self.Name )
end

function NumberSheet:writeText ( strCell, strText, objFormat )
	isString(strCell) isString(strText)	

	self.Cobj:write_string(strCell, strText, objFormat)
end

-- set_column(first_col, last_col, width, format, options)
function NumberSheet:setColWidth( first_col, last_col, width, format, options )
	self.Cobj:set_column(first_col -1, last_col -1, width, format, options)
end

function NumberSheet:setRowHeight( nRow, nHeight )
	self.Cobj:set_row(nRow -1, nHeight, nil, nil)
end

-- merge_range(first_row, first_col, last_row, last_col, format)
--- ("B3:D4", "Merged Cells", merge_format)
function NumberSheet:mergeCells ( strCells, strText, oFormat )
	self.Cobj:merge_range(strCells, strText, oFormat)
end

function NumberSheet:getCobj( )
	return self.Cobj
end

return NumberSheet