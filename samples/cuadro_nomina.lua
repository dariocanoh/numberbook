-- lide version : 0.01
-- 
--

package.path = './?/init.lua;' .. package.path

Numberbook = require 'numberbook'

report = Numberbook:new { Name = 'report', Filename ='./report.xlsx' }

sheet1 = report:addSheet 'sheet1';

format1 = report:newFormat {
	FontName = 'Arial', Bold = true, FontSize = 16,
	Align = 'center',
}

format2 = report:newFormat {
	FontName = 'Arial', Bold = true, FontSize = 10,
	--> another option Align = 'vcenter',
}

format3 = report:newFormat {
	FontName = 'Arial', FontSize = 9,
	--> another option Align = 'vcenter',
}

format3:set_align 'center'
format3:set_italic()
format2:set_align('vcenter') -- combine aligns
format2:set_align('justify' ) -- combine aligns
format2:set_align('center' ) -- combine aligns
format2:set_text_wrap()

sheet1:setColWidth(1, 1, 3)

sheet1:setRowHeight(3, 20)
sheet1:mergeCells('A3:S3', ' NOMINA PARA EL PAGO DE SUELDOS', format1)

sheet1:mergeCells('B4:B7', 'NUM'   , format2)
sheet1:setColWidth(2, 2, 5)

sheet1:mergeCells('C4:C7', 'CEDULA', format2)
sheet1:setColWidth(3, 3, 15)

sheet1:mergeCells('D4:D7', 'NOMBRE', format2)
sheet1:setColWidth(4, 4, 30)

sheet1:mergeCells('E4:E7', 'SUELDO BASICO MENSUAL', format2)
sheet1:setColWidth(5, 5, 15)

sheet1:mergeCells('F4:L4', 'DEVENGADOS', format2)

sheet1:mergeCells('F5:F7', 'DIAS LABORADOS', format2)
sheet1:setColWidth(6, 6, 10)

sheet1:mergeCells('G5:G7', 'VALOR QUINCENA', format2)
sheet1:setColWidth(7, 7, 15)

sheet1:mergeCells('H5:H7', 'EXTRAS Y RECARGOS', format2)
sheet1:setColWidth(8, 8, 13)

sheet1:mergeCells('I5:I7', 'COMISIONES', format2)
sheet1:setColWidth(9, 9, 13)

sheet1:mergeCells('J5:J7', 'BONIFICACIONES', format2)
sheet1:setColWidth(10, 10, 15)

sheet1:mergeCells('K5:K7', 'AUXILIO DE TRANSPORTE', format2)
sheet1:setColWidth(11, 11, 15)

sheet1:mergeCells('L5:L7', 'TOTAL DEVENGADOS', format2)
sheet1:setColWidth(12, 12, 15)

sheet1:mergeCells('M4:S4', 'DEDUCCIONES', format2)

sheet1:mergeCells('M5:M7', 'SALUD\n\n4%', format2)
sheet1:setColWidth(13, 13, 13)

sheet1:mergeCells('N5:N7', 'PENSION\n\n4%', format2)
sheet1:setColWidth(14, 14, 13)

sheet1:mergeCells('O5:O7', 'FONDO DE SOLIDARIDAD', format2)
sheet1:setColWidth(15, 15, 13)

sheet1:mergeCells('P5:P7', 'RETENCION EN LA FUENTE SALARIOS', format2)
sheet1:setColWidth(16, 16, 13)

sheet1:mergeCells('Q5:Q7', 'APORTE A SINDICATO\n\n2%', format2)
sheet1:setColWidth(17, 17, 15)

sheet1:mergeCells('R5:R7', 'TOTAL DEDUCCIONES', format2)
sheet1:setColWidth(18, 18, 15)

sheet1:mergeCells('S5:S7', 'NETO A PAGAR', format2)
sheet1:setColWidth(19, 19, 17)

--writeText ( string_ strCell, string_ strText, object_ objFormat )
sheet1:writeText('G8', '001', format3)
sheet1:writeText('H8', '007', format3)
sheet1:writeText('I8', '030', format3)
sheet1:writeText('J8', '031', format3)
sheet1:writeText('K8', '002', format3)
sheet1:writeText('M8', '010', format3)
sheet1:writeText('N8', '011', format3)
sheet1:writeText('O8', '012', format3)
sheet1:writeText('P8', '050', format3)
sheet1:writeText('Q8', '020', format3)

os.execute 'rm ./report.xlsx'

report:close()