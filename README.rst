Numberbook
==========

Constructor de libros de cálculo desde Lua.

=======================  =================================================================================
  Métodos de la clase      Descripción
=======================  =================================================================================
 Numberbook:newFormat     Crea un nuevo formato
 Numberbook:addSheet      Agrega una hoja al libro
 Numberbook:close    	  Cierra el libro
 Numberbook:getCobj 	  Obtiene el objeto real
=======================  =================================================================================

Numbersheet
===========

Clase auxiliar que representa una hoja del libro.

==========================  ===============================================================================
  Métodos de la clase         Descripción
==========================  ===============================================================================
 Numbersheet:writeText       Escribe un texto en la celda especificada
 Numbersheet:setColWidth     Modifica el tamaño de la columna
 Numbersheet:setRowHeight    Modifica el tamaño de la fila
 Numbersheet:mergeCells      Combina las celdas
 Numbersheet:getCobj 	     Obtiene el objeto real
==========================  ===============================================================================