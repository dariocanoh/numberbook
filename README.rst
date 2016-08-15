.. // Required values for html docs visualization
.. include:: docs/ids.rst

Para poder funcionar, ésta libreria requiere tener instalado el `framework Lide <http://github.com/lidesdk/framework>`_ 
en su máquina.

Esta libreria está basada en `xlsxwriter.lua <https://github.com/jmcnamara/xlsxwriter.lua>`_ , todos los creditos a su autor.



Numberbook
==========

Constructor de libros de cálculo desde Lua.

=======================  =================================================================================
  Métodos de la clase      Descripción
=======================  =================================================================================
 Numberbook:newFormat_    Crea un nuevo formato
 Numberbook:addSheet_     Agrega una hoja al libro
 Numberbook:close_    	  Cierra el libro
 Numberbook:getCobj_ 	  Obtiene el objeto real
=======================  =================================================================================



Numbersheet
===========

Clase auxiliar que representa una hoja del libro.

==========================  ===============================================================================
  Métodos de la clase         Descripción
==========================  ===============================================================================
 Numbersheet:writeText_      Escribe un texto en la celda especificada
 Numbersheet:setColWidth_    Modifica el tamaño de la columna
 Numbersheet:setRowHeight_   Modifica el tamaño de la fila
 Numbersheet:mergeCells_     Combina las celdas
 Numbersheet:getCobj_ 	     Obtiene el objeto real
==========================  ===============================================================================



Licencia
========

The MIT/X11 License.