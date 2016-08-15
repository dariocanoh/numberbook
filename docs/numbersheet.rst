.. /////// 2016/04/05 - Dario Cano [thdkano@gmail.com]

Numbersheet
==========

Clase auxiliar que permite escribir hojas dentro de los libros lua.



Constructor
-----------

.. code-block:: lua

 object Numberbook:new ( object objBook, string strName, number nID )


Argumentos
^^^^^^^^^^

Estos argumentos son recibidos por el constructor de la clase:

=============  =====================================================================================
  Argumento     Descripcion
=============  ===================================================================================== 
 objBook 		El objeto padre, el libro
 strName        El nombre de la hoja
 nID			El identificador unico de la hoja
=============  =====================================================================================

----------------------------------------------------------------------------------------------------


Metodos de la clase
-------------------
----------------------------------------------------------------------------------------------------

Numberbook:writeText
^^^^^^^^^^^^^^^^^^^^
   
   Escribir un texto en la celda especificada.

==========  ========================================================================================
  nil_ 		  Numberbook:writeText ( string_ strCell, string_ strText, object_ objFormat )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numberbook:mergeCells
^^^^^^^^^^^^^^^^^^^
   
   This method allows cells to be merged together so that they act as a single area.

   .. code-block:: lua

   		Numberbook:mergeCells('B3:D4', 'Merged Cells', merge_format)

==========  ========================================================================================
  nil_ 		  Numberbook:mergeCells( string_ strCells, string_ strText, object_ oFormat )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numberbook:setRowHeight
^^^^^^^^^^^^^^^^^^^^^^^
   
	Modifica el tamaño de la fila.


==========  ========================================================================================
  nil_ 		  Numberbook:setRowHeight( number_ nRow, number_ nHeight )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numberbook:setColWidth
^^^^^^^^^^^^^^^^^^^^^^
   
	Modifica el tamaño de la columna


==========  ========================================================================================
  nil_ 		  Numberbook:setColWidth( number_ first_col, number_ last_col, number_ width, object_ format )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numberbook:getCobj
^^^^^^^^^^^^^^^^^^
   
	Obtiene el objeto real.


==========  ========================================================================================
  nil_ 		  Numberbook:setRowHeight( number_ nRow, number_ nHeight )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

.. // Required values for html docs visualization
.. include:: ../directives.rst