.. /////// 2016/04/05 - Dario Cano [thdkano@gmail.com]

Numbersheet
==========

Clase auxiliar que permite escribir hojas dentro de los libros lua.



Constructor
-----------

.. code-block:: lua

 object Numbersheet:new ( object objBook, string strName, number nID )


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

Numbersheet:writeText
^^^^^^^^^^^^^^^^^^^^
   
   Escribir un texto en la celda especificada.

==========  ========================================================================================
  nil_ 		  Numbersheet:writeText ( string_ strCell, string_ strText, object_ objFormat )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numbersheet:mergeCells
^^^^^^^^^^^^^^^^^^^
   
   This method allows cells to be merged together so that they act as a single area.

   .. code-block:: lua

   		Numbersheet:mergeCells('B3:D4', 'Merged Cells', merge_format)

==========  ========================================================================================
  nil_ 		  Numbersheet:mergeCells( string_ strCells, string_ strText, object_ oFormat )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numbersheet:setRowHeight
^^^^^^^^^^^^^^^^^^^^^^^
   
	Modifica el tamaño de la fila.


==========  ========================================================================================
  nil_ 		  Numbersheet:setRowHeight( number_ nRow, number_ nHeight )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numbersheet:setColWidth
^^^^^^^^^^^^^^^^^^^^^^
   
	Modifica el tamaño de la columna


==========  ========================================================================================
  nil_ 		  Numbersheet:setColWidth( number_ first_col, number_ last_col, number_ width, object_ format )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------


Numbersheet:getCobj
^^^^^^^^^^^^^^^^^^
   
	Obtiene el objeto real.


==========  ========================================================================================
  nil_ 		  Numbersheet:setRowHeight( number_ nRow, number_ nHeight )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

.. // Required values for html docs visualization
.. include:: ../directives.rst