.. /////// 2016/04/05 - Dario Cano [thdkano@gmail.com]

Numberbook
==========

Esta clase permite escribir archivos xlsx desde Lide.

----------------------------------------------------------------------------------------------------


Constructor
***********

La clase Numberbook tiene un constructor complejo:


.. code-block:: lua

 object Numberbook:new { 
 	string Name, 	number ID,
 	string Filename,
 }


Argumentos
^^^^^^^^^^

Estos argumentos son recibidos por la clase:

=============  =====================================================================================
  Argumento     Descripcion
=============  =====================================================================================
 ID 			The object identificator (optional)
 Name 		    The object name
 Filename       The object file to save
=============  =====================================================================================

----------------------------------------------------------------------------------------------------


Metodos de la clase
*******************

Estos metodos son definidos por la clase:

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
   


==========  ========================================================================================
  nil_ 		  Numberbook:setRowHeight( number_ nRow, number_ nHeight )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

.. // Required values for html docs visualization
.. include:: ../directives.rst