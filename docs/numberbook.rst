.. /////// 2016/04/05 - Dario Cano [thdkano@gmail.com]

Numberbook
==========

Esta clase permite escribir archivos xlsx utilizando el framework Lide.

*Ejemplo:*

.. code-block: lua
	
 numberbook = require 'numberbook'
 book = numberbook : new { Name = 'book', Filename = 'test.xlsx'}
 book:writeText ('A1', 'Hola mundo')
 book:close()



Constructor
-----------

La clase Numberbook tiene un constructor complejo:


.. code-block:: lua

 object Numberbook:new { 
 	string Name, 	string Filename,
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



Metodos de la clase
-------------------
----------------------------------------------------------------------------------------------------

Numberbook:newFormat
^^^^^^^^^^^^^^^^^^^^
   
   Crea un nuevo formato de celda para el libro.

   * Posibles opciones *

   .. code-block: lua

   		fields = { string FontName, number FontSize, bool Bold }

==========  ========================================================================================
 object_ 	 Numberbook:newFormat( table_ fields )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

Numberbook:addSheet
^^^^^^^^^^^^^^^^^^^
   
   Agrega una hoja al documento

==========  ========================================================================================
 object_ 	 Numberbook:addSheet( string_ SheetName )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

Numberbook:close
^^^^^^^^^^^^^^^^
   
   Cierra el libro

=======  ===========================================================================================
 nil_ 	  Numberbook:close()
=======  ===========================================================================================

----------------------------------------------------------------------------------------------------

Numberbook:getCobj
^^^^^^^^^^^^^^^^^^
   
   Obtiene el objeto real

=========  =========================================================================================
 object_ 	 Numberbook:getCobj()
=========  =========================================================================================



.. // Required values for html docs visualization
.. include:: ../directives.rst