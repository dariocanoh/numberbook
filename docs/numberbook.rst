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

Numberbook:addSheet
^^^^^^^^^^^^^^^^^^^
   
   Agrega una hoja al documento

==========  ========================================================================================
 object_ 	 Numberbook:addSheet( string_ SheetName )
==========  ========================================================================================

----------------------------------------------------------------------------------------------------

Numberbook:close
^^^^^^^^^^^^^^^^
   
   xcclose

=======  =========================================================================================
 bool_ 	  Numberbook:setID( number_ nID )
=======  =========================================================================================

----------------------------------------------------------------------------------------------------


.. // Required values for html docs visualization
.. include:: ../directives.rst