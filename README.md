# App Mod Excel Juancho

aplicacion de escritorio para hacer modificacion de un archivo de excel con una estructura ya definida.

inicialmente se intento una solucion en python pero las librerias disponibles en python para modificar archivos de
excel fueron muy precarias para copiar imagenes en celdas por lo que cambie a usar macros con VBA, pero como el
usa WPS en mac, la solucion funciono bien en windows pero presento inconvenientes para instalar el soporte de VBA
para WPS en mac. asi que aunque los macros de excel funcionaron mucho mejor para copiar imagenes en las celdas,
se requiere una solucion mejor y que tenga interfaz grafica tambien para que sea mas amigable.

en este momento parece que la mejor opcion es usar .net, avalon y c#, ya que se puede usar la libreria EPPplus
que promete ser mas robusta que los intentos anteriores.
