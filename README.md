# Dis.Reportes.Automatizados.

Esta es una aplicación de biblioteca de clases desarrollada con el lenguaje de programación C# usando la versión de DotNet Framework 4.7.2, que tiene el propósito de facilitar la tarea de crear reportes automatizados en forma de documentos de Word.


## Librerías que implementa el programa:
* ### OpenXml:
  Es una biblioteca que te permite trabajar con documentos Office (Word, Excel, PowerPoint) en su formato XML sin necesidad de tener Office instalado. Interactúa directamente con el formato de archivo Office XML (DOCX, XLSX, PPTX), lo que significa que no necesita iniciar una instancia de Word, Excel o PowerPoint para trabajar con los archivos. Te permite crear, modificar y leer documentos directamente a nivel de archivo.
  
* ### Microsoft Office Interop Word:
  El conjunto de ensamblados de Office Interop permite a las aplicaciones .NET interactuar con aplicaciones de Microsoft Office, incluido Word. Estos ensamblados funcionan como un puente entre la aplicación .NET y Office.


# Clases y métodos del programa Dis.Reportes.Automatizados.
## Clase FuncionesCreacion.
En esta clase se definen los métodos que deben ser implementados desde un inicio para todas las tareas que tienen que ver con la creación del documento de Word y el guardado de la ruta dentro del directorio del equipo.

  * ### Método GuardarRuta.
    Se usan los métodos propios del sistema para acceder a las rutas o directorios del explorador de archivos, es importante resaltar que la biblioteca de clases lo que va a hacer es abrir una ventana emergente del explorador de archivos de Windows con una dirección que se puede modificar y un nombre predeterminado que viene en el método, por ende si la biblioteca de clases se implementa desde por ejemplo una aplicación de consola esta no abrirá la ventana emergente del explorador de archivos, así que la mejor forma de inicializar el proyecto es por medio de proyectos que si puedan tener estas funcionalidades como por ejemplo una aplicación de WPF.
    
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/861ecb3d-173e-4974-a4bb-95f2dccb4434)

    
  * ### Método GenerarDocumentoWord.
    En este método se definen las propiedades que tendrá el documento de word, para ello se recibe como parámetros la ruta en la cual está el documento a modificar y un enum son las dimensiones de la hoja que puede tener
 el documento en este caso: A3, A4, A5, B3, B4.

    ![generar_documento](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/481c9db2-806a-4d03-98dc-778d64c721db)

    
  * ### Método CambiarOrientacion.
    Método que usa la librería de OpenXML para agregar una sección con propiedades específicas en este caso el cambio de orientación de las hojas del documento, en este punto es importante aclarar que como la librería OpenXML no funciona modificando por medio de la herramienta Word en términos rigurosos no se está haciendo un cambio de orientación como tal sino que se hace un reajuste a las dimensiones de la página del documento, estos valores se pueden cambiar modificando el método, motivo por el cual está la posibilidad de pérdida de formato, por ese motivo con este método es recomentable no hacer uso de pies de páginas y encabezados personalizados, lo recomendable es uno general o en blanco.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/be326e30-c0e3-480c-a5e6-ecf4776e9b04)

    
  * ### Método CambiarOrientacionPaginaEnDocumento.
    A diferencia del método anterior este está hecho en su totalidad con la librería de Microsoft Office Interop Word, ya que con esta se abre la herramienta de Word en segundo plano para cambiar la orientación de una sección específica del documento, este método no se cambian las dimensiones de las página alternando los valores de ancho y alto, asegurando así que no haya una pérdida de formato del documento, además de eso hace el uso de dos métodos los cuales con cada cambio de orientación que se haga se puede añadir un pie de página y encabezado personalizado, lo más recomendable es que no haya un encabezado previo para evitar conflictos entre las dos librerías.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/dec29002-de76-4e9d-8669-812e50428bfa)
    

  * ### Método ActualizarCamposEnWord.
    Este método está hecho con el propósito de ser implementado cuando ya se hayan hecho absolutamente todas las modificaciones del documento, ya que como con el uso de ambas librerías el documento se abre y cierra para su edición ocurre el problema de campos actualizables como lo pueden ser las tablas de contenido, ilustraciones y tablas que se generan con OpenXML no contengan todos los elementos que son referenciables, por lo que el usuario tendrá la tarea manual de actualizar todos los campos varias veces para asegurarse que los índices funcionen. Teniendo lo anterior en cuenta se hace nuevamente uso de la librería de Interop para que esta, una vez hecho el documento, lo abra en segundo plano y refresque los campos del documento.

    ![actualizar_campos_en_word](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/58fca339-070c-472d-ba77-84dbc8d963c1)


## Clase PropiedadesParrafo.
En esta clase de definen la mayoría de los métodos que editan el contenido del cuerpo del documento de word, también por eso mismo son los métodos que más se implementan y están más generalizados para la comodidad del usuario final que haga uso del aplicativo, para generar los reportes y no adentrarse mucho en las configuraciones de cosas básicas.

  * ### Método AgregarParrafo.
    Con la instancia de este método se inserta directamente en el documento de word un párrafo que puede variar y personalizarse según las necesidades del usuario, para ello se reciben como parámetros la ruta del documento en el cual se va insertar el párrafo, el texto que este va a tener, el tamaño de la fuente, un enum con estilos predefinidos (normal, negrita, itálica, subrayado) y otro enum con las alineaciones que puede tener el párrafo (izquierda, derecha, centrado, justificado).
    
    ![agregar_parrfo](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/b4755e6e-52b6-4fef-877c-75c6285cb083)

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/557dc173-3275-44f1-966c-ec943010a6cd)


  * ### Método AsegurarDefinicionNumeracion.
    Con este método para todo el documento se define la numeración, este podría decirse que es un método complementario en el cual se establecen los niveles de títulos en base a los que ya existen en word, el nivel  de numeración va del 1 al 9. Este método asegura que haya una numeración multinivel para tener en cuenta la jerarquía del nivel del título que se emplea, esto con el fin de que cuando se implemente el título este va a poder ser referenciado desde una tabla de contenido.

    ![asegurar_definicion_numeracion](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/234f080d-fb3b-4a2e-907c-317f55533be1)

    
  * ### Método AgregarTitulo.
    Este método tiene una funcionalidad muy similar a la del método de agregar párrafo ya que este es muy general y permite editar el formato del texto del título para personalizarlo tal y como el usuario lo requiera, hace el uso del método que define la numeración y según el nivel de título que se pase por parámetro se va a mostrar la numeración siguiendo niveles de jerarquía y una numeración automática que hace posible la referencia del título en la tabla de contenido.

    ![agregar_titulo](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/aabb4458-ec90-4921-b0f7-b6fec751980c)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/998a4770-f359-4b39-a4c2-b670c31de788)

    
  * ### Método AgregarSaltosDeLinea.
    Con este método se añade un espaciado entre los párrafos del documento, para la implementación del método solo es necesario pasar por parámetros la ruta del documento y el número de espacios en blanco que se van a insertar.

    ![agregar_saltos_de_linea](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/3e84e73f-66ef-455b-ae9f-037b126bc47b)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/84d45805-f545-44fd-9f08-e42e0bd015ac)

    
  * ### Método AgregarSaltoDePagina.
    La implementación de este método es sencilla, ya que este solo necesita de un parámetro para funcionar el cual es la ruta del documento en la cual va a insertar el salto de pagina, el metodo siempre se inicia se ubica en la última parte del documento, por lo cual se se está editando un documento, no importa desde cualquier espacio de la página se llame este insertará un salto de página.

    ![agregar_salto_de_pagina](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/5d43f94a-0c75-4f68-aa6a-8351c2d6a7f2)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/16196885-3d7b-4c03-9872-a1603148aecf)

    
  * ### Método TablaContenido.
    En este método se busca la numeración especial que word usa para los títulos que ya tiene predefinidos, de esta forma se puede tener una lista numerada que puede ser diferenciada de un título normal que se inserta con el método de agregar título, para la implementación de este método solo es necesaria la ruta en la cual se va a insertar y además de esto un texto con el título que tendrá la tabla de contenido del documento.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/7b9eb7fd-3500-4ef8-b65c-d4fcf474209d)

    
  * ### Método TablaTablas.
    Este método busca etiquetas específicas dentro de todo el documento, en este caso se busca específicamente la etiqueta “Tabla”.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/c4d06a92-467a-4307-834d-32981637ee06)
    

  * ### Método TablaIlustraciones.
    Este método busca etiquetas específicas dentro de todo el documento, en este caso se busca específicamente la etiqueta “Ilustración”.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/a40c8d2c-b9b2-42ca-af7b-d2d17508a8c2)
    

  * ### Método AgregarParrafoConCita.
    Para insertar una cita bibliográfica que pueda ser referencia dentro de una tabla de referencias, el método hace uso de ingeniería inversa con un string el cual recibe por parámetros la información que debe ir para ser referenciada, la cual es: nombre de la cita (que actuará como un id permite insertar la cita ya existente varias veces dentro del documento), el nombre del autor, apellido del autor, el año de la publicación, el título del libro o la fuente. El método lo que hace es añadir un párrafo al igual que el método agregar párrafo guarda internamente dentro de las configuraciones de word la propiedad de la cita además de en este caso darle el formato a la cita “IEE”.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/c6d8304f-9eae-4414-bca3-7f43d5ee2e6f)

    
  * ### Método InsertarBibliografia.
    Este método tiene un comportamiento muy similar al anterior ya que este lo que hace es buscar las referencias bibliográficas que se encuentren dentro del documento, estas referencias las presenta en el formato IEE, en la bibliografía se puede visualizar las propiedades que se añadieron por parámetros dentro del método que inserta la bibliografía. Al ser una bibliografía que se inserta manualmente se le dio formato a la bibliografía tomando en cuenta los párrafos y mostrando la información en una tabla que no es dinámica por lo cual es recomendable insertar la bibliografía al final del documento.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/2b840f73-565a-407a-8e05-739e569ad61e)


  * ### Método AgregarListado.
    Como anteriormente se había definido una numeración para los títulos dentro de todo el documento de word usando la librería de OpenXML tendrían que cambiarse las propiedades para agregar un listado sencillo a partir del código sin que este se confunda y cambie la numeración de los títulos por viñetas, teniendo presente este hecho se optó por el uso de Interop para que en segundo plano con la herramienta de word abierta este pueda insertar un párrafo dentro de una lista de viñetas con esto lo estilos de numeración definidos con OpenXML no interfieren con las viñetas que genera Interop.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/8f917e6b-4eef-4b30-a4d0-f99dfa564b0b)


## Clase EditarEncabezadoPie.
En esta clase se definen todos los métodos que están relacionados con la creación de los pie de páginas y encabezados que se insertan en el documento automatizado.

  * ### Método nuevoEncabezado.
    Método reutilizable, este tipo de métodos no insertan contenido directamente dentro del documento de word, ya que estos se encargan de definir componentes que se van  a insertar dentro del documento con la implementación de otros métodos, en este caso este método lo que hace es definir un objeto de tipo encabezado específicamente una tabla con un formato espacial y retorna este objeto.
    
  * ### Método EditarEncabezado.
    Este método está hecho para implementar un objeto de OpenXML de tipo encabezado que no se recibe por parámetro sino que se implementa internamente dentro de código con el uso de una variable, esta variable se manipula para insertar el encabezado dentro del documento de word, para ello como parámetros solo recibe la ruta del documento, el texto que va a estar en la tabla del encabezado  y la altura del encabezado, para lo cual es recomendable siempre usar el valor de 2 ya que la librería OpenXML trabaja con la unidad de puntos y no de cm. Además de la inserción de encabezado el método define un encabezado en blanco y otro general para todo el documento en este caso la tabla estilizada.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/aeada18d-364b-4c4f-b330-6b55d74c7a0d)


  * ### Método nuevoPie.
    Método reutilizable que retorna una tabla con un formato estilizado, este es un método reutilizable el cual solo tiene el propósito de dar la definición de una tabla con un formato que va a ser usada dentro de los pie de página del documento.
    
  * ### Método EditarPieDePagina.
    Este método hace la implementación de la tabla que retorna el método anterior, pero con una serie de validaciones que pueden ser modificadas a nivel del código dentro del método, en este caso las tablas que van dentro del pie de página que se va a insertar con este método tienen un orden específico por lo cual por medio de validaciones se define una tabla específica para la primera página que actúa como una portada, mientras que la tabla hecha por el método anterior se implementa como un pie de página general de todo el documento. Dado lo anterior es importante aclarar que en OpenXML las secciones se toman como un documento nuevo, por lo tanto si se quiere implementar el método varias veces se va a estar sobreescribiendo, por ende el formato no puede ser el esperado, motivo por el cual si se crean otras secciones con propiedades específicas dentro del documento se repetirá el pie de página que se definió para solo la primera página, por este hecho es recomendable implementar estos métodos de forma general para un documento el cual no va a tener secciones para evitar este inconveniente.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/e8995607-4753-4d37-b805-76130e6aac15)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/a464712b-cd18-42c1-8504-dbdee9ff392f)


  * ### Método CrearPieDePagina y CrearEncabezado.
    Teniendo en cuenta el inconveniente que tiene la librera de OpenXML con las secciones cuando ya están establecidos los pie de página y encabezados previamente para el documento, se optó por hacer una imitación de formato del encabezado y pie de página  hecha con interop que reciba los mismos parámetros para ser implementado dentro del documento, pero con la pequeña diferencia que también recibe una sección puntal en la cual se va a implementar la tabla, este método tiene lo lógica para eliminar los encabezados y pies de página previos, ya que con OpenXML solo se pueden escoger tres opciones para la visualización de los encabezados y pies de páginas las cuales son: default para todas las páginas, even para las páginas pares y first para la primera página. por ende si se intenta implementar la misma lógica de eliminar el encabezado anterior que ya existe para insertar uno nuevo no solo lo va a tomar la sección específica sino que lo hará todo el documento, por eso estos métodos usan Interop para que una vez se pase una sección el encabezado y pie de página  puedan ser implementados dentro de esa sección eliminado previamente el encabezado y pie  que ya existen, pero solo dentro de esa sección. Por comodidad del usuario y con el propósito de evitar errores en el formato se sugiere que si se van a insertar estos métodos de agregar pies de página y encabezados dentro de una sección específica se evite la implementación de los método que hacen lo mismo con OpenXML.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/d2c6cb29-6155-4080-8216-2487f649476f)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/8daedfe1-eeb7-4dc6-b91b-8f197aa514f0)


## Clase PropiedadesTabla.
Dentro de esta clase se encuentran todos los métodos referentes a tablas dentro de los documentos de word.

  * ### Método TituloTablas.
    Dentro de este método se define un “caption”, que en OpenXML es una forma de crear una etiqueta para un atributo en específico dentro del documento de word, en este caso el método genera un párrafo normal, pero con el prefijo que en realidad es una etiqueta Tabla, esto con el propósito de hacer la función de título a las tablas que se generen dentro del documento, para que estas puedan ser referenciadas dentro de un índice de tablas.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/d1c5870f-d5a3-48f5-be3c-265728530a2c)
    

  * ### Método AgregarTablaDesdeLista.
    En este se inserta una tabla directamente en el documento de word, para ello el método recibe por parámetros la ruta del documento en el cual va a insertar la tabla, una lista con otras listas de strings (una matriz nxn) que son los datos que y las posiciones de la tabla, un valor entero con las filas de fondo, estas filas serán un encabezado de la tabla y un valor booleano para determinar si la tabla tantra bordes o no, esta es una forma eficiente de trabajar con grandes datos y mostrarlos en el documento en forma de tabla la cual puede personalizarse según las necesidades del usuario final. Este método también emplea una lógica especial para hacer la combinación de celdas para que así que dentro de un método queden abarcados varios aspectos que pueden ser modificados por el usuario, para hacer la combinación de celdas dentro de los valores de las listas deben ir símbolos especiales que deben estar solos, para que el método detecte el símbolo y sepa cómo combinar las celdas si se usa el símbolo “|” la celda se combinara de arriba a abajo y si se usa el símbolo “~” se combinara de izquierda a derecha.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/0296c36f-382f-4d42-8194-5d1698ff39e7)

  * ### Método AgregarTablaDesdeLista (Sobrecarga de Método).
    C# es un lenguaje que permite el uso de la sobrecarga de métodos dentro de sus clases, en este caso se creó una sobrecarga del método anterior de modo que el usuario tenga  la posibilidad de agregar una tabla que se ajuste a sus necesidades y darle título a esta tabla, por lo cual la sobrecarga lo que hace es la misma lógica que el método anterior, pero con el ligero ajuste de que la tabla creada viene con un texto y la etiqueta de título para ser referenciada en el índice.

    
  * ### Método AgregarTablaConImagen.
    Este método también hubiera podido ser una sobrecarga de los anteriores, pero se decidió hacer de forma distinta, para evitar pasarle muchos parámetros a una solo función y de esta forma ahorrándonos el problema de tener un método muy extenso, ya que éste tendría una lógica un poco distinta. Para la implementación de este método se pasan lo parámetros de la ruta en al cual se va a agregar la tabla, una lista de listas que también tiene las validaciones de combinación de celdas, pero además de eso hay una validación que facilita el proceso de inserción de imágenes dentro de las celdas, el cual es la identificación de un conjunto de caracteres, para determinar que se pase dentro de las lista un string de base 64, el método internamente procesa ese string y lo pasara a un formato de imagen ya dentro de la tabla, para que esto funcione antes de toda la cadena dle base 64 de debe agregar el prefijo “[B64]” para que agregue la imagen dentro de la celda or el contrario solo va añadir texto, este método conserva internamente las dimensiones de en cm de las imágenes ya con un valor preestablecido, por lo cual si se quiere modificar el método para personalizar esos valores se puede hacer agregando los parámetros y reemplazando los valores constantes que ya se encuentran.
    
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/9270bb4a-45d5-4a37-8771-75ebe89031fb)


  * ### Método CrearTablaConImagen
    El funcionamiento de este método es exactamente igual al método anterior, la diferencia es que en lugar de insertar la tabla con imagen dentro del documento directamente, lo que hace es que crea la tabla con la imagen y la retorna, esta tabla a diferencia de la anterior no está hecha para que se inserte en el cuerpo de documento sino en el pie de página. Este es el método que se encuentra dentro del método OpenXML de editar pie de página, en el cual está la lógica para que esta tabla solo sea visible dentro de la primera página de la sección principal del documento.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/c90d3e49-30bd-459a-a9c1-cb10a999b36a)


## Clase PropiedadesImagen
En esta clase se encuentran todos los métodos relacionados con la inserción de imágenes dentro de los reportes automáticos en word.

  * ### Método TituloImagen.
    Este método genera un párrafo que hace la función de título para las imágenes, este es un método reutilizable ya que no inserta contenido directamente en el documento usa las propiedades del caption de OpenXML para que de esta forma el título que se genera tenga la etiqueta “Ilustración”, la cual va a hacer la imagen referenciable en una tabla ilustraciones.

  * ### Método AgregarImagenDesdeArchivo.
    Este método inserta directamente una imagen dentro de un documento de word, este recibe por parámetros la ruta del documento, la ruta de la imagen dentro del equipo del usuario que se va a insertar, el valor entero del ancho y el alto, un enum con los valores constantes de la alineación de la imagen dentro del cuerpo del documento y el título de la imagen en caso de que el usuario lo requiera, este parámetro se puede pasar por alto.

    ![agregar_imagen_desde_archivo](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/fe9940af-7701-4ef8-a59c-3a7f11f786aa)
    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/28069ae6-5abb-4361-90d5-d17c9711e669)


  * ### Método AgregarImagenDesdeBase64
    Este método tiene exactamente la misma lógica que el anterior, solo que se presenta desde un enfoque diferente , en este caso la imagen no se inserta en el documento de word a partir de una imagen que existe dentro del disco del equipo,con el fin de ahorrar recursos del equipo, con solo pasar por parametro un string con un base 64 de la imagen, se va a insertar la imagen dentro del documento con exactamente las mismas personalizaciones que el método anterior.

    ![imagen](https://github.com/andresgutierrez73/avances_proyecto_reporte_automatizado/assets/139988681/39de493a-bc90-43b1-a080-a69137cfab88)

    
  * ### Método ObtenerImagenDesdeBase64 y su método sobreconstruido.
    Esos métodos lo que hacen es procesar los base64 que le llegan como parámetros, para crear un objeto de imagen de la librería de OpenXML, ambos retornan este objeto de imagen, lo que varía en cada uno es el contexto en el que se usan ya que un método genera la imagen para ser manipulada en el cuerpo principal del documento, esta por ejemplo es usada en las tablas que se crean con las imágenes, mientras que la sobrecarga del método hace posible que los objetos de las imágenes puedan ser usadas dentro de los pie de página y los encabezados del documento. 







