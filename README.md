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








