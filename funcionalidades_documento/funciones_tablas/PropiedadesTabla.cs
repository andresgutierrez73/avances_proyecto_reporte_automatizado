using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;
using funcionalidades_documento.funciones_imagenes;

namespace funcionalidades_documento.funciones_tablas
{
    public class PropiedadesTabla
    {
        /// <summary>
        /// Método para validar la ruta y extensión del archivo
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del docuemento de word</param>
        /// <exception cref="ArgumentException"></exception>
        private static void ValidarRutaArchivo(string ruta)
        {
            if (string.IsNullOrEmpty(ruta))
            {
                throw new ArgumentException("La ruta no puede estar vacía.");
            }

            string extension = System.IO.Path.GetExtension(ruta);
            if (extension != ".docx")
            {
                throw new ArgumentException("La ruta debe tener una extensión .docx");
            }
        }

        /// <summary>
        /// Método encargado de hacel el prefijo de los titulos de las tablas
        /// ademas de agregar una numeración especial para que los titulos de las tablas puedan ser referenciados 
        /// dentro de una tabla de contenido
        /// </summary>
        /// <param name="mensaje">Aquí se pasa el titulo que va a ir ubicado encima de la tabla que se agrega al contenido</param>
        /// <returns>retorna el parrafo con estilo para que pueda ser insertado en el documento</returns>
        public static Paragraph TituloTablas(string mensaje)
        {
            // Crear y establecer las propiedades del párrafo
            Paragraph paragraph = new Paragraph();

            // Propiedades para centrar el párrafo
            Justification justification = new Justification() { Val = JustificationValues.Center };
            ParagraphProperties paragraphProperties = new ParagraphProperties(justification, new ParagraphStyleId() { Val = "IEBNormal1" });
            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties(new Languages() { Val = "es-CO" });

            paragraphProperties.Append(paragraphMarkRunProperties);
            paragraph.Append(paragraphProperties);

            // Propiedades comunes de Run con negrita añadida
            RunProperties runProperties = new RunProperties(
                new Languages() { Val = "es-CO" },
                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" },
                new FontSize() { Val = "24" },  // Tamaño 12 en Word
                new Bold()  // Negrita
            );

            // Añadir los runs al párrafo con propiedades únicas
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text("Tabla ")));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldCode(" SEQ Tabla \\* ARABIC ")));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.Separate }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.End }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text(": " + mensaje.Trim())));

            return paragraph;
        }

        /// <summary>
        /// Método para la creacion de tablas personalizadas en el documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del directorio donde se encuentra ubicado el documento</param>
        /// <param name="datos">Aquí va una lista de listas, eso con el proposito de hacer mas dinámica la longitud
        /// de las tablas que se crean</param>
        /// <param name="filasConFondo">Aquí se pasa la cantidad de filas que van a actuar como encabezado la diferencia
        /// con respecto a las otras es que estas van a estar centradas y con un color de fondo de celda asigando</param>
        /// <param name="sinBordes">Aquí se pasa un booleando como parámetro por defecto las tables siempre van a tener bordes, pero si se quiere tener una tabla la cual no tenga bordes se pasa el valor de true</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarTablaDesdeLista(string ruta, List<List<string>> datos, int filasConFondo = 0, bool sinBordes = false)
        {
            if (datos == null || !datos.Any())
            {
                throw new ArgumentNullException(nameof(datos), "Los datos no pueden ser nulos o vacíos.");
            }

            int maxColumnCount = datos.Max(row => row.Count);
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                // Crea la tabla
                DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

                // Define el ancho de la tabla al 100% del ancho del documento
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                // Define los bordes de la tabla
                DocumentFormat.OpenXml.Wordprocessing.TableBorders tblBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                    new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 }
                );

                DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
                tblProperties.Append(tableWidth);
                if (!sinBordes)
                {
                    tblProperties.Append(tblBorders);
                }
                table.Append(tblProperties);

                // Añadir las filas desde datos
                for (int rowIndex = 0; rowIndex < datos.Count; rowIndex++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    int currentColumnCount = datos[rowIndex].Count;

                    for (int colIndex = 0; colIndex < currentColumnCount; colIndex++)
                    {
                        var cellText = datos[rowIndex][colIndex];
                        DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();

                        // Agrega el texto a la celda
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                        DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                        // Establece la fuente en Arial y el tamaño de fuente en 10 para todas las celdas
                        run.RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }, new FontSize() { Val = "20" });

                        // Si se especificó un número de filas con fondo, darle fondo gris y poner la letra en negrita y centrada
                        if (filasConFondo > 0 && rowIndex < filasConFondo)
                        {
                            cellProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Shading() { Val = ShadingPatternValues.Clear, Fill = "f0f0f0" });
                            paragraph.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                            run.RunProperties.Append(new Bold());
                        }

                        run.Append(new DocumentFormat.OpenXml.Wordprocessing.Text(cellText));
                        paragraph.Append(run);
                        cell.Append(paragraph);

                        // Verifica si la celda debe ser combinada horizontalmente
                        if (cellText.Contains("~"))
                        {
                            cellText = cellText.Replace("~", "");
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                        }

                        // Verifica si la celda debe ser combinada verticalmente
                        if (cellText.Contains("|"))
                        {
                            cellText = cellText.Replace("|", "");
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                        }

                        cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                        cell.Append(cellProperties);
                        row.Append(cell);
                    }

                    // Aquí es donde agregas la propiedad TableHeader a las filas que actuarán como encabezados.
                    if (filasConFondo > 0 && rowIndex < filasConFondo)
                    {
                        row.Append(new TableHeader());
                    }

                    table.Append(row);
                }

                // Agrega la tabla al documento
                body.Append(table);

                // Guarda los cambios en el documento
                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Se agregó una tabla al documento.");
        }

        /// <summary>
        /// Método sobreconstruido en el cual se pasa el parametro del titulo de la tabla
        /// </summary>
        /// <param name="ruta">Aquí va el directorio donde este ubicado el documento de word</param>
        /// <param name="datos">Aquí va una lista de listas, eso con el proposito de hacer mas dinámica la longitud
        /// de las tablas que se crean</param>
        /// <param name="titulo">Aquí se pasa el titullo que va a ir en el tabla, este será el parametro del método encargado de dar titulos a las tablas</param>
        /// <param name="filasConFondo">Aquí se pasa la cantidad de filas que van a actuar como encabezado la diferencia
        /// con respecto a las otras es que estas van a estar centradas y con un color de fondo de celda asigando</param>
        /// <param name="sinBordes">Aquí se pasa un booleando como parámetro por defecto las tables siempre van a tener bordes, pero si se quiere tener una tabla la cual no tenga bordes se pasa el valor de true</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarTablaDesdeLista(string ruta, List<List<string>> datos, string titulo, int filasConFondo = 0, bool sinBordes = false)
        {
            if (datos == null || !datos.Any())
            {
                throw new ArgumentNullException(nameof(datos), "Los datos no pueden ser nulos o vacíos.");
            }

            int maxColumnCount = datos.Max(row => row.Count);
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                // Añadir el título a la tabla usando el método NewCaption
                if (!string.IsNullOrEmpty(titulo))
                {
                    Paragraph caption = TituloTablas(titulo);
                    body.Append(caption);
                }

                // Crea la tabla
                DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

                // Define el ancho de la tabla al 100% del ancho del documento
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                // Define los bordes de la tabla
                DocumentFormat.OpenXml.Wordprocessing.TableBorders tblBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                    new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 }
                );

                DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
                tblProperties.Append(tableWidth);
                if (!sinBordes)
                {
                    tblProperties.Append(tblBorders);
                }
                table.Append(tblProperties);

                // Añadir las filas desde datos
                for (int rowIndex = 0; rowIndex < datos.Count; rowIndex++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    int currentColumnCount = datos[rowIndex].Count;

                    for (int colIndex = 0; colIndex < currentColumnCount; colIndex++)
                    {
                        var cellText = datos[rowIndex][colIndex];
                        DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();

                        // Agrega el texto a la celda
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                        DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();

                        // Establece la fuente en Arial y el tamaño de fuente en 10 para todas las celdas
                        run.RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }, new FontSize() { Val = "20" });

                        // Si se especificó un número de filas con fondo, darle fondo gris y poner la letra en negrita y centrada
                        if (filasConFondo > 0 && rowIndex < filasConFondo)
                        {
                            cellProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Shading() { Val = ShadingPatternValues.Clear, Fill = "f0f0f0" });
                            paragraph.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                            run.RunProperties.Append(new Bold());
                        }

                        run.Append(new DocumentFormat.OpenXml.Wordprocessing.Text(cellText));
                        paragraph.Append(run);
                        cell.Append(paragraph);

                        // Verifica si la celda debe ser combinada horizontalmente
                        if (cellText.Contains("~"))
                        {
                            cellText = cellText.Replace("~", "");
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                        }

                        // Verifica si la celda debe ser combinada verticalmente
                        if (cellText.Contains("|"))
                        {
                            cellText = cellText.Replace("|", "");
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                        }

                        cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                        cell.Append(cellProperties);

                        row.Append(cell);
                    }

                    // Aquí es donde agregas la propiedad TableHeader a las filas que actuarán como encabezados.
                    if (filasConFondo > 0 && rowIndex < filasConFondo)
                    {
                        row.Append(new TableHeader());
                    }

                    table.Append(row);
                }

                // Agrega la tabla al documento
                body.Append(table);

                // Guarda los cambios en el documento
                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Se agregó una tabla con el título \"{titulo}\" al documento.");
        }

        /// <summary>
        /// Método para crear la tabla de firmas del documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del directorio donde se encuentra ubicado el documento</param>
        /// <param name="datos">Aquí va una lista de listas, eso con el proposito de hacer mas dinámica la longitud
        /// de las tablas que se crean</param>
        /// <param name="sinBordes">Aquí se pasa un booleando como parámetro por defecto las tables siempre van a tener bordes, pero si se quiere tener una tabla la cual no tenga bordes se pasa el valor de true</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarTablaConImagen(string ruta, List<List<string>> datos, bool sinBordes = false)
        {
            if (datos == null || !datos.Any())
            {
                throw new ArgumentNullException(nameof(datos), "Los datos no pueden ser nulos o vacíos.");
            }

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                var body = document.MainDocumentPart.Document.Body;
                var mainPart = document.MainDocumentPart;

                // Crea la tabla
                DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

                // Define el ancho de la tabla al 100% del ancho del documento
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                // Define los bordes de la tabla
                DocumentFormat.OpenXml.Wordprocessing.TableBorders tblBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                    new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 }
                );

                DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
                tblProperties.Append(tableWidth);
                if (!sinBordes)
                {
                    tblProperties.Append(tblBorders);
                }
                table.Append(tblProperties);

                for (int rowIndex = 0; rowIndex < datos.Count; rowIndex++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    int currentColumnCount = datos[rowIndex].Count;

                    for (int colIndex = 0; colIndex < currentColumnCount; colIndex++)
                    {
                        var cellText = datos[rowIndex][colIndex];
                        DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();

                        // Si es una imagen base64
                        if (cellText.StartsWith("[B64]") && Regex.IsMatch(cellText.Substring(5), @"^[a-zA-Z0-9+/]*={0,2}$"))
                        {
                            // MODIFICACIÓN: Cambia las dimensiones en función de si es la primera fila o no
                            int imgAncho = rowIndex == 0 ? 3 : 2;  // Supongamos que para la primera fila quieres que sea 5
                            int imgAlto = rowIndex == 0 ? 2 : 1;   // Y aquí, por ejemplo, que sea 4

                            Drawing imageElement = PropiedadesImagen.ObtenerImagenDesdeBase64(mainPart, cellText.Substring(5), imgAncho, imgAlto);

                            // Crear un párrafo
                            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

                            // Establecer las propiedades de alineación del párrafo para centrar
                            DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                                new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center }
                            );
                            paragraph.Append(paragraphProperties);

                            // Agregar la imagen al párrafo
                            DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run(imageElement);
                            paragraph.Append(run);

                            // Agregar el párrafo a la celda
                            cell.Append(paragraph);
                        }
                        else // Si es texto
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                            DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cellText));

                            // Establecer la fuente a Arial, el tamaño de la fuente a 10, y centrar el texto
                            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                            runProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });
                            runProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "20" });
                            run.PrependChild(runProperties);

                            // Centrar el texto
                            DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                            paragraphProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = JustificationValues.Center });
                            paragraph.PrependChild(paragraphProperties);

                            paragraph.Append(run);
                            cell.Append(paragraph);
                        }

                        // Combinación de celdas
                        if (cellText.Contains("~"))
                        {
                            cellText = cellText.Replace("~", "");
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                        }

                        if (cellText.Contains("|"))
                        {
                            cellText = cellText.Replace("|", "");
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                        }

                        cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                        cell.Append(cellProperties);

                        row.Append(cell);
                    }

                    table.Append(row);
                }

                body.Append(table);
                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Se agregó una tabla al documento.");
        }

        /// <summary>
        /// Método encargado de insertar directamente en el documento de word una tabla
        /// que puede contener imagenes dentro, las imágenes deben estar decodificadas en un base64
        /// </summary>
        /// <param name="footerPart">Aquí se pasa la seccion del footer o el pie de pagina en el cual va a ir ubicada la tabla</param>
        /// <param name="datos">Se pasa una lista con una lista de strings para configuracion de las celdas de la tabla</param>
        /// <param name="sinBordes">Aquí se pasa un booleando como parámetro por defecto las tables siempre van a tener bordes, pero si se quiere tener una tabla la cual no tenga bordes se pasa el valor de true</param>
        /// <returns>retorna una tabla que puede ser insertada solo en el pie de pagina, esto con el proposito de que se van las imagenes dentro de la tabla
        /// que esta ubicada en el pie de pagina</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static DocumentFormat.OpenXml.Wordprocessing.Table CrearTablaConImagen(FooterPart footerPart, List<List<string>> datos, bool sinBordes = false)
        {
            if (datos == null || !datos.Any())
            {
                throw new ArgumentNullException(nameof(datos), "Los datos no pueden ser nulos o vacíos.");
            }

            // Crea la tabla
            DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

            // Define el ancho de la tabla al 100% del ancho del documento
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

            // Define los bordes de la tabla
            DocumentFormat.OpenXml.Wordprocessing.TableBorders tblBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 },
                new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single, Size = 6 }
            );

            DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
            tblProperties.Append(tableWidth);
            if (!sinBordes)
            {
                tblProperties.Append(tblBorders);
            }
            table.Append(tblProperties);

            for (int rowIndex = 0; rowIndex < datos.Count; rowIndex++)
            {
                DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                int currentColumnCount = datos[rowIndex].Count;

                for (int colIndex = 0; colIndex < currentColumnCount; colIndex++)
                {
                    var cellText = datos[rowIndex][colIndex];
                    DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();

                    // Si es una imagen base64
                    if (cellText.StartsWith("[B64]") && Regex.IsMatch(cellText.Substring(5), @"^[a-zA-Z0-9+/]*={0,2}$"))
                    {
                        // MODIFICACIÓN: Cambia las dimensiones en función de si es la primera fila o no
                        int imgAncho = rowIndex == 0 ? 3 : 2;
                        int imgAlto = rowIndex == 0 ? 2 : 1;

                        Drawing imageElement = PropiedadesImagen.ObtenerImagenDesdeBase64(footerPart, cellText.Substring(5), imgAncho, imgAlto);

                        // Crear un párrafo
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

                        // Establecer las propiedades de alineación del párrafo para centrar
                        DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center }
                        );
                        paragraph.Append(paragraphProperties);

                        // Agregar la imagen al párrafo
                        DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run(imageElement);
                        paragraph.Append(run);

                        // Agregar el párrafo a la celda
                        cell.Append(paragraph);
                    }
                    else // Si es texto
                    {
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                        DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cellText));

                        // Establecer la fuente a Arial, el tamaño de la fuente a 10, y centrar el texto
                        DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                        runProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });
                        runProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "20" });
                        run.PrependChild(runProperties);

                        // Centrar el texto
                        DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                        paragraphProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = JustificationValues.Center });
                        paragraph.PrependChild(paragraphProperties);

                        paragraph.Append(run);
                        cell.Append(paragraph);
                    }

                    // Combinación de celdas
                    if (cellText.Contains("~"))
                    {
                        cellText = cellText.Replace("~", "");
                        cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cellProperties.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                    }

                    if (cellText.Contains("|"))
                    {
                        cellText = cellText.Replace("|", "");
                        cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                    }

                    cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                    cell.Append(cellProperties);

                    row.Append(cell);
                }

                table.Append(row);
            }

            return table;
        }
    }
}
