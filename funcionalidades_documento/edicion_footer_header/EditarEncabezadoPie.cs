﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using funcionalidades_documento.funciones_tablas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using Word = Microsoft.Office.Interop.Word;


namespace funcionalidades_documento.edicion_footer_header
{
    public class EditarEncabezadoPie
    {
        /// <summary>
        /// Método para validar la ruta y extensión del archivo
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
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


        #region Métodos que usan la librería openxml
        /// <summary>
        /// Método para establecer el encabezado con el formato del documento
        /// </summary>
        /// <param name="titulo">Aquí va a el texto que se puede ver arriba en el encabezado</param>
        /// <param name="preTitulo">Aquí va a el texto que se puede ver abajo en el encabezado</param>
        /// <returns></returns>
        public static Header nuevoEncabezado(string titulo, string preTitulo)
        {
            try
            {
                Header header = new Header();

                #region NameSpaces
                header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                #endregion

                Table headerTable = new Table(new TableProperties(
                    new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                    ),
                    new TableCellMarginDefault(
                        new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                    )
                ));

                TableRow headerRow1 = new TableRow();
                TableRow headerRow2 = new TableRow();

                TableCell headerCell11 = new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Left },
                            new SpacingBetweenLines() { Before = "0", After = "22" },
                            new Languages() { Val = "es-ES" }
                        ),
                        new Run(
                            new RunProperties(new FontSize() { Val = "18" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                            new Text(preTitulo)
                        )
                    )
                );
                TableCell headerCell21 = new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Left },
                            new SpacingBetweenLines() { Before = "0", After = "22" },
                            new Languages() { Val = "es-ES" }
                        ),
                        new Run(
                            new RunProperties(new FontSize() { Val = "18" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                            new Text(titulo)
                        )
                    )
                );
                TableCell headerCell22 = new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Right },
                            new SpacingBetweenLines() { Before = "0", After = "22" },
                            new Languages() { Val = "es-ES" }
                        ),
                        new Run(
                            new StyleRunProperties(new FontSize() { Val = "18" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                            new Text("Página ") { Space = SpaceProcessingModeValues.Preserve },
                            new SimpleField() { Instruction = "PAGE" },
                            new Text(" de ") { Space = SpaceProcessingModeValues.Preserve },
                            new SimpleField() { Instruction = "NUMPAGES" }
                        )
                    )
                );

                headerRow1.Append(headerCell11);
                headerRow2.Append(headerCell21, headerCell22);

                headerTable.Append(headerRow1, headerRow2);
                header.Append(headerTable);

                return header;
            }
            catch (Exception ex)
            {
                // Puedes manejar la excepción aquí. Por simplicidad, simplemente la lanzaré nuevamente.
                throw ex;
            }
        }


        /// <summary>
        /// Método para añadir en encabezado en el documento de word a partir del formato del metodo anterior
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="textoAbajo">Aquí va el texto que esta en la parte baja del texto</param>
        /// <param name="textoArriba">Aquí va el texto que esta en la parte alta del texto</param>
        /// <param name="altura">Aquí va la altura que va a tener el encabezado, lo mejor es usar el valor de 2</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void EditarEncabezado(string ruta, string textoAbajo, string textoArriba, double altura)
        {
            try
            {
                ValidarRutaArchivo(ruta);

                // Conversion de cm a puntos (1 cm = 28.3465 puntos)
                int alturaEnPuntos = (int)(360000 * altura);

                using (var document = WordprocessingDocument.Open(ruta, true))
                {
                    if (document == null)
                    {
                        throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                    }

                    var mainPart = document.MainDocumentPart;

                    HeaderPart headerPart;
                    if (mainPart.HeaderParts.Count() > 0)
                    {
                        headerPart = mainPart.HeaderParts.First();
                    }
                    else
                    {
                        headerPart = mainPart.AddNewPart<HeaderPart>();
                        mainPart.Document.Save();
                    }

                    var header = nuevoEncabezado(textoArriba, textoAbajo);
                    var paragraphProperties = header.Descendants<ParagraphProperties>().FirstOrDefault();
                    if (paragraphProperties == null)
                    {
                        paragraphProperties = new ParagraphProperties();
                        header.Descendants<Paragraph>().First().PrependChild(paragraphProperties);
                    }
                    var spacingBetweenLines = paragraphProperties.Descendants<SpacingBetweenLines>().FirstOrDefault();
                    if (spacingBetweenLines == null)
                    {
                        spacingBetweenLines = new SpacingBetweenLines();
                        paragraphProperties.AppendChild(spacingBetweenLines);
                    }
                    spacingBetweenLines.After = "0";
                    spacingBetweenLines.Line = alturaEnPuntos.ToString();
                    spacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;

                    headerPart.Header = header;

                    // Agregar la referencia al encabezado en el documento principal
                    if (mainPart.Document.Body.Elements<SectionProperties>().Any())
                    {
                        SectionProperties sectionProperties = mainPart.Document.Body.Elements<SectionProperties>().First();
                        HeaderReference headerReference = new HeaderReference { Id = mainPart.GetIdOfPart(headerPart), Type = HeaderFooterValues.Default }; // Tipo por defecto
                        sectionProperties.RemoveAllChildren<HeaderReference>();
                        sectionProperties.PrependChild(headerReference);

                        // Configurar el encabezado diferente para la primera página
                        TitlePage titlePage = new TitlePage();
                        sectionProperties.RemoveAllChildren<TitlePage>();
                        sectionProperties.PrependChild(titlePage);
                    }
                    else
                    {
                        mainPart.Document.Body.Append(new SectionProperties(new HeaderReference { Id = mainPart.GetIdOfPart(headerPart), Type = HeaderFooterValues.Default }, new TitlePage()));
                    }

                    document.Save();
                }
            }
            catch (Exception ex)
            {
                // Aquí puedes manejar la excepción o lanzarla nuevamente, según lo que necesites.
                // Por simplicidad, simplemente la lanzaré nuevamente.
                throw ex;
            }
        }



        /// <summary>
        /// Método que retorna un objeto de tipo Footer con un formato específico como en este caso que es un tabla personalizada con un texto
        /// </summary>
        /// <param name="textoPie">Aquí se recibe un texto, le cual va a ser el que se vera en el pie de página del documento</param>
        /// <returns>Esté método retorna un pie de página, por lo cual no lo agrega directemente al documento</returns>
        public static Footer nuevoPie(string textoPie)
        {
            try
            {
                Footer footer = new Footer();

                #region NameSpaces
                footer.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                footer.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                footer.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                footer.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                footer.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                footer.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                footer.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                footer.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                footer.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                footer.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                footer.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                footer.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                footer.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                footer.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                #endregion

                Table footerTable = new Table(new TableProperties(
                    new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                    ),
                    new TableCellMarginDefault(
                        new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                    )
                ));

                TableRow footerRow1 = new TableRow();
                TableCell footerCell11 = new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Right },
                            new SpacingBetweenLines() { Before = "0", After = "22" },
                            new Languages() { Val = "es-ES" }
                        ),
                        new Run(
                            new RunProperties(new FontSize() { Val = "18" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                            new Text(textoPie)
                        )
                    )
                );

                footerRow1.Append(footerCell11);
                footerTable.Append(footerRow1);
                footer.Append(footerTable);

                return footer;
            }
            catch (Exception ex)
            {
                // Aquí puedes manejar la excepción o lanzarla nuevamente, según lo que necesites.
                // Por simplicidad, simplemente la lanzaré nuevamente.
                throw ex;
            }
        }


        /// <summary>
        /// Método encargado de recibir un texto y de manipula un footer para insertarlo en el documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="texto">Aquí va el texto que se mostrará en todos los pie de página del docuento exceptuando el inicio</param>
        /// <param name="textoPrimeraPagina"></param>
        /// <exception cref="ArgumentNullException">Aquí va el texto que irá en el pie de la primera página del documento</exception>
        public static void EditarPieDePagina(string ruta, string texto, List<List<string>> datosTabla)
        {
            try
            {
                ValidarRutaArchivo(ruta);

                using (var document = WordprocessingDocument.Open(ruta, true))
                {
                    if (document == null)
                    {
                        throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                    }

                    var mainPart = document.MainDocumentPart;

                    // Pie de página para las demás páginas
                    FooterPart footerPart;
                    if (mainPart.FooterParts.Count() > 0)
                    {
                        footerPart = mainPart.FooterParts.First();
                    }
                    else
                    {
                        footerPart = mainPart.AddNewPart<FooterPart>();
                        mainPart.Document.Save();
                    }

                    var footer = nuevoPie(texto);  // Asegúrate de que esta función pueda trabajar con FooterPart si es necesario.
                    footerPart.Footer = footer;

                    // Pie de página para la primera página
                    FooterPart firstPageFooterPart = mainPart.AddNewPart<FooterPart>();

                    // Crear la tabla con imágenes para el pie de página de la primera página
                    Table tablaConImagen = PropiedadesTabla.CrearTablaConImagen(firstPageFooterPart, datosTabla);  // Modificado para pasar FooterPart

                    // Crea el pie de página para la primera página y añade la tabla directamente a él
                    var firstPageFooter = new Footer();
                    firstPageFooter.Append(tablaConImagen);

                    firstPageFooterPart.Footer = firstPageFooter;
                    firstPageFooterPart.Footer.Save();

                    SectionProperties sectionProperties;
                    if (mainPart.Document.Body.Elements<SectionProperties>().Any())
                    {
                        sectionProperties = mainPart.Document.Body.Elements<SectionProperties>().First();
                    }
                    else
                    {
                        sectionProperties = new SectionProperties();
                        mainPart.Document.Body.Append(sectionProperties);
                    }

                    // Elimina referencias existentes y agrega nuevas referencias
                    sectionProperties.RemoveAllChildren<FooterReference>();
                    sectionProperties.Append(new FooterReference { Id = mainPart.GetIdOfPart(footerPart), Type = HeaderFooterValues.Default }); // Para las demás páginas
                    sectionProperties.Append(new FooterReference { Id = mainPart.GetIdOfPart(firstPageFooterPart), Type = HeaderFooterValues.First }); // Para la primera página

                    document.Save();
                }
            }
            catch (Exception ex)
            {
                // Aquí puedes manejar la excepción o lanzarla nuevamente, según lo que necesites.
                // Por simplicidad, simplemente la lanzaré nuevamente.
                throw ex;
            }
        }
        #endregion


        #region Métodos que usan la librería microsoft interop word
        /// <summary>
        /// Método que inserta directamente en el documento de word un encabezado a partir de una sección del documento 
        /// </summary>
        /// <param name="section">Aquí se pasa una sección del documento en la cual se van a aplicar los cambios</param>
        /// <param name="textoPie">Aquí se pasa un string con el texto que estará en el pie de página</param>
        /// <param name="firstPage">Aquí se pasa un valor booleano en caso de que la primera página sea diferente</param>
        /// <exception cref="ApplicationException"></exception>
        public static void CrearPieDePagina(Word.Section section, string textoPie, bool firstPage = false)
        {
            try
            {
                Word.HeaderFooter footer;
                if (firstPage)
                {
                    footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                }
                else
                {
                    footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                }

                // Verificar si el pie de página contiene una tabla
                if (footer.Range.Tables.Count > 0)
                {
                    // Eliminar la tabla del pie de página
                    footer.Range.Tables[1].Delete();
                }

                // Agregar el nuevo contenido al pie de página
                Word.Paragraph paragraph = footer.Range.Paragraphs.Add();
                paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                Word.Border border = paragraph.Borders[Word.WdBorderType.wdBorderTop];
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.LineWidth = Word.WdLineWidth.wdLineWidth150pt;

                Word.Font font = paragraph.Range.Font;
                font.Size = 9;
                font.Name = "Arial";

                paragraph.Range.Text = textoPie;
            }
            catch (Exception ex)
            {
                // Aquí puedes manejar la excepción o lanzarla nuevamente.
                throw new ApplicationException("Ocurrió un error al crear el pie de página.", ex);
            }
        }


        /// <summary>
        /// Método en el cual se establece un encabezado que se inserta directamente en el documento de word a partir de una sección
        /// </summary>
        /// <param name="section">Aquí va la sección en la cual estará ubicado el encabezado</param>
        /// <param name="preTitulo">Aquí se pasa un string con el valor del pre-título</param>
        /// <param name="titulo">Aquí se pasa un string con el valor del título</param>
        /// <exception cref="ApplicationException">Aquí está el manejo de expciones que puedne ocurrir si se combian encabezados generados por openxml</exception>
        public static void CrearEncabezado(Word.Section section, string preTitulo, string titulo)
        {
            try
            {
                foreach (Word.HeaderFooter header in section.Headers)
                {
                    // Limpiar completamente el encabezado.
                    header.Range.Text = "";

                    // Eliminar todas las tablas existentes en el encabezado
                    while (header.Range.Tables.Count > 0)
                    {
                        header.Range.Tables[1].Delete();
                    }

                    // Crear una tabla en el encabezado con 2 filas y 2 columnas
                    Word.Table table = header.Range.Tables.Add(header.Range, 2, 2);

                    // Formatear la tabla según sea necesario
                    table.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    table.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    // Configuración y llenado de celdas
                    Word.Cell cell1 = table.Cell(1, 1);
                    cell1.Merge(table.Cell(1, 2)); // Combinar celdas de la primera fila
                    cell1.Range.Text = preTitulo;
                    cell1.Range.Font.Size = 9;
                    cell1.Range.Font.Name = "Arial";

                    Word.Cell cell2 = table.Cell(2, 1);
                    cell2.Range.Text = titulo;
                    cell2.Range.Font.Size = 9;
                    cell2.Range.Font.Name = "Arial";
                    cell2.Range.Font.Bold = 1; // Hacer el texto negrita

                    Word.Cell cell3 = table.Cell(2, 2);
                    Word.Range cellRange = cell3.Range;

                    cellRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    cellRange.Text = "Página ";
                    cellRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    cellRange.Fields.Add(cellRange, Word.WdFieldType.wdFieldPage);

                    cellRange.InsertAfter(" de ");
                    cellRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    cellRange.Fields.Add(cellRange, Word.WdFieldType.wdFieldNumPages);
                    cell3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    cell3.Range.Font.Size = 9;
                    cell3.Range.Font.Name = "Arial";
                }
            }
            catch (Exception ex)
            {
                // Aquí puedes manejar la excepción o lanzarla nuevamente.
                throw new ApplicationException("Ocurrió un error al crear el encabezado.", ex);
            }
        }

        #endregion

    }
}