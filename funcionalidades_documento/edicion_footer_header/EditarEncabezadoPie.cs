using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using funcionalidades_documento.funciones_tablas;
using System;
using System.Collections.Generic;
using System.Linq;
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

        /// <summary>
        /// Método para establecer el encabezado con el formato del documento
        /// </summary>
        /// <param name="titulo">Aquí va a el texto que se puede ver arriba en el encabezado</param>
        /// <param name="preTitulo">Aquí va a el texto que se puede ver abajo en el encabezado</param>
        /// <returns></returns>
        public static Header nuevoEncabezado(string titulo, string preTitulo)
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

        /// <summary>
        /// Método que retorna un objeto de tipo Footer con un formato específico como en este caso que es un tabla personalizada con un texto
        /// </summary>
        /// <param name="textoPie">Aquí se recibe un texto, le cual va a ser el que se vera en el pie de página del documento</param>
        /// <returns>Esté método retorna un pie de página, por lo cual no lo agrega directemente al documento</returns>
        public static Footer nuevoPie(string textoPie)
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

        /// <summary>
        /// Método encargado de recibir un texto y de manipula un footer para insertarlo en el documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="texto">Aquí va el texto que se mostrará en todos los pie de página del docuento exceptuando el inicio</param>
        /// <param name="textoPrimeraPagina"></param>
        /// <exception cref="ArgumentNullException">Aquí va el texto que irá en el pie de la primera página del documento</exception>
        public static void EditarPieDePagina(string ruta, string texto, List<List<string>> datosTabla)
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
    }
}