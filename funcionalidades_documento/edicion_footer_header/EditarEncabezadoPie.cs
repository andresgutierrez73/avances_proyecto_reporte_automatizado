using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using funcionalidades_documento.funciones_parrafo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;

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

        public static Header NewHeader(string titulo, string preTitulo)
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

                var header = NewHeader(textoArriba, textoAbajo);
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
                    HeaderReference headerReference = new HeaderReference { Id = mainPart.GetIdOfPart(headerPart) };
                    sectionProperties.RemoveAllChildren<HeaderReference>();
                    sectionProperties.PrependChild(headerReference);
                }
                else
                {
                    mainPart.Document.Body.Append(new SectionProperties(new HeaderReference { Id = mainPart.GetIdOfPart(headerPart) }));
                }

                document.Save();
            }
        }

        public static Footer NewFooter(string texto)
        {
            Footer footer = new Footer();

            // Agregar declaraciones de espacio de nombres si es necesario
            // ...

            // Crear un párrafo con el texto del pie de página
            Paragraph paragraph = new Paragraph(
                new Run(
                    new Text(texto)
                )
            );

            // Agregar el párrafo al pie de página
            footer.Append(paragraph);

            return footer;
        }

        public static void EditarPieDePagina(string ruta, string texto)
        {
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var mainPart = document.MainDocumentPart;

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

                var footer = NewFooter(texto);

                footerPart.Footer = footer;

                // Agregar la referencia al pie de página en el documento principal
                if (mainPart.Document.Body.Elements<SectionProperties>().Any())
                {
                    SectionProperties sectionProperties = mainPart.Document.Body.Elements<SectionProperties>().First();
                    FooterReference footerReference = new FooterReference { Id = mainPart.GetIdOfPart(footerPart) };
                    sectionProperties.RemoveAllChildren<FooterReference>();
                    sectionProperties.PrependChild(footerReference);
                }
                else
                {
                    mainPart.Document.Body.Append(new SectionProperties(new FooterReference { Id = mainPart.GetIdOfPart(footerPart) }));
                }

                document.Save();
            }
        }
    }
}
