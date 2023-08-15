using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


        public static void AgregarTablaDatosGrandes(string ruta, List<List<string>> listaDatos, List<string> listaColumnas)
        {
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

                // Define el borde de la tabla
                DocumentFormat.OpenXml.Wordprocessing.TableBorders tblBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                    new DocumentFormat.OpenXml.Wordprocessing.TopBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.BottomBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.LeftBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.RightBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 },
                    new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single), Size = 1 }
                );

                DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(tableWidth, tblBorders);
                table.Append(tblProperties);

                // Agrega una fila de encabezado con los valores de listaColumnas
                DocumentFormat.OpenXml.Wordprocessing.TableRow headerRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                foreach (var header in listaColumnas)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                        new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                            new DocumentFormat.OpenXml.Wordprocessing.Run(
                                new DocumentFormat.OpenXml.Wordprocessing.Text(header)
                            )
                        )
                    );
                    headerRow.Append(cell);
                }
                table.Append(headerRow);

                // Crea las filas y columnas restantes
                int maxRowCount = listaDatos.Max(list => list.Count);
                for (int i = 0; i < maxRowCount; i++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    for (int j = 0; j < listaColumnas.Count; j++)
                    {
                        var cellValue = i < listaDatos[j].Count ? listaDatos[j][i] : string.Empty;

                        DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                new DocumentFormat.OpenXml.Wordprocessing.Run(
                                    new DocumentFormat.OpenXml.Wordprocessing.Text(cellValue)
                                )
                            )
                        );
                        row.Append(cell);
                    }
                    table.Append(row);
                }

                // Añade la tabla al documento
                body.Append(table);

                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Se agregó una tabla al documento.");
        }

        public static void AgregarTablaDesdeLista(string ruta, List<List<string>> datos)
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
                tblProperties.Append(tblBorders);
                table.Append(tblProperties);

                // Añadir las filas desde datos
                foreach (var rowData in datos)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                    // Si hay solo un elemento en la fila, hacer que ocupe todas las columnas
                    if (rowData.Count == 1)
                    {
                        DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        cell.TableCellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                            new DocumentFormat.OpenXml.Wordprocessing.GridSpan() { Val = maxColumnCount }
                        );

                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                        DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center }
                        );
                        paragraph.Append(paragraphProperties);

                        paragraph.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(rowData[0])));

                        cell.Append(paragraph);
                        row.Append(cell);
                    }
                    else
                    {
                        for (int i = 0; i < maxColumnCount; i++)
                        {
                            DocumentFormat.OpenXml.Wordprocessing.TableCell cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();

                            string cellData = i < rowData.Count ? rowData[i] : string.Empty;

                            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                            DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                                new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center }
                            );
                            paragraph.Append(paragraphProperties);

                            paragraph.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cellData)));

                            cell.Append(paragraph);
                            cell.TableCellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                                new DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                            );

                            row.Append(cell);
                        }
                    }

                    table.Append(row);
                }

                // Añade la tabla al documento
                body.Append(table);

                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Se agregó una tabla al documento.");
        }

    }
}
