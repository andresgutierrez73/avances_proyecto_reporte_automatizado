using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using funcionalidades_documento.edicion_footer_header;
using System.Xml.Linq;

namespace funcionalidades_documento.crear_documento
{
    public class FuncionesCreacion
    {
        /// <summary>
        /// Método para guardar el archivo de word en una ruta especifica
        /// </summary>
        /// <returns>Retorna la ruta del directorio en el que se va a guardar el documento</returns>
        public static string GuardarRuta()
        {
            // Obtenemos la fecha actual
            DateTime fechaActual = DateTime.Now;

            // Damos formato a la fecha actual
            string ferchaConFormato = fechaActual.ToString("yyyyMMdd");

            var createFile = new Microsoft.Win32.SaveFileDialog()
            {
                FileName = $"documento_prueba.docx",
                Filter = "Word Files (*.docx)|*.docx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                DefaultExt = "docx"
            };

            var res = createFile.ShowDialog();
            if (res != true) return "";

            return createFile.FileName;
        }

        #region Métodos que usan la librería de openxml
        /// <summary>
        /// Método para generar un documento de Word en una ruta específica
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del docuemento de wordAquí va la ruta del docuemento de word</param>
        /// <exception cref="ArgumentException"></exception>
        public static void GenerarDocumentoWord(string ruta, DimesionHoja tamanoHoja)
        {
            // Validar que la ruta no esté vacía
            if (string.IsNullOrEmpty(ruta))
            {
                throw new ArgumentException("La ruta no puede estar vacía.");
            }

            // Validar la extensión del archivo
            string extension = Path.GetExtension(ruta);
            if (extension != ".docx")
            {
                throw new ArgumentException("La ruta debe tener una extensión .docx");
            }

            // Crear el documento de Word
            using (var document = WordprocessingDocument.Create(ruta, WordprocessingDocumentType.Document))
            {
                // Agregar el MainDocumentPart y establecer el contenido del documento
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document();

                // Asegurarse de que Body está inicializado
                if (mainPart.Document.Body == null)
                {
                    mainPart.Document.Body = new Body();
                }

                // Agregar las propiedades de sección al cuerpo
                SectionProperties sectionProps = new SectionProperties();

                // Variables con la dimensión de la hoja
                double ancho, alto;

                switch (tamanoHoja)
                {
                    case DimesionHoja.A3:
                        ancho = 29.7;
                        alto = 42.0;
                        break;
                    case DimesionHoja.A4:
                        ancho = 21.0;
                        alto = 29.7;
                        break;
                    case DimesionHoja.A5:
                        ancho = 14.8;
                        alto = 21.0;
                        break;
                    case DimesionHoja.B3:
                        ancho = 35.3;
                        alto = 50.0;
                        break;
                    case DimesionHoja.B4:
                        ancho = 25.0;
                        alto = 35.3;
                        break;
                    default:
                        ancho = 21.0;
                        alto = 29.7;
                        break;
                }

                // Definir el tamaño de la hoja como A4
                PageSize pageSize = new PageSize()
                {
                    Width = (UInt32Value)(ancho * 567),  // Ancho para A4 en vigésimos de punto
                    Height = (UInt32Value)(alto * 567)  // Alto para A4 en vigésimos de punto
                };
                sectionProps.Append(pageSize);

                // Agregar las propiedades de sección al cuerpo del documento
                mainPart.Document.Body.Append(sectionProps);

                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// En este método crear una secccion específica dentro dentro del documento para que a partir de este punto en en el que se implemente el método las hojas cambien de orientación
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="orientacion">Este es un enum a partir del cual se pasa un valor que será leido en el método para daterminar cúal es el cambio de orientación a implementar</param>
        public static void CambiarOrientacion(string ruta, Orientacion orientacion)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(ruta, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;

                    if (mainPart == null)
                    {
                        throw new InvalidOperationException("El documento no contiene una parte principal.");
                    }

                    SectionProperties lastSectPr = mainPart.Document.Body.Elements<SectionProperties>().LastOrDefault();

                    // Si no hay ninguna sección en el documento, la creamos.
                    if (lastSectPr == null)
                    {
                        lastSectPr = new SectionProperties();
                        mainPart.Document.Body.Append(lastSectPr);
                    }

                    // Creamos un nuevo salto de sección.
                    Paragraph breakParagraph = new Paragraph();
                    ParagraphProperties paraProps = new ParagraphProperties();
                    SectionProperties newSectPr = (SectionProperties)lastSectPr.CloneNode(true); // Copiamos las propiedades de la sección anterior
                    breakParagraph.Append(paraProps);
                    paraProps.Append(newSectPr);
                    mainPart.Document.Body.Append(breakParagraph);

                    // Definimos la nueva orientación en la sección original
                    PageSize pageSize = new PageSize();

                    switch (orientacion)
                    {
                        case Orientacion.Horizontal:
                            pageSize.Width = (UInt32Value)15840U; // 11 pulgadas
                            pageSize.Height = (UInt32Value)12240U; // 8.5 pulgadas
                            pageSize.Orient = PageOrientationValues.Landscape;
                            break;
                        case Orientacion.Vertical:
                        default:
                            pageSize.Width = (UInt32Value)12240U; // 8.5 pulgadas
                            pageSize.Height = (UInt32Value)15840U; // 11 pulgadas
                            pageSize.Orient = PageOrientationValues.Portrait;
                            break;
                    }

                    // Si la sección anterior ya tiene un elemento PageSize, lo eliminamos para reemplazarlo
                    PageSize oldPageSize = lastSectPr.GetFirstChild<PageSize>();
                    if (oldPageSize != null)
                    {
                        lastSectPr.RemoveChild(oldPageSize);
                    }
                    lastSectPr.Append(pageSize);

                    doc.Save(); // Guardamos los cambios
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Se produjo un error al cambiar la orientación del documento: {ex.Message}");
            }
        }
        #endregion

        #region Métodos que usan la librería de microsoft interop word
        /// <summary>
        /// Método para cambiar la orientación de las hojas de un documento a partir del punto en el cual se llame, este método al crear las secciones hace la implmentacion de otros dos métodos los cuales
        /// retornar un encabezado y pie de página con el propósito de cada que se haga una instancia de este método se puedan tener valores personalizados dentro de los encabezados y pie de página
        /// </summary>
        /// <param name="rutaArchivo">Aquí va la ruta del documento de word</param>
        /// <param name="aHorizontal">Aquí se pasa un valor booleano para determinar sí se hace un cambio de orientación horizontal o vertical</param>
        /// <param name="textoPie">Aquí se pasa un string con el texto que tendrá el pie de página</param>
        /// <param name="preTituloEncabezado">Aquí se pasa un string para personalizar el pre-título del encabezado</param>
        /// <param name="tituloEncabezado">Aquí se pasa un string para personalizar el título del encabezado</param>
        /// <exception cref="ApplicationException"></exception>
        public static void CambiarOrientacionPaginaEnDocumento(string rutaArchivo, bool aHorizontal, string textoPie = "modificarPie", string preTituloEncabezado = "preTituloEncabezado", string tituloEncabezado = "tituloEncabezado")
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                // Iniciar la aplicación Word.
                wordApp = new Word.Application();

                // Abrir el documento.
                doc = wordApp.Documents.Open(rutaArchivo);

                // Cambio de orientación.
                // Agregar un salto de sección al final del documento.
                Word.Range endRange = doc.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endRange.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);

                // Obtener la sección del salto insertado (sería la última sección).
                Word.Section newSection = doc.Sections[doc.Sections.Count];
                newSection.PageSetup.Orientation = aHorizontal ? Word.WdOrientation.wdOrientLandscape : Word.WdOrientation.wdOrientPortrait;

                // Desvincular encabezados y pies de página de la sección anterior en la nueva sección creada.
                foreach (Word.HeaderFooter headerFooter in newSection.Headers)
                {
                    headerFooter.LinkToPrevious = false;
                }

                foreach (Word.HeaderFooter headerFooter in newSection.Footers)
                {
                    headerFooter.LinkToPrevious = false;
                }

                // Eliminar el contenido del encabezado y del pie de página SOLO de la primera página de la nueva sección.
                newSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
                newSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";

                // Llamar al método para crear el pie de página personalizado.
                EditarEncabezadoPie.CrearPieDePagina(newSection, textoPie); // Pie de página para las demás páginas
                EditarEncabezadoPie.CrearPieDePagina(newSection, textoPie, true); // Pie de página para la primera página

                // Crear el encabezado personalizado.
                EditarEncabezadoPie.CrearEncabezado(newSection, preTituloEncabezado, tituloEncabezado);

                // Actualizar los campos del documento.
                doc.Repaginate();
                doc.Fields.Update();

                // Guardar los cambios.
                doc.Save();
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Ocurrió un error al cambiar la orientación de la página.", ex);
            }
            finally
            {
                // Cerrar el documento y liberar recursos.
                if (doc != null) doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                if (wordApp != null) wordApp.Quit();
            }
        }
        
        /// <summary>
        /// Método que funciona como complemento y debe insertarse al final de la creación del documento, este tiene la finalidad de refrescar todos los campos de word, para evitar la tarea manual por
        /// parte del usuario para actualzar campos uno por uno
        /// </summary>
        /// <param name="rutaArchivo">Aquí se pasa la ruta del archivo</param>
        public static void ActualizarCamposEnWord(string rutaArchivo)
        {
            // Crear una nueva aplicación Word
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                // Abrir el documento
                doc = wordApp.Documents.Open(rutaArchivo);

                // Actualizar la Tabla de Tablas
                foreach (Word.TableOfContents toc in doc.TablesOfContents)
                {
                    toc.Update();
                }

                // Actualizar la Tabla de Ilustraciones
                foreach (Word.TableOfFigures tof in doc.TablesOfFigures)
                {
                    tof.Update();
                }

                // Seleccionar todo el contenido del documento
                Word.Range range = doc.Content;

                // Actualizar todos los campos en el rango seleccionado
                range.Fields.Update();

                // Guardar y cerrar el documento
                doc.Save();
                doc.Close();
            }
            finally
            {
                // Cerrar Word
                if (doc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }
        #endregion

        #region Enumeraciones para los valores constantes
        /// <summary>
        /// Enum para los valores constantes de la decoracion de los textos
        /// </summary>
        public enum EstiloParrafo
        {
            Normal,
            Negrita,
            Italico,
            Subrayado
        }


        /// <summary>
        /// Enum para los valores constantes de la alineación de textos
        /// </summary>
        public enum AlineacionTexto
        {
            Izquierda,
            Derecha,
            Centro,
            Justificado
        }

        /// <summary>
        /// Enum para los valores constantes de la alineación de la imagen
        /// </summary>
        public enum AlineacionImagen
        {
            Izquierda,
            Centro,
            Derecha
        }

        /// <summary>
        /// Enum con los valores se los tamaños de hoja más comunes
        /// </summary>
        public enum DimesionHoja
        {
            A3,
            A4,
            A5,
            B3,
            B4
        }

        /// <summary>
        /// Enum con los valores de orientacion de las hojas del documento
        /// </summary>
        public enum Orientacion
        {
            Vertical,
            Horizontal
        }
        #endregion

    }
}