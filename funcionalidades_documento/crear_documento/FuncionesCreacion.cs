using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = Microsoft.Office.Interop.Word;

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
        /// Método para cambiar la orientacion de una pagina determinada en un docmento de word,
        /// antes de cambiar la orientacion es mejor usar antes un salto de pagina,
        /// para que de esta forma no se cambie paginas las cuales no se necesitan
        /// </summary>
        /// <param name="ruta">Aquí va la ubicacion del archivo de word</param>
        /// <param name="orientacion">Aquí se pasa un enum con los valores de la orientacion del documento</param>
        /// <param name="tamanoHoja">Aquí se pasa una enum con lo tamaños de la hoja que se ncesiten</param>
        public static void CambiarOrientacionDePaginaActual(string rutaDocumento)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document document = null;

            try
            {
                // Abrir el documento
                document = wordApp.Documents.Open(rutaDocumento);

                // Insertar un salto de sección en la posición actual del cursor
                wordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);

                // Mover la selección al inicio de la próxima sección
                wordApp.Selection.MoveDown(Word.WdUnits.wdLine, 1, Word.WdMovementType.wdMove);

                // Cambiar la orientación de esta sección (la página actual)
                Word.Section section = document.Sections[wordApp.Selection.Sections[1].Index];
                section.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;  // Cambia a wdOrientPortrait para orientación vertical

                // Insertar otro salto de sección para separar la página actual del resto del documento
                wordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);

                // Cambiar la orientación de la sección siguiente (resto del documento) a la original
                Word.Section nextSection = document.Sections[wordApp.Selection.Sections[1].Index];
                nextSection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait; // o lo que sea la orientación original

                // Guardar y cerrar
                document.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Cerrar el documento y Word
                document?.Close();
                wordApp.Quit();
            }
        }

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
    }
}