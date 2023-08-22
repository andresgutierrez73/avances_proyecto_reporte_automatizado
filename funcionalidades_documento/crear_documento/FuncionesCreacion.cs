using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

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
        public static void GenerarDocumentoWord(string ruta)
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
                mainPart.Document = new Document(new Body());
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
            Derecha,
            Arriba,
            Abajo
        }
    }
}