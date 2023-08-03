using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;
using System.Security.AccessControl;

namespace funcionalidades_documento.funciones_imagenes
{
    public class PropiedadesImagen
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
        /// Método para crear un directorio temporal para guardar las imágenes
        /// </summary>
        /// <returns>Retorna la ruta del explorador de archivos en la cual se encuentra la carpeta</returns>
        public static string CrearDirectorioTemporal()
        {
            // Creamos una variable con el nombre del directorio temporal
            string nombreDirectorioTemporal = "temp_imagenes";

            // Creamos el directorio temporal con los permisos adecuados
            Directory.CreateDirectory(nombreDirectorioTemporal);
            var security = Directory.GetAccessControl(nombreDirectorioTemporal);
            security.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow));
            Directory.SetAccessControl(nombreDirectorioTemporal, security);

            return nombreDirectorioTemporal;
        }

        /// <summary>
        /// Método para eliminar el directorio temporal en el que se guardan las imágenes
        /// </summary>
        /// <param name="directorioTemporal">Esta es la ruta en la cual está el direcotorio que se va a eliminar
        /// lo ideal es que sea la ruta del directorio temporal que se retorno en el método anterior
        /// </param>
        public static void EliminarDirectorioTemporal(string directorioTemporal)
        {
            // Creamos una variable con el nombre del directorio temporal
            string nombreDirectorioTemporal = "temp_imagenes";

            // Validamos que el directorio exista
            if (Directory.Exists(nombreDirectorioTemporal))
            {
                Directory.Delete(nombreDirectorioTemporal, true);

                // Si el directorio en este punto no existe, entonces se eliminó correctamente
                if (!Directory.Exists(nombreDirectorioTemporal))
                {
                    Console.WriteLine("el directorio temporal se eliminó correctamente");
                }
            }
        }

        /// <summary>
        /// Método para agregar una imagen en base64
        /// </summary>
        /// <param name="rutaDocumento">Aquí va la ruta del documento de word</param>
        /// <param name="base64Image">Aquí va un string con la imagen decodificada en base64</param>
        /// <param name="ancho">Aquí va la longitud en cm del ancho de la imágen</param>
        /// <param name="alto">Aquí va la longitud en cm del alto de la imágen</param>
        /// <param name="alineacion">Aquí se pasa un enum con los valores de alineación de la imágen que se pueden editar con la librería</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarImagenArchivoBase64(string rutaDocumento, string base64Image, int ancho, int alto, AlineacionImagen alineacion)
        {
            // Aquí se multiplican los valores por esta cantidad, para hacer la convesión de EMU a cm
            ancho *= 360000;
            alto *= 360000;

            ValidarRutaArchivo(rutaDocumento);

            using (var document = WordprocessingDocument.Open(rutaDocumento, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                MainDocumentPart mainPart = document.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                // Convierte el código base64 a bytes
                byte[] imageBytes = Convert.FromBase64String(base64Image);

                using (MemoryStream stream = new MemoryStream(imageBytes))
                {
                    imagePart.FeedData(stream);
                }

                string relationshipId = mainPart.GetIdOfPart(imagePart);

                var element = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = ancho, Cy = alto },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })
                                        )
                                        { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = ancho, Cy = alto }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

                // Resto del código sin cambios
                // ...
            }

            Console.WriteLine($"Imagen añadida al documento {rutaDocumento}");
        }

        /// <summary>
        /// Método para agregar una imágen a partir de una ruta del escritorio
        /// </summary>
        /// <param name="rutaDocumento">Aquí va la ruta del documento de word</param>
        /// <param name="rutaImagen">Aquí se pasa un string con la ruta del imágen en el explorador de archivos</param>
        /// <param name="ancho"></param>
        /// <param name="alto"></param>
        /// <param name="alineacion"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarImagenDesdeArchivo(string rutaDocumento, string rutaImagen, int ancho, int alto, AlineacionImagen alineacion)
        {
            // Aquí se multiplican los valores por esta cantidad, para hacer la convesión de EMU a cm
            ancho *= 360000;
            alto *= 360000;

            ValidarRutaArchivo(rutaDocumento);

            using (var document = WordprocessingDocument.Open(rutaDocumento, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                MainDocumentPart mainPart = document.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(rutaImagen, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                string relationshipId = mainPart.GetIdOfPart(imagePart);

                var element = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = ancho, Cy = alto },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })
                                        )
                                        { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = ancho, Cy = alto }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

                // Creamos la variable aling la cual va a cambiar dependiendo de los valores que se pasen por parametro en el enum
                string align = "";

                // Usamos la estructura de control switch para modificar la variable anterior
                switch (alineacion)
                {
                    case AlineacionImagen.Izquierda:
                        align = "Left";
                        break;
                    case AlineacionImagen.Derecha:
                        align = "Right";
                        break;
                    case AlineacionImagen.Centro:
                        align = "Center";
                        break;
                    default:
                        align = "Left";
                        break;
                }

                var paragraph = new Paragraph(new Run(element));
                JustificationValues jv;
                if (!Enum.TryParse<JustificationValues>(align, out jv))
                {
                    jv = JustificationValues.Left;
                }
                paragraph.ParagraphProperties = new ParagraphProperties(new Justification() { Val = jv });
                document.MainDocumentPart.Document.Body.AppendChild(paragraph);
                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Imagen {rutaImagen} añadida al documento {rutaDocumento}");
        }

        /// <summary>
        /// Método que pasa una imagen a base64
        /// </summary>
        /// <param name="rutaImagen">Aquí se pasa un string con la ruta del explorador de archivos en la cual se encuentra la imágen</param>
        /// <returns>Retorna un string con la imagen decodificada en base64</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="Exception"></exception>
        public static string ImagenABase64(string rutaImagen)
        {
            if (string.IsNullOrEmpty(rutaImagen))
            {
                throw new ArgumentException("La ruta de la imagen no puede estar vacía.");
            }

            byte[] imagenBytes;
            try
            {
                imagenBytes = System.IO.File.ReadAllBytes(rutaImagen);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al leer el archivo de la imagen. Detalles: " + ex.Message);
            }

            string imagenBase64 = Convert.ToBase64String(imagenBytes);

            return imagenBase64;
        }

        /// <summary>
        /// Método que pasa string de base64 a imagen
        /// </summary>
        /// <param name="base64String">Aquí se pasa por parametro un string de base64 con una imágen</param>
        /// <param name="rutaSalida">Aquí se pasa la ruta del explorador de archivos en la cual se va a guardar las imagen cuando se codifique el base 64</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="Exception"></exception>
        public static string Base64AImagen(string base64String, string rutaSalida)
        {
            if (string.IsNullOrEmpty(base64String))
            {
                throw new ArgumentException("La cadena Base64 no puede estar vacía.");
            }

            byte[] imageBytes;
            try
            {
                imageBytes = Convert.FromBase64String(base64String);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al convertir la cadena Base64 en bytes. Detalles: " + ex.Message);
            }

            try
            {
                System.IO.File.WriteAllBytes(rutaSalida, imageBytes);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al escribir los bytes de la imagen en el archivo. Detalles: " + ex.Message);
            }

            return rutaSalida;
        }
    }
}