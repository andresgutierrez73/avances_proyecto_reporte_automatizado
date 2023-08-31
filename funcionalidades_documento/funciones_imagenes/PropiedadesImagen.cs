using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.IO;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;
using System.Linq;

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
        /// Método para generar el título con contador para las imágenes
        /// </summary>
        /// <param name="mensaje">Aquí se pasa el título que va a ir ubicado encima de la imagen que se agrega al contenido</param>
        /// <returns>Retorna el párrafo con estilo para que pueda ser insertado en el documento</returns>
        public static Paragraph TituloImagen(string mensaje)
        {
            Paragraph paragraph = new Paragraph();

            Justification justification = new Justification() { Val = JustificationValues.Center };
            ParagraphProperties paragraphProperties = new ParagraphProperties(justification);
            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties(new Languages() { Val = "es-CO" });

            paragraphProperties.Append(paragraphMarkRunProperties);
            paragraph.Append(paragraphProperties);

            RunProperties runProperties = new RunProperties(
                new Languages() { Val = "es-CO" },
                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" },
                new FontSize() { Val = "24" },
                new Bold()
            );

            paragraph.Append(new Run(runProperties.CloneNode(true), new Text("Ilustración ")));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldCode(" SEQ Ilustración \\* ARABIC ")));
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.Separate }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text(" "))); // Esto generará el número de secuencia.
            paragraph.Append(new Run(runProperties.CloneNode(true), new FieldChar() { FieldCharType = FieldCharValues.End }));
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text(": " + mensaje.Trim())));

            return paragraph;
        }


        /// <summary>
        /// Método para agregar una imágen a partir de una ruta del escritorio
        /// </summary>
        /// <param name="rutaDocumento">Aquí va la ruta del documento de word</param>
        /// <param name="rutaImagen">Aquí se pasa un string con la ruta del imágen en el explorador de archivos</param>
        /// <param name="ancho">Aquí se pasa la longitud en cm del ancho de la imágen del documento</param>
        /// <param name="alto">Aquí se pasa la longitud en cm del alto de la imágen del documento</param>
        /// <param name="alineacion">Aquí se pasa un enum con un valor el cual determinará la alineación de la imágen dentro
        /// del documento</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarImagenDesdeArchivo(string rutaDocumento, string rutaImagen, int ancho, int alto, AlineacionImagen alineacion, string tituloImagen = null)
        {
            // Aquí se multiplican los valores por esta cantidad, para hacer la conversión de EMU a cm
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

                // Obtiene el último Id utilizado para imágenes en el documento
                UInt32Value lastId = 1U;
                if (document.MainDocumentPart.Document.Descendants<DW.DocProperties>().Any())
                {
                    lastId = document.MainDocumentPart.Document.Descendants<DW.DocProperties>().Max(p => p.Id.Value);
                    lastId++;
                }

                var element = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = ancho, Cy = alto },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = lastId, Name = "Picture " + lastId },
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

                // Creamos la variable align la cual va a cambiar dependiendo de los valores que se pasen por parametro en el enum
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

                // Después de agregar la imagen, añade el título si es proporcionado
                if (!string.IsNullOrEmpty(tituloImagen))
                {
                    Paragraph caption = TituloImagen(tituloImagen);
                    document.MainDocumentPart.Document.Body.AppendChild(caption);
                }

                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Imagen {rutaImagen} añadida al documento {rutaDocumento}");
        }

        /// <summary>
        /// Método para insertar una imágen decodificada en base64, este método inserta directamente la imágen dentro del
        /// contenido del documento, por lo que si se quiere insertar algo inmediato dentro del documento este es el método a usar
        /// </summary>
        /// <param name="rutaDocumento">Aquí va la ruta del documento de word</param>
        /// <param name="imagenBase64">Aquí se pasa un string que tenga toda la cadena de caracteres en base64 de la imágen</param>
        /// <param name="ancho">Aquí se pasa la longitud en cm del ancho de la imágen del documento</param>
        /// <param name="alto">Aquí se pasa la longitud en cm del alto de la imágen del documento</param>
        /// <param name="alineacion">Aquí se pasa un enum con un valor el cual determinará la alineación de la imágen dentro
        /// del documento</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarImagenDesdeBase64(string rutaDocumento, string imagenBase64, int ancho, int alto, AlineacionImagen alineacion, string tituloImagen = null)
        {
            // Aquí se multiplican los valores por esta cantidad, para hacer la conversión de EMU a cm
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

                byte[] imageBytes = Convert.FromBase64String(imagenBase64);

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

                // Creamos la variable align la cual va a cambiar dependiendo de los valores que se pasen por parámetro en el enum
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
                        align = "Center";
                        break;
                }

                var paragraph = new Paragraph(new Run(element));
                JustificationValues jv;
                if (!Enum.TryParse<JustificationValues>(align, out jv))
                {
                    jv = JustificationValues.Left;
                }
                paragraph.ParagraphProperties = new ParagraphProperties(new Justification() { Val = jv });

                // Primero añadimos la imagen al cuerpo del documento.
                document.MainDocumentPart.Document.Body.AppendChild(paragraph);

                // Luego, si se proporcionó un título para la imagen, se añade después de la imagen.
                if (!string.IsNullOrEmpty(tituloImagen))
                {
                    Paragraph caption = TituloImagen(tituloImagen);
                    document.MainDocumentPart.Document.Body.AppendChild(caption);
                }

                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Imagen añadida al documento {rutaDocumento}. Título: \"{tituloImagen ?? "Sin título"}\"");
        }



        /// <summary>
        /// Método para insertar una imágen decodificada en base64, este método necesita una mayor manupulación ya que no inserta
        /// directemente la imágen en el contenido del documento, este en el caso del reporte se usa como complemento de un método
        /// que inserta una tabla con imágenes
        /// </summary>
        /// <param name="rutaDocumento">Aquí va la ruta del documento de word</param>
        /// <param name="imagenBase64">Aquí se pasa un string que tenga toda la cadena de caracteres en base64 de la imágen</param>
        /// <param name="ancho">Aquí se pasa la longitud en cm del ancho de la imágen del documento</param>
        /// <param name="alto">Aquí se pasa la longitud en cm del alto de la imágen del documento</param>
        /// <param name="alineacion">Aquí se pasa un enum con un valor el cual determinará la alineación de la imágen dentro
        /// del documento</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static Drawing ObtenerImagenDesdeBase64(MainDocumentPart mainPart, string imagenBase64, int ancho, int alto)
        {
            // Conversión de dimensiones de centímetros a EMU.
            ancho *= 360000;
            alto *= 360000;

            // Creación de una nueva parte de imagen en el documento de tipo JPEG.
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            byte[] imageBytes = Convert.FromBase64String(imagenBase64);

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
            );

            return element;
        }

        /// <summary>
        /// Método sobreconstruido para solucionar solucionar el problema de agregar las imagenes solo al cuerpo del documento
        /// con este se puede insertar tanto en el documento como en el encabezado y pie de pagina
        /// </summary>
        /// <param name="part">Aquí se pasa la parte principal del documento, ya sea dle cuerpo, cabecera o pie de pagina para retornar las imagen</param>
        /// <param name="imagenBase64">Aquí se pasa un string que tenga toda la cadena de caracteres en base64 de la imágen</param>
        /// <param name="ancho">Aquí se pasa la longitud en cm del ancho de la imágen del documento</param>
        /// <param name="alto">Aquí se pasa la longitud en cm del alto de la imágen del documento</param>
        /// <returns>Retorna un tipo de objeto de la libreria que permite instar una imagen para que puede ser usada por
        /// otro método que inserte la información directamente en el documento</returns>
        /// <exception cref="ArgumentException"></exception>
        public static Drawing ObtenerImagenDesdeBase64(OpenXmlPart part, string imagenBase64, int ancho, int alto)
        {
            // Conversión de dimensiones de centímetros a EMU.
            ancho *= 360000;
            alto *= 360000;

            ImagePart imagePart;

            if (part is MainDocumentPart mainPart)
            {
                imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            }
            else if (part is HeaderPart headerPart)
            {
                imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
            }
            else if (part is FooterPart footerPart)
            {
                imagePart = footerPart.AddImagePart(ImagePartType.Jpeg);
            }
            else
            {
                throw new ArgumentException("Tipo de parte no compatible", nameof(part));
            }

            byte[] imageBytes = Convert.FromBase64String(imagenBase64);

            using (MemoryStream stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            string relationshipId = part.GetIdOfPart(imagePart);

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
            );

            return element;
        }
    }
}