using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;

namespace funcionalidades_documento.funciones_parrafo
{
    public class PropiedadesParrafo
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
        /// Método para agregar un párrafo con estilo y alineación indicados al documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="texto">Aquí va el texto del párrafo que se va a mostrar</param>
        /// <param name="tamanoFuente">Aquí se pasa el valor numérico de los tamaños de fuente</param>
        /// <param name="estilo">Aquí se pasa un enum con los estilos de letra predefinidos del word</param>
        /// <param name="alineacion">Aquí se pasa un enum con los tipos de alineación que hay en word</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarParrafo(string ruta, string texto, int tamanoFuente, EstiloParrafo estilo, AlineacionTexto alineacion)
        {
            // Por defecto, al librería de OpenXML divide a la mitad el valor que se ingresa
            // como tamaño de la fuente, por ello en el método debe multiplicarse este valor por 2
            tamanoFuente *= 2;

            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                var runProperties = new RunProperties(new FontSize { Val = tamanoFuente.ToString() });

                // Aplicar el estilo correspondiente
                switch (estilo)
                {
                    case EstiloParrafo.Normal:
                        // No se aplica ningún estilo adicional
                        break;
                    case EstiloParrafo.Negrita:
                        runProperties.Append(new Bold());
                        break;
                    case EstiloParrafo.Italico:
                        runProperties.Append(new Italic());
                        break;
                    case EstiloParrafo.Subrayado:
                        runProperties.Append(new Underline { Val = UnderlineValues.Single });
                        break;
                    default:
                        // Si se proporciona un estilo no reconocido, se asume el estilo normal
                        break;
                }

                var run = new Run(runProperties, new Text(texto));
                var paragraph = new Paragraph(run);

                // Aplicar la alineación correspondiente
                switch (alineacion)
                {
                    case AlineacionTexto.Izquierda:
                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Left });
                        break;
                    case AlineacionTexto.Derecha:
                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Right });
                        break;
                    case AlineacionTexto.Centro:
                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
                        break;
                    case AlineacionTexto.Justificado:
                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Both });
                        break;
                    default:
                        // Si se proporciona una alineación no reconocida, se asume alineación izquierda
                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Left });
                        break;
                }

                body.AppendChild(paragraph);
            }

            Console.WriteLine("Agregando párrafo al documento: " + texto);
        }

        /// <summary>
        /// Método para definir la numeración multinivel
        /// </summary>
        /// <param name="document">Aquí se recibe como parametro el documento que se se va a editar.</param>
        private static void AsegurarDefinicionNumeracion(WordprocessingDocument document)
        {
            NumberingDefinitionsPart numberingPart;
            if (document.MainDocumentPart.NumberingDefinitionsPart == null)
            {
                numberingPart = document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }
            else
            {
                numberingPart = document.MainDocumentPart.NumberingDefinitionsPart;
            }

            if (numberingPart.Numbering.Elements<AbstractNum>().Count() == 0)
            {
                var abstractNum = new AbstractNum(
                    Enumerable.Range(1, 9).Select(i =>
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = string.Join("", Enumerable.Range(1, i).Select(j => "%" + j + ".")) }, // Ajuste aquí
                            new LevelJustification { Val = LevelJustificationValues.Left })
                        { LevelIndex = i - 1 }))
                { AbstractNumberId = 1 };

                var numberingInstance = new NumberingInstance(
                    new AbstractNumId { Val = 1 })
                { NumberID = 1 };

                numberingPart.Numbering.Append(abstractNum);
                numberingPart.Numbering.Append(numberingInstance);
            }
        }

        /// <summary>
        /// Método para agregar un título con nivel específico y estilo personalizado
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="titulo">Aquí va el texto que se va a mostrar en el título</param>
        /// <param name="nivelTitulo">Aquí va el nivel de título que ofrece el office</param>
        /// <param name="tamanoFuente">Aquí se pasa un valor numérico con el valor del tamaño de la fuente</param>
        /// <param name="estilo">Aquí se pasa un enum con los estilos de letra predefinidos del word</param>
        /// <param name="alineacion">Aquí se pasa un enum con los tipos de alineación que hay en word</param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarTitulo(string ruta, string titulo, int nivelTitulo, int tamanoFuente, EstiloParrafo estilo, AlineacionTexto alineacion)
        {
            if (nivelTitulo < 1 || nivelTitulo > 9)
            {
                throw new ArgumentException("El nivel del título debe estar entre 1 y 9.");
            }

            ValidarRutaArchivo(ruta);

            // Multiplicar el tamaño de la fuente por 2, ya que OpenXML usa unidades de media punta
            tamanoFuente *= 2;

            // Creamos la variable que va a tener los valores de justiicacion (alineacion de texto)
            var alineacionTexto = JustificationValues.Start;

            // Switch case para para evaluar el valor de la alineacion.
            switch (alineacion)
            {
                case AlineacionTexto.Izquierda:
                    alineacionTexto = JustificationValues.Start;
                    break;
                case AlineacionTexto.Centro:
                    alineacionTexto = JustificationValues.Center;
                    break;
                case AlineacionTexto.Derecha:
                    alineacionTexto = JustificationValues.End;
                    break;
                default:
                    alineacionTexto = JustificationValues.Start;
                    break;
            }

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                // Asegurarse de que el documento tiene una definición de numeración
                AsegurarDefinicionNumeracion(document);

                string styleId = "Titulo" + nivelTitulo;
                string styleName = "Titulo " + nivelTitulo;

                // Crea el estilo "TituloX"
                Style style = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    StyleName = new StyleName() { Val = styleName },
                    BasedOn = new BasedOn() { Val = "Titulo" + (nivelTitulo - 1) }, // Basado en el estilo del nivel anterior
                    NextParagraphStyle = new NextParagraphStyle() { Val = "Titulo" + (nivelTitulo + 1) }, // Siguiente estilo de párrafo
                    PrimaryStyle = new PrimaryStyle(),
                    UnhideWhenUsed = new UnhideWhenUsed(),
                    StyleRunProperties = new StyleRunProperties()
                    {
                        Bold = estilo == EstiloParrafo.Negrita ? new Bold() : null,
                        Italic = estilo == EstiloParrafo.Italico ? new Italic() : null,
                        Underline = estilo == EstiloParrafo.Subrayado ? new Underline() : null,
                        FontSize = new FontSize() { Val = tamanoFuente.ToString() },
                        Color = new Color() { Val = "000000" } // Establecer el color de fuente a negro
                    },
                    StyleParagraphProperties = new StyleParagraphProperties()
                    {
                        OutlineLevel = new OutlineLevel() { Val = nivelTitulo - 1 },  // Añade un nivel de esquema al estilo
                        Justification = new Justification() { Val = alineacionTexto } // Alineación del texto
                    }
                };

                // Comprobar si ya existen definiciones de estilo en el documento
                StyleDefinitionsPart stylesPart;
                if (document.MainDocumentPart.StyleDefinitionsPart != null)
                {
                    // Si existen, obtener la primera
                    stylesPart = document.MainDocumentPart.StyleDefinitionsPart;
                }
                else
                {
                    // Si no existen, crear una nueva
                    stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

                    // Crea un nuevo conjunto de estilos si no existe
                    stylesPart.Styles = new Styles();
                }

                // Añadir el nuevo estilo a las definiciones de estilo
                stylesPart.Styles.Append(style);
                stylesPart.Styles.Save();

                // Configurar propiedades de numeración
                NumberingProperties numberingProperties = new NumberingProperties(
                    new NumberingLevelReference() { Val = nivelTitulo - 1 },  // Los niveles de numeración en Word comienzan en 0, por lo que restamos 1 del nivel del título
                    new NumberingId() { Val = 1 }  // Cambia este valor a la ID de la definición de numeración que estás utilizando
                );

                // Crear y configurar las propiedades del párrafo
                ParagraphProperties paraProps = new ParagraphProperties(
                    new ParagraphStyleId() { Val = styleId },
                    numberingProperties,
                    new Justification() { Val = alineacionTexto } // Alineación del texto
                );

                // Crear el párrafo con el título y sus propiedades
                Paragraph para = new Paragraph(paraProps,
                    new Run(
                        new Text(titulo)
                    )
                );

                // Agregar párrafo al cuerpo del documento
                body.Append(para);
                document.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Agregando título numerado nivel {nivelTitulo}: {titulo}");
        }

        /// <summary>
        /// Método para agregar saltos de linea
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="numSaltos">Aquí se pasa un valor numérico con los saltos de línea</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarSaltosDeLinea(string ruta, int numSaltos)
        {
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                for (int i = 0; i < numSaltos; i++)
                {
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }

            Console.WriteLine($"Agregando {numSaltos} saltos de línea al documento");
        }

        /// <summary>
        /// Método para agregar un salto de pagina en el documento de word
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarSaltoDePagina(string ruta)
        {
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
            }

            Console.WriteLine("Agregando un salto de página al documento");
        }

        /// <summary>
        /// Método para agregar una cita bibliográfica al final del documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="autor">Aquí se pasa el nombre del autor</param>
        /// <param name="titulo">Aquí se pasa el título dle artículo o de la fuente de información</param>
        /// <param name="publicador">Aquí se pone el nombre de la empresa o lugar donde trabaja el autor</param>
        /// <param name="fechaPublicacion">Aquí se pasa la fecha de publicación</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarCitaBibliografica(string ruta, string autor, string titulo, string publicador, string fechaPublicacion)
        {
            ValidarRutaArchivo(ruta);

            // Definimos la variable que contendrá la cita
            string citaBibliografica = "";

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                // Formato de cita bibliográfica APA
                citaBibliografica = $"{autor}. ({fechaPublicacion}). {titulo}. {publicador}.";

                var run = new Run(new Text(citaBibliografica));
                var paragraph = new Paragraph(run);

                body.AppendChild(paragraph);
            }

            Console.WriteLine("Agregando cita bibliográfica al documento: " + citaBibliografica);
        }

        /// <summary>
        /// Método para la creación de una tabla de contenido
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="tituloTabla">Aquí va el título que va a tener la tabla de contenido</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void TablaContenido(string ruta, string tituloTabla)
        {
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var doc = document.MainDocumentPart.Document;

                SdtBlock block = new SdtBlock();

                SdtProperties sdtProperties = new SdtProperties(
                    new SdtContentDocPartObject(
                        new DocPartGallery() { Val = "Table of Contents" },
                        new DocPartUnique())
                );
                block.Append(sdtProperties);

                SdtContentBlock sdtContent = new SdtContentBlock();

                // Establecer el título de la tabla de contenido en negrita y centrado
                Run tituloRun = new Run(new Text(tituloTabla));
                RunProperties tituloRunProperties = new RunProperties();
                tituloRunProperties.Append(new Bold());
                tituloRun.PrependChild<RunProperties>(tituloRunProperties);

                ParagraphProperties tituloParaProperties = new ParagraphProperties();
                tituloParaProperties.Append(new Justification() { Val = JustificationValues.Center });
                Paragraph Titulo = new Paragraph(tituloRun);
                Titulo.PrependChild<ParagraphProperties>(tituloParaProperties);

                sdtContent.Append(Titulo);

                Paragraph Contenido = new Paragraph(
                    new ParagraphProperties(
                        new RunProperties(
                            new NoProof())
                        ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true }
                        ),
                    new Run(
                        new FieldCode(@"TOC \o ""1-3"" \h \z \u") { Space = SpaceProcessingModeValues.Preserve }
                        ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Separate }
                        )
                    );
                sdtContent.Append(Contenido);

                Paragraph ContenEnd = new Paragraph(
                    new Run(
                        new RunProperties(
                            new Bold(),
                            new NoProof()
                            ),
                         new FieldChar { FieldCharType = FieldCharValues.End }
                        )

                    );
                sdtContent.Append(ContenEnd);

                block.Append(sdtContent);
                doc.Body.AppendChild(block);

                var docSettings = document.MainDocumentPart.DocumentSettingsPart;
                if (docSettings == null)
                {
                    docSettings = document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    docSettings.Settings = new Settings();
                    docSettings.Settings.Append(new UpdateFieldsOnOpen() { Val = true });
                }
            }

            Console.WriteLine("Agregando tabla de contenido al documento.");
        }
    }
}