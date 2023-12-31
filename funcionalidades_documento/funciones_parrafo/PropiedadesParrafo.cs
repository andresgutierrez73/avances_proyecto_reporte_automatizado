﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;
using System.Runtime.InteropServices;

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


        #region Métodos que usan la librería de openxml
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

                // Establecer la fuente a Arial
                runProperties.RunFonts = new RunFonts { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

                switch (estilo)
                {
                    case EstiloParrafo.Normal:
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
                        break;
                }

                var run = new Run(runProperties, new Text(texto));
                var paragraph = new Paragraph(run);

                // Configuración para evitar saltos de línea entre párrafos consecutivos
                var spacing = new SpacingBetweenLines() { After = "0" };
                paragraph.ParagraphProperties = new ParagraphProperties(spacing);

                switch (alineacion)
                {
                    case AlineacionTexto.Izquierda:
                        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Left });
                        break;
                    case AlineacionTexto.Derecha:
                        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Right });
                        break;
                    case AlineacionTexto.Centro:
                        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Center });
                        break;
                    case AlineacionTexto.Justificado:
                        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Both });
                        break;
                    default:
                        paragraph.ParagraphProperties.Append(new Justification { Val = JustificationValues.Left });
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
                        Italic = new Italic(),
                        Underline = estilo == EstiloParrafo.Subrayado ? new Underline() : null,
                        FontSize = new FontSize() { Val = tamanoFuente.ToString() },
                        Color = new Color() { Val = "000000" }, // Establecer el color de fuente a negro
                        RunFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
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
        public static void AgregarSaltosDeLinea(string ruta, int numEspacios)
        {
            ValidarRutaArchivo(ruta);

            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                for (int i = 0; i < numEspacios; i++)
                {
                    body.AppendChild(new Paragraph(new Run(new Text("\n"))));
                }
            }

            Console.WriteLine($"Agregando {numEspacios} espacios al documento");
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
                tituloRunProperties.RunFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };  // Fuente Arial
                tituloRun.PrependChild<RunProperties>(tituloRunProperties);

                ParagraphProperties tituloParaProperties = new ParagraphProperties();
                tituloParaProperties.Append(new Justification() { Val = JustificationValues.Center });
                Paragraph Titulo = new Paragraph(tituloRun);
                Titulo.PrependChild<ParagraphProperties>(tituloParaProperties);

                sdtContent.Append(Titulo);

                Paragraph Contenido = new Paragraph(
                    new ParagraphProperties(
                        new RunProperties(
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }  // Fuente Arial
                            )
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
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }  // Fuente Arial
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


        /// <summary>
        /// Método para la creación de una tabla de tablas, en la cual se hace referencia a los titulos de las tablas del documento
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="tituloTabla">Aquí va el título que va a tener la tabla de contenido</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void TablaTablas(string ruta, string tituloTabla)
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
                        new DocPartGallery() { Val = "Table of Tables" },
                        new DocPartUnique())
                );
                block.Append(sdtProperties);

                SdtContentBlock sdtContent = new SdtContentBlock();

                // Establecer el título de la "Tabla de Tablas" en negrita y centrado
                Run tituloRun = new Run(new Text(tituloTabla));
                RunProperties tituloRunProperties = new RunProperties();
                tituloRunProperties.Append(new Bold());
                tituloRunProperties.RunFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };  // Fuente Arial
                tituloRun.PrependChild<RunProperties>(tituloRunProperties);

                ParagraphProperties tituloParaProperties = new ParagraphProperties();
                tituloParaProperties.Append(new Justification() { Val = JustificationValues.Center });
                Paragraph Titulo = new Paragraph(tituloRun);
                Titulo.PrependChild<ParagraphProperties>(tituloParaProperties);

                sdtContent.Append(Titulo);

                Paragraph Contenido = new Paragraph(
                    new ParagraphProperties(
                        new RunProperties(
                            new Bold() { Val = false }, // Establecer explícitamente no-negrita
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }  // Fuente Arial
                        )
                    ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true }
                    ),
                    new Run(
                        new FieldCode(@"TOC \c ""Tabla"" \h \z \u") { Space = SpaceProcessingModeValues.Preserve }
                    ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Separate }
                    )
                );
                sdtContent.Append(Contenido);

                Paragraph ContenEnd = new Paragraph(
                    new Run(
                        new RunProperties(
                            new Bold() { Val = false }, // Establecer explícitamente no-negrita
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }  // Fuente Arial
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
        }


        /// <summary>
        /// Método para crear una tabla de contenido en la cual se hace referencia a los titulos de las imagenes (leyenda)
        /// </summary>
        /// <param name="ruta">Aquí va la ruta del documento de word</param>
        /// <param name="tituloIlustraciones"></param>
        /// <exception cref="ArgumentNullException">Aquí va el título que va a tener la tabla de contenido</exception>
        public static void TablaIlustraciones(string ruta, string tituloTabla)
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
                        new DocPartGallery() { Val = "Table of Figures" },
                        new DocPartUnique())
                );
                block.Append(sdtProperties);

                SdtContentBlock sdtContent = new SdtContentBlock();

                // Establecer el título de la "Tabla de Ilustraciones" en negrita y centrado
                Run tituloRun = new Run(new Text(tituloTabla));
                RunProperties tituloRunProperties = new RunProperties();
                tituloRunProperties.Append(new Bold());
                tituloRunProperties.RunFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
                tituloRun.PrependChild<RunProperties>(tituloRunProperties);

                ParagraphProperties tituloParaProperties = new ParagraphProperties();
                tituloParaProperties.Append(new Justification() { Val = JustificationValues.Center });
                Paragraph Titulo = new Paragraph(tituloRun);
                Titulo.PrependChild<ParagraphProperties>(tituloParaProperties);

                sdtContent.Append(Titulo);

                Paragraph Contenido = new Paragraph(
                    new ParagraphProperties(
                        new RunProperties(
                            new Bold() { Val = false },
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                        )
                    ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true }
                    ),
                    new Run(
                        new FieldCode(@"TOC \c ""Ilustración"" \h \z \u") { Space = SpaceProcessingModeValues.Preserve }
                    ),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.Separate }
                    )
                );
                sdtContent.Append(Contenido);

                Paragraph ContenEnd = new Paragraph(
                    new Run(
                        new RunProperties(
                            new Bold() { Val = false },
                            new NoProof(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
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
        }
        #endregion


        #region Métodos que usan la librería de microsoft interop word
        /// <summary>
        /// Método para agregar una referencia / cita a un parrafo en el caso que se quiera añadir una referencia que ya existe entonces no es
        /// necesario agregar los demas parametros siempre que se vaya a crear una referencia esta debe tener un id unico el cual es el 
        /// nombre se la cita
        /// </summary>
        /// <param name="ruta">Aquí va a la ruta del documento de word</param>
        /// <param name="texto">Aquí va a el texto del parrafo que se va a mostrar</param>
        /// <param name="tamanoFuente">Aquí se inserta el tamaño de la fuente del parrafo</param>
        /// <param name="estilo">Aquí se pasa un enum con los valores constantes de estilos</param>
        /// <param name="alineacion">Aquí se pasa un enum con los valores constantes de alineación</param>
        /// <param name="nombreAutor">Aquí va el nombre del autor al cual se esta citando</param>
        /// <param name="apellidoAutor">Aquí va el apellido del autor al cual se esta citando</param>
        /// <param name="año">Aquí se inserta el año de a referencia</param>
        /// <param name="tituloLibro">Aquí se inserta un valor opcional con el nombre del libro</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarParrafoConCita(string ruta, string texto, int tamanoFuente, EstiloParrafo estilo, AlineacionTexto alineacion, string nombreCita, string nombreAutor = "", string apellidoAutor = "", string año = "", string tituloLibro = "")
        {
            tamanoFuente *= 2;

            // Suponiendo que tienes un método ValidarRutaArchivo que verifica la ruta del archivo
            ValidarRutaArchivo(ruta);

            // Primero, abre el documento con OpenXML y agrega el párrafo
            using (var document = WordprocessingDocument.Open(ruta, true))
            {
                if (document == null)
                {
                    throw new ArgumentNullException(nameof(document), "El documento no puede ser nulo.");
                }

                var body = document.MainDocumentPart.Document.Body;

                var runProperties = new RunProperties(new FontSize { Val = tamanoFuente.ToString() }, new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }); // Establecer la fuente a Arial

                // Aplicar el estilo correspondiente
                switch (estilo)
                {
                    case EstiloParrafo.Normal:
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
                        break;
                }

                var run = new Run(runProperties, new Text(texto));
                var paragraph = new Paragraph(run);

                // Evitar salto de línea después del párrafo
                var spacing = new SpacingBetweenLines() { Before = "0", After = "0" };
                var paragraphProps = new ParagraphProperties(spacing);

                // Aplicar la alineación correspondiente
                switch (alineacion)
                {
                    case AlineacionTexto.Izquierda:
                        paragraphProps.Append(new Justification { Val = JustificationValues.Left });
                        break;
                    case AlineacionTexto.Derecha:
                        paragraphProps.Append(new Justification { Val = JustificationValues.Right });
                        break;
                    case AlineacionTexto.Centro:
                        paragraphProps.Append(new Justification { Val = JustificationValues.Center });
                        break;
                    case AlineacionTexto.Justificado:
                        paragraphProps.Append(new Justification { Val = JustificationValues.Both });
                        break;
                    default:
                        paragraphProps.Append(new Justification { Val = JustificationValues.Left });
                        break;
                }

                paragraph.PrependChild(paragraphProps);
                body.AppendChild(paragraph);
                document.MainDocumentPart.Document.Save();
            }

            // Luego, abre el documento con Interop y agrega la cita
            Word.Application wordApp = new Word.Application();
            Word.Document docInterop = wordApp.Documents.Open(ruta);

            // Configurar el formato de las citas y bibliografía a IEEE
            docInterop.Bibliography.BibliographyStyle = "IEEE";

            // Buscar el final del documento
            Word.Range citationRange = docInterop.Content;
            citationRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Verificar si la cita ya existe
            bool sourceExists = false;
            foreach (Word.Source source in docInterop.Bibliography.Sources)
            {
                if (source.Tag == nombreCita)
                {
                    sourceExists = true;
                    break;
                }
            }

            // Si la fuente bibliográfica no existe, agrégala
            if (!sourceExists)
            {
                string sourceXML = $@"<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">
                        <b:Tag>{nombreCita}</b:Tag>
                        <b:SourceType>Book</b:SourceType>
                        <b:Author>
                            <b:Author>
                                <b:NameList>
                                    <b:Person>
                                        <b:Last>{apellidoAutor}</b:Last>
                                        <b:First>{nombreAutor}</b:First>
                                    </b:Person>
                                </b:NameList>
                            </b:Author>
                        </b:Author>
                        <b:Title>{tituloLibro}</b:Title>
                        <b:Year>{año}</b:Year>
                    </b:Source>";
                docInterop.Bibliography.Sources.Add(sourceXML);
            }

            // Insertar la cita (ya sea que la fuente bibliográfica se haya agregado en este método o previamente)
            citationRange.Fields.Add(citationRange, Word.WdFieldType.wdFieldCitation, nombreCita, true);

            docInterop.Save();
            docInterop.Close();
            wordApp.Quit();
        }


        /// <summary>
        /// Método para hacer la bibliografía con el formato IEEE
        /// </summary>
        /// <param name="ruta">Aquí va a la ruta del documento</param>
        public static void InsertarBibliografia(string ruta)
        {
            ValidarRutaArchivo(ruta);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(ruta);

                // Buscar el final del documento para insertar la bibliografía
                Word.Range bibliographyRange = doc.Content;
                bibliographyRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // Cambiar el estilo de la bibliografía a IEEE
                Word.Bibliography bib = doc.Bibliography;
                bib.BibliographyStyle = "IEEE";
                Word.Field bibField = bibliographyRange.Fields.Add(bibliographyRange, Word.WdFieldType.wdFieldBibliography);

                // Define un estilo con fuente Arial y aplícalo a la bibliografía
                Word.Style bibStyle = doc.Styles.Add("BibliographyArialStyle", Word.WdStyleType.wdStyleTypeParagraph);
                bibStyle.Font.Name = "Arial";

                Word.Range bibFieldRange = bibField.Code;
                bibFieldRange.Next().set_Style(bibStyle);

                doc.Save();
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }


        /// <summary>
        /// Este método usa la libreria de microsoft office interop para hacer las listas con viñetas, de este modo la numeración
        /// no se de estropea porque openxml solo deja hacer una numeracion
        /// </summary>
        /// <param name="ruta">Aquí va a el directorio donde esta el documento de word</param>
        /// <param name="items">Aquí se recibe una lista con los elementos de las listas</param>
        /// <param name="tamanoFuente">Aquí se inserta un valor numerico con el tamaño de la fuente</param>
        /// <param name="estilo">Aquí se pasa un enum con el estilo del texto</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AgregarListado(string ruta, List<string> items, int tamanoFuente, EstiloParrafo estilo)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentNullException(nameof(items), "La lista de items no puede ser nula o vacía.");
            }

            // Inicia una nueva aplicación de Word y abre el documento.
            Word.Application wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Open(ruta);

            try
            {
                // Desplazarse al final del documento.
                wordApp.Selection.EndKey(Word.WdUnits.wdStory);

                // Establecer fuente a Arial.
                wordApp.Selection.Font.Name = "Arial";

                // Establecer estilo.
                switch (estilo)
                {
                    case EstiloParrafo.Negrita:
                        wordApp.Selection.Font.Bold = 1;
                        break;
                    case EstiloParrafo.Italico:
                        wordApp.Selection.Font.Italic = 1;
                        break;
                    case EstiloParrafo.Subrayado:
                        wordApp.Selection.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                        break;
                    default:
                        break;
                }

                // Establecer tamaño de fuente.
                wordApp.Selection.Font.Size = tamanoFuente;

                // Crear una lista con viñetas.
                wordApp.Selection.Range.ListFormat.ApplyBulletDefault();

                // Agregar los elementos de la lista.
                foreach (string item in items)
                {
                    wordApp.Selection.TypeText(item);
                    wordApp.Selection.TypeParagraph();
                }

                // Guardar y cerrar el documento.
                document.Save();
                document.Close();
            }
            finally
            {
                // Asegurarse de liberar los recursos y cerrar Word.
                if (document != null) Marshal.ReleaseComObject(document);
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }

            Console.WriteLine("Agregando listado con viñetas al documento.");
        }
        #endregion
    }
}