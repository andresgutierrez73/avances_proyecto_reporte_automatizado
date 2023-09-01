using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using funcionalidades_documento.crear_documento;
using funcionalidades_documento.funciones_parrafo;
using funcionalidades_documento.funciones_tablas;
using funcionalidades_documento.funciones_imagenes;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;

namespace funcionalidades_documento.componentes_reporte
{
    public class SeccionesCuerpoReporte
    {

        public static void Objeto(string ruta)
        {
			// Controlamos las excepciones
			try
			{
				// Definicion de los titulos y parrafos
				string titulo = "objeto";
				string parrafo1 = "Presentar los procedimientos, criterios y resultados de los análisis efectuados para el diseño estructural de los pórticos metálicos requeridos para el cambio rápido del nuevo reactor de repuesto de 12.5 Mvar que será instalado en la subestación Banadía 230 kV, ubicada en el municipio de Saravena, en el departamento de Arauca.";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void Alcance(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "alcance";
                string parrafo1 = "En los siguientes capítulos se detallarán los procedimientos, criterios y resultados de los análisis efectuados para el diseño de la estructura metálica de los pórticos. Se incluye una descripción de las cargas aplicadas producto del peso de los equipos, cables, y de las acciones ambientales que inciden directamente sobre las estructuras metálicas. Además, se presentan los resultados del análisis y diseño realizado usando el software SAP2000, para cada uno de los elementos que conforman las estructuras atendiendo las solicitaciones más desfavorables que exijan las distintas combinaciones de carga.";
                string parrafo2 = "Los diseños han sido realizados teniendo en cuenta todos los requerimientos de las especificaciones técnicas del proyecto [10] y [2]. Los resultados del diseño se ilustran en el plano “CO-RBAN-14113-S-01-K1525: Planos de diseño estructuras metálicas de pórticos”, en dicho plano se presenta la guía para la fabricación de las estructuras.";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void DescripcionPorticos(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "descripción de los pórticos";
                string parrafo1 = "Los pórticos se diseñan como estructuras en celosía con diagonales, estos elementos soportan en la parte superior las cargas de templas y equipos dependiendo de la configuración del sistema. Además, los pórticos se encargan de transmitir las solicitaciones a la fundación y posteriormente al suelo.";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void EspecificacionMateriales(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "especificaciones de los materiales";

                //Definimos la lista con los valores que van a ir en la tabla
                List<List<string>> datos = new List<List<string>> {
                    new List<string> { "ítem".ToUpper(), "descripción".ToUpper(), "criterio".ToUpper()  },
                    new List<string> { "elemento", "Perfiles", "ASTM A-572 Gr50 ó ASTM A-36"  },
                    new List<string> { "|", "Platinas", "ASTM A-36"  },
                    new List<string> { "|", "Soldadura", "E60, E70"  },
                    new List<string> { "|", "Tornillos", "ASTM A-394"  },
                    new List<string> { "|", "Pernos de anclaje", "ASTM F1554 Gr55. Resistencia mínima \r\n\r\nfy = 380 MPa y fu =517 MPa "  },
                    new List<string> { "|", "Arandelas", "ASTM F-436"  },
                    new List<string> { "|", "Tuercas", "ASTM A-563"  },
                    new List<string> { "|", "Galvanización", "ASTM A-123, ASTM A-153"  },
                    new List<string> { "|", "Columnas ", "Celosía "  },
                    new List<string> { "|", "Vigas", "Celosía "  },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos, "Especificación de los materiales", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CriteriosDiseno(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "criterios de diseño";
                string parrafo1 = "El diseño de la estructura metálica para los pórticos se lleva a cabo teniendo en cuenta los criterios de diseño de estructuras metálicas [10], documento en el que se referencian las especificaciones de los planos del fabricante de los equipos, la geometría básica, distancias eléctricas, y cargas de conexión.";
                string parrafo2 = "El análisis estructural se realizó en el software SAP2000, versión 24.0.0, mediante un modelo tridimensional, en el cual, la estructura está idealizada como un conjunto de celosías planas, con una configuración de diagonales tipo “X”.";
                string parrafo3 = "Para el estado límite de resistencia, el diseño de los elementos se realizó con la aplicación IEB “Diseño de Estructura Metálica de pórticos y Equipos”, la cual con base en la información de entrada (resultados del SAP 2000), realiza el diseño por compresión, tracción, flexión, la interacción entre estas solicitaciones y el diseño de las conexiones para la cantidad mínima requerida de pernos.";
                string parrafo4 = "La determinación de los esfuerzos máximos a compresión, tensión, flexión, cortante y aplastamiento se hace siguiendo los lineamientos de las normas AISC 360 – 16 (American Institute of Steel Construction), referencia [11], y ASCE 10-15 (American Society of Civil Engineers) “Design of Latticed Steel Transmission Structures” referencia [12] y siguiendo las recomendaciones del manual ASCE N°52 “Guide for Design of Steel Transmission Towers”, referencia [13]; con ayuda del programa SAP2000.";
                string parrafo5 = "Para la definición de los elementos metálicos los límites de las relaciones de esbeltez serán los presentados en la Tabla 2:";
                string parrafo6 = "La dimensión mínima de los perfiles que componen las estructuras debe responder a la Tabla 3.";

                // Definicion de las listas con lso datos de las tablas
                List<List<string>> datos1 = new List<List<string>> {
                    new List<string> { "ítem".ToUpper(), "descripción".ToUpper(), "criterio".ToUpper()  },
                    new List<string> { "Relación de esbeltez - ASCE 10-15", "Otros miembros", "L/r ≤ 200"  },
                    new List<string> { "|", "Redundantes", "L/r ≤ 250"  },
                    new List<string> { "|", "Solo a tensión", "L/r ≤ 350"  },
                    new List<string> { "|", "Miembros a compresión", "Montantes L/r ≤ 150"  },
                    new List<string> { "Relación w/t - ASCE 10-15", "Ángulos a 90° Numeral 3.7.1", "Máximo w/t ≤ 25"  },
                    new List<string> { "|", "Compacto", "w/t ≤ (w/t) lím "  },
                    new List<string> { "|", "Esbelto Ecuación 3.7-2", "(w/t) lím< w/t ≤144Ψ/Fy1/2"  },
                    new List<string> { "|", "Esbelto Ecuación 3.7-3", "w/t >144Ψ/Fy1/2 "  },
                };

                List<List<string>> datos2 = new List<List<string>> {
                    new List<string> { "ítem".ToUpper(), "descripción".ToUpper(), "criterio".ToUpper()  },
                    new List<string> { "Espesor mínimo - ASCE 10-15 ", "Miembros", "3/16\" (4.8mm)"  },
                    new List<string> { "|", "Miembros secundarios redundantes", "1/8\" (3.2mm)"  },
                    new List<string> { "|", "Platinas de conexión", "L3/16\" (4.8mm)"  },
                    new List<string> { "|", "Criterio de espesor exposición a corrosión", "3/16\" (4.8mm)"  },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo4, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo5, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, "Relaciones de esbeltez y ancho-espesor", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo6, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2, "Dimensiones mínimas de elementos estructurales", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CriteriosDeflecciones(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "criterios de deflexiones";
                string parrafo1 = "Las deformaciones de la estructura metálica, se limitarán a los valores presentados en la Tabla 4 (Los valores fueron tomados del capítulo 4 de la norma ASCE 113 “Substation Structure Design Guide” referencia [4])";
                string parrafo2 = "Los elementos de las estructuras de pórticos se deben clasificar como tipo A, cuando hay equipos sobre los pórticos y tipo B cuando no se tienen equipos sobre los pórticos.";
                string parrafo3 = "Los elementos de las estructuras de soporte de los seccionadores e interruptores se deben clasificar como tipo A y como tipo B las estructuras de soporte de los demás equipos.";
                string parrafo4 = "Notas:";
                string parrafo5 = "La luz para los miembros horizontales debe ser medida como la luz libre entre miembros verticales o para miembros en cantiléver como la distancia al punto vertical más cercano. Luego la deflexión debe ser el desplazamiento neto, vertical u horizontal, relativo al punto de soporte.";
                string parrafo6 = "La luz para miembros verticales debe ser la distancia vertical desde el punto de conexión de la fundación al punto de investigación.";

                // Definicion de la lista con los datos que van a tener las tablas
                List<List<string>> datos1 = new List<List<string>> {
                    new List<string> { "tipo de deflexión".ToUpper(), "estructuras de clase a".ToUpper(), "~", "estructuras de clase b".ToUpper(), "~" },
                    new List<string> { "|", "Elementos horizontales", "Elementos verticales", "Elementos horizontales ", "Elementos verticales" },
                    new List<string> { "Horizontal", "1/200 ", "1/100 ", "1/100 ", "1/100" },
                    new List<string> { "Vertical", "1/200 ", "", "1/200", "" },
                };
                List<List<string>> datos2 = new List<List<string>> {
                    new List<string> { "clase A", "Interruptores y seccionadores" },
                    new List<string> { "clase B", "Transformadores de corriente, transformadores de tensión, descargadores de sobretensión, aisladores poste y trampas de onda  \r\n\r\nColumnas de pórticos - Vigas de pórticos " },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, "Deformaciones permisibles", 2);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2, "Clasificación de miembros, según ASCE-113");
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo4, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo5, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo6, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void Cargas(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "cargas";
                string parrafo1 = "Para el diseño de la estructura se considera el peso propio entre las cargas actuantes. En las cargas de diseño presentadas en los planos no se incluyen factores de sobrecarga, por lo tanto, en el análisis de la estructura metálica realizado se incluyen estos factores. ";
                string parrafo2 = "Las cargas sobre los pórticos y las dimensiones generales son tomadas de los documentos de referencia del [20] al [22].";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void PesoPropioEstructura(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "peso propio de la estructura";
                string parrafo1 = "Cargas debidas al peso de la estructura metálica, cables, templas, aisladores, herrajes, accesorios, y todos los elementos que componen el conjunto analizado. Se afecta en un 20% adicional para considerar el peso de los elementos estructural no modelados tales como: platinas, pernos, tuercas, arandelas, galvanizado, etc.";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CargasConexion(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "cargas de conexión";
                string parrafo1 = "Se refiere a las tensiones mecánicas y cargas de cortocircuito. Considerando las tensiones mecánicas, esta es aplicable a barraje flexible en templas, barras, cable guardas, conexión entre equipos, etc. Metodología según Overhead Power Lines, referencia [15]. Flecha máxima para condición EDS del barraje del 3% ";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CargasViento(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "cargas de viento";
                string parrafo1 = "Se considera las cargas de vientos sobre templas, equipos y estructuras en dirección X y Y. La velocidad del viento se toma de la NSR-10 [2] y el cálculo de estas fuerzas se realiza bajo la metodología del manual ASCE-74 “Guidelines for Electrical Transmission Line Structural Loading”, referencia [3]. La fuerza del viento sobre la estructura debida a la presión del viento sobre los conductores se calcula como:";

                // Buscamos la ruta de la imágen
                string rutaSalidaImagen = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp_imagenes\\formulas_cargas_viento.jpg";

                // Creamos la lista con los datos que van en la tabla
                List<List<string>> datos1 = new List<List<string>> {
                    new List<string> { "F", "Fuerza debida al viento " },
                    new List<string> { "γW", "Factor de carga, función del periodo de retorno, Tabla 1-1 ó 1-2 de referencia [3]" },
                    new List<string> { "V50", "Velocidad del viento, para un periodo de retorno de 50 años." },
                    new List<string> { "A", "Área frontal efectiva de la estructura, en m2." },
                    new List<string> { "KZ", "Coeficiente de exposición, Tabla 2-2 de referencia [3]" },
                    new List<string> { "KZT", "Factor de Topografía, Ec. 2-14 de referencia [3]" },
                    new List<string> { "Q", "Constante numérica, en función de la densidad del aire, sección 2.1.2 de referencia [3]" },
                    new List<string> { "G", "Factor de ráfaga. Sección 2.1.5 de referencia [3] " },
                    new List<string> { "CF", "Coeficiente de fuerza, sección 2.1.6 de referencia [3] " },
                    new List<string> { "qZ ", "Presión del viento" },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                //PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 10, 2, FuncionesCreacion.AlineacionImagen.Centro, "hoos");
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, "Donde:", 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, 0, true);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);


            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CargasSismo(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "cargas de sismo";
                string parrafo1 = "El cálculo de estas fuerzas se realiza bajo la metodología del Reglamento Colombiano de Construcción Sismo Resistente NSR-10, referencia [2]. El sismo vertical se define como Ez = 2/3 E(x,y). Los parámetros sísmicos se indican en la siguiente referencia [9].";
                string parrafo2 = "Ez: \tSismo vertical";
                string parrafo3 = "Ex,y: Sismo horizontal";
                string parrafo4 = "Nota: para el análisis estructural se utilizó coeficiente de capacidad de disipación de energía R de 3.00 y un factor de sobre-resistencia Ω de 3.00.";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo4, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CargasMontajeMantenimiento(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "cargas de montaje y mantenimiento";
                string parrafo1 = "Todos los miembros de las estructuras en análisis cuyo eje longitudinal forme un ángulo con la horizontal menor que 45 grados tendrán suficiente sección para resistir una carga adicional de 150 daN vertical, aplicada en cualquier punto de su eje longitudinal.";
                string parrafo2 = "Considerando las cargas de montaje y mantenimiento para columnas: el castillete será diseñado para resistir la acción de un hombre con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 150 daN.";
                string parrafo3 = "Considerando las cargas de montaje y mantenimiento para vigas: el nodo donde llega cada barraje, será diseñado para resistir la acción de dos hombres con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 250 daN. ";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void CombinacionesCarga(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "combinaciones de carga";
                string parrafo1 = "Para el diseño de la estructura metálica se utilizan las combinaciones de carga que se listan de la Tabla 6 a la Tabla 8; estas combinaciones de carga provienen del documento CO-RBAN-14113-S-01-D1181 “Criterios de diseño - Estructuras metálicas” [10].";
                string parrafo2 = "Nota: La fuerza de peso propio y Ez se toman positivas en el sentido de la gravedad.";
                string parrafo3 = "Nota: La fuerza de peso propio y Ez se toman positivas en el sentido de la gravedad.";

                // Definimos una lista que se va a mostrar en el docuemento
                List<string> datos = new List<string>
                {
                    "PP: Peso propio de la estructura (Ps), equipos (Pe) y conductores de la conexión (Pc).",
                    "CT: Cargas por tensión mecánica de los conductores de conexión y cables guarda, se debe considerar tiro unilateral (un solo sentido, caso más desfavorable).",
                    "CMM: Carga de montaje y mantenimiento",
                    "VD: Carga viento de diseño sobre equipos, cables y estructuras. Ver velocidad de viento de diseño en referencia [2]. Está conformado por el Viento sobre la estructura (VSx,y), sobre equipos (VEx,y) y sobre conductores de la conexión (VCx,y).",
                    "VS: Carga viento de servicio sobre equipos, cables y estructuras. Ver velocidad de viento de servicio en referencia [2].",
                    "CTVDL o CTVSL: Carga de sobretensión en el cable debido al viento de diseño o viento de servicio (solo actúa en el sentido de la tensión del cable). Correspondiente a (VCTx,y)",
                    "CC: Cargas sobre conductores por efecto de cortocircuito.",
                    "EX,Y: Cargas por sismo horizontal sobre equipos y estructuras, obtenidos con coeficientes sísmicos últimos. Para verificación de deflexiones, la carga sísmica se debe emplear sin dividir por R.",
                    "EZ: Cargas por sismo vertical sobre equipos y estructuras, obtenidos con coeficientes sísmicos últimos.",
                    "AC: Carga de accionamiento que aplica solo para interruptores.",
                };

                //Definimos las listas que contienen los datos que van en las tablas

                //Definimos una variable contador para la numeracion
                int con = 1;

                List<List<string>> datos1 = new List<List<string>> {
                    new List<string> { "combinaciones de carga".ToUpper(), "~" },
                    new List<string> { "Diseño Estructural", $"{con++}) 1,2 PP + 1,3 CT+ 1,0 CMM " },
                    new List<string> { "|", $"{con++}) 1,1 PP + 1,1 CT ± 1,0 VD(X,Y) + 1,0 CTVDL" },
                    new List<string> { "|", $"{con++}) 1,1 PP + 1,1 CT ± 1,0 E(X,Y) ± 0,3 E(Y,X) + 1,0 E(Z)" },
                    new List<string> { "|", $"{con++}) 0,9 PP + 1,1 CT ± 1,0 E(X,Y) ± 0,3 E(Y,X) - 1,0 E(Z)" },
                    new List<string> { "|", $"{con++}) 1,1 PP + 1,1 CT + 1,0 CC + 1,0 AC" },
                };

                con = 1;
                List<List<string>> datos2 = new List<List<string>> {
                    new List<string> { "combinaciones de carga".ToUpper(), "~" },
                    new List<string> { "Diseño Estructural", $"{con++}) 1,0 PP + 1,0 CT + 1,0 CMM " },
                    new List<string> { "|", $"{con++}) 1,0 PP + 1,0 CT ± 1,0 VS(X,Y) + 1,0 CTVSL" },
                    new List<string> { "|", $"{con++}) 1,0 PP + 1,0 CT ± 0,7 E(X,Y) ± 0,21 E(Y,X) + 0,7 E(Z) " },
                    new List<string> { "|", $"{con++}) 0,6 PP + 1,0 CT ± 0,7 E(X,Y) ± 0,21 E(Y,X) - 0,7 E(Z) " },
                    new List<string> { "|", $"{con++}) 1,0 PP + 1,0 CT + 1,0 CC + 1,0 AC " },
                };

                con = 1;
                List<List<string>> datos3 = new List<List<string>> {
                    new List<string> { "Condición", "Combinación" },
                    new List<string> { "Viento máximo esperado", $"{con++}) 1.0PP + 0.78 V + 1.0 CT " },
                    new List<string> { "Accionamiento de equipos", $"{con++}) 1.0PP + AC + 1.0 CT " },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, "Combinaciones de carga - Últimas", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2, "Combinaciones de carga – Servicio", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos3, "Combinaciones de carga – Deflexiones", 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, "Donde,", 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarListado(ruta, datos, 12, FuncionesCreacion.EstiloParrafo.Normal);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        public static void NomenclaturaReporte(string ruta)
        {
            // Controlamos las excepciones
            try
            {
                // Definicion de los titulos y parrafos
                string titulo = "Nomenclatura del reporte";
                string parrafo1 = "A continuación, se indica la nomenclatura del reporte del diseño de ángulos del soporte crítico que será presentado posteriormente.";

                // Buscamos la ruta de la imágen
                string rutaSalidaImagen = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\temp_imagenes\\nomenclaturaReporte.jpeg";

                // Base 64 de la imagen
                string base64Imagen = "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Ao++2O3aitz7Zcoi6Nxa9j2nbJDI05ZJG4cWva4Ahw4ghSCIK1ZL3caG4jS2qXMNaWl1DXNGI7hGBx4fImbjymciMOacFzWWVR99sVu1FbpLZc4S6NxDmPY4skhkBy2SN44se04IcCCCFFWW93GhuLdLaqcDWlpdRVwaGxXCMc+A4Mmb8pnIghzcjcGBZUREBQWsOzIu3fTt7H9N5rvO+Z3+3ap1QWsOzIu3fTt7H9N5rvO+Z3+3agnUREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAVZvV7uFyuEmlNKSFlY1gNdcNgdHbmO5AZ4PnIyWswQANz8AtD/l7vdwudfNpPSk5jrWMHh1wEYey3McOGNwLXzkcWsOQ0Yc8bS0PmbLZbfYKBlutsTmRNJc5z3l8kjz5z3vdkvcTxLickoPlkslu09b47bbInNiZlznSSOkkleeLnyPcS573HiXOJJPNd6IgIiICIiCJ1d8VL19XVP2bl6aZ+Llq/MYPs2rz1d8VL19XVP2bl6aZ+Llq/MYPs2oJJERAREQFH3yx27UNvfbblE50biHMfG8skieOLZI3tw5j2niHAggqQRBWbJe7hQXBmldVPzWuYXUVftDYrixvPlwZM0YLmYAIO5mQHBlmXBerJb9QUD7dconOjcQ5j43mOSJ4817Htw5jgeIcDkKIsd7uFur4tKarlL65zCaKv2BsdxY3ny8lswHFzMAHi5g2hwYFmUFrDsyLt307ex/Tea7zvmd/t2qSku9pic5ktzpGOYSHB0zQWkc88eCjdYdmRduenb2P6bzXed8zv9u1BOoiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiICIiAiIgIiIIbU11q7U21mkLB4Xc6elk3Nz9zeTnHt4L7d7pVUV5slDAW9VXzyxzZGThsTnDHdxAXlrG03C7WmI2gxGvoKynr6dkri1kropA50RcM7d7A5m7B2lwdg4weGGPUF/v1sr7hYXWmltRllImqI5ZJpHsLA1ojJAaAXEknuAHEkBZaajpaJr2UlPHCJJHSvDG43Pcclx7ye9eyIgIiICIiAiIgib9eKS1upKa50pdQ3GTwOWc4McT34axrx3PJ2g8skD1r5b71S1N3qbFbaN7oLZG1k9Q3AijlONsDfWXBvE44Ny0czgd9fQUV0op7bcaaOopamN0U0Ugy17CMEEKu9FlF4B0c6bjfHI2eW1009SZSTI+d8TXSPeTxLy4kknjlBakREBERAREQF41NHS1gjFVTxyiKRsse9udrxycO4jvXsiCgal01pzVt9OlIdN22SEObV3yrNJGTscdzafOMl8p4u7mZzxe3Ni1eALXEAL4AJ29j+m813P5nf7dq4KPQlbbpayWg19qGHw6rlrJGiKgf5b3ZwC6mLtrRhrckkNaBkrv1f2ZFxvp+7t7H9N5rvO+Z3+3agnUREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAVN6VK+ut2nqea31s9NI6sY0vhkLHFux5xkergP1IiD/9k=";

                // Definimos las lista con los valores que van a ir en la tabla
                List<List<string>> datos1 = new List<List<string>> {
                    new List<string> { "nomenclatura del reporte".ToUpper(), "~" },
                    new List<string> { "L:", $"Longitud no arriostrada del elemento" },
                    new List<string> { "Rx,y:", $"Radio de giro del elemento respecto a los ejes geométricos X y Y" },
                    new List<string> { "Ru:", $"Radio de giro del elemento respecto al eje principal menor U" },
                    new List<string> { "L/R:", $"Relación de esbeltez" },
                    new List<string> { "kL/R:", $"Longitud efectiva" },
                    new List<string> { "Curva:", $"Ecuación empleada para la estimación de la longitud efectiva" },
                    new List<string> { "(L/R)LIM:", $"Máxima relación de esbeltez" },
                    new List<string> { "Cc:", $"Relación de esbeltez critica" },
                    new List<string> { "Fa:", $"Esfuerzo de compresión" },
                    new List<string> { "Comb:", $"Combinación de carga empleada en el diseño" },
                    new List<string> { "P", $"Fuerza axial de tracción o compresión*" },
                    new List<string> { "Puc:", $"Fuerza axial de compresión" },
                    new List<string> { "Put:", $"Fuerza axial de tracción" },
                    new List<string> { "V2:", $"Fuerza cortante en el plano 1-2" },
                    new List<string> { "V3:", $"Fuerza cortante en el plano 1-3" },
                    new List<string> { "M2:", $"Momento flector en el plano 1-3 (alrededor del eje 2)" },
                    new List<string> { "M3:", $"Momento flector en el plano 1-2 (alrededor del eje 3)" },
                    new List<string> { "Mr:", $"Momento actuante resultante, debido a M2 y M3" },
                    new List<string> { "θ°:", $"Angulo del momento resultante con respecto a la horizontal" },
                    new List<string> { "Uso:", $"Relación de uso total del elemento (interacción de todas las solicitaciones)" },
                    new List<string> { "Ecu:", $"Ecuación empleada para estimar el uso" },
                    new List<string> { "Puc/Pac:", $"Relación de uso del elemento en compresión" },
                    new List<string> { "Put/Pat-v:", $"Relación de uso del elemento en tracción" },
                    new List<string> { "Mr/Ma:", $"Relación de uso del elemento en flexión" },
                    new List<string> { "Pac:", $"Capacidad a compresión del elemento" },
                    new List<string> { "Pat-g:", $"Capacidad a tracción en el área bruta del elemento" },
                    new List<string> { "Pat-v:", $"Capacidad a tracción en el área neta del elemento o por bloque de cortante" },
                    new List<string> { "Pat:", $"Capacidad a tracción del elemento" },
                    new List<string> { "Ma:", $"Capacidad a flexión del elemento" },
                    new List<string> { "Pe:", $"Carga critica de pandeo de Euler" },
                    new List<string> { "ØPyc:", $"Resistencia axial del elemento en el área bruta" },
                    new List<string> { "Myt:", $"Momento que produce esfuerzos de tracción en la fibra extrema" },
                    new List<string> { "Myc:", $"Momento que produce compresión en la fibra extrema" },
                    new List<string> { "Me:", $"Momento crítico elástico" },
                    new List<string> { "Me.Ecu:", $"Ecuación empleada para el cálculo de Me," },
                    new List<string> { "Mb:", $"Momento que produce pandeo lateral" },
                    new List<string> { "Mb.Ecu:", $"Ecuación empleada para el cálculo de Mb," },
                    new List<string> { "K:", $"Factor que depende de la condición de apoyo del elemento" },
                    new List<string> { "Cm:", $"Factor que depende de la distribución de momento en la sección" },
                    new List<string> { "Ø:", $"Diámetro del perno en pulgadas" },
                    new List<string> { "emin:", $"Distancia mínima al borde del elemento cortado" },
                    new List<string> { "fmin:", $"Distancia mínima al borde del elemento" },
                    new List<string> { "smin:", $"Distancia mínima entre centros de perforaciones" },
                };

                List<List<string>> datos2 = new List<List<string>> {
                    new List<string> { "z", "~", "~", "~", "z", "~", "~", "~" },
                    new List<string> { "nombres", "~", "firma", "matricula", "total paginas", "1006", "fecha emision", "2022.09.16" },
                    new List<string> { "elaboro", "c.castaño", "firma", "267773 ANT", "nombre proyecto", "~", "~", "~" },
                    new List<string> { "reviso", "C.METRIO", "firma", "357197 ANT", "RENOVACIÓN SUBESTACIÓN BANADÍA 230 kV", "~", "~", "~" },
                    new List<string> { "|", "|", "|", "|", "Código del Documento", "~", "~", "~" },
                    new List<string> { "Aprobó", "I. VILLALBA", "firma", "196375 ANT", "CO-RBAN-14113-S-01-D1531", "~", "~", "~" },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, 1, true);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                //PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 11, 5, FuncionesCreacion.AlineacionImagen.Centro, "dmff");
                //PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 11, 5, FuncionesCreacion.AlineacionImagen.Centro, "dmfeeef");

                for (int i = 0; i < 5; i++)
                {
                    PropiedadesImagen.AgregarImagenDesdeBase64(ruta, base64Imagen, 5, 5, FuncionesCreacion.AlineacionImagen.Centro, $"hola {i}");
                    PropiedadesImagen.AgregarImagenDesdeBase64(ruta, base64Imagen, 5, 5, FuncionesCreacion.AlineacionImagen.Centro, $"hola {i++}");
                }
                

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }
    }
}
