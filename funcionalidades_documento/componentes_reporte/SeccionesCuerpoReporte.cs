using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using funcionalidades_documento.crear_documento;
using funcionalidades_documento.funciones_parrafo;
using funcionalidades_documento.funciones_tablas;
using funcionalidades_documento.funciones_imagenes;

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
                    new List<string> { "elementos", "Perfiles", "ASTM A-572 Gr50 ó ASTM A-36"  },
                    new List<string> { "elementos", "Platinas", "ASTM A-36"  },
                    new List<string> { "elementos", "Soldadura", "E60, E70"  },
                    new List<string> { "elementos", "Tornillos", "ASTM A-394"  },
                    new List<string> { "elementos", "Pernos de anclaje", "ASTM F1554 Gr55. Resistencia mínima \r\n\r\nfy = 380 MPa y fu =517 MPa "  },
                    new List<string> { "elementos", "Arandelas", "ASTM F-436"  },
                    new List<string> { "elementos", "Tuercas", "ASTM A-563"  },
                    new List<string> { "elementos", "Galvanización", "ASTM A-123, ASTM A-153"  },
                    new List<string> { "elementos", "Columnas ", "Celosía "  },
                    new List<string> { "elementos", "Vigas", "Celosía "  },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos);
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
                    new List<string> { "Relación de esbeltez - ASCE 10-15", "Redundantes", "L/r ≤ 250"  },
                    new List<string> { "Relación de esbeltez - ASCE 10-15", "Solo a tensión", "L/r ≤ 350"  },
                    new List<string> { "Relación de esbeltez - ASCE 10-15", "Miembros a compresión", "Montantes L/r ≤ 150"  },
                    new List<string> { "Relación w/t - ASCE 10-15", "Ángulos a 90° Numeral 3.7.1", "Máximo w/t ≤ 25"  },
                    new List<string> { "Relación w/t - ASCE 10-15", "Compacto", "w/t ≤ (w/t) lím "  },
                    new List<string> { "Relación w/t - ASCE 10-15", "Esbelto Ecuación 3.7-2", "(w/t) lím< w/t ≤144Ψ/Fy1/2"  },
                    new List<string> { "Relación w/t - ASCE 10-15", "Esbelto Ecuación 3.7-3", "w/t >144Ψ/Fy1/2 "  },
                };

                List<List<string>> datos2 = new List<List<string>> {
                    new List<string> { "ítem".ToUpper(), "descripción".ToUpper(), "criterio".ToUpper()  },
                    new List<string> { "Espesor mínimo - ASCE 10-15 ", "Miembros", "3/16\" (4.8mm)"  },
                    new List<string> { "Espesor mínimo - ASCE 10-15 ", "Miembros secundarios redundantes", "1/8\" (3.2mm)"  },
                    new List<string> { "Espesor mínimo - ASCE 10-15 ", "Platinas de conexión", "L3/16\" (4.8mm)"  },
                    new List<string> { "Espesor mínimo - ASCE 10-15 ", "Criterio de espesor exposición a corrosión", "3/16\" (4.8mm)"  },
                };

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo4, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo5, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo6, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2);
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
                    new List<string> { "tipo de deflexión".ToUpper(), "estructuras de clase a".ToUpper(), "estructuras de clase a".ToUpper(), "estructuras de clase b".ToUpper(), "estructuras de clase b".ToUpper() },
                    new List<string> { "tipo de deflexión".ToUpper(), "Elementos horizontales", "Elementos verticales", "Elementos horizontales ", "Elementos verticales" },
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
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo4, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo5, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
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
                PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 10, 2, FuncionesCreacion.AlineacionImagen.Centro);
                PropiedadesParrafo.AgregarParrafo(ruta, "Donde:", 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1);
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
                string titulo = "cargas de sismo";
                string parrafo1 = "Todos los miembros de las estructuras en análisis cuyo eje longitudinal forme un ángulo con la horizontal menor que 45 grados tendrán suficiente sección para resistir una carga adicional de 150 daN vertical, aplicada en cualquier punto de su eje longitudinal.";
                string parrafo2 = "Considerando las cargas de montaje y mantenimiento para columnas: el castillete será diseñado para resistir la acción de un hombre con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 150 daN.";
                string parrafo3 = "Considerando las cargas de montaje y mantenimiento para vigas: el nodo donde llega cada barraje, será diseñado para resistir la acción de dos hombres con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 250 daN. ";

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 2, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
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

                // Llamado a los métodos para editar el documento con la información
                PropiedadesParrafo.AgregarTitulo(ruta, titulo.ToUpper(), 1, 12, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo1, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarListado(ruta, datos, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }
    }
}
