using funcionalidades_documento.crear_documento;
using funcionalidades_documento.funciones_imagenes;
using funcionalidades_documento.funciones_parrafo;
using funcionalidades_documento.funciones_tablas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace funcionalidades_documento.componentes_reporte
{
    public class CreacionReporteAutomatizado
    {
        // Esta es la propiedad que se va a reutilizar en toda la clase con la ubicación del documento
        public string Ruta { get; set; }

        // Este es el constructor de la clase que asigna la ruta apenas es instanciado
        public CreacionReporteAutomatizado(string ruta)
        {
            this.Ruta = ruta;
        }

        // Este es el método que crea y pasa el archivo de word al inicializador del proyecto
        public void GeneradorDocumento()
        {
            // Controlamos las excepciones del programa
            try
            {
                // Generamos ele documento de word llamando al método
                FuncionesCreacion.GenerarDocumentoWord(Ruta);

                CreacionPortada(Ruta);
                CreacionCuerpoInforme(Ruta);
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        // Este es el método encargado de generar la portada del documento
        public void CreacionPortada(string ruta)
        {
            // Controlamos las excepciones del programa
            try
            {
                // Agregar los textos de los la portada
                string tituloPortada = "renovación subestación";
                string tituloPortada2 = "ingeniería del detalle para el montaje de un reactor de repuesto 12,5 Mvar, en la subestación banadía 230kV";
                string tituloPortada3 = "memoria del diseño de estructuras metálicas de pórticos";

                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 2);

                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada2.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada3.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);

                // Creamos la lista que va a contener los datos de la tabla que va en el cuerpo de la portada
                List<List<string>> datos = new List<List<string>> {
                    new List<string> { "", "", "", "", "", "", ""  },
                    new List<string> { "", "", "", "", "", "", ""  },
                    new List<string> { "", "", "", "", "", "", ""  },
                    new List<string> { "PA", "1", "Emisión Inicial", "2022.09.16", "C.CASTAÑO", "C,METRIO", "I.VILLALBA"  },
                    new List<string> { "Estado/fase", "Rev", "Comentarios/Modificaciones", "Fecha de Act", "Elaboró", "Revisó", "Aprobó"  },
                };

                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos);

                // Insertamos el salto de pagina
                PropiedadesParrafo.AgregarSaltoDePagina(ruta);

                //Insertamos la Tabla de Contenido
                PropiedadesParrafo.TablaContenido(ruta, "TABLA DE CONTENDIO");
                PropiedadesParrafo.AgregarSaltoDePagina(ruta);
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }

        // Este es el método encargado de hacer el cuerpo del informe
        public void CreacionCuerpoInforme(string ruta)
        {
            // Controlamos las excepciones del programa
            try
            {
                SeccionesCuerpoReporte.Objeto(ruta);
                SeccionesCuerpoReporte.Alcance(ruta);
                SeccionesCuerpoReporte.DescripcionPorticos(ruta);
                SeccionesCuerpoReporte.EspecificacionMateriales(ruta);
                SeccionesCuerpoReporte.CriteriosDiseno(ruta);
                SeccionesCuerpoReporte.CriteriosDeflecciones(ruta);
                SeccionesCuerpoReporte.Cargas(ruta);
                SeccionesCuerpoReporte.PesoPropioEstructura(ruta);
                SeccionesCuerpoReporte.CargasConexion(ruta);
                SeccionesCuerpoReporte.CargasViento(ruta);
                SeccionesCuerpoReporte.CargasSismo(ruta);
                SeccionesCuerpoReporte.CargasMontajeMantenimiento(ruta);
                SeccionesCuerpoReporte.CombinacionesCarga(ruta);
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }
    }
}
