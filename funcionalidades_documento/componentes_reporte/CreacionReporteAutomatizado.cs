using funcionalidades_documento.crear_documento;
using funcionalidades_documento.funciones_imagenes;
using funcionalidades_documento.funciones_parrafo;
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
                string tituloPortada3 = "memoria del diseño de estrcuturas metálicas de pórticos";

                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 2);

                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada2.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);

                PropiedadesParrafo.AgregarParrafo(ruta, tituloPortada3.ToUpper(), 20, FuncionesCreacion.EstiloParrafo.Negrita, FuncionesCreacion.AlineacionTexto.Centro);
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }
    }
}
