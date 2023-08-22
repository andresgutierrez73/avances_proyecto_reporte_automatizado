using System;
using System.Windows;
using funcionalidades_documento.componentes_reporte;
using funcionalidades_documento.crear_documento;

namespace inicializador_proyecto
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            // Obtener la ruta del archivo de Word
            string ruta = FuncionesCreacion.GuardarRuta();

            try
            {
                // Creamos la instancia de la clase que se encarga de crear el documento de word
                CreacionReporteAutomatizado nuevoDocumento = new CreacionReporteAutomatizado(ruta);
                nuevoDocumento.GeneradorDocumento();
            }
            catch (Exception ex)
            {
                // Mostrar mensaje de error en caso de excepción
                Console.WriteLine("Error al crear el documento de Word: " + ex.Message);
            }

            InitializeComponent();
        }
    }
}