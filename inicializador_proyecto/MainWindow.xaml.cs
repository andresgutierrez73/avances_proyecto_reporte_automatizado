using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using funcionalidades_documento.componentes_reporte;
using funcionalidades_documento.crear_documento;
using funcionalidades_documento.edicion_footer_header;
using funcionalidades_documento.funciones_imagenes;
using funcionalidades_documento.funciones_parrafo;
using funcionalidades_documento.funciones_tablas;
using LoremNET;
using static funcionalidades_documento.crear_documento.FuncionesCreacion;

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