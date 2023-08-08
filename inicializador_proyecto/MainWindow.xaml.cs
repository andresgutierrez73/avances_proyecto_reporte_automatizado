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
            string rutaImagen = "";
            string rutaSalidaImagen = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string textoAleatorio = "";

            try
            {
                // Generar el documento de Word en la ruta especificada
                FuncionesCreacion.GenerarDocumentoWord(ruta);

                EditarEncabezadoPie.EditarEncabezado(ruta, "Este es el encabezado", "Esta es la parte baja del texto", 2);
                EditarEncabezadoPie.EditarPieDePagina(ruta, "Diseño de estructura metalmecánica");

                #region Aquí se va a crear la portada APA
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 5);
                PropiedadesParrafo.AgregarParrafo(ruta, "Este es el ejemplo del titulo de la portada", 12, EstiloParrafo.Negrita, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, "Andrés Juan Gutiérrez Castro", 12, EstiloParrafo.Normal, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarParrafo(ruta, "Ingeniería Especializada (IEB)", 12, EstiloParrafo.Normal, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarParrafo(ruta, "Área de desarrollo de proyectos de ingenieria", 12, EstiloParrafo.Normal, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarParrafo(ruta, "Jhefferson Rios", 12, EstiloParrafo.Normal, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarParrafo(ruta, "2 de agosto de 2023", 12, EstiloParrafo.Normal, AlineacionTexto.Centro);
                PropiedadesParrafo.AgregarSaltoDePagina(ruta);
                #endregion

                #region Aquí se va a llamar a la tabla de contenido
                PropiedadesParrafo.TablaContenido(ruta, "Tabla de contenido IEB");
                PropiedadesParrafo.AgregarSaltoDePagina(ruta);
                #endregion

                #region Aquí va esta el contenido del cuerpo del documento
                PropiedadesParrafo.AgregarTitulo(ruta, "Noticias IEB", 1, 12, EstiloParrafo.Negrita, AlineacionTexto.Izquierda);
                PropiedadesParrafo.AgregarTitulo(ruta, "IEB presente en la cita con el ministro de minas y energía en Medellín", 2, 12, EstiloParrafo.Negrita, AlineacionTexto.Izquierda);
                textoAleatorio = Lorem.Paragraph(4, 20);
                PropiedadesParrafo.AgregarParrafo(ruta, textoAleatorio, 12, EstiloParrafo.Normal, AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                rutaImagen = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\documento_aleatorio\\1.jpg";
                PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaImagen, 10, 10, AlineacionImagen.Centro);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                textoAleatorio = Lorem.Paragraph(4, 8);

                PropiedadesParrafo.AgregarParrafo(ruta, textoAleatorio, 12, EstiloParrafo.Normal, AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltoDePagina(ruta);

                PropiedadesParrafo.AgregarParrafoConCita(ruta, "este es un texto con referencia", 12, EstiloParrafo.Italico, AlineacionTexto.Izquierda, "Andres Juan", "gutierrez", "2022");
                #endregion
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