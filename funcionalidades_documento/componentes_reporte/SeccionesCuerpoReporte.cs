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
        public static void TablaEjemplo(string ruta)
        {
            // base 64 de las imagenes
            string iebLogo = "iVBORw0KGgoAAAANSUhEUgAAAGcAAABrCAYAAABqg5yCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAA3SSURBVHhe7ZwJdFRFFoYbBxzUWdwR2SGyJ+wKDiKiGDdAWQZUQJCBAzijg8igg4woogJyXGBUEHcQHccNCMeFTRZBFFCGLelOJ530kq2zdafTnXT3P/dWv0hCHsl7naYp5f3n3ENO+lWlqr5Xt27dqsYUDof3GyanmWBIWhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCRWTOGEFTMUG8UETvnen+Be+iZyJswTVrB4NXx7flQ+PXMKHLeg9LNN8KzfQra5tn0e+de38weEyvxKKXnUIDiVOQXImbEQ5iY9kGZqD7Opi7BUUweYf5OInGkLUJnrVp6Ov/L/tQxHTL+j9iSINqnZcVNbansn2Ic9gAp7rlJSDkUNJ1hYDPvw6dTBlgSkOyymnjWMf5dqagX70GkIFpUopeIr91MraPBbUHt61WpfTUui51oh75FlSkk5FBWccDBIb+Vy6lBL0TH1DrNxp1sjfw51OhRSSsdP7oVa4fDL1AVZiaMRrqhUSp95RQWnwuZE+oUDqEPdVDta3dJMnWG9bDAC/zMrpeMnvXBs7W5D2F+hlD7zigqO5+OvaNA7Uqd61Oqkmh03tUHphylK6fhJN5wOt//y4RQteUuZNdrgpNGiW/jSu0rp+OnshLP8fRrwLtQprXBao/i1D5TS8dNZCce3fR91phN1SgscDrM7wrvxG6W0PoUrgwjmFaIiPRt+Wrf8Px4XFjhsRkWGHaFij/JkbZ2VcIKFpbD1HEODfpVqR6tbmqkdbP3GIZijfb/jP3AcJav+i/x/LoPjzr8ho+twWC7oR6F5Ag12a1Enh+rWZkNgv/5+5M1ZjNL3NtC+q+bfOCvhsEo++IIGqRt1qjN1Tq3zvejzzkg7txtK121SSqkrXFYO/yEz3ItWwT5iKrL6jEFW0mhk3zQZrsnzaPCfR8H85cJy/74ErrH/QFbPcbCc31cAO2a6jP5tSy/MCPFMML9I1Ote9MrZCQfhMAqfo8DgvB40MG2oc4nVOppIYGjn3bQH3M+uVgrUVsjrQ9m2fch/fAWcY+fC/czrKNu6J+KuikppzxFUnqwp4epKvDQb8+Fdvw2u8Y/B2mYIQWpGMNois9sIlO89BPfi16ht+uBUV6W7GIG0TJrJx1D+3SGU7zuMwBELKl35cckhRg9HkWfDNjjumglLU3Y7bYVZmvaFY/h0eD7bqjxVW4Ej6ShZk4LC5etE5xva2YAliyAvR/qlfyJAzWG9dBBsvUeJQdcKJ6vTCFRkOlD06ofImbkA2YMmIaNVMtLPu4Y+p/X1nF5U72Bk970brkmPonDJG/B9f1hpQezVYDgsdgXl1Ejvtj305n8rfg6VB5RPa6vCmQffD0dpIJzKb2Kn8oPHkZ38FzGbOWdmMfWuBULdeiD9/P7I6DyMynF+kNe2jgSlKxl7Bc6EJNHPEVfO9aea2sN6xfVw3D4Dnk07EArE1iXGBI4ehckdhisrT6tbCHrK4JrwGA0uByzawv2IJSlAGUR9s40/70Gguoi/Y27cHY4RDwgvECtphsM5J9/WfeTjt6Ns4456jZ/z7TyA0BlaYEO+cjhHP0xvdwcxiOoDHDvj2cXBieXSa1C44n1ekhsszXCChSWwdRtFEVIfpF8woH4jF2G5uD+yrh0H3/7T55frEi/otl4jhXtSG9DYmxKh0uzLnbGwTteuRdrhFBQjo8VQ+sMJit/VYl1pcb4S9lumIVTqVWrSJx7gIEVu0cq7aSfM5/DAnYgmT6+xu+PjkgTkTJyHELnYaKUDThEy2iRH/KuIgNQaVtvEWc95V6P8W+0no8EiD9zPrYZ95Ayaefcie+B4CrVnw5OyQ3lCu9i7OO+aRe2of8McW+PtRALy570caUgU0r7m+AMUNm9H4YtrRHaAZ4V6o042jnC6ouzL3UpNdcufmiFcIb8EkYgrQRhHT+YLeqPopXW0oOhz6F5qd1ojrcEBp5v4b7cm460BR2WRyE3v2iW8R+NEsd2IRprhVJf/4DHay3CIqsVVMJzuKPvqW6X0qcUbz6zrJpIrbC3KnRwxcT7PfFE/+PYdUkpoU4XNBVuXO6k8ZzOqt62m8WCmUVuzr52AnBkLkP/w88iftQQ5U+cjqy+/MBxi64PEoDMTRyLo1n8aHBWcyiwXrBcOFIOu1qCaph1OyXvrqfMcXanVw9aLPm9Hi+3TCOs4WQ3TTHNOfFTMAPV6I2Csl18Pz8btP6d/qquS9mae9dtgv3mqmEXa+s71JiG1UUcULn1bqUm7ooLDGWLr76+NKRze/7hGPUzPnvrt5l06D4z1kuvEHQY9yp+/QoA91Vsv0jdXDVOePrVCPj8Knl2JNNrX8MKvVtfJxpvV7AETdN+lkAcO7a4zWnPA0YGe51155CYPw+A1J3JLpjvSLxkIW9/RqHDlKSW1qeS9jbA07k11qrtiAUdH4tO95E1qj9ZjExqDc7vTPvE7pbQ2yQPH70dGl1thvrwfMukNdgyZipxJ81Gw4BWUvP0pvFt2U8R3UNxFqMjK0X0Rw7frAKwXX0dtUQ9k9MLhTaZz+EPixVGr72TjwKLwhXeU0tokD5xgEAGrHcHiUoTKysXgxzLFE8h0wNpmKLVF3W3qhcPyUmhvbsozp/7AiD1Azn2Pi35qlTRwTreCJR5kdryF2sKuqHY7o4FTSWtI1jX3aJo9XD8fDIa82jelvwo4nA3m9FJldg4qzDb49hxEydoUuJ9+Hbn3P0GDMgWZSSMo/O9D7akjINAJh+UcN0u4LLU6qxu7U1vSGITIM2jVLwZOiDbBvNaU7z8q3EnJu+tRuOJdcWHRNW6uiIb42Np8ThJtOLuK4CGSQuKgopNwKxEw6tnmaOHkjp+rA87oXy4cngG8Nvh27EfJmo1wL1xJG8AnxdtpHzoFth5jkNH8JnHoxWEx3zjl+wSRfQdHd3xd68TZSwSGlmgqejiu8XM0w8keOElXrk0qOJ6NW2Fu1gvpfxwAS5M+okO8R4ikT9orEDjaqlqA6ztz0W5RwaGQzXEH3xfn/ZN6vVXGgYjrnkfFEbtWSQWnfO+P1FHOevMs4Lr57VerU48xyPpnTzRw/EcsyGh5syirVmd14zyde/Gp71OoSSo4wWIv7IOniNnS0FnBAyZmWmOqhzafas9Ut2jguJe+obRVvc4TFslseFL03d2TLiDwfPw1Uht1ojJas941jQeZ14D0i/vDOeohGpCdcNw6s976foYT0La59f+UivTmg0U5tfqqGyc/bd1HimhSj6SDw8pf8KroEJ8q1u2WInXzwEcisgRkdrgNebMWo2z790ptQN70p8Xn6nVELALnDqVE3QocTUdWb75UqeWEla+OtUfeI0uV0tolJRxOghav/gQZCcniPCRyvsK5NY7M2ik/cw6O/v75fWFtOQTOe+ai9NPN1Da7UssJuRetFM+rt6/KaH1r2luE5gFrtmp6iE9zS9ZsgLXtLVSftqMD8RWYywfBf9Sq1KJdUsKpUtBdLPYznGNzDJtO4fRkskni5xyKfNwLV8G75bvIEXgdNypK1nxOQE99XFBlnN7nweTvHrnGzkHR6x/B++kWYe5lb4lsgLkRn/nUfS50wnjWdID7iVeUluiT1HBiJe+mHcK1qLfvZOPrTt0JQCcxQzlMrrLINyvqz6NVGZfJHnQfgl6f0hJ9ihOcRJRt3qOUjr94U2s5l1M32ge2ocZbAmvCzSg/lKq0Qr/iAIffxM4EZ69SOv4KHOb9CK1fUUaAeo3XN2uLG1G2dZ/SgugUBzjsyzvBMeLBqKd3Q8Vfyc+++l4aNC3RVbTWS4wHuzJb77vE/83QUMUFDs8ejricY2eJ7PGZkOOOmeKNVm9fQ437115ElnxXocKh75T2VIoTHJ49HOW0R9age2kN+EGp6fQqWFqGspRdKHjy38hsx2c5dbk1zkhE7krXTKCqhcv8u0SlT7wfuwrZN0xE6dqUmB4QRgfHbBONiyQkO9RjnLikN7YJn9/zDdAW4ppu7l+fEZBjIb4XzRcv+JZNdblffFN8Z+eo6SJqQ6uf26NmIrP8275w/XkOrM0GwdwkEmVWbW5P9IfvNHSlWZKE9Iso5B41C6WffC1ST7FWVHAqHblw3Dad3pYJsN8wqR67D9lDJlMHNoude/qVgwhQcxq05rD8oT9ypjwBb8o3CBzL0HS3OFRZKf4fhPJdByg83wPf7oMoo70On/Pw8XZ1Fa1ei4yByeKt5rbWZVk3jIXz7tkEOSBut3q/3AX3UyvhGj8PzuQZ1I8psN84jWDMRv7cF+DhDa8zV/ddBj2KCk5D5KdBZDfDwMyNehKkKwQo/o8kHMnTkcuX+WiX7n7yNbifWS1OMwv++TLyZj8nfl9AGzr38++g8OW1KFq+Dt4vdtMmtPYZSSzdy5lS3OFUiS9y8Nvu2bRd3ErJnbkIrjGzaff/ABx3PgjXpMdppi0lOKtQtPI/KP3oC3g3fCO++sfR19mgMwanLv0a3vpYSEo4hiIy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOtAL+D0PoHdDmCF5sAAAAAElFTkSuQmCC";

            string isaLogo = "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";

            string firma1 = "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";

            string firma2 = "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";

            string firma3 = "/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCAArAGYDASIAAhEBAxEB/8QAHAAAAwEAAwEBAAAAAAAAAAAAAAcIBgMECQUB/8QAKhAAAQMEAgIBAwUBAQAAAAAAAQIDBAUGBxEAEgghEyIxQRQVIzJRUnH/xAAWAQEBAQAAAAAAAAAAAAAAAAAAAQL/xAAcEQEAAwACAwAAAAAAAAAAAAAAAREhAjESQVH/2gAMAwEAAhEDEQA/APVPhw4cD59auCh25HYlV+rRaezKlMwmVyHQ2HJDqwhppJP3UpRAA+5J59Dks+bjFWvC5MD4ot5PyT6xkiDXH0pd0UQqa2t59akj2UDsk7+wUEj7kcqbhRw4cOEHDhzEZXyvRMR0yg1KtwZcsXDcdNtqM3G69hImPBpC1diPpTsqOvegdDgbfhw4cA4cOHAOHDhwDnHID5YcEVSEvFCvjLgJSFa9bA9kb5ycxF25vw/Yv7k3duTbYpcikNF2ZGk1Zht9odewBbUoKBI1oa97HARODMQ1s+VV55Rv+6l3RcVuW9DoL85SCiO1OlqMp9mGz2IYYaY/SoSP7qLrpUT291Zsb67G/wDOQ23na0E+JGT7qx5frrl83JS63esyZRw69+1yXgPgacktgtsutMCM0EKWFbb9D8cVuVLoyJdk6+skWrkGs0NVBRZWIIFbiSl9/nkPsSaq8okj+RJkIQV72dn37HL0s6u69o2SpGa8ZuWzNlsWhGYrj1zoQE/C+r4GUQ0L377Bxa1p1/wrfGTzySh5NytWsZZepNjZfuifTq5ftKxdjiS/V3n5AAnPPOPpeCg4olh5tJWNkpKAdgDVveDlx35V7Fve379yBIvl+0b8rFvxLhkkfJOYYUj2QPSQFqWANqGtAHQHJZR5X5c8eybHuK8pRQGKDSpdTcK99QlhlTh3r3rSfxyUYkW45GH/ABUo191ioVq4ZlxxLqqcye67LkBpimzZzyisElfTu20nZJIKQBvWnV5dNvTPHW8qJHluR3q9GYobSmk9lrVMkNRghI/JV8vXX57cQeacszoNfyVXMcT0MVPElLo2P7W6R0OhNZq8thElxtCtpW4022wgI6k6DoH9jvMy1xj231c8tYdzzcT3JjCqBu07np1x3PW1T4fR4Uqlx1pUCkgqaJklA3r2EK0fwWj4y3Nfl64Gsy88mSG3rir9OFVkluOlhKW31qdYQEJAH0sraTv89dn2eSD5MvLuTMdwY0tJUKNKqVJt7FMZUZkNht+qzHKhVVISjQ0mLGbKx90l1H27b518w5JyTZlHzdd1n3bXGaZArrGHrZoMeT8bEVxVPYCZTTaQSl9t5f0qSR2Q4oEbSkhelfFYZvyHeNsZJw7ZVnVFmOLuuSSmsNqjodW5S4sN15/r2/oO3xArT7HYf+HqeHtz3vfeHl5BvmtO1B+6bhrFTpyVnYi01UxxEVlHs/xhttKk/wCBYB97PEpkq802Hfdy3HXK6H2/H7EbcJD8uT8zz9fqbfVC+y9la1Nx2k9jpSi772FesBQPIvOVB8ZKfhWzsfLx1kenN2XQrYkvPMS25LdTV1Q9ICm+rLrgjvqWkpJHyp9A8Xp449F1LQgbWtKd/wCnXDkSY8xZXs7Z9u6DkfJtdvC2MYUxq3USnm4zRcr0n4X56ELZZbQtDPwpQAUdk/Jon1tRzUJUR3K3eIrNnhfgTOUyPWLjsikQax+9wqzPqsKmRhMqYj+v00h5TZWpladJUAoEhKffrXHrw4ZZy7seWhfFiVLGlw0dty26tBVTZMFhSo6TGUnqUJLZSUDXodSNfji1R4Z+OzOGZGAI1irYsmXUBVZEFuoyUuPSwpKvlW/8nyqO0I+6vskD7Djt4cCc2fArANNpVvUC3I1w0SlWuurSKbEh1l7TMmoMfC9JS4sqcDqE6LagodCAQOMzBeEbJ8e8cwMaWG1KNPhuOyHpMxwOSZkh1ZU488vQ7LUTrevQCQPQHN/z94C8zpj66sl2Qxb1mXJT6FVItapdYZmToJltJVCltSUgthSSdqaT+RzFU7w2xVT49J6yaw9Nh3mzf1TluyA4us1ptKgH5IUkgJ+pRDbfRKSQQNgafHDgJy1PEvCVo3Obyh26/LrKLqn3jGlzJa3FxahMZS0900QFN9U/SlfbqSSCOYS9PBC0r6yXNv8AqmTbsjQnrvg3tFt+GtlEBipMNNNuOLSpKi6XfhTskgpBITrZPKdHDhblNuZ/BnGWUKFdTFEfXQbgvKvxK9VqzIQqorfWwkoSypDqx/ClK1FDQUEJWEkDQ1zv3b4c21cWP41pwL9uSmV0XNCu2fdiVtu1SbUYySlDqlKT0R1HUIShIQ2EJCUgDlCcOSi5LnAeD7b8fMfN49tip1Kps/rpVSk1CpuJcly5EhwrW48sAd1ewnsRvSRvfDjG4cqP/9k=";

            // Lista con la estructura de celdas de la tabla
            List<List<string>> datos3 = new List<List<string>> {
                    new List<string> { "[B64]" + iebLogo, "~", "~", "~", "[B64]" + isaLogo, "~", "~", "~",  },
                    new List<string> { "Nombres", "~", "Firma", "Matricula", "Total de paginas", "1006", "Fecha de Emisión", "2022.09.16",  },
                    new List<string> { "Elaboró", "c.castaño".ToUpper(), "[B64]" + firma1, "267773 ANT", "Nombre del proyecto", "~", "~", "~",  },
                    new List<string> { "Revisó", "c.metrio".ToUpper(), "[B64]" + firma2, "357197 ANT", "RENOVACIÓN SUBESTACIÓN BANADÍA 230 kV", "~", "~", "~",  },
                    new List<string> { "Aprobó", "i.vullalba".ToUpper(), "[B64]" + firma3, "196375 ANT", "Código del Documento", "~", "~", "~",  },
                    new List<string> { "|", "|", "|", "|", "CO-RBAN-14113-S-01-D1531", "~", "~", "~",  },
                };

            PropiedadesTabla.AgregarTablaConImagen(ruta, datos3);
        }

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
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos, 1);
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
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo6, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2, 1);
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
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, 2);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo2, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesParrafo.AgregarParrafo(ruta, parrafo3, 12, FuncionesCreacion.EstiloParrafo.Normal, FuncionesCreacion.AlineacionTexto.Justificado);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2);
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
                PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 10, 2, FuncionesCreacion.AlineacionImagen.Centro);
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
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos1, 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos2, 1);
                PropiedadesParrafo.AgregarSaltosDeLinea(ruta, 1);
                PropiedadesTabla.AgregarTablaDesdeLista(ruta, datos3, 1);
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
                PropiedadesImagen.AgregarImagenDesdeArchivo(ruta, rutaSalidaImagen, 11, 5, FuncionesCreacion.AlineacionImagen.Centro);
                //PropiedadesImagen.AgregarImagenDesdeBase64(ruta, base64Imagen, 5, 5, FuncionesCreacion.AlineacionImagen.Centro);
                
            }
            catch (Exception ex)
            {
                // Mostrar el mensaje de error en caso de que se de alguna excepcion
                Console.WriteLine("Error al crear el documento de word" + ex.Message);
            }
        }
    }
}
