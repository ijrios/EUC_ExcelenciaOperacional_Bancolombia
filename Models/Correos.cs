using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Qualitas.Models
{
    public class Correos
    {
        public void EnviarCorreo(string correo, string nombreasesor, string errores) { 
            Outlook.Application outlook = new Outlook.Application();
            Outlook.MailItem mensaje = outlook.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Accounts cuentas;
            DateTime fecha = DateTime.Now;

            DateTime Hoy = DateTime.Now;
            //DateTime Hoy = Hoyy.AddDays(-4);
            var DiaActualSemana = Hoy.DayOfWeek;


            cuentas = outlook.Session.Accounts;

            foreach (Outlook.Account cuentatmp in cuentas)
            {

                if (cuentatmp.SmtpAddress == "jaramos@bancolombia.com.co" || cuentatmp.SmtpAddress == "joarios@bancolombia.com.co")
                {
                    string asunto = "";
                    string cuerpo = "";

                   
                    asunto = "¡Es momento de una pausa para la excelencia! – Valida tus reprocesos.";
                 
                    //Inicio cuerpo del correo.
                    string firma = $@"C:\Users\{Environment.UserName}\Documents\Productividades\firma.png";

                    // Convertir la imagen a base64
                    string imagenBase64 = "";
                    using (Image imagen = Image.FromFile(firma))
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            imagen.Save(ms, imagen.RawFormat);
                            byte[] imagenBytes = ms.ToArray();
                            imagenBase64 = Convert.ToBase64String(imagenBytes);
                        }
                    }


                    cuerpo = $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Hola {nombreasesor},</font><br><br>" +
                                  $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Te invitamos a revisar la página de excelencia operacional; en la última semana tienes {errores} nuevos reprocesos reportados. En caso de tener alguna duda con la información por favor consulta con tu coordinador o jefe inmediato." +
                                  $"<br><br>Estamos seguros que la revisión oportuna de tus novedades de calidad, te ayudan a generar planes de acción transformadores.</font>" +
                                  $"<br><br><font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Contamos contigo,</font>" +
                                 $"<br><br><img src='data:image/png;base64,{imagenBase64}' />";
                    mensaje.To = correo;
                  
                    mensaje.HTMLBody = cuerpo;
                    mensaje.CC = "sebmoral@bancolombia.com.co;kjpadill@bancolombia.com.co";
                    mensaje.Subject = asunto;
                    mensaje.Sensitivity = Outlook.OlSensitivity.olConfidential;
                    mensaje.Display();
                    mensaje.SendUsingAccount = cuentatmp;
                    mensaje.Send();
                    Console.WriteLine("Correo enviado");
                    InsertarRepro(Hoy.ToString("yyyyMMdd"),1, DiaActualSemana.ToString());

                }
            }
        }


        public void InsertarRepro(string fecha, int bit, string dia)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "INSERT INTO CorreosEnvio (bit_env, fecha_env, dia_env) " +
                       "VALUES (@bit_env, @fecha_env, @dia_env)";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Agregar parámetros para evitar la inyección SQL
                        command.Parameters.AddWithValue("@bit_env", Convert.ToInt32(bit));
                        command.Parameters.AddWithValue("@fecha_env", fecha);
                        command.Parameters.AddWithValue("@dia_env", dia);

                        // Ejecutar la consulta de inserción
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al insertar datos en la tabla reprocesos:", ex);
                // Puedes lanzar la excepción nuevamente para manejarla en el nivel superior
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al insertar datos en la tabla reprocesos:", ex);
                // Puedes lanzar la excepción nuevamente para manejarla en el nivel superior
                throw;
            }
        }


    }
}
