using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Qualitas.Models
{
    public class Correos
    {
        public void EnviarCorreo(string correo, string nombreasesor, string errores, string fechaini, string fechafin) { 
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

                if (cuentatmp.SmtpAddress == "GEROPINT@BANCOLOMBIA.COM.CO" || cuentatmp.SmtpAddress == "joarios@bancolombia.com.co")
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

                    if (errores == "1")
                    {
                        cuerpo = $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Hola {nombreasesor},</font><br><br>" +
                                 $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Te invitamos a revisar la página de excelencia operacional; en la última semana tienes {errores} nuevo reproceso reportado. En caso de tener alguna duda con la información por favor consulta con tu coordinador o jefe inmediato." +
                                 $"<br><br>Estamos seguros que la revisión oportuna de tus novedades de calidad, te ayudan a generar planes de acción transformadores.</font>" +
                                 $"<br><br><font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Contamos contigo,</font>" +
                                $"<br><br><img src='data:image/png;base64,{imagenBase64}' />";
                    }
                    else
                    {
                        cuerpo = $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Hola {nombreasesor},</font><br><br>" +
                                 $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Te invitamos a revisar la página de excelencia operacional; en la última semana tienes {errores} nuevos reprocesos reportados. En caso de tener alguna duda con la información por favor consulta con tu coordinador o jefe inmediato." +
                                 $"<br><br>Estamos seguros que la revisión oportuna de tus novedades de calidad, te ayudan a generar planes de acción transformadores.</font>" +
                                 $"<br><br><font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>Contamos contigo,</font>" +
                                $"<br><br><img src='data:image/png;base64,{imagenBase64}' />";
                    }


                    Outlook.Application outlookApp = null;
                    Outlook._NameSpace outlookNamespace = null;

                    try
                    {
                        // Intenta obtener la instancia actual de Outlook
                        outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    }
                    catch (Exception)
                    {
                        // Si no se puede obtener la instancia, crea una nueva
                        outlookApp = new Outlook.Application();
                    }

                    // Obtiene el namespace MAPI
                    outlookNamespace = outlookApp.GetNamespace("MAPI");
                    outlookNamespace.Logon("", "", Missing.Value, Missing.Value);

                    // Verifica si la bandeja de correos (Outlook) está abierta
                    if (outlookApp != null && outlookNamespace != null)
                    {
                        // Verifica si la cuenta especificada está disponible
                        Outlook.Accounts accounts = outlookNamespace.Accounts;
                        Outlook.Account cuenta = null;
                        foreach (Outlook.Account acc in accounts)
                        {
                            if (acc.SmtpAddress == "geropint@bancolombia.com.co" || acc.SmtpAddress == "joarios@bancolombia.com.co")
                            {
                                cuenta = acc;
                                break;
                            }
                        }

                        if (cuenta != null)
                        {
                            mensaje.To = correo;

                            mensaje.HTMLBody = cuerpo;
                            mensaje.Subject = asunto;
                            mensaje.Sensitivity = Outlook.OlSensitivity.olConfidential;
                            mensaje.Display();
                            mensaje.SendUsingAccount = cuentatmp;
                            mensaje.Send();
                            Console.WriteLine("Correo enviado");
                          
                        }
                        else
                        {
                            // La cuenta especificada no está disponible
                            Console.WriteLine("La cuenta de correo especificada no está disponible.");
                        }
                    }
                    else
                    {
                        // Outlook no está abierto
                        Console.WriteLine("Outlook no está abierto.");
                    }

                  

                }
            }
        }


        public void EnviarCorreoDuo(string correo, string nombreasesor, string fechaini, string fechafin)
        {
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

                if (cuentatmp.SmtpAddress == "GEROPINT@BANCOLOMBIA.COM.CO" || cuentatmp.SmtpAddress == "joarios@bancolombia.com.co")
                {
                    string asunto = "";
                    string cuerpo = "";


                    asunto = "¡Felicitaciones! Estás brillando.";

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

                   
                   cuerpo = $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>¡Felicitaciones! {nombreasesor},</font><br><br>" +
                                 $"<font color='#7D7C7C' face='CIBFont Sans' style='font-size:15px;'>En la última semana brillaste por tu calidad, no cuentas con reprocesos. Gracias por tu compromiso; eres la muestra de que juntos podemos lograr resultados extraordinarios." +
                                 $"<br><br>Continúa luciéndote con tu excelencia operacional.</font>" +
                                $"<br><br><img src='data:image/png;base64,{imagenBase64}' />";
                   


                    Outlook.Application outlookApp = null;
                    Outlook._NameSpace outlookNamespace = null;

                    try
                    {
                        // Intenta obtener la instancia actual de Outlook
                        outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    }
                    catch (Exception)
                    {
                        // Si no se puede obtener la instancia, crea una nueva
                        outlookApp = new Outlook.Application();
                    }

                    // Obtiene el namespace MAPI
                    outlookNamespace = outlookApp.GetNamespace("MAPI");
                    outlookNamespace.Logon("", "", Missing.Value, Missing.Value);

                    // Verifica si la bandeja de correos (Outlook) está abierta
                    if (outlookApp != null && outlookNamespace != null)
                    {
                        // Verifica si la cuenta especificada está disponible
                        Outlook.Accounts accounts = outlookNamespace.Accounts;
                        Outlook.Account cuenta = null;
                        foreach (Outlook.Account acc in accounts)
                        {
                            if (acc.SmtpAddress == "geropint@bancolombia.com.co" || acc.SmtpAddress == "joarios@bancolombia.com.co")
                            {
                                cuenta = acc;
                                break;
                            }
                        }

                        if (cuenta != null)
                        {
                            mensaje.To = correo;

                            mensaje.HTMLBody = cuerpo;
                            mensaje.Subject = asunto;
                            mensaje.Sensitivity = Outlook.OlSensitivity.olConfidential;
                            mensaje.Display();
                            mensaje.SendUsingAccount = cuentatmp;
                            mensaje.Send();
                            Console.WriteLine("Correo enviado");

                        }
                        else
                        {
                            // La cuenta especificada no está disponible
                            Console.WriteLine("La cuenta de correo especificada no está disponible.");
                        }
                    }
                    else
                    {
                        // Outlook no está abierto
                        Console.WriteLine("Outlook no está abierto.");
                    }



                }
            }
        }





    }
}
