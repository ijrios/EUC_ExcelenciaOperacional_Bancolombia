using Newtonsoft.Json;
using Qualitas.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Qualitas.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;

            return View();
        }
        public ActionResult CerrarSesion()
        {
            Session.Clear();
            Session.Abandon();

            return RedirectToAction("Index", "Home");
        }

        public ActionResult Registro()
        {

            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;

            AccessDatabase connection = new AccessDatabase();
            var con = connection.Inicio(windowsUsername);

            if (con == null)
            {
                ViewBag.MensajeDuo = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Home");
            }
            else if (con == "Administrador")
            {
                ViewBag.MensajeUnus = "Bienvenido al ingreso de reprocesos";
                return View();
            } else if (con == "Consulta")
            {
                ViewBag.MensajeTris = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Home");
            }
            else
            {
                string mensajeRechazo = char.ToUpper(windowsUsername[0]) + windowsUsername.Substring(1) + " no cuentas con permisos para ingresar a esta vista";

                return RedirectToAction("Index", "Home", new { mensajeRechazo = mensajeRechazo });

            }

        }

        public ActionResult Consultas()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }

        public ActionResult ConsultasnoAdmin()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }

        public ActionResult ImpactoInfo()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }


        public ActionResult Asesores()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }

        public ActionResult Informes()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }


        public ActionResult Perdidas()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }

        public ActionResult ReporteTotal()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;

            AccessDatabase connection = new AccessDatabase();
            var con = connection.Inicio(windowsUsername);

            if (con == null)
            {
                ViewBag.MensajeDuo = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Home");
            }
            else if (con == "Administrador")
            {
                ViewBag.MensajeUnus = "Bienvenido al histórico de errores por usuarios";
                return View();
            }
            else if (con == "Consulta")
            {
                ViewBag.MensajeTris = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Asesores");
            }
            else
            {
                string mensajeRechazo = char.ToUpper(windowsUsername[0]) + windowsUsername.Substring(1) + " no cuentas con permisos para ingresar a esta vista";

                return RedirectToAction("Index", "Asesores", new { mensajeRechazo = mensajeRechazo });

            }
        }

        public ActionResult Causas()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }

        public ActionResult AreasInfo()
        {
            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;
            return View();
        }



        public ActionResult Administrador()
        {

            string windowsUsername = User.Identity.Name;
            if (!string.IsNullOrEmpty(windowsUsername) && windowsUsername.StartsWith("BANCOLOMBIA\\"))
            {
                windowsUsername = windowsUsername.Substring("BANCOLOMBIA\\".Length);
            }

            ViewBag.WindowsUsername = windowsUsername;

            AccessDatabase connection = new AccessDatabase();
            var con = connection.Inicio(windowsUsername);

            if (con == null)
            {
                ViewBag.MensajeDuo = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Home");
            }
            else if (con == "Administrador")
            {
                ViewBag.MensajeUnus = "Bienvenido a la ventana de administración de usuarios";
                ViewBag.WindowsUsername = windowsUsername;
                return View();
            }
            else if (con == "Consulta")
            {
                ViewBag.MensajeTris = "No cuenta con permisos para esta vista";
                return RedirectToAction("Index", "Home");
            }
            else
            {
                string mensajeRechazo = char.ToUpper(windowsUsername[0]) + windowsUsername.Substring(1) + " no cuentas con permisos para ingresar a esta vista";

                return RedirectToAction("Index", "Home", new { mensajeRechazo = mensajeRechazo });

            }

        }

        [HttpPost]
        public ActionResult Buscar(string consecutivo, string fecha, string code)
        {
            try
            {
                AS400Connection connection = new AS400Connection();
                var dataTable = connection.ConsultarDatos(consecutivo, fecha, code);

                if (dataTable == null)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                var data = dataTable.AsEnumerable().Select(
                    row => dataTable.Columns.Cast<DataColumn>().ToDictionary(
                        column => column.ColumnName,
                        column => row[column]
                    )
                ).ToList();

                return Json(data, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                // Registrar la excepción para su diagnóstico
                Console.WriteLine("Excepción en la búsqueda: " + ex.Message);
                return Content("Error en la búsqueda");
            }

        }



        [HttpPost]
        public ActionResult BuscarProductoEvento(string producto, string evento)
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                var dataTable = causa.GetProdEvento(producto, evento);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (dataTable == null)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }
                var data = dataTable.AsEnumerable().Select(
                                   row => dataTable.Columns.Cast<DataColumn>().ToDictionary(
                                       column => column.ColumnName,
                                       column => row[column]
                                   )
                               ).ToList();

                return Json(data, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }

        }

        [HttpPost]
        public ActionResult EliminarReproceso(int id)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EliminarRepro(id);

                ViewBag.Mensaje = "Reproceso eliminado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al eliminar el reproceso: " + ex.Message;
            }

            return Json(new { success = true });
        }

        [HttpPost]
        public ActionResult EliminarCausa(int id)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EliminarCausas(id);

                ViewBag.Mensaje = "Causa eliminada correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al eliminar la causa: " + ex.Message;
            }

            return Json(new { success = true });
        }


        [HttpPost]
        public ActionResult EliminarUsuario(int id)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EliminarUsu(id);

                ViewBag.Mensaje = "Reproceso eliminado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al eliminar el reproceso: " + ex.Message;
            }

            return Json(new { success = true });
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


        [HttpPost]
        public ActionResult EnviarCorreos()
        {
            try
            {
                AccessDatabase correoServ = new AccessDatabase();

                DateTime fechaActual = DateTime.Now;
                string diaActual = fechaActual.DayOfWeek.ToString();
                int limiteinf = 0;
                int limitesup = 0;

                string fechaInicial = null;
                string fechaFinal = null;
                string fechaInicialCorreo = null;
                string fechaFinalCorreo = null;
                string fechaInicialUsuarios = null;
                string fechaFinalUsuario = null;
                string confirmacion = null;
                if (diaActual == "Monday")
                {
                    limiteinf = -7;
                    limitesup = -3;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                    fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                    fechaInicialCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = correoServ.GetCorreoConfirma(fechaInicialCorreo, fechaFinalCorreo);
                }
                else if (diaActual == "Tuesday")
                {
                    limiteinf = -8;
                    limitesup = -4;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                    fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                    fechaInicialCorreo = fechaActual.AddDays(-1).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = correoServ.GetCorreoConfirma(fechaInicialCorreo, fechaFinalCorreo);
                }
                else if (diaActual == "Wednesday")
                {
                    limiteinf = -9;
                    limitesup = -5;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                    fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                    fechaInicialCorreo = fechaActual.AddDays(-2).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = correoServ.GetCorreoConfirma(fechaInicialCorreo, fechaFinalCorreo);
                }
                else if (diaActual == "Thursday")
                {
                    limiteinf = -10;
                    limitesup = -6;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                    fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                    fechaInicialCorreo = fechaActual.AddDays(-3).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = correoServ.GetCorreoConfirma(fechaInicialCorreo, fechaFinalCorreo);
                }
                else if (diaActual == "Friday")
                {
                    limiteinf = -11;
                    limitesup = -7;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                    fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                    fechaInicialCorreo = fechaActual.AddDays(-4).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = correoServ.GetCorreoConfirma(fechaInicialCorreo, fechaFinalCorreo);
                }
                else
                {
                    limiteinf = 0;
                    limitesup = 0;
                    fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                    fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                    fechaInicialCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                    confirmacion = "NO";
                }

                List<ErrorUsuario> errores = ObtenerUsuariosError(fechaInicialUsuarios,fechaFinalUsuario);
                List<Usuarios> usuarios = ObtenerUsuariosDesdeLaClase();

             

                if (errores != null && errores.Count > 0)
                {
                    Correos mail = new Correos();
                    

                    

                    if (confirmacion == "Enviado")
                    {
                        //Yo Alexander Rios no se que pueda a hacer aquí realmente, ya que no hay opción 
                        ViewBag.Mensaje = "Ya se enviaron";
                        return Json(new { success = "Ya" });
                    }
                    else if (confirmacion == "Error")
                    {
                        ViewBag.Mensaje = "No se ha enviado correos esta semana";
                        return Json(new { success = "Error" });
                    }
                    else if (confirmacion == "OK")
                    {

                        foreach (var errorUsuario in errores)
                        {
                            foreach (var errorUsuarioTotal in usuarios)
                            {
                                if (errorUsuarioTotal.Usuario.Contains(errorUsuario.Usuario))
                                {
                                    // Asumiendo que ErrorUsuario tiene propiedades para el correo y el nombre del usuario
                                    string correo = errorUsuarioTotal.Correo;
                                    string nombre = errorUsuario.Usuario;
                                    string idError = errorUsuario.Errores.ToString();
                                    mail.EnviarCorreo(correo, nombre, idError, fechaInicial, fechaFinal);
                                }

                            }


                        }
                        InsertarRepro(fechaActual.ToString("yyyyMMdd"), 1, diaActual.ToString());
                        ViewBag.Mensaje = "Correos enviados correctamente";
                        return Json(new { success = "Enviado" });
                    }
                    else if (confirmacion == "NO") 
                    {
                        //Yo Alexander Rios no se que pueda a hacer aquí realmente, ya que no hay opción 
                        ViewBag.Mensaje = "Ya se enviaron";
                        return Json(new { success = "Ya" });
                    }
                
                }
                else
                {
                    ViewBag.Mensaje = "No se encontraron usuarios con errores";
                }
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al enviar los correos: " + ex.Message;
            }

            return Json(new { success = false });
        }


        [HttpPost]
        public ActionResult InsertarReproceso(
           string FechaOp,
           string FechaReg,
           string Consecutivo,
           string Nit,
           string Moneda,
           decimal Valor,
           string Cliente,
           string ProdEvento,
           string Responsable,
           string UsuarioReproceso,
           string Perdida,
           string Impacto,
           string Causa,
           string Descripcion,
           string Area,
           string TipoError,
           string Mes,
           string Año,
           string QuejaCliente)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.InsertarRepro(FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes,Año, QuejaCliente);

            }
            catch (Exception ex)
            {
                return Json(new { success = false });
            }

            return Json(new { success = true });
        }

        [HttpPost]
        public ActionResult EditarReproceso(
           int id,
           string FechaOp,
           string FechaReg,
           string Consecutivo,
           string Nit,
           string Moneda,
           decimal Valor,
           string Cliente,
           string ProdEvento,
           string Responsable,
           string UsuarioReproceso,
           string Perdida,
           string Impacto,
           string Causa,
           string Descripcion,
           string Area,
           string TipoError, 
           string Mes,
           string Año,
           string QuejaCliente)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EditarReprocesos(id, FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes,Año, QuejaCliente);

                ViewBag.Mensaje = "Reproceso registrado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el reproceso: " + ex.Message;
            }

            return Json(new { success = true });
        }

        private List<Reproceso> ObtenerReprocesosDesdeLaClase()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDBPerdida = new AccessDatabase();
            return reprocesoDBPerdida.ObtenerReprocesos();
        }

        private List<Reproceso> ObtenerReprocesosFechaDesdeLaClase(string fechaini, string fechafin, string area)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDBPerdida = new AccessDatabase();
            return reprocesoDBPerdida.ObtenerReprocesosFechaArea(fechaini, fechafin, area);
        }

        [HttpPost]
        public ActionResult EditarUsuarios(
         int id,
           string usuario,
           string nombre,
           string correo,
           string perfil)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EditarUsuario(id, usuario, nombre, correo, perfil);

                ViewBag.Mensaje = "Usuario editado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return Json(new { success = true });
        }


        [HttpPost]
        public ActionResult EditarCausas(
         int id,
           string area,
           string tipoerror,
           string causas)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EditarCausa(id, area, tipoerror, causas);

                ViewBag.Mensaje = "Usuario editado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return Json(new { success = true });
        }


        [HttpGet]
        public ActionResult ObtenerReprocesosFechaArea(string fechaini, string fechafin, string area)
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<Reproceso> reprocesos = ObtenerReprocesosFechaDesdeLaClase(fechaini, fechafin, area);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (reprocesos == null || reprocesos.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(reprocesos);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }
        }


        [HttpGet]
        public ActionResult ObtenerReprocesos()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<Reproceso> reprocesos = ObtenerReprocesosDesdeLaClase();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (reprocesos == null || reprocesos.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(reprocesos);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }
        }

       

        private List<ReprocesoPerdida> ObtenerReprocesosPerdidasDesdeLaClase(string area, string mes)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDB = new AccessDatabase();
            return reprocesoDB.ObtenerReprocesosPerdidaEconomica(area, mes);
        }

        private List<ReprocesoPerdida> ObtenerReprocesosNoPerdidasDesdeLaClase(string area, string mes)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDB = new AccessDatabase();
            return reprocesoDB.ObtenerReprocesosPerdidaNoEconomica(area, mes);
        }

        [HttpGet]
        public ActionResult ObtenerReprocesosPerdidas(string area, string mes)
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<ReprocesoPerdida> reprocesos = ObtenerReprocesosPerdidasDesdeLaClase(area, mes);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (reprocesos == null || reprocesos.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(reprocesos);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }
        }


        [HttpPost]
        public ActionResult LlenarReprocesos(string jsonData)
        {
            // Deserializar el JSON en una lista de objetos Reproceso
            List<Reproceso> reprocesos = JsonConvert.DeserializeObject<List<Reproceso>>(jsonData);

            Excel.Application myexcelApplication = new Excel.Application();
            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

            myexcelWorksheet.Cells[1, 1].Value = "Filial";
            myexcelWorksheet.Cells[1, 2].Value = "Fecha Inicial";
            myexcelWorksheet.Cells[1, 3].Value = "Fecha de descubrimiento";
            myexcelWorksheet.Cells[1, 4].Value = "Fecha final";
            myexcelWorksheet.Cells[1, 5].Value = "Periodo al que corresponde la información";
            myexcelWorksheet.Cells[1, 6].Value = "Gerencia que reporta";
            myexcelWorksheet.Cells[1, 7].Value = "Nit";
            myexcelWorksheet.Cells[1, 8].Value = "Descripción del evento";
            myexcelWorksheet.Cells[1, 9].Value = "Queja del cliente";
            myexcelWorksheet.Cells[1, 10].Value = "Subproceso";
            myexcelWorksheet.Cells[1, 11].Value = "Area geográfica";
            myexcelWorksheet.Cells[1, 12].Value = "Producto";
            myexcelWorksheet.Cells[1, 13].Value = "Tipo de Falla";
            myexcelWorksheet.Cells[1, 14].Value = "Valor moneda del país";
            myexcelWorksheet.Cells[1, 15].Value = "Hora inicial";
            myexcelWorksheet.Cells[1, 16].Value = "Hora Descubirimiento";
            myexcelWorksheet.Cells[1, 17].Value = "Hora final";
            myexcelWorksheet.Cells[1, 18].Value = "Cuenta contable";
            myexcelWorksheet.Cells[1, 19].Value = "Centro de costos";
            myexcelWorksheet.Cells[1, 20].Value = "Responsable";

            int rowunus = 2;

            DateTime fechaActual = DateTime.Now;
            string fechaFinal = fechaActual.AddDays(0).ToString("yyyy-MM-dd");
            string periodo = fechaActual.AddDays(0).ToString("MM");

            foreach (var reproceso in reprocesos)
            {

                myexcelWorksheet.Cells[rowunus, 1].Value = "Bancolombia";
                myexcelWorksheet.Cells[rowunus, 2].Value = reproceso.FechaOp;
                myexcelWorksheet.Cells[rowunus, 3].Value = reproceso.FechaReg;
                myexcelWorksheet.Cells[rowunus, 4].Value = fechaFinal;
                myexcelWorksheet.Cells[rowunus, 5].Value = periodo;
                myexcelWorksheet.Cells[rowunus, 6].Value = "Gerencia Servicios Comercio Internacional";
                myexcelWorksheet.Cells[rowunus, 7].Value = reproceso.Nit;
                myexcelWorksheet.Cells[rowunus, 8].Value = reproceso.Descripcion;
                myexcelWorksheet.Cells[rowunus, 9].Value = reproceso.Queja;
                myexcelWorksheet.Cells[rowunus, 10].Value = "Cumplir operaciones de compra venta de divisas";
                myexcelWorksheet.Cells[rowunus, 11].Value = "Medellin";
                myexcelWorksheet.Cells[rowunus, 12].Value = "Cumplir operaciones de compra venta de divisas";
                myexcelWorksheet.Cells[rowunus, 13].Value = "Cumplir operaciones de compra venta de divisas";
                myexcelWorksheet.Cells[rowunus, 14].Value = reproceso.Valor;
                myexcelWorksheet.Cells[rowunus, 15].Value = "08:00 a.m";
                myexcelWorksheet.Cells[rowunus, 16].Value = "10:00 a.m";
                myexcelWorksheet.Cells[rowunus, 17].Value = "11:00 a.m";
                myexcelWorksheet.Cells[rowunus, 18].Value = " ";
                myexcelWorksheet.Cells[rowunus, 19].Value = "103700500";
                myexcelWorksheet.Cells[rowunus, 20].Value = reproceso.Responsable;
            }


            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string downloadsFolder = System.IO.Path.Combine(userProfile, "Downloads");
            string informe = System.IO.Path.Combine(downloadsFolder, "Informe de pérdidas económicas.xls");

            //string informe = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Informes Recepción y Envio\Perdidas económicas y no económicas\" + "Informe de pérdidas económicas.xls";
            myexcelApplication.ActiveWorkbook.SaveAs(informe, Excel.XlFileFormat.xlWorkbookNormal);
            Console.WriteLine("Archivo generado");

            myexcelWorkbook.Close();
            myexcelApplication.Quit();


            return Json(new { success = true });
        }

        [HttpPost]
        public ActionResult LlenarReprocesosInformes(string jsonData)
        {
            // Deserializar el JSON en una lista de objetos Reproceso
            List<ReprocesosInformes> reprocesos = JsonConvert.DeserializeObject<List<ReprocesosInformes>>(jsonData);

            Excel.Application myexcelApplication = new Excel.Application();
            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

            // Nombres de columnas basados en las propiedades de la clase Reproceso
            string[] columnNames = { "Id", "FechaOp", "FechaReg", "Consecutivo", "Nit", "Moneda", "Valor", "Cliente", "ProdEvento", "Responsable", "UsuarioReproceso", "Perdida", "Impacto", "Causa", "Descripcion", "Area", "TipoError", "Mes", "Año","QuejaCLiente"};

            // Llenar la primera fila con los nombres de columnas
            for (int i = 0; i < columnNames.Length; i++)
            {
                myexcelWorksheet.Cells[1, i + 1].Value = columnNames[i];
            }

            // Llenar las filas siguientes con los datos de la lista de reprocesos
            int rowunus = 2;

            foreach (var reproceso in reprocesos)
            {
                myexcelWorksheet.Cells[rowunus, 1].Value = reproceso.Id;
                myexcelWorksheet.Cells[rowunus, 2].Value = reproceso.FechaOp;
                myexcelWorksheet.Cells[rowunus, 3].Value = reproceso.FechaReg;
                myexcelWorksheet.Cells[rowunus, 4].Value = reproceso.Consecutivo;
                myexcelWorksheet.Cells[rowunus, 5].Value = reproceso.Nit;
                myexcelWorksheet.Cells[rowunus, 6].Value = reproceso.Moneda;
                myexcelWorksheet.Cells[rowunus, 7].Value = reproceso.Valor;
                myexcelWorksheet.Cells[rowunus, 8].Value = reproceso.Cliente;
                myexcelWorksheet.Cells[rowunus, 9].Value = reproceso.ProdEvento;
                myexcelWorksheet.Cells[rowunus, 10].Value = reproceso.Responsable;
                myexcelWorksheet.Cells[rowunus, 11].Value = reproceso.UsuarioReproceso;
                myexcelWorksheet.Cells[rowunus, 12].Value = reproceso.Perdida;
                myexcelWorksheet.Cells[rowunus, 13].Value = reproceso.Impacto;
                myexcelWorksheet.Cells[rowunus, 14].Value = reproceso.Causa;
                myexcelWorksheet.Cells[rowunus, 15].Value = reproceso.Descripcion;
                myexcelWorksheet.Cells[rowunus, 16].Value = reproceso.Area;
                myexcelWorksheet.Cells[rowunus, 17].Value = reproceso.TipoError;
                myexcelWorksheet.Cells[rowunus, 18].Value = reproceso.Mes;
                myexcelWorksheet.Cells[rowunus, 19].Value = reproceso.Año;
                myexcelWorksheet.Cells[rowunus, 20].Value = reproceso.Queja;

                rowunus++; 
            }

            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string downloadsFolder = System.IO.Path.Combine(userProfile, "Downloads");
            string informe = System.IO.Path.Combine(downloadsFolder, "Informe de calidad.xls");

            //string informe = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Informes Recepción y Envio\Perdidas económicas y no económicas\" + "Informe Calidad.xls";
            myexcelApplication.ActiveWorkbook.SaveAs(informe, Excel.XlFileFormat.xlWorkbookNormal);
            myexcelWorkbook.Close();
            myexcelApplication.Quit();

            return Json(new { success = true });
        }


        [HttpGet]
        public ActionResult ObtenerReprocesosNoPerdidas(string area, string mes)
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<ReprocesoPerdida> reprocesos = ObtenerReprocesosNoPerdidasDesdeLaClase(area, mes);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (reprocesos == null || reprocesos.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(reprocesos);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }
        }

        private List<Usuarios> ObtenerUsuariosDesdeLaClase()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDB = new AccessDatabase();
            return reprocesoDB.ObtenerUsuarios();
        }


        private List<ErrorUsuario> ObtenerUsuariosError(string fechaini, string fechafin)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            
            return error.ObtenerReprocesoUsuarios(fechaini, fechafin);
        }

        private List<ErrorUsuario> ObtenerUsuariosErrorTotus()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();

            return error.ObtenerReprocesoUsuariosTodos();
        }

        private List<ErrorArea> ObtenerAreaError()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoArea();
        }

        private List<ErrorMes> ObtenerErroresGenerales(string code, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoErroresArea(code, año);
        }

        private List<ErrorCausa> ObtenerErroresCausas(string code, string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorCausas(code,mes, año);
        }

        private List<ImpactoError> ObtenerErroresImpacto(string code, string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorImpacto(code, mes, año);
        }

        private List<ErrorPerdida> ObtenerErroresPerdidas(string code, string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorPerdidas(code, mes, año);
        }

        private List<ErrorQueja> ObtenerErroresQuejas(string code, string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorQuejas(code, mes, año);
        }

        private List<ErrorTipo> ObtenerTipoError(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoError(mes, año);
        }

        private List<ErrorTipo> ObtenerTipoErrorSwift(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Swift", mes, año);
        }

        private List<ErrorTipo> ObtenerTipoErrorTrade(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Trade", mes, año);
        }
        private List<ErrorTipo> ObtenerTipoErrorCartera(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Cartera", mes, año);
        }
        private List<ErrorTipo> ObtenerTipoErrorCV(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Compraventa", mes, año);
        }


        private List<ErrorTipo> ObtenerTipoErrorBalanza(string mes, string año)
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Balanza", mes, año);
        } 

      
        [HttpGet]
        public ActionResult ObtenerError()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<ErrorUsuario> errores = ObtenerUsuariosErrorTotus();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public ActionResult ObtenerErrorArea()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<ErrorArea> errores = ObtenerAreaError();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }
        

        [HttpGet]
        public ActionResult ObtenerErrorTipo(string code, string fecha, string año)
        {
            try
            {
                List<ErrorTipo> errores = new List<ErrorTipo>();
                if (code.Equals("1"))
                {
                    errores = ObtenerTipoErrorTrade(fecha, año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerTipoErrorCartera(fecha, año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerTipoErrorBalanza(fecha, año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerTipoErrorCV(fecha, año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerTipoErrorSwift(fecha, año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerTipoError(fecha, año);
                }

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public ActionResult ObtenerErrorCausa(string code, string mes, string año)
        {
            try
            {
                List<ErrorCausa> errores = new List<ErrorCausa>();
                if (code.Equals("1"))
                {
                    errores = ObtenerErroresCausas("Trade",mes, año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerErroresCausas("Cartera", mes, año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerErroresCausas("Balanza", mes, año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerErroresCausas("Compraventa",mes, año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerErroresCausas("Swift",mes, año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerErroresCausas("Todas",mes, año);
                }
                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }


        [HttpGet]
        public ActionResult ObtenerErrorImpacto(string code, string mes, string año)
        {
            try
            {
                List<ImpactoError> errores = new List<ImpactoError>();
                if (code.Equals("1"))
                {
                    errores = ObtenerErroresImpacto("Trade", mes,año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerErroresImpacto("Cartera", mes, año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerErroresImpacto("Balanza", mes, año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerErroresImpacto("Compraventa", mes, año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerErroresImpacto("Swift",mes, año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerErroresImpacto("Todas",mes, año);
                }

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }


        [HttpGet]
        public ActionResult ObtenerErrorPerdida(string code,string mes, string año)
        {
            try
            {
                List<ErrorPerdida> errores = new List<ErrorPerdida>();
                if (code.Equals("1"))
                {
                    errores = ObtenerErroresPerdidas("Trade",mes, año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerErroresPerdidas("Cartera",mes, año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerErroresPerdidas("Balanza",mes, año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerErroresPerdidas("Compraventa",mes, año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerErroresPerdidas("Swift",mes, año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerErroresPerdidas("Todas",mes, año);
                }
                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public ActionResult ObtenerErrorQueja(string code, string mes, string año)
        {
            try
            {
                List<ErrorQueja> errores = new List<ErrorQueja>();

                if (code.Equals("1"))
                {
                    errores = ObtenerErroresQuejas("Trade", mes,año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerErroresQuejas("Cartera",mes, año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerErroresQuejas("Balanza",mes, año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerErroresQuejas("Compraventa",mes, año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerErroresQuejas("Swift",mes, año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerErroresQuejas("Todas",mes, año);
                }

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public ActionResult ObtenerErrorMes(string code, string año)
        {
            try
            {
                List<ErrorMes> errores = new List<ErrorMes>();

                
                if (code.Equals("1"))
                {
                    errores = ObtenerErroresGenerales("Trade", año);
                }
                else if (code.Equals("2"))
                {
                    errores = ObtenerErroresGenerales("Cartera", año);
                }
                else if (code.Equals("3"))
                {
                    errores = ObtenerErroresGenerales("Balanza", año);
                }
                else if (code.Equals("4"))
                {
                    errores = ObtenerErroresGenerales("Compraventa", año);
                }
                else if (code.Equals("6"))
                {
                    errores = ObtenerErroresGenerales("Swift", año);
                }
                else if (code.Equals("7"))
                {
                    errores = ObtenerErroresGenerales("Todas", año);
                }
                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (errores == null || errores.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(errores);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Json(new { error = "Error al obtener los reprocesos" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public ActionResult ObtenerUsuario()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<Usuarios> usuarios = ObtenerUsuariosDesdeLaClase();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (usuarios == null || usuarios.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(usuarios);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }
        }


        [HttpPost]
        public ActionResult InsertarUsuarios(
      string usuario,
      string nombre,
      string correo,
      string perfil)
        {
            try
            {
                List<Usuarios> usuarios = ObtenerUsuariosDesdeLaClase();
                AccessDatabase db = new AccessDatabase();

                // Verificar si el usuario ya existe en la lista
                bool usuarioExiste = usuarios.Any(u => u.Usuario == usuario);

                if (usuarioExiste)
                {
                    ViewBag.Mensaje = "Usuario ya existe";
                    return Json(new { success = false });
                }
                else
                {
                    ViewBag.Mensaje = "Usuario editado correctamente";
                    db.InsertarUsuario(usuario, nombre, correo, perfil);
                }
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return Json(new { success = true });
        }

        [HttpPost]
        public ActionResult InsertarCausas(
     string area,
     string tipoerror,
     string causa)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();

                ViewBag.Mensaje = "Usuario editado correctamente";
                db.InsertarCausa(area,tipoerror, causa);
               
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return Json(new { success = true });
        }



        [HttpGet]
        public ActionResult GetTipoErrores(string area)
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<TipoError> causas = causa.GetTipos(area);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }


        [HttpGet]
        public ActionResult GetCausaErrores(string tipoerror, string area)
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetTiposCausas(tipoerror, area);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }

        [HttpGet]
        public ActionResult GetCausa()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<CausasTotal> causas = causa.GetCausas();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }

        [HttpGet]
        public ActionResult GetCausaErroresFecha(string area)
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetTiposCausasArea(area);

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }


        [HttpGet]
        public ActionResult GetMoneda()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Moneda> causas = causa.GetMonedas();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }

        [HttpGet]
        public ActionResult GetAño()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Años> causas = causa.GetAños();

                // Si no se encontraron reprocesos, devolver un JSON vacío
                if (causas == null || causas.Count == 0)
                {
                    return Json(null, JsonRequestBehavior.AllowGet);
                }

                // Convertir la lista de reprocesos a un formato JSON
                string jsonReprocesos = JsonConvert.SerializeObject(causas);

                // Devolver el JSON como resultado
                return Content(jsonReprocesos, "application/json");
            }
            catch (Exception ex)
            {
                // Manejar la excepción y devolver un mensaje de error si es necesario
                Console.WriteLine("Excepción al obtener los reprocesos: " + ex.Message);
                return Content("Error al obtener los reprocesos");
            }


        }

    }
}
