using Newtonsoft.Json;
using Qualitas.Models;
using System;
using System.Collections.Generic;
using System.Data;
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
            else if (con == "Administrador" || con == "IngresoConsulta")
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
            return View();
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
            else if (con == "Administrador" || con == "IngresoConsulta")
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

            return View("Consultas");
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

            return View("Administrador");
        }

        [HttpPost]
        public ActionResult EnviarCorreos()
        {
            try
            {
                List<ErrorUsuario> errores = ObtenerUsuariosError(); // Asumiendo que esta función devuelve la lista de usuarios con errores

                if (errores != null && errores.Count > 0)
                {
                    Correos mail = new Correos();

                    foreach (var errorUsuario in errores)
                    {
                        // Asumiendo que ErrorUsuario tiene propiedades para el correo y el nombre del usuario
                        string correo = errorUsuario.Usuario+"@bancolombia.com.co";
                        string nombre = errorUsuario.Usuario;
                        string idError = errorUsuario.Errores.ToString();

                        mail.EnviarCorreo(correo, nombre, idError);
                    }

                    ViewBag.Mensaje = "Correos enviados correctamente";
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

            return View("Index");
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
           string Mes)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.InsertarRepro(FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes);

                ViewBag.Mensaje = "Reproceso registrado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el reproceso: " + ex.Message;
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
           string Mes)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EditarReproceso(id, FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes);

                ViewBag.Mensaje = "Reproceso registrado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el reproceso: " + ex.Message;
            }

            return View("Consultas");
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
           string perfil,
           string nombre,
           string correo)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.EditarUsuario(id, usuario, perfil, nombre, correo);

                ViewBag.Mensaje = "Usuario editado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return View("Consultas");
        }

        [HttpPost]
        public ActionResult InsertarUsuarios(
      string usuario,
      string perfil,
      string nombre,
      string correo)
        {
            try
            {
                AccessDatabase db = new AccessDatabase();
                db.InsertarUsuario(usuario, perfil, nombre, correo);

                ViewBag.Mensaje = "Usuario editado correctamente";
            }
            catch (Exception ex)
            {
                ViewBag.Mensaje = "Error al registrar el usuario: " + ex.Message;
            }

            return View("Consultas");
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

       

        private List<Reproceso> ObtenerReprocesosPerdidasDesdeLaClase()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDB = new AccessDatabase();
            return reprocesoDB.ObtenerReprocesosPerdidaEconomica();
        }

        private List<Reproceso> ObtenerReprocesosNoPerdidasDesdeLaClase()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase reprocesoDB = new AccessDatabase();
            return reprocesoDB.ObtenerReprocesosPerdidaNoEconomica();
        }

        [HttpGet]
        public ActionResult ObtenerReprocesosPerdidas()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<Reproceso> reprocesos = ObtenerReprocesosPerdidasDesdeLaClase();

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


            foreach (var reproceso in reprocesos)
            {

                myexcelWorksheet.Cells[rowunus, 1].Value = "Bancolombia";
                myexcelWorksheet.Cells[rowunus, 2].Value = reproceso.FechaOp;
                myexcelWorksheet.Cells[rowunus, 3].Value = reproceso.FechaReg;
                myexcelWorksheet.Cells[rowunus, 4].Value = " ";
                myexcelWorksheet.Cells[rowunus, 5].Value = " ";
                myexcelWorksheet.Cells[rowunus, 6].Value = "Gerencia Servicios Comercio Internacional";
                myexcelWorksheet.Cells[rowunus, 7].Value = reproceso.Nit;
                myexcelWorksheet.Cells[rowunus, 8].Value = reproceso.Descripcion;
                myexcelWorksheet.Cells[rowunus, 9].Value = " ";
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
                myexcelWorksheet.Cells[rowunus, 20].Value =reproceso.Responsable;
            }

            string informe = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Informes Recepción y Envio\Perdidas económicas y no económicas\" + "Informe de pérdidas económicas.xls";
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
            string[] columnNames = { "Id", "FechaOp", "FechaReg", "Consecutivo", "Nit", "Moneda", "Valor", "Cliente", "ProdEvento", "Responsable", "UsuarioReproceso", "Perdida", "Impacto", "Causa", "Descripcion", "Area", "TipoError", "Mes"};

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
                myexcelWorksheet.Cells[rowunus, 18].Value = reproceso.TipoError;

                rowunus++; 
            }

            string informe = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Informes Recepción y Envio\Perdidas económicas y no económicas\" + "Informe Calidad.xls";
            myexcelApplication.ActiveWorkbook.SaveAs(informe, Excel.XlFileFormat.xlWorkbookNormal);
            myexcelWorkbook.Close();
            myexcelApplication.Quit();

            return Json(new { success = true });
        }


        [HttpGet]
        public ActionResult ObtenerReprocesosNoPerdidas()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<Reproceso> reprocesos = ObtenerReprocesosNoPerdidasDesdeLaClase();

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
        private List<ErrorUsuario> ObtenerUsuariosError()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoUsuarios();
        }

        private List<ErrorArea> ObtenerAreaError()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoArea();
        }


        private List<ErrorTipo> ObtenerTipoError()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoError();
        }

        private List<ErrorTipo> ObtenerTipoErrorSwift()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Swift");
        }

        private List<ErrorTipo> ObtenerTipoErrorTrade()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Trade");
        }
        private List<ErrorTipo> ObtenerTipoErrorCartera()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Cartera");
        }
        private List<ErrorTipo> ObtenerTipoErrorCV()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Par comercial - Compraventa");
        }

        private List<ErrorTipo> ObtenerTipoErrorOCV()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Otros segmentos - Compraventa");
        }

        private List<ErrorTipo> ObtenerTipoErrorBalanza()
        {
            // Aquí llamas al método ObtenerReprocesos de tu clase ReprocesoDB
            AccessDatabase error = new AccessDatabase();
            return error.ObtenerReprocesoTipoErrorArea("Balanza");
        } 

      
        [HttpGet]
        public ActionResult ObtenerError()
        {
            try
            {
                // Aquí obtienes la lista de reprocesos desde tu clase aparte
                List<ErrorUsuario> errores = ObtenerUsuariosError();

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
        public ActionResult ObtenerErrorTipo(string code)
        {
            try
            {
                List<ErrorTipo> errores = new List<ErrorTipo>();
                if (code.Equals("1"))
                {
                    //Todas
                    errores = ObtenerTipoErrorTrade();
                }
                else if (code.Equals("2"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorCartera();
                }
                else if (code.Equals("3"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorBalanza();
                }
                else if (code.Equals("4"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorCV();
                }
                else if (code.Equals("5"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorOCV();
                }
                else if (code.Equals("6"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorSwift();
                }
                else if (code.Equals("7"))
                {
                    //Cartera
                    errores = ObtenerTipoError();
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
        public ActionResult ObtenerErrorTipoImp(string code)
        {
            try
            {
                List<ErrorTipo> errores = new List<ErrorTipo>();
                if (code.Equals("1"))
                {
                    //Todas
                    errores = ObtenerTipoErrorTrade();
                }
                else if (code.Equals("2"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorCartera();
                }
                else if (code.Equals("3"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorBalanza();
                }
                else if (code.Equals("4"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorCV();
                }
                else if (code.Equals("5"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorOCV();
                }
                else if (code.Equals("6"))
                {
                    //Cartera
                    errores = ObtenerTipoErrorSwift();
                }
                else if (code.Equals("7"))
                {
                    //Cartera
                    errores = ObtenerTipoError();
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

        [HttpGet]
        public ActionResult GetCausasVerificacion()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetCausasVeri();

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
        public ActionResult GetCausasAsignacion()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetCausasAigna();

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
        public ActionResult GetCausasCumplimiento()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetCausasCumpli();

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
        public ActionResult GetCausasOrientacion()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetCausasOrienta();

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
        public ActionResult GetCausasAprobacion()
        {
            try
            {
                AccessDatabase causa = new AccessDatabase();
                List<Causas> causas = causa.GetCausasAproba();

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
