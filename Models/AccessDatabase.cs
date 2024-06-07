using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.OleDb;
using System.IO;
using System.Data;

namespace Qualitas.Models
{
    public class AccessDatabase
    {

        public string Inicio(string usuario)
        { 
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");

            // Combinar la ruta de la carpeta con el nombre de la base de datos
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");

            // Cadena de conexión a la base de datos Access usando la ruta relativa
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para obtener todos los registros de una tabla
            string query = "SELECT * FROM Persona WHERE Usuario = @Usuario";

            int posicionUnus = 0;
            int posicionDuo = 1;
            int posicionTris = 2;
            int posicionQuattuor = 3;

            try
            {
                // Crear una conexión a la base de datos
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    // Abrir la conexión
                    connection.Open();

                    // Crear un comando SQL
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Usuario", usuario);
                        // Ejecutar la consulta y obtener un lector de datos
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            // Iterar sobre los resultados y mostrarlos en la consola
                            while (reader.Read())
                            {
                                string perfil = reader["Perfil"].ToString().Trim();

                                if (perfil.Equals("Administrador"))
                                {
                                    return "Administrador";
                                }
                                else if (perfil.Equals("Consultas"))
                                {
                                    return "Consultas";
                                }
                               
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Registrar la excepción en un archivo de registro
                Console.WriteLine("Error al conectar a la base de datos", ex);

                // Devolver un mensaje de error significativo al usuario

            }
            catch (Exception ex)
            {
                // Manejar otras excepciones no específicas de la base de datos
                // Registrar la excepción en un archivo de registro
                Console.WriteLine("Error desconocido", ex);

                // Devolver un mensaje de error genérico al usuario
            }
            return "UsuarioDesconocido";
        }

        public void InsertarRepro(
        string fechaOp,
        string fechaReg,
        string consecutivo,
        string nit,
        string moneda,
        decimal valor,
        string cliente,
        string prodEvento,
        string responsable,
        string usuarioReproceso,
        string perdida,
        string impacto,
        string causa,
        string descripcion,
        string area,
        string tipoError,
        string mes,
        string año,
        string queja)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "INSERT INTO Reproceso (FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes, Año, QuejaCliente) " +
                       "VALUES (@FechaOp, @FechaReg, @Consecutivo, @Nit, @Moneda, @Valor, @Cliente, @ProdEvento, @Responsable, @UsuarioReproceso, @Perdida, @Impacto, @Causa, @Descripcion, @Area, @TipoError, @Mes, @Año, @QuejaCliente)";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Agregar parámetros para evitar la inyección SQL
                        command.Parameters.AddWithValue("@FechaOp", fechaOp);
                        command.Parameters.AddWithValue("@FechaReg", fechaReg);
                        command.Parameters.AddWithValue("@Consecutivo", consecutivo);
                        command.Parameters.AddWithValue("@Nit", nit);
                        command.Parameters.AddWithValue("@Moneda", moneda);
                        command.Parameters.AddWithValue("@Valor", valor);
                        command.Parameters.AddWithValue("@Cliente", cliente);
                        command.Parameters.AddWithValue("@ProdEvento", prodEvento);
                        command.Parameters.AddWithValue("@Responsable", responsable);
                        command.Parameters.AddWithValue("@UsuarioReproceso", usuarioReproceso);
                        command.Parameters.AddWithValue("@Perdida", perdida);
                        command.Parameters.AddWithValue("@Impacto", impacto);
                        command.Parameters.AddWithValue("@Causa", causa);
                        command.Parameters.AddWithValue("@Descripcion", descripcion);
                        command.Parameters.AddWithValue("@Area", area);
                        command.Parameters.AddWithValue("@TipoError", tipoError);
                        command.Parameters.AddWithValue("@Mes", mes);
                        command.Parameters.AddWithValue("@Año", año);
                        command.Parameters.AddWithValue("@QuejaCliente", queja);

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

        public void EliminarRepro(int id)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string query = "DELETE FROM Reproceso WHERE Id = @ID"; 

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID", id);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

        public void EliminarCausas(int id)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string query = "DELETE FROM Causas WHERE Id = @ID";

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID", id);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
        public void EliminarUsu(int id)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string query = "DELETE FROM Usuarios WHERE Id = @ID";

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ID", id);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
        public List<Reproceso> ObtenerReprocesos()
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Reproceso reproceso = new Reproceso();
                                reproceso.Id = Convert.ToInt32(reader["Id"]);
                                reproceso.FechaOp = reader["FechaOp"] == DBNull.Value ? null : reader["FechaOp"].ToString();
                                reproceso.FechaReg = reader["FechaReg"] == DBNull.Value ? null : reader["FechaReg"].ToString();
                                reproceso.Consecutivo = reader["Consecutivo"] == DBNull.Value ? null : reader["Consecutivo"].ToString();
                                reproceso.Nit = reader["Nit"] == DBNull.Value ? null : reader["Nit"].ToString();
                                reproceso.Moneda = reader["Moneda"] == DBNull.Value ? null : reader["Moneda"].ToString();
                                reproceso.Valor = reader["Valor"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["Valor"]);
                                reproceso.Cliente = reader["Cliente"] == DBNull.Value ? null : reader["Cliente"].ToString();
                                reproceso.ProdEvento = reader["ProdEvento"] == DBNull.Value ? null : reader["ProdEvento"].ToString();
                                reproceso.Responsable = reader["Responsable"] == DBNull.Value ? null : reader["Responsable"].ToString();
                                reproceso.UsuarioReproceso = reader["UsuarioReproceso"] == DBNull.Value ? null : reader["UsuarioReproceso"].ToString();
                                reproceso.Perdida = reader["Perdida"] == DBNull.Value ? null : reader["Perdida"].ToString();
                                reproceso.Impacto = reader["Impacto"] == DBNull.Value ? null : reader["Impacto"].ToString();
                                reproceso.Causa = reader["Causa"] == DBNull.Value ? null : reader["Causa"].ToString();
                                reproceso.Descripcion = reader["Descripcion"] == DBNull.Value ? null : reader["Descripcion"].ToString();
                                reproceso.Area = reader["Area"] == DBNull.Value ? null : reader["Area"].ToString();
                                reproceso.TipoError = reader["TipoError"] == DBNull.Value ? null : reader["TipoError"].ToString();
                                reproceso.Mes = reader["Mes"] == DBNull.Value ? null : reader["Mes"].ToString();
                                reproceso.Año = reader["Año"] == DBNull.Value ? null : reader["Año"].ToString();
                                reproceso.Queja = reader["QuejaCliente"] == DBNull.Value ? null : reader["QuejaCliente"].ToString();
                                reprocesos.Add(reproceso);
                            }
                        }
                    }
                }

            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return reprocesos;
        }

        public List<Reproceso> ObtenerReprocesosInd(string usuario)
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE Responsable = @Responsable";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Responsable", usuario);
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Reproceso reproceso = new Reproceso();
                                reproceso.Id = Convert.ToInt32(reader["Id"]);
                                reproceso.FechaOp = reader["FechaOp"] == DBNull.Value ? null : reader["FechaOp"].ToString();
                                reproceso.FechaReg = reader["FechaReg"] == DBNull.Value ? null : reader["FechaReg"].ToString();
                                reproceso.Consecutivo = reader["Consecutivo"] == DBNull.Value ? null : reader["Consecutivo"].ToString();
                                reproceso.Nit = reader["Nit"] == DBNull.Value ? null : reader["Nit"].ToString();
                                reproceso.Moneda = reader["Moneda"] == DBNull.Value ? null : reader["Moneda"].ToString();
                                reproceso.Valor = reader["Valor"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["Valor"]);
                                reproceso.Cliente = reader["Cliente"] == DBNull.Value ? null : reader["Cliente"].ToString();
                                reproceso.ProdEvento = reader["ProdEvento"] == DBNull.Value ? null : reader["ProdEvento"].ToString();
                                reproceso.Responsable = reader["Responsable"] == DBNull.Value ? null : reader["Responsable"].ToString();
                                reproceso.UsuarioReproceso = reader["UsuarioReproceso"] == DBNull.Value ? null : reader["UsuarioReproceso"].ToString();
                                reproceso.Perdida = reader["Perdida"] == DBNull.Value ? null : reader["Perdida"].ToString();
                                reproceso.Impacto = reader["Impacto"] == DBNull.Value ? null : reader["Impacto"].ToString();
                                reproceso.Causa = reader["Causa"] == DBNull.Value ? null : reader["Causa"].ToString();
                                reproceso.Descripcion = reader["Descripcion"] == DBNull.Value ? null : reader["Descripcion"].ToString();
                                reproceso.Area = reader["Area"] == DBNull.Value ? null : reader["Area"].ToString();
                                reproceso.TipoError = reader["TipoError"] == DBNull.Value ? null : reader["TipoError"].ToString();
                                reproceso.Mes = reader["Mes"] == DBNull.Value ? null : reader["Mes"].ToString();
                                reproceso.Año = reader["Año"] == DBNull.Value ? null : reader["Año"].ToString();
                                reproceso.Queja = reader["QuejaCliente"] == DBNull.Value ? null : reader["QuejaCliente"].ToString();
                                reprocesos.Add(reproceso);
                            }
                        }
                    }
                }

            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return reprocesos;
        }

        public List<Reproceso> ObtenerReprocesosFechaAreaInd(string fechaini, string fechafin, string usuario)
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE Responsable = @Responsable FechaOp BETWEEN @FechaIni AND @FechaFin";
          
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);

                   
                    command.Parameters.AddWithValue("@FechaIni", fechaini);
                    command.Parameters.AddWithValue("@FechaFin", fechafin);
                    command.Parameters.AddWithValue("@Responsable", usuario);


                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Reproceso reproceso = new Reproceso();
                            reproceso.Id = Convert.ToInt32(reader["Id"]);
                            reproceso.FechaOp = reader["FechaOp"].ToString();
                            reproceso.FechaReg = reader["FechaReg"].ToString();
                            reproceso.Consecutivo = reader["Consecutivo"].ToString();
                            reproceso.Nit = reader["Nit"].ToString();
                            reproceso.Moneda = reader["Moneda"].ToString();
                            reproceso.Valor = Convert.ToDecimal(reader["Valor"]);
                            reproceso.Cliente = reader["Cliente"].ToString();
                            reproceso.ProdEvento = reader["ProdEvento"].ToString();
                            reproceso.Responsable = reader["Responsable"].ToString();
                            reproceso.UsuarioReproceso = reader["UsuarioReproceso"].ToString();
                            reproceso.Perdida = reader["Perdida"].ToString();
                            reproceso.Impacto = reader["Impacto"].ToString();
                            reproceso.Causa = reader["Causa"].ToString();
                            reproceso.Descripcion = reader["Descripcion"].ToString();
                            reproceso.Area = reader["Area"].ToString();
                            reproceso.TipoError = reader["TipoError"].ToString();
                            reproceso.Mes = reader["Mes"].ToString();
                            reproceso.Año = reader["Año"].ToString();
                            reproceso.Queja = reader["QuejaCliente"].ToString();
                            reprocesos.Add(reproceso);
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos: " + ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos: " + ex.Message);
                throw;
            }

            return reprocesos;
        }

        public List<Reproceso> ObtenerReprocesosFechaArea(string fechaini, string fechafin, string area)
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (area == "Todas")
            {
                query = "SELECT * FROM Reproceso WHERE FechaOp BETWEEN @FechaIni AND @FechaFin";
            }
            else
            {
                query = "SELECT * FROM Reproceso WHERE FechaOp BETWEEN @FechaIni AND @FechaFin AND Area = @Area";
            }
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);

                    if (area == "Todas")
                    {
                        command.Parameters.AddWithValue("@FechaIni", fechaini);
                        command.Parameters.AddWithValue("@FechaFin", fechafin);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@FechaIni", fechaini);
                        command.Parameters.AddWithValue("@FechaFin", fechafin);
                        command.Parameters.AddWithValue("@Area", area);
                    }
                  

                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Reproceso reproceso = new Reproceso();
                            reproceso.Id = Convert.ToInt32(reader["Id"]);
                            reproceso.FechaOp = reader["FechaOp"].ToString();
                            reproceso.FechaReg = reader["FechaReg"].ToString();
                            reproceso.Consecutivo = reader["Consecutivo"].ToString();
                            reproceso.Nit = reader["Nit"].ToString();
                            reproceso.Moneda = reader["Moneda"].ToString();
                            reproceso.Valor = Convert.ToDecimal(reader["Valor"]);
                            reproceso.Cliente = reader["Cliente"].ToString();
                            reproceso.ProdEvento = reader["ProdEvento"].ToString();
                            reproceso.Responsable = reader["Responsable"].ToString();
                            reproceso.UsuarioReproceso = reader["UsuarioReproceso"].ToString();
                            reproceso.Perdida = reader["Perdida"].ToString();
                            reproceso.Impacto = reader["Impacto"].ToString();
                            reproceso.Causa = reader["Causa"].ToString();
                            reproceso.Descripcion = reader["Descripcion"].ToString();
                            reproceso.Area = reader["Area"].ToString();
                            reproceso.TipoError = reader["TipoError"].ToString();
                            reproceso.Mes = reader["Mes"].ToString();
                            reproceso.Año = reader["Año"].ToString();
                            reproceso.Queja = reader["QuejaCliente"].ToString();
                            reprocesos.Add(reproceso);
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos: " + ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos: " + ex.Message);
                throw;
            }

            return reprocesos;
        }


        public List<ErrorUsuario> ObtenerReprocesoUsuarios(string fechaInicio, string fechaFin)
        {
            List<ErrorUsuario> erroresPorUsuario = new List<ErrorUsuario>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT Responsable, COUNT(*) AS TotalErrores FROM Reproceso WHERE FechaReg BETWEEN @fechaInicio AND @fechaFin GROUP BY Responsable";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@fechaInicio", fechaInicio);
                        command.Parameters.AddWithValue("@fechaFin", fechaFin);

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorUsuario error = new ErrorUsuario();
                                error.Usuario = reader["Responsable"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 16 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }


        public List<ReprocesoUsuario> ObtenerReprocesoUsuariosSemana(string fechaInicio, string fechaFin, string usuario)
        {
            List<ReprocesoUsuario> reprocesos = new List<ReprocesoUsuario>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE FechaReg BETWEEN @fechaInicio AND @fechaFin AND Responsable = @Responsable";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@fechaInicio", fechaInicio);
                        command.Parameters.AddWithValue("@fechaFin", fechaFin);
                        command.Parameters.AddWithValue("@Responsable", usuario);

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ReprocesoUsuario reproceso = new ReprocesoUsuario();
                                reproceso.Id = Convert.ToInt32(reader["Id"]);
                                reproceso.FechaOp = reader["FechaOp"].ToString();
                                reproceso.FechaReg = reader["FechaReg"].ToString();
                                reproceso.Consecutivo = reader["Consecutivo"].ToString();
                                reproceso.Nit = reader["Nit"].ToString();
                                reproceso.Moneda = reader["Moneda"].ToString();
                                reproceso.Valor = Convert.ToDecimal(reader["Valor"]);
                                reproceso.Cliente = reader["Cliente"].ToString();
                                reproceso.ProdEvento = reader["ProdEvento"].ToString();
                                reproceso.Responsable = reader["Responsable"].ToString();
                                reproceso.UsuarioReproceso = reader["UsuarioReproceso"].ToString();
                                reproceso.Perdida = reader["Perdida"].ToString();
                                reproceso.Impacto = reader["Impacto"].ToString();
                                reproceso.Causa = reader["Causa"].ToString();
                                reproceso.Descripcion = reader["Descripcion"].ToString();
                                reproceso.Area = reader["Area"].ToString();
                                reproceso.TipoError = reader["TipoError"].ToString();
                                reproceso.Mes = reader["Mes"].ToString();
                                reproceso.Año = reader["Año"].ToString();
                                reproceso.Queja = reader["QuejaCliente"].ToString();
                                reprocesos.Add(reproceso);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return reprocesos;
        }

        public List<ErrorUsuario> ObtenerReprocesoUsuariosTodos()
        {
            List<ErrorUsuario> erroresPorUsuario = new List<ErrorUsuario>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT Responsable, COUNT(*) AS TotalErrores FROM Reproceso GROUP BY Responsable";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorUsuario error = new ErrorUsuario();
                                error.Usuario = reader["Responsable"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 16 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorUsuario> ObtenerReprocesoUsuariosArea(string fechaInicio, string fechaFin, string areas)
        {
            List<ErrorUsuario> erroresPorUsuario = new List<ErrorUsuario>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT Responsable, COUNT(*) AS TotalErrores FROM Reproceso WHERE FechaOp BETWEEN @fechaInicio AND @fechaFin AND Area = @area GROUP BY Responsable";


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@fechaInicio", fechaInicio);
                        command.Parameters.AddWithValue("@fechaFin", fechaFin);
                        command.Parameters.AddWithValue("@area", areas);

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorUsuario error = new ErrorUsuario();
                                error.Usuario = reader["Responsable"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }


        public List<ErrorArea> ObtenerReprocesoArea()
        {
            List<ErrorArea> erroresPorUsuario = new List<ErrorArea>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT Area, COUNT(*) AS TotalErrores FROM Reproceso GROUP BY Area";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorArea error = new ErrorArea();
                                error.Area = reader["Area"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                error.Fecha = reader["FechaOp"].ToString();
                                error.Mes = reader["Mes"].ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorTipo> ObtenerReprocesoTipoError(string mes, string año)
        {
            List<ErrorTipo> erroresPorUsuario = new List<ErrorTipo>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (mes == "Todos")
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY TipoError";
            }
            else
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Mes = @Mes AND Año = @Año GROUP BY TipoError";
            }

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                      
                        if (mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Mes", mes);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        command.Parameters.AddWithValue("@Año", año);
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorTipo error = new ErrorTipo();
                                error.Tipo = reader["TipoError"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString(); 
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ImpactoError> ObtenerReprocesoTipoErrorImpacto(string Area, string mes, string año)
        {
            List<ImpactoError> erroresPorUsuario = new List<ImpactoError>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;


            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT Impacto, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY Impacto";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT Impacto, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY Impacto";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT Impacto, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año and Mes = @Mes GROUP BY Impacto";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT Impacto, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año AND Mes = @Mes GROUP BY Impacto";
            }
           

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {

                        if (Area == "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                       
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ImpactoError error = new ImpactoError();
                                error.Impacto = reader["Impacto"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorCausa> ObtenerReprocesoTipoErrorCausas(string Area, string mes, string año)
        {
            List<ErrorCausa> erroresPorUsuario = new List<ErrorCausa>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;


            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT Causa, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY Causa";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT Causa, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY Causa";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT Causa, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año and Mes = @Mes GROUP BY Causa";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT Causa, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año AND Mes = @Mes GROUP BY Causa";
            }


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorCausa error = new ErrorCausa();
                                error.Causa = reader["Causa"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorPerdida> ObtenerReprocesoTipoErrorPerdidas(string Area, string mes, string año)
        {
            List<ErrorPerdida> erroresPorUsuario = new List<ErrorPerdida>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;


            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT Perdida, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY Perdida";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT Perdida, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY Perdida";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT Perdida, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año and Mes = @Mes GROUP BY Perdida";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT Perdida, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año AND Mes = @Mes GROUP BY Perdida";
            }


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorPerdida error = new ErrorPerdida();
                                error.Perdida = reader["Perdida"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorQueja> ObtenerReprocesoTipoErrorQuejas(string Area, string mes, string año)
        {
            List<ErrorQueja> erroresPorUsuario = new List<ErrorQueja>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT QuejaCliente, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY QuejaCliente";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT QuejaCliente, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY QuejaCliente";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT QuejaCliente, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año and Mes = @Mes GROUP BY QuejaCliente";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT QuejaCliente, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año AND Mes = @Mes GROUP BY QuejaCliente";
            }


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorQueja error = new ErrorQueja();
                                error.Queja = reader["QuejaCliente"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorTipo> ObtenerReprocesoTipoErrorArea(string Area, string mes, string año)
        {
            List<ErrorTipo> erroresPorUsuario = new List<ErrorTipo>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY TipoError";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY TipoError";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año and Mes = @Mes GROUP BY TipoError";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año AND Mes = @Mes GROUP BY TipoError";
            }

            
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Año", año);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }

                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorTipo error = new ErrorTipo();
                                error.Tipo = reader["TipoError"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }

        public List<ErrorMes> ObtenerReprocesoErroresArea(string Area, string año)
        {
            List<ErrorMes> erroresPorUsuario = new List<ErrorMes>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;


            if (Area == "Todas")
            {
                query = "SELECT Area,Mes, COUNT(*) AS TotalErrores FROM Reproceso WHERE Año = @Año GROUP BY Area, Mes";
            }
            else
            {
                query = "SELECT Area,Mes, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area AND Año = @Año GROUP BY Area, Mes";
            }
           

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area != "Todas")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                        }
                        command.Parameters.AddWithValue("@Año", año);
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ErrorMes error = new ErrorMes();
                                error.Mes = reader["Mes"].ToString();
                                error.Area = reader["Area"].ToString();
                                int totalErrores = Convert.ToInt32(reader["TotalErrores"]);
                                double porcentaje = (double)(totalErrores) / 15 * 100;
                                error.Errores = totalErrores.ToString();
                                error.Porcentaje = porcentaje.ToString();
                                erroresPorUsuario.Add(error);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return erroresPorUsuario;
        }


        public void EditarReprocesos(
            int id,
            string fechaOp,
            string fechaReg,
            string consecutivo,
            string nit,
            string moneda,
            decimal valor,
            string cliente,
            string prodEvento,
            string responsable,
            string usuarioReproceso,
            string perdida,
            string impacto,
            string causa,
            string descripcion,
            string area,
            string tipoError, 
            string mes,
            string año,
            string clientequeja)
        {
            
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "UPDATE Reproceso SET FechaOp = @FechaOp, FechaReg = @FechaReg, Consecutivo = @Consecutivo, Nit = @Nit, Moneda = @Moneda, " +
                           "Valor = @Valor, Cliente = @Cliente, ProdEvento = @ProdEvento, Responsable = @Responsable, UsuarioReproceso = @UsuarioReproceso, " +
                           "Perdida = @Perdida, Impacto = @Impacto, Causa = @Causa, Descripcion = @Descripcion, Area = @Area, TipoError = @TipoError, Mes = @Mes, Año = @Año, QuejaCliente = @QuejaCliente  " +
                           "WHERE Id = @Id";

            try
            {
                // Crear conexión y comando
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Asignar parámetros
                        command.Parameters.AddWithValue("@FechaOp", fechaOp);
                        command.Parameters.AddWithValue("@FechaReg", fechaReg);
                        command.Parameters.AddWithValue("@Consecutivo", consecutivo);
                        command.Parameters.AddWithValue("@Nit", nit);
                        command.Parameters.AddWithValue("@Moneda", moneda);
                        command.Parameters.AddWithValue("@Valor", valor);
                        command.Parameters.AddWithValue("@Cliente", cliente);
                        command.Parameters.AddWithValue("@ProdEvento", prodEvento);
                        command.Parameters.AddWithValue("@Responsable", responsable);
                        command.Parameters.AddWithValue("@UsuarioReproceso", usuarioReproceso);
                        command.Parameters.AddWithValue("@Perdida", perdida);
                        command.Parameters.AddWithValue("@Impacto", impacto);
                        command.Parameters.AddWithValue("@Causa", causa);
                        command.Parameters.AddWithValue("@Descripcion", descripcion);
                        command.Parameters.AddWithValue("@Area", area);
                        command.Parameters.AddWithValue("@TipoError", tipoError);
                        command.Parameters.AddWithValue("@Mes", mes);
                        command.Parameters.AddWithValue("@Año", año);
                        command.Parameters.AddWithValue("@QuejaCliente", clientequeja);
                        command.Parameters.AddWithValue("@Id", id);
                        // Ejecutar la consulta de actualización
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
        }
        public List<TipoError> GetTipos(string area)
        {
            List<TipoError> causas = new List<TipoError>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT DISTINCT tipoerror FROM Causas WHERE area = @area";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@area", area);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new TipoError
                        {
                            Tipo = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }

        public List<Causas> GetTiposCausas(string tipoerror, string area)
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            try
            {
                string query = "SELECT * FROM Causas WHERE tipoerror = @tipoerror AND area = @area";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.Parameters.AddWithValue("@tipoerror", tipoerror);
                    command.Parameters.AddWithValue("@area", area);
                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            causas.Add(new Causas
                            {
                                Causa = reader["causas"].ToString(),
                                Area = reader["area"].ToString()
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener tipos de causas: " + ex.Message);
                // Maneja el error según tus necesidades (registra, notifica, etc.)
            }

            return causas;
        }

        public List<CausasTotal> GetCausas()
        {
            List<CausasTotal> causas = new List<CausasTotal>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            try
            {
                string query = "SELECT * FROM Causas";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);
                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            causas.Add(new CausasTotal
                            {
                                Id = reader["Id"].ToString(),
                                Area = reader["area"].ToString(),
                                Tipo = reader["tipoerror"].ToString(),
                                Causa = reader["causas"].ToString()
                            }); ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener tipos de causas: " + ex.Message);
                // Maneja el error según tus necesidades (registra, notifica, etc.)
            }

            return causas;
        }


        public List<Causas> GetTiposCausasArea(string areas)
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;
            try
            {
                if (areas == "Todas")
                {
                    query = "SELECT DISTINCT Causa FROM Reproceso";
                }
                else if (areas != "Todas")
                {
                    query = "SELECT DISTINCT Causa FROM Reproceso WHERE Area = @Area";
                }
               
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);

                    if (areas != "Todas")
                    {
                        command.Parameters.AddWithValue("@Area", areas);
                    }
                    
                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            causas.Add(new Causas
                            {
                                Causa = reader["Causa"].ToString()
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener tipos de causas: " + ex.Message);
                // Maneja el error según tus necesidades (registra, notifica, etc.)
            }

            return causas;
        }

        public List<CausasTotal> GetTiposCausasUsuarios(string usuario)
        {
            List<CausasTotal> causas = new List<CausasTotal>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;
            try
            {
                
                query = "SELECT DISTINCT Causa FROM Reproceso WHERE Responsable = @Responsable";
                

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);

                 
                   command.Parameters.AddWithValue("@Responsable", usuario);
                   

                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            causas.Add(new CausasTotal
                            {
                                Causa = reader["Causa"].ToString()
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener tipos de causas: " + ex.Message);
                // Maneja el error según tus necesidades (registra, notifica, etc.)
            }

            return causas;
        }

        public DataTable GetProdEvento(string producto, string evento)
        {
            DataTable dataTable = new DataTable();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            try
            {
                string query = "SELECT * FROM Producto WHERE producto = @producto AND evento = @evento";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.Parameters.AddWithValue("@producto", producto);
                    command.Parameters.AddWithValue("@evento", evento);
                    connection.Open();

                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        dataTable.Load(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener productos de evento: " + ex.Message);
                // Maneja el error según tus necesidades (registra, notifica, etc.)
            }

            return dataTable;
        }


        public string GetCorreoConfirma(string fechaInicio, string fechaFin)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM CorreosEnvio WHERE fecha_env BETWEEN @FechaInicio AND @FechaFin ";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@fechaInicio", fechaInicio);
                command.Parameters.AddWithValue("@fechaFin", fechaFin);

                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int recordCount = 0;

                        while (reader.Read())
                        {
                            recordCount++;
                        }

                        if (recordCount > 0)
                        {
                            return "Enviado";
                        }
                    }
                    else
                    {
                        return "OK";
                    }
                }
            }

            return "Error";

        }

        public List<Moneda> GetMonedas()
        {
            List<Moneda> mon = new List<Moneda>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Moneda";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        mon.Add(new Moneda
                        {
                            Monedas = reader["moneda"].ToString(),
                            Ref = reader["etiqueta"].ToString()
                        });
                    }
                }
            }


            return mon;
        }

        public List<Años> GetAños()
        {
            List<Años> mon = new List<Años>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT DISTINCT Año FROM Reproceso";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        mon.Add(new Años
                        {
                            Año = reader["Año"].ToString()
                        });
                    }
                }
            }


            return mon;
        }


        public List<ReprocesoPerdida> ObtenerReprocesosPerdidaEconomica(string Area, string mes)
        {
            List<ReprocesoPerdida> reprocesos = new List<ReprocesoPerdida>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='Si' ";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='Si' AND Area = @Area";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='Si' AND Area = @Area AND Mes = @Mes";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='Si' AND Mes = @Mes";
            }


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {

                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Mes", mes);

                        }
                       using (OleDbDataReader reader = command.ExecuteReader())
                        {

                            while (reader.Read())
                            {
                                ReprocesoPerdida reproceso = new ReprocesoPerdida();
                                reproceso.Id = Convert.ToInt32(reader["Id"].ToString());
                                reproceso.FechaOp = reader["FechaOp"].ToString();
                                reproceso.FechaReg = reader["FechaReg"].ToString();
                                reproceso.Consecutivo = reader["Consecutivo"].ToString();
                                reproceso.Nit = reader["Nit"].ToString();
                                reproceso.Moneda = reader["Moneda"].ToString();
                                reproceso.Valor = Convert.ToDecimal(reader["Valor"]);
                                reproceso.Cliente = reader["Cliente"].ToString();
                                reproceso.ProdEvento = reader["ProdEvento"].ToString();
                                reproceso.Responsable = reader["Responsable"].ToString();
                                reproceso.UsuarioReproceso = reader["UsuarioReproceso"].ToString();
                                reproceso.Perdida = reader["Perdida"].ToString();
                                reproceso.Impacto = reader["Impacto"].ToString();
                                reproceso.Causa = reader["Causa"].ToString();
                                reproceso.Descripcion = reader["Descripcion"].ToString();
                                reproceso.Area = reader["Area"].ToString();
                                reproceso.TipoError = reader["TipoError"].ToString();
                                reproceso.Mes = reader["Mes"].ToString();
                                reproceso.Año = reader["Año"].ToString();
                                reproceso.Queja = reader["QuejaCliente"].ToString();
                                reprocesos.Add(reproceso);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return reprocesos;
        }

        public List<ReprocesoPerdida> ObtenerReprocesosPerdidaNoEconomica(string Area, string mes)
        {
            List<ReprocesoPerdida> reprocesos = new List<ReprocesoPerdida>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            string query = null;

            if (Area == "Todas" && mes == "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='No' ";
            }
            else if (Area != "Todas" && mes == "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='No' AND Area = @Area";
            }
            else if (Area != "Todas" && mes != "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='No' AND Area = @Area AND Mes = @Mes";
            }
            else if (Area == "Todas" && mes != "Todos")
            {
                query = "SELECT * FROM Reproceso WHERE Perdida ='No' AND Mes = @Mes";
            }


            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (Area == "Todas" && mes == "Todos")
                        {

                        }
                        else if (Area != "Todas" && mes == "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                        }
                        else if (Area != "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Area", Area);
                            command.Parameters.AddWithValue("@Mes", mes);
                        }
                        else if (Area == "Todas" && mes != "Todos")
                        {
                            command.Parameters.AddWithValue("@Mes", mes);

                        }
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ReprocesoPerdida reproceso = new ReprocesoPerdida();
                                reproceso.Id = Convert.ToInt32(reader["Id"].ToString());
                                reproceso.FechaOp = reader["FechaOp"].ToString();
                                reproceso.FechaReg = reader["FechaReg"].ToString();
                                reproceso.Consecutivo = reader["Consecutivo"].ToString();
                                reproceso.Nit = reader["Nit"].ToString();
                                reproceso.Moneda = reader["Moneda"].ToString();
                                reproceso.Valor = Convert.ToDecimal(reader["Valor"]);
                                reproceso.Cliente = reader["Cliente"].ToString();
                                reproceso.ProdEvento = reader["ProdEvento"].ToString();
                                reproceso.Responsable = reader["Responsable"].ToString();
                                reproceso.UsuarioReproceso = reader["UsuarioReproceso"].ToString();
                                reproceso.Perdida = reader["Perdida"].ToString();
                                reproceso.Impacto = reader["Impacto"].ToString();
                                reproceso.Causa = reader["Causa"].ToString();
                                reproceso.Descripcion = reader["Descripcion"].ToString();
                                reproceso.Area = reader["Area"].ToString();
                                reproceso.TipoError = reader["TipoError"].ToString();
                                reproceso.Mes = reader["Mes"].ToString();
                                reproceso.Año = reader["Año"].ToString();
                                reproceso.Queja = reader["QuejaCliente"].ToString();
                                reprocesos.Add(reproceso);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return reprocesos;
        }

        public void EditarUsuario(
            int id,
            string usuario,
            string nombre,
            string correo,
            string perfil)
        {

            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "UPDATE Persona SET Usuario = @Usuario, Nombre = @Nombre, Correo = @Correo, Perfil = @Perfil "+
                           "WHERE Id = @Id";

            try
            {
                // Crear conexión y comando
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Asignar parámetros

                        command.Parameters.AddWithValue("@Usuario", usuario);
                        command.Parameters.AddWithValue("@Nombre", nombre);
                        command.Parameters.AddWithValue("@Correo", correo);
                        command.Parameters.AddWithValue("@Perfil", perfil);
                        command.Parameters.AddWithValue("@Id", id);
                        // Ejecutar la consulta de actualización
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
        }

        public void EditarCausa(
            int id,
            string area,
            string tipoerror,
            string causas)
        {

            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "UPDATE Causas SET area = @area, tipoerror = @tipoerror, causas = @causas " +
                           "WHERE Id = @Id";

            try
            {
                // Crear conexión y comando
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Asignar parámetros

                        command.Parameters.AddWithValue("@area", area);
                        command.Parameters.AddWithValue("@tipoerror", tipoerror);
                        command.Parameters.AddWithValue("@causas", causas);
                        command.Parameters.AddWithValue("@Id", id);
                        // Ejecutar la consulta de actualización
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
        }


        public void InsertarUsuario(
            string usuario,
            string perfil,
            string nombre,
            string correo)
        {

            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "INSERT INTO Usuarios (Usuario, Perfil, Correo, Nombre) VALUES (@Usuario, @Perfil, @Correo, @Nombre)";
                   

            try
            {
                // Crear conexión y comando
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Asignar parámetros
                        command.Parameters.AddWithValue("@Usuario", usuario);
                        command.Parameters.AddWithValue("@Perfil", perfil);
                        command.Parameters.AddWithValue("@Nombre", nombre);
                        command.Parameters.AddWithValue("@Correo", correo);

                        // Ejecutar la consulta de actualización
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
        }

        public void InsertarCausa(
           string area,
           string tipoerror,
           string causas)
        {

            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "INSERT INTO Causas (area, tipoerror, causas) VALUES (@area, @tipoerror, @causas)";


            try
            {
                // Crear conexión y comando
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        // Asignar parámetros
                        command.Parameters.AddWithValue("@area", area);
                        command.Parameters.AddWithValue("@tipoerror", tipoerror);
                        command.Parameters.AddWithValue("@causas", causas);

                        // Ejecutar la consulta de actualización
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al actualizar el registro en la tabla reprocesos:", ex);
                throw;
            }
        }


        public List<Usuarios> ObtenerUsuarios()
        {
            List<Usuarios> usuarios = new List<Usuarios>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Persona";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Usuarios usuario = new Usuarios();
                                usuario.Usuario = reader["Usuario"].ToString();
                                usuario.Perfil = reader["Perfil"].ToString();
                                usuario.Correo = reader["Correo"].ToString();
                                usuario.Nombre = reader["Nombre"].ToString();
                                usuario.Id = Convert.ToInt32(reader["Id"].ToString());
                                usuarios.Add(usuario);
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                // Manejar excepciones de base de datos
                Console.WriteLine("Error al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones
                Console.WriteLine("Error desconocido al obtener los datos de la tabla reprocesos:", ex);
                throw;
            }

            return usuarios;
        }

    }
}


