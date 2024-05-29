using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.OleDb;
using System.IO;

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
            string query = "SELECT * FROM Usuarios WHERE Usuario = @usuario";

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
                        command.Parameters.AddWithValue("@usuario", usuario);
                        // Ejecutar la consulta y obtener un lector de datos
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            // Iterar sobre los resultados y mostrarlos en la consola
                            while (reader.Read())
                            {
                                string perfil = reader["Perfil"].ToString().Trim();

                                if (perfil[posicionUnus] == '1' && perfil[posicionDuo] == '1' && perfil[posicionTris] == '1' && perfil[posicionQuattuor] == '1')
                                {
                                    return "Administrador";
                                }
                                else if (perfil[posicionUnus] == '1' && perfil[posicionDuo] == '1' && perfil[posicionTris] == '0' && perfil[posicionQuattuor] == '1')
                                {
                                    return "IngresoConsulta";
                                }
                                else if (perfil[posicionUnus] == '0' && perfil[posicionDuo] == '1' && perfil[posicionTris] == '0' && perfil[posicionQuattuor] == '0')
                                {
                                    return "Consulta";
                                }
                                else if (perfil[posicionUnus] == '0' && perfil[posicionDuo] == '1' && perfil[posicionTris] == '0' && perfil[posicionQuattuor] == '1')
                                {
                                    return "Informe";
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
        string mes)
        {
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "INSERT INTO Reproceso (FechaOp, FechaReg, Consecutivo, Nit, Moneda, Valor, Cliente, ProdEvento, Responsable, UsuarioReproceso, Perdida, Impacto, Causa, Descripcion, Area, TipoError, Mes) " +
                       "VALUES (@FechaOp, @FechaReg, @Consecutivo, @Nit, @Moneda, @Valor, @Cliente, @ProdEvento, @Responsable, @UsuarioReproceso, @Perdida, @Impacto, @Causa, @Descripcion, @Area, @TipoError, @Mes)";
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

        public List<Reproceso> ObtenerReprocesosFechaArea(string fechaini, string fechafin, string area)
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE FechaOp BETWEEN @FechaIni AND @FechaFin AND Area = @Area";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.Parameters.AddWithValue("@FechaIni", fechaini);
                    command.Parameters.AddWithValue("@FechaFin", fechafin);
                    command.Parameters.AddWithValue("@Area", area);

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


        public List<ErrorUsuario> ObtenerReprocesoUsuarios()
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

        public List<ErrorTipo> ObtenerReprocesoTipoError()
        {
            List<ErrorTipo> erroresPorUsuario = new List<ErrorTipo>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso GROUP BY TipoError";

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

        public List<ImpactoError> ObtenerReprocesoTipoErrorImpacto(string Area)
        {
            List<ImpactoError> erroresPorUsuario = new List<ImpactoError>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT Impacto, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area GROUP BY Impacto";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Area", Area);
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
        public List<ErrorTipo> ObtenerReprocesoTipoErrorArea(string Area)
        {
            List<ErrorTipo> erroresPorUsuario = new List<ErrorTipo>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT TipoError, COUNT(*) AS TotalErrores FROM Reproceso WHERE Area = @Area GROUP BY TipoError";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Area", Area);
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

        public void EditarReproceso(
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
            string mes)
        {
            
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "UPDATE Reproceso SET FechaOp = @FechaOp, FechaReg = @FechaReg, Consecutivo = @Consecutivo, Nit = @Nit, Moneda = @Moneda, " +
                           "Valor = @Valor, Cliente = @Cliente, ProdEvento = @ProdEvento, Responsable = @Responsable, UsuarioReproceso = @UsuarioReproceso, " +
                           "Perdida = @Perdida, Impacto = @Impacto, Causa = @Causa, Descripcion = @Descripcion, Area = @Area, TipoError = @TipoError, Mes = @Mes  " +
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
                        command.Parameters.AddWithValue("@Id", id);
                        command.Parameters.AddWithValue("@Mes", mes);

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
        public List<Causas> GetCausasVeri()
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Causas WHERE tipoerror ='Verificacion'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new Causas
                        {
                            Causa = reader["causas"].ToString(),
                            TipoError = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }
        public List<Causas> GetCausasAigna()
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Causas WHERE tipoerror ='Asignacion'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new Causas
                        {
                            Causa = reader["causas"].ToString(),
                            TipoError = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }
        public List<Causas> GetCausasAproba()
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Causas WHERE tipoerror ='Aprobacion'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new Causas
                        {
                            Causa = reader["causas"].ToString(),
                            TipoError = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }

        public List<Causas> GetCausasCumpli()
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Causas WHERE tipoerror ='Cumplimiento'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new Causas
                        {
                            Causa = reader["causas"].ToString(),
                            TipoError = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }

        public List<Causas> GetCausasOrienta()
        {
            List<Causas> causas = new List<Causas>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Causas WHERE tipoerror ='Orientacion'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);
                connection.Open();
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        causas.Add(new Causas
                        {
                            Causa = reader["causas"].ToString(),
                            TipoError = reader["tipoerror"].ToString()
                        });
                    }
                }
            }


            return causas;
        }


        public List<Reproceso> ObtenerReprocesosPerdidaEconomica()
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE Perdida ='Si'";

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

        public List<Reproceso> ObtenerReprocesosPerdidaNoEconomica()
        {
            List<Reproceso> reprocesos = new List<Reproceso>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Reprocesos.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Reproceso WHERE Perdida ='No'";

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
            string perfil,
            string nombre,
            string correo)
        {

            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";


            // Consulta SQL para la actualización
            string query = "UPDATE Usuarios SET Usuario = @Usuario, Nombre = @Nombre, Correo = @Correo, Perfil = @Perfil "+
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
                        command.Parameters.AddWithValue("@Perfil", perfil);
                        command.Parameters.AddWithValue("@Nombre", nombre);
                        command.Parameters.AddWithValue("@Correo", correo);
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
        public List<Usuarios> ObtenerUsuarios()
        {
            List<Usuarios> usuarios = new List<Usuarios>();
            string dbFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
            string databasePath = Path.Combine(dbFolderPath, "Administracion.accdb");
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";

            string query = "SELECT * FROM Usuarios";
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


