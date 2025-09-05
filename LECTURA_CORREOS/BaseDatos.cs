using LECTURA_CORREOS.Clases;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace LECTURA_CORREOS
{
    public class BaseDatos
    {

        public static SqlConnection abrirConexionSQL(string servidor,string baseDatos,string usuario,string contrasena)
        {
            string connectionString = $"Data Source={servidor};Initial Catalog={baseDatos};User ID={usuario};Password={contrasena}";
            SqlConnection connection = null;
            try
            {
                connection = new SqlConnection(connectionString);
                connection.Open();
                Console.WriteLine("Conexión exitosa a la base de datos.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al conectar a la base de datos: " + ex.Message);
            }
            return connection;
        }

        // Función para cerrar la conexión a la base de datos
        public static void CloseSqlConnection(SqlConnection connection)
        {
            if (connection != null && connection.State == System.Data.ConnectionState.Open)
            {
                try
                {
                    connection.Close();
                    Console.WriteLine("Conexión cerrada exitosamente.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error al cerrar la conexión: " + ex.Message);
                }
            }
        }

        // Función para ejecutar un procedimiento almacenado
        static DataTable ExecuteStoredProcedure(SqlConnection connection, string storedProcedureName, SqlParameter[] parameters)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlCommand command = new SqlCommand(storedProcedureName, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // Agregar parámetros al comando
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }
                    
                    // Ejecutar el procedimiento almacenado y llenar el DataTable con los resultados
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al ejecutar el procedimiento almacenado: " + ex.Message);
            }
            return dataTable;
        }

        public static int ingresarCorreo(SqlConnection connection,string identificador,string titulo,string de,string fechaCorreo,string fechaInicioLectura,string fechaFinLectura)
        {
            int idCorreo = 0;
            string mensajeSalida = "";
            SqlParameter[] parameters = new SqlParameter[]
            {
                new SqlParameter("@identificador", SqlDbType.VarChar) { Value = identificador },
                new SqlParameter("@titulo", SqlDbType.VarChar) { Value = titulo },
                new SqlParameter("@de", SqlDbType.VarChar) { Value = de },
                new SqlParameter("@fechaCorreo", SqlDbType.VarChar) { Value = fechaCorreo },
                new SqlParameter("@fechaInicioLectura", SqlDbType.VarChar) { Value = fechaInicioLectura },
                new SqlParameter("@fechaFinLectura", SqlDbType.VarChar) { Value =fechaFinLectura }
            };
            DataTable result = ExecuteStoredProcedure(connection, "dbo.INGRESAR_DATOS_CORREOS", parameters);

            // Procesar resultados
            foreach (DataRow row in result.Rows)
            {
                idCorreo = int.Parse(row["idCorreo"].ToString());
                mensajeSalida = row["mensaje"].ToString();
            }
            if (mensajeSalida != "")
            {
                Console.WriteLine("ERROR AL INGRESAR CORREO A LA BASE DE DATOS " + mensajeSalida);
            }
            return idCorreo;
        }

        public static int ingresarAdjuntoCorreo(SqlConnection connection, int idCorreo, string rutaXML, string tipoDocRef, string tipoDTE, string folio, string fechEmis,string rutEmisor,string rutRecep,decimal mntTotal,decimal iva,decimal mntNeto,string folioRef,string razonRef,string tsTed, string RznSocRecep, string mntExe, string RutReceptor)
        {
            int idFactura = 0;
            string mensajeSalida = "";
            SqlParameter[] parameters = new SqlParameter[]
            {
                new SqlParameter("@idCorreo", SqlDbType.Int) { Value = idCorreo },
                new SqlParameter("@rutaXML", SqlDbType.VarChar) { Value = rutaXML },
                new SqlParameter("@tipoDocRef", SqlDbType.VarChar) { Value = tipoDocRef },
                new SqlParameter("@tipoDTE", SqlDbType.VarChar) { Value = tipoDTE },
                new SqlParameter("@folio", SqlDbType.VarChar) { Value = folio },
                new SqlParameter("@fechEmis", SqlDbType.VarChar) { Value =fechEmis },
                new SqlParameter("@rutEmisor", SqlDbType.VarChar) { Value =rutEmisor },
                new SqlParameter("@rutRecep", SqlDbType.VarChar) { Value =rutRecep },
                new SqlParameter("@mntTotal", SqlDbType.Decimal) { Value =mntTotal },
                new SqlParameter("@iva", SqlDbType.Decimal) { Value =iva },
                new SqlParameter("@mntNeto", SqlDbType.VarChar) { Value =mntNeto },
                new SqlParameter("@folioRef", SqlDbType.VarChar) { Value =folioRef },
                new SqlParameter("@razonRef", SqlDbType.VarChar) { Value =razonRef },
                new SqlParameter("@tsTed", SqlDbType.VarChar) { Value =tsTed },
                new SqlParameter("@RznSocRecep", SqlDbType.VarChar) { Value =RznSocRecep },
                new SqlParameter("@MntExe", SqlDbType.VarChar) { Value =mntExe },
                new SqlParameter("@rutReceptor", SqlDbType.VarChar) { Value =RutReceptor }

            };
            DataTable result = ExecuteStoredProcedure(connection, "dbo.INGRESAR_DATOS_XML", parameters);

            // Procesar resultados
            foreach (DataRow row in result.Rows)
            {
                idFactura = int.Parse(row["idFactura"].ToString());
                mensajeSalida = row["mensaje"].ToString();
            }
            if (mensajeSalida !="")
            {
                Console.WriteLine("ERROR AL INGRESAR ADJUNTO XML "+mensajeSalida);
            }
            return idFactura;
        }

        public static int ingresarReferencia(SqlConnection connection, int idFactura, string TpoDocRef, string FolioRef, string FchRef)
        {
            int idReferencia = 0;
            string mensajeSalida = "";
            SqlParameter[] parameters = new SqlParameter[]
            {
                new SqlParameter("@idCorreo", SqlDbType.Int) { Value = idFactura },
                new SqlParameter("@TpoDocRef", SqlDbType.VarChar) { Value = TpoDocRef },
                new SqlParameter("@FolioRef", SqlDbType.VarChar) { Value = FolioRef },
                new SqlParameter("@FchRef", SqlDbType.VarChar) { Value = FchRef }

            };
            DataTable result = ExecuteStoredProcedure(connection, "dbo.INGRESAR_REFERENCIAS", parameters);

            // Procesar resultados
            foreach (DataRow row in result.Rows)
            {
                idReferencia = int.Parse(row["idReferencia"].ToString());
                mensajeSalida = row["mensaje"].ToString();
            }
            if (mensajeSalida != "")
            {
                Console.WriteLine("ERROR AL INGRESAR REFERENCIA XML " + mensajeSalida);
            }
            return idReferencia;
        }
    }
}
