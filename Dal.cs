using System.Data;
using System.Data.SqlClient;
using System.Reflection;

namespace SAC.CertificadosLibraryI
{
    public class Dal
    {
        private string _cadenaConexion;

        public string Procedimiento { get; set; }
        public List<SqlParameter> Parametros { get; set; }

        public Dal(string cadenaConexion)
        {
            _cadenaConexion = cadenaConexion;
        }

        public DataSet ObtenerDataSets()
        {
            DataSet ds = new DataSet();

            SqlConnection connection = new SqlConnection(_cadenaConexion);
            using (SqlCommand cmd = connection.CreateCommand())
            {
                cmd.Connection.Open();

                cmd.CommandText = Procedimiento;
                cmd.CommandType = CommandType.StoredProcedure;

                if (Parametros != null)
                {
                    foreach (SqlParameter param in Parametros)
                    {
                        cmd.Parameters.Add(param);
                    }
                }

                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(ds);
                }
            }


            return ds;
        }

        public async Task<List<T>> ObtenerDatos<T>() where T : new()
        {
            List<T> resul = new List<T>();

            SqlConnection connection = new SqlConnection(_cadenaConexion);
            using (SqlCommand cmd = connection.CreateCommand())
            {
                cmd.Connection.Open();

                cmd.CommandText = Procedimiento;
                cmd.CommandType = CommandType.StoredProcedure;

                if (Parametros != null)
                {
                    foreach (SqlParameter param in Parametros)
                    {
                        cmd.Parameters.Add(param);
                    }
                }

                using (var rd = await cmd.ExecuteReaderAsync())
                {
                    while (rd.Read())
                    {
                        T t = new T();

                        for (int inc = 0; inc < rd.FieldCount; inc++)
                        {
                            var colName = rd.GetName(inc);
                            Type type = t.GetType();
                            PropertyInfo prop = type.GetProperty(colName);
                            if (prop != null)
                            {
                                object value = rd.GetValue(inc);
                                if (value == DBNull.Value)
                                {
                                    prop.SetValue(t, null);
                                }
                                else
                                {
                                    prop.SetValue(t, value);
                                }
                            }
                        }

                        resul.Add(t);
                    }
                }
            }

            return resul;
        }

        public async Task<int> EjecutarNoQuery()
        {
            int rc;
            SqlConnection connection = new SqlConnection(_cadenaConexion);
            using (SqlCommand cmd = connection.CreateCommand())
            {
                cmd.Connection.Open();

                cmd.CommandText = Procedimiento;
                cmd.CommandType = CommandType.StoredProcedure;

                if (Parametros != null)
                {
                    foreach (SqlParameter param in Parametros)
                    {
                        cmd.Parameters.Add(param);
                    }
                }

                rc = await cmd.ExecuteNonQueryAsync();
            }

            return rc;
        }

        public async Task<object?> ObtenerEscalar()
        {
            object? rc;
            SqlConnection connection = new SqlConnection(_cadenaConexion);
            using (SqlCommand cmd = connection.CreateCommand())
            {
                cmd.Connection.Open();

                cmd.CommandText = Procedimiento;
                cmd.CommandType = CommandType.StoredProcedure;

                if (Parametros != null)
                {
                    foreach (SqlParameter param in Parametros)
                    {
                        cmd.Parameters.Add(param);
                    }
                }

                rc = await cmd.ExecuteScalarAsync();
            }

            return rc;
        }
    }
}
