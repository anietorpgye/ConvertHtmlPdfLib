using System.Data.SqlClient;
using System.Data;

namespace SAC.CertificadosLibraryI
{
    public interface ICertificado
    {
        DataSet ObtenerDatos(string cadenaConexion, string procedimiento, List<SqlParameter> parametros);
        Task<string> ObtenerParametro(string cadenaConexion, string parametro);
        Task<string> ObtenerRepositorioOpenKm(string cadenaConexion, int anio, int mes, int tipoServicio);
    }
}
