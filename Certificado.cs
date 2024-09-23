using System.Data.SqlClient;
using System.Data;
using SAC.CertificadosLibraryI.Auxiliares;

namespace SAC.CertificadosLibraryI
{
    public enum TipoCertificado
    {
        PoseerBienes = 1,
        HistoriaDominio = 3,
        NotaInscripcion = 4,
        NegativaInscripcion = 5,
        HistoriaOtroCanton = 7,
        PosesionEfectiva = 8,
        OrganizacionReligiosa = 9,
        EspecificoInmobiliario = 10,
        CertificacionRegistral = 11,
        Interdiccion = 12,
        NoInscripcion = 13,
        VigenciaAsientoRegistral = 14,
        InexistenciaAsientoRegistral = 15,
        AsientoRegistral = 16,
        Varios = 17,
        CopiaAsientoRegistral = 19,
        NegativaDefinitivaInscripcion = 20,
        CertificadoNotaInscripcion = 25,
        Gravamen = 28
    };

    public class Certificado : ICertificado
    {
        public const string PARAM_HOST_OPENKM = "HOSTOPENKM";
        public const string PARAM_USER_OPENKM = "USEROPENKM";
        public const string PARAM_PWD_OPENKM = "PWDOPENKM";

        public DataSet ObtenerDatos(string cadenaConexion, string procedimiento, List<SqlParameter> parametros)
        {
            Dal dal = new Dal(cadenaConexion)
            {
                Procedimiento = procedimiento,
                Parametros = parametros
            };

            var ds = dal.ObtenerDataSets();

            return ds;
        }

        public async Task<string> ObtenerParametro(string cadenaConexion, string parametro)
        {
            Dal dal = new Dal(cadenaConexion)
            {
                Procedimiento = "dbo.SP_ParametroSistemas",
                Parametros = new List<SqlParameter>
                    {
                        new SqlParameter("@NombreParametro", parametro),
                        new SqlParameter("@MensajeErrorControlado", SqlDbType.VarChar, -1) { Direction = ParameterDirection.Output }
                    }
            };

            var rowsAffected = await dal.EjecutarNoQuery();

            string? valorParametro = (dal.Parametros[1].Value == DBNull.Value) ? string.Empty : dal.Parametros[1].Value.ToString();

            return string.IsNullOrEmpty(valorParametro) ? string.Empty: valorParametro;
        }

        public async Task<string> ObtenerRepositorioOpenKm(string cadenaConexion, int anio, int mes, int tipoServicio)
        {
            Dal dal = new Dal(cadenaConexion)
            {
                Procedimiento = "dbo.SP_DocumentosTramite",
                Parametros = new List<SqlParameter>
                {
                    new SqlParameter("@opcion", 1),
                    new SqlParameter("@ANIO", anio),
                    new SqlParameter("@MES", mes),
                    new SqlParameter("@TIPOSERVICIO", tipoServicio),
                    new SqlParameter("@MensajeErrorControlado", 0) { Direction = ParameterDirection.Output }
                }
            };

            var ds = await dal.ObtenerEscalar();

            return (ds == null) ? string.Empty: (string)ds;
        }
    }
}
