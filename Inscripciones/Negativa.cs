using SAC.CertificadosLibraryI.Certificados;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SAC.CertificadosLibraryI.Oficios;
using SAC.CertificadosLibraryI.Auxiliares;
using SISPRA.Entidades.Modelo.Catalago;

namespace SAC.CertificadosLibraryI.Inscripciones
{
    public class Negativa : Certificado
    {

        #region"**** Variables Globales ****"
        private string _cadenaConexion;
        private string _workPath;
        private string _reportPath;
        private string _imagesPath;
        private string _cssPath;
        private string _signedDocPath;
        private string _urlDescarga;
        private string _hostOpenKm;
        private string _userOpenKm;
        private string _passwordOpenKm;
        // Defino una lista para almacenar los datos, antes de ordenarlos
        List<Propietario> _listaDatoPropietarios = new List<Propietario>();
        string _estadocivil =  string.Empty;
        int _contador = 0;
        string _dataCss = "";
        #endregion

        /// <summary>
        /// Constructor para generar el PDF a partir de una consulta  a la base de datos
        /// </summary>
        /// <param name="cadenaConexion"></param>
        /// <param name="workPath"></param>
        /// <param name="reportPath"></param>
        /// <param name="cssPath"></param>
        /// <param name="imagesPath"></param>
        /// <param name="signedDocPath"></param>
        /// <param name="urlDescarga"></param>
        public Negativa(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
        {
            _cadenaConexion = cadenaConexion;
            _workPath = workPath;
            _reportPath = reportPath;
            _cssPath = cssPath;
            _imagesPath = imagesPath;
            _signedDocPath = signedDocPath;
            _urlDescarga = urlDescarga;
            _hostOpenKm = ObtenerParametro(_cadenaConexion, PARAM_HOST_OPENKM).Result;
            _userOpenKm = ObtenerParametro(_cadenaConexion, PARAM_USER_OPENKM).Result;
            _passwordOpenKm = ObtenerParametro(_cadenaConexion, PARAM_PWD_OPENKM).Result;
        }

        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int tipoServicio,int tipoCertificado, bool generarFileHtml = false)
        {
            bool procesoOk = true;
            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataSet ds = ObtenerDatos(tramite, tipoCertificado);
                DataTable ActosNoInscritos = ds.Tables[0];
                DataTable Intervinientes = ds.Tables[1];
                DataTable Soliciutd = ds.Tables[2];
                DataTable Negativa = ds.Tables[3];
                DataTable dtAclaratorias = ds.Tables[4];

                // info solicitud
                string nombreFirmante = Soliciutd.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim() ;
                string apellidosFirmante = Soliciutd.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = Soliciutd.Rows[0]["CARGO_QUIENFIRMA"].ToString();
                string empresa = Soliciutd.Rows[0]["EMPRESA"].ToString();
                string nombreSolicitante= Soliciutd.Rows[0]["NOMBRE_SOLICITANTE"].ToString();
                string cedula = Soliciutd.Rows[0]["IDEN_SOLICITANTE"].ToString();
                //string uso = Soliciutd.Rows[0]["USO"].ToString();
                string fechaVigencia = (Soliciutd.Rows[0]["FECHA_FIN"] == DBNull.Value) ? "N/A" : Convert.ToDateTime(Soliciutd.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm");
                DateTime fecEmision = Convert.ToDateTime(Soliciutd.Rows[0]["FECHA"]);
                string certificado = Soliciutd.Rows[0]["CERTIFICADO"].ToString();
                string texto = Soliciutd.Rows[0]["TEXTO"].ToString().Trim();
                string texto2 = Soliciutd.Rows[0]["TEXTO2"].ToString().Trim();
                string notaAclaratoria = Soliciutd.Rows[0]["NOTA_ACLARATORIA"].ToString().Trim();


                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombreFirmante + " " + apellidosFirmante +
                   "\nRAZON: CERTIFICADO " + certificado +
                   "\nLOCALIZACION: GUAYAQUIL" +
                   "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                   "\nNo. de Solicitud: " + tramite.ToString() +
                   "\nSolicitante: " + cedula + " " + nombreSolicitante +
                   "\nFecha de Emisión: " + Convert.ToDateTime(Soliciutd.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                   "\nFecha de Vencimiento: " + fechaVigencia +
                   $"\nDescargue su certificado en {_urlDescarga}");


                #region Crear HTML
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaSinBorde, .tablaSinBorde th, .tablaSinBorde tr, .tablaSinBorde td {border: 0px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaSinBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-size: 18.0pt;\"><b>NOTA DE NEGATIVA</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-size: 16.0pt;\"><b>{certificado}</b></p>");
                sb.AppendLine("<br/>");
                sb.AppendLine("<br/>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "</br>")}</p>");

                #region Actos no preocesados

                var ActosNoInscritosList = ActosNoInscritos.AsEnumerable()
                                 .Select(row => new
                                 {
                                     Repertorio = row.Field<int>("REPERTORIO"),
                                     Acto = row.Field<string>("ACTO"),
                                     FechaRepertorio = row.Field<DateTime>("FECHA_REPERTORIO"),
                                     Libro = row.Field<string>("LIBRO"),
                                     Origen = row.Field<string>("ORIGEN")
                                 })
                                 .Distinct();


                sb.AppendLine($"<table class=\"tablaSinBorde justificar\">");
                sb.AppendLine($"<thead ><tr class=\"tablaConBorde\">");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Repertorio</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Acto</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Fecha Emisión</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Libro</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Origen</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Partes</th>");
                sb.AppendLine($"</tr></thead>");
                sb.AppendLine($"<tbody>");
                foreach (var actos in ActosNoInscritosList)
                {
                    var IntervinientesList = Intervinientes.AsEnumerable()
                                     .Select(row => new
                                     {
                                         Repertorio = row.Field<int>("REPERTORIO"),
                                         Cedula = row.Field<string>("CEDULA"),
                                         Cliente = row.Field<string>("CLIENTE"),
                                         Papel = row.Field<string>("PAPEL")
                                     }).Where(x => x.Repertorio == actos.Repertorio);

                    sb.AppendLine($"<tr><td style=\"border: 1px solid black;\"><p class=\"parrafospequeno_sj\">{actos.Repertorio}</p></td>");
                    sb.AppendLine($"<td style=\"border: 1px solid black;\"><p class=\"parrafospequeno_sj\">{actos.Acto}</p></td>");
                    sb.AppendLine($"<td style=\"border: 1px solid black;\"><p class=\"parrafospequeno_sj\">{actos.FechaRepertorio}</p></td>");
                    sb.AppendLine($"<td style=\"border: 1px solid black;\"><p class=\"parrafospequeno_sj\">{actos.Libro}</p></td>");
                    sb.AppendLine($"<td style=\"border: 1px solid black;\"><p class=\"parrafospequeno_sj\">{actos.Origen}</p></td>");
                    sb.AppendLine($"<td style=\"border: 1px solid black;\">");
                    sb.AppendLine($"<table class=\"tablaSinBorde justificar\">");
                    sb.AppendLine($"<tbody>");
                    foreach (var interviniente in IntervinientesList)
                    {
                        sb.AppendLine($"<tr><td><p class=\"parrafospequeno_sj\">{interviniente.Cedula}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{interviniente.Cliente}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{interviniente.Papel}</p></td></tr>");
                    }
                    sb.AppendLine($"</tbody></table></br>");
                    sb.AppendLine($"</td></tr>");
                }
                sb.AppendLine($"</tbody></table></br>");

                #endregion

                sb.AppendLine($"<p class=\"parrafosmediano\">{texto2.Replace("\n", "</br>")}</p>");

                // info negativa
                foreach (DataRow neg in Negativa.Rows)
                {
                    sb.AppendLine("<br/>");
                    var body = neg["Comentario"].ToString();
                    sb.AppendLine(body.Replace("\n", "</br>"));
                    sb.AppendLine("<br/><br/><br/>");
                }

                sb.AppendLine($"<p class=\"parrafosmediano\" >Guayaquil, {fecEmision.ToString("D")}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargoQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<br/> <p class=\"parrafosmediano\" ><b>NOTAS ACLARATORIAS:</b></p>");

                foreach (DataRow nota in dtAclaratorias.Rows)
                {
                    sb.AppendLine($"<p class=\"subrayado\" style=\"font-family: 'Helvetica'; font-size: 9px;\"><b>{nota["TITULO_NOTA"].ToString()}</b></p>");
                    sb.AppendLine($"<p class=\"justificar\" style=\"font-family: 'Helvetica'; font-size: 9px;\">{nota["NOTA"].ToString()}</p>");
                }
                sb.AppendLine("</body></html>");
                #endregion

                #region Crear PDF y subir a OKM
                var pathToHtml = $"{_reportPath}\\{certificado}.html";
                var pathToPdf = $"{_reportPath}\\{certificado}.pdf";

                if (generarFileHtml)
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                var pathFirma = $"{_reportPath}\\{certificado}.png";

                var dataFooter = new DataFooter
                {
                    Solicitante = $"CED-{cedula}-{nombreSolicitante}",
                   // Uso = uso,
                    FechaEmision = Convert.ToDateTime(Soliciutd.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = fechaVigencia,
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(Soliciutd.Rows[0]["VALOR"]).ToString("N2")}"
                };

                //genera el pdf del documento a partir del reporte en html
                 Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

                string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, fecEmision.Year, fecEmision.Month, tipoServicio).Result.ToString();

                Dictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "okp:sispra.tramite", tramite.ToString() },
                    { "okp:sispra.tipo_documento", "" },
                    { "okp:sispra.tipo_servicio", tipoServicio.ToString() },
                    { "okp:sispra.anio", fecEmision.Year.ToString() },
                    { "okp:sispra.etapa", "6" },
                    { "okp:sispra.notificado", "NO" },
                    { "okp:sispra.nombre_documento", certificado }
                };

                //genera el archivo de la firma p12
                Comun.Upload(firmaRegistrador.Firma, firmaRegistrador.PathKeyStore);

                //firma el documento
                procesoOk = Firma.FirmarDocumento(
                    new KeyStore
                    {
                        Password = firmaRegistrador.Password,
                        Firma = firmaRegistrador.Firma,
                        Path = firmaRegistrador.PathKeyStore
                    },
                    new Documento
                    {
                        Nombre = $"{certificado}",
                        PathSinFirmar = _reportPath,
                        PathFirmado = _signedDocPath,
                        CargarEnGestor = true,
                        RepositorioOkm = repositorioOkm,
                        Razon = "Firma de Certificado",
                        Localizacion = "E.P.M.R.P.G.",
                        GrupoOkm = "okg:sispra",
                        Metadata = metadata
                    },
                    new GestorDocumental
                    {
                        Host = _hostOpenKm,
                        Usuario = _userOpenKm,
                        Contrasena = _passwordOpenKm
                    });
                #endregion

            }
            catch (Exception ex)
            {
                procesoOk = false;
                throw new CertificadosLibraryException("Error al generar el pdf.\n" + ex.Message, ex);
            }
            return procesoOk;
        }
        private DataSet ObtenerDatos(int tramite, int tipoCertificado)
        {
            // TODO: REVISAR LA DATA
            List<SqlParameter> parametros = new List<SqlParameter>
            {
                new SqlParameter("@SOLINUMTRA", tramite),
                new SqlParameter("@SOLICODCER", tipoCertificado)
            };

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_Negativa", parametros);

            return ds;
        }

        static bool EstaLleno(DataTable dataTable, int filaIndice, string nombreCampo)
        {
            // Verifica si la fila y el campo existen en el DataTable
            if (filaIndice >= 0 && filaIndice < dataTable.Rows.Count && dataTable.Columns.Contains(nombreCampo))
            {
                // Obtiene el valor del campo en la fila especificada
                object valorCampo = dataTable.Rows[filaIndice][nombreCampo];
                // Verifica si el valor del campo es DBNull o nulo
                return valorCampo != DBNull.Value && valorCampo != null && !string.IsNullOrWhiteSpace(valorCampo.ToString());
            }
            // Devuelve false si la fila o el campo no existen
            return false;

        }
    }
}
