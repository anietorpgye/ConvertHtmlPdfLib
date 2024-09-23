using SAC.CertificadosLibraryI.Auxiliares;
using SAC.CertificadosLibraryI.Certificados;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI.Inscripciones
{
    public class Inscripcion : Certificado
    {
        #region"**** Variables Globales ****"
        private string _cadenaConexion = string.Empty;
        private string _workPath = string.Empty;
        private string _reportPath = string.Empty;
        private string _imagesPath = string.Empty;
        private string _cssPath = string.Empty;
        private string _signedDocPath = string.Empty;
        private string _urlDescarga = string.Empty;
        private string _hostOpenKm = string.Empty;
        private string _userOpenKm = string.Empty; 
        private string _passwordOpenKm = string.Empty;
        // Defino una lista para almacenar los datos, antes de ordenarlos
        List<Propietario> _listaDatoPropietarios = new List<Propietario>();
        string _estadocivil = string.Empty;
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
        public Inscripcion(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int tipoServicio, int tipoCertificado, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataSet ds = ObtenerDatos(tramite, tipoCertificado);

                DataTable dtData = ds.Tables[0];
                DataTable dtNotas = ds.Tables[1];
                DataTable dtAclaratorias = ds.Tables[2];

                string certificado = dtData.Rows[0]["CERTIFICADO"].ToString().Trim();
                string cedula = dtData.Rows[0]["IDEN_SOLICITANTE"].ToString().Trim();
                string nombreSolicitante = dtData.Rows[0]["NOMBRE_SOLICITANTE"].ToString().Trim();
                string nombresFirmante = dtData.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim();
                string apellidosFirmante = dtData.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = dtData.Rows[0]["CARGO_QUIENFIRMA"].ToString().Trim();
                string empresa = dtData.Rows[0]["EMPRESA"].ToString().Trim();
                string texto = dtData.Rows[0]["TEXTO"].ToString().Trim();
                string texto2 = dtData.Rows[0]["TEXTO2"].ToString().Trim();
                string notaAclaratoria = dtData.Rows[0]["NOTA_ACLARATORIA"].ToString().Trim();

                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombresFirmante + " " + apellidosFirmante +
                    "\nRAZON: CERTIFICADO " + certificado +
                    "\nLOCALIZACION: GUAYAQUIL" +
                    "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                string fechaVigencia = (dtData.Rows[0]["FECHA_FIN"] == DBNull.Value) ? "N/A" : Convert.ToDateTime(dtData.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm");

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                    "\nNo. de Solicitud: " + tramite.ToString() +
                    "\nSolicitante: " + cedula + " " + nombreSolicitante +
                    "\nFecha de Emisión: " + Convert.ToDateTime(dtData.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + fechaVigencia +
                    $"\nDescargue su certificado en {_urlDescarga}");

                DateTime fecEmision = Convert.ToDateTime(dtData.Rows[0]["FECHA"]);

                var imgFirmaQr = Comun.ImageFromByteArray(firmaQr);

                var pathImagenFirma = $"{_reportPath}\\{certificado}.png";

                Comun.DibujarFirma2(new ImagenFirma
                {
                    ColorTexto = System.Drawing.Color.Black,
                    ColorBackground = System.Drawing.Color.White,
                    Fuente = "Consolas",
                    TamanoFuente = 12,
                    FirmaQr = imgFirmaQr,
                    NombresFirmante = nombresFirmante,
                    ApellidosFirmante = apellidosFirmante,
                    Alto = 90,
                    Ancho = 370,
                    RutaImagen = pathImagenFirma
                });

                var documentos = dtNotas.AsEnumerable()
                                 .Select(row => new
                                 {
                                     Documento = row.Field<string>("DOCUMENTO"),
                                     Fecha = row.Field<DateTime>("FECHA"),
                                     Oficina = row.Field<string>("OFICINA")
                                 })
                                 .Distinct();

                var repertorios = dtNotas.AsEnumerable()
                                 .Select(row => new
                                 {
                                     Repertorio = row.Field<int>("REPERTORIO").ToString(),
                                     Acto = row.Field<string>("ACTO"),
                                     FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                     Libro = row.Field<string>("LIBRO"),
                                     NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION").ToString(),
                                     Tomo = row.Field<Int16>("TOMO").ToString(),
                                     FolioInicial = row.Field<int>("FOLIO_INICIAL").ToString(),
                                     FolioFinal = row.Field<int>("FOLIO_FIN").ToString(),
                                     Catastral = row.Field<string>("CODIGO_CATASTRAL"),
                                     Descripcion = row.Field<string>("DESCRIPCION"),
                                     Matricula = row.Field<int>("MATRICULA").ToString(),
                                     Estado = row.Field<string>("ESTADO")
                                 })
                                 .Distinct();

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Arial; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>NOTA DE INSCRIPCIÓN</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtData.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>REGISTRO DE ACTOS</u></b></p>");

                #region tabla documento
                sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                sb.AppendLine($"<thead ><tr><th class=\"parrafospequeno_sj\">DOCUMENTO</th><th class=\"parrafospequeno_sj\">FECHA</th><th class=\"parrafospequeno_sj\">OFICINA</th></tr></thead>");
                sb.AppendLine($"<tbody>");
                foreach (var doc in documentos)
                {
                    sb.AppendLine($"<tr><td><p class=\"parrafospequeno_sj\">{((doc.Documento == "") ? "-----" : doc.Documento.Trim())}</p></td><td><p class=\"parrafospequeno_sj\">{doc.Fecha.ToString("dd/MM/yyyy")}</p></td><td><p class=\"parrafospequeno_sj\">{doc.Oficina}</p></td></tr>");
                }
                sb.AppendLine($"</tbody></table></br>");
                #endregion

                #region repertorios
                foreach (var reper in repertorios)
                {
                    #region tabla repertorio
                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead><tr><th class=\"parrafospequeno_sj\">Repertorio</th><th class=\"parrafospequeno_sj\">Acto</th><th class=\"parrafospequeno_sj\">Fecha de</br>Inscripción</th><th class=\"parrafospequeno_sj\">Libro</th><th class=\"parrafospequeno_sj\">Número de Inscripción</th><th class=\"parrafospequeno_sj\">Tomo</th><th class=\"parrafospequeno_sj\">Folio<br/>Inicial</th><th class=\"parrafospequeno_sj\">Folio<br/>Final</th></tr></thead>");
                    sb.AppendLine($"<tbody>");
                    sb.AppendLine($"<tr><td><p class=\"parrafospequeno_sj\">{reper.Repertorio}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.Acto}</td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.FechaInscripcion.ToString("dd/MM/yyyy")}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.NumeroInscripcion}</p></td><td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.Libro}</p></td><p class=\"parrafospequeno_sj\">{reper.Tomo}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.FolioInicial}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{reper.FolioFinal}</p></td></tr>");
                    sb.AppendLine($"</tbody></table></br>");
                    #endregion

                    #region tabla inmueble
                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead><tr><th class=\"parrafospequeno_sj\">Código Catastral</th>");
                    //sb.AppendLine($"<th class=\"parrafospequeno_sj\">Matrícula INM./</br>Ficha Registral</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Descripción del Inmueble</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Estado</th></tr></thead>");

                    sb.AppendLine($"<tbody>");
                    sb.AppendLine($"<tr><td style=\"width: 25%;\" ><p class=\"parrafospequeno_sj\">{reper.Catastral}</p></td>");
                    //sb.AppendLine($"<td style=\"width: 20%;\"><p class=\"parrafospequeno_sj\">{reper.Matricula}</p></td>"); COMENTA NUMERO DE MATRICULA HASTA NUEVA ORDEN
                    sb.AppendLine($"<td class=\"justificar\" style=\"width: 65%; \"><p class=\"parrafospequeno\">{reper.Descripcion}</p></td>");
                    sb.AppendLine($"<td style=\"width: 10%;\"><p class=\"parrafospequeno_sj\">{reper.Estado}</p></td></tr>");
                    sb.AppendLine($"</tbody></table></br>");
                    #endregion

                    #region tabla partes
                    var partes = dtNotas.AsEnumerable()
                                    .Where(n => n["REPERTORIO"].ToString() == reper.Repertorio)
                                     .Select(row => new
                                     {
                                         Cedula = row.Field<string>("CEDULA"),
                                         Cliente = row.Field<string>("CLIENTE"),
                                         Papel = row.Field<string>("PAPEL"),
                                         EstadoCivil = row.Field<string>("ESTADO_CIVIL")
                                     });

                    if (partes != null)
                    {
                        sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                        sb.AppendLine($"<thead><tr><th class=\"parrafospequeno_sj\">Partes</th><th class=\"parrafospequeno_sj\">Doc.de Identidad</th><th class=\"parrafospequeno_sj\">Nombre</th><th class=\"parrafospequeno_sj\">Estado Civil</th></tr></thead>");
                        sb.AppendLine($"<tbody>");

                        foreach (var parte in partes)
                        {
                            string estadoCivil = (string.IsNullOrEmpty(parte.EstadoCivil)) ? "-----" : parte.EstadoCivil;
                            sb.AppendLine($"<tr><td class=\"parrafospequeno_sj\">{parte.Papel}</td><td><p class=\"parrafospequeno_sj\">{parte.Cedula}</p></td><td><p class=\"parrafospequeno_sj\">{parte.Cliente}</p></td><td><p class=\"parrafospequeno_sj\"> {estadoCivil}</p></td></tr>");
                        }

                        sb.AppendLine($"</tbody></table></br>");
                    }
                    #endregion
                }
                #endregion

                sb.AppendLine($"<p class=\"parrafosmediano\">{texto2.Replace("\\n", "</br></br>")}</p>");

                sb.AppendLine($"<p class=\"parrafosmediano\" >Guayaquil, {fecEmision.ToString("D")}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                //sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombresFirmante} {apellidosFirmante}</b></td></tr>");
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

                var pathToHtml = $"{_reportPath}\\{certificado}.html";
                var pathToPdf = $"{_reportPath}\\{certificado}.pdf";

                if (generarFileHtml)
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                var pathFirma = $"{_reportPath}\\{certificado}.png";
                //byte[] firmaReg = (Byte[])dtXml.Rows[0]["ImagenFirma"];

                //var imgFirmaManual = Comun.ImageFromByteArray(firmaReg);
                //Comun.Upload(firmaReg, pathFirma);

                var dataFooter = new DataFooter
                {
                    Solicitante = $"CED-{cedula}-{nombreSolicitante}",
                    FechaEmision = Convert.ToDateTime(dtData.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = fechaVigencia,
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(dtData.Rows[0]["VALOR"]).ToString("N2")}"
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
                        Contrasena = _passwordOpenKm,
                    });

            }
            catch (Exception ex)
            {
                procesoOk = false;
                throw new CertificadosLibraryException("Error al generar el pdf.\n" + ex.Message, ex);
            }

            return procesoOk;
        }

        private DataSet ObtenerDatos(int tramite, int tipoDocumento)
        {
            // TODO: REVISAR LA DATA
            List<SqlParameter> parametros = new List<SqlParameter>
            {
                new SqlParameter("@SOLINUMTRA", tramite),
                new SqlParameter("@SOLICODCER", tipoDocumento),
            };

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_NotaInscripcion", parametros);

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
