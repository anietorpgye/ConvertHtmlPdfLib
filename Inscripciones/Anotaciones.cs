using SAC.CertificadosLibraryI.Auxiliares;
using SAC.CertificadosLibraryI.Certificados;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SISPRA.Entidades.Modelo.Helper.EnumList;

namespace SAC.CertificadosLibraryI.Inscripciones
{
    public class Anotaciones : Certificado
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
        public Anotaciones(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

                DataTable dtDataCertificado = ds.Tables[0];
                DataTable dtDatosAnotacion = ds.Tables[1];
                DataTable dtPropietarios = ds.Tables[2];
                DataTable dtPredios = ds.Tables[3];
                DataTable dtDatosSolicitud = ds.Tables[4];
                DataTable dtAclaratorias = ds.Tables[5];

                string certificado = dtDataCertificado.Rows[0]["CERTIFICADO"].ToString().Trim();
                string cedula = dtDataCertificado.Rows[0]["IDEN_SOLICITANTE"].ToString().Trim();
                string nombreSolicitante = dtDataCertificado.Rows[0]["NOMBRE_SOLICITANTE"].ToString().Trim();
                string nombresFirmante = dtDataCertificado.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim();
                string apellidosFirmante = dtDataCertificado.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = dtDataCertificado.Rows[0]["CARGO_QUIENFIRMA"].ToString().Trim();
                string empresa = dtDataCertificado.Rows[0]["EMPRESA"].ToString().Trim();
                string texto = dtDataCertificado.Rows[0]["TEXTO"].ToString().Trim();
                string texto2 = dtDataCertificado.Rows[0]["TEXTO2"].ToString().Trim();
                string Valor = dtDataCertificado.Rows[0]["VALOR"].ToString().Trim();

                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombresFirmante + " " + apellidosFirmante +
                    "\nRAZON: CERTIFICADO " + certificado +
                    "\nLOCALIZACION: GUAYAQUIL" +
                    "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                string fechaVigencia = (dtDataCertificado.Rows[0]["FECHA_FIN"] == DBNull.Value) ? "N/A" : Convert.ToDateTime(dtDataCertificado.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm");

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                    "\nNo. de Solicitud: " + tramite.ToString() +
                    "\nSolicitante: " + cedula + " " + nombreSolicitante +
                    "\nFecha de Emisión: " + Convert.ToDateTime(dtDataCertificado.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + fechaVigencia +
                    $"\nDescargue su certificado en {_urlDescarga}");

                DateTime fecEmision = Convert.ToDateTime(dtDataCertificado.Rows[0]["FECHA"]);

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


                var Repertorio = dtDatosAnotacion.AsEnumerable()
                                 .Select(row => new
                                 {
                                     Reperotorio = row.Field<int>("REPERTORIO"),
                                     NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION"),
                                     FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                     CodLibMov = row.Field<int>("COD_LIBRO_MOV"),
                                 })
                                 .Distinct();


                var predios = dtPredios.AsEnumerable()
                                 .Select(row => new
                                 {
                                     CodigoCatastral = row.Field<string>("CODIGO_CATASTRAL"),
                                     Descirpcion = row.Field<string>("DESCRIPCION"),
                                     Estado = row.Field<string>("ESTADO")
                                 })
                                 .Distinct();

                var documentos = dtDatosSolicitud.AsEnumerable()
                                 .Select(row => new
                                 {
                                     Documento = row.Field<string>("DOCUMENTO"),
                                     Fecha = row.Field<DateTime>("FECHA"),
                                     Oficina = row.Field<string>("OFICINA")
                                 })
                                 .Distinct();

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Arial; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>NOTA DE INSCRIPCIÓN</b></p>");
                switch (tipoCertificado)
                {
                    case 21:
                        sb.AppendLine($"<p style=\"\"text-align: center; font-family: Helvetica; font-size: 20px;\" ><b><u>REPOSICIÓN DE FOLIOS O REPOSICIÓN DE ASIENTO REGISTRAL</u></b></p>");
                        break;
                    case 22:
                        sb.AppendLine($"<p style=\"\"text-align: center; font-family: Helvetica; font-size: 20px;\" ><b><u>ANOTACIÓN MARGINAL POR EXTRANJERÍA</u></b></p>");
                        break;
                    case 23:
                        sb.AppendLine($"<p style=\"\"text-align: center; font-family: Helvetica; font-size: 20px;\" ><b><u>ANOTACIÓN DE CESIÓN DE DERECHOS FIDUCIARIOS</u></b></p>");
                        break;
                }
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtDataCertificado.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "</br>")}</p>");

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

                #region Reperotorios
                foreach (var reper in Repertorio)
                {
                    var AnotacionRepertorio = dtDatosAnotacion.AsEnumerable()
                                    .Where(n => n["NUM_INSCRIPCION"].ToString() == reper.NumeroInscripcion.ToString()
                                            && n["FECHA_INSCRIPCION"].ToString() == reper.FechaInscripcion.ToString()
                                            && n["COD_LIBRO_MOV"].ToString() == reper.CodLibMov.ToString())
                                     .Select(row => new
                                     {
                                         LibroAnotacion = row.Field<string>("LIBRO_ANOTACION"),
                                         FechaAnotacion = row.Field<DateTime>("FECHA_ANOTACION"),
                                         NumeroAnotacion = row.Field<int>("NUMERO_ANOTACION"),
                                         NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION"),
                                         FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                         CodLibMov = row.Field<int>("COD_LIBRO_MOV"),
                                     })
                                     .Distinct();

                    var RepertorioDetalle = dtDatosAnotacion.AsEnumerable()
                                    .Where(n => n["NUM_INSCRIPCION"].ToString() == reper.NumeroInscripcion.ToString()
                                            && n["FECHA_INSCRIPCION"].ToString() == reper.FechaInscripcion.ToString()
                                            && n["COD_LIBRO_MOV"].ToString() == reper.CodLibMov.ToString())
                                     .Select(row => new
                                     {
                                         Reperotorio = row.Field<int>("REPERTORIO"),
                                         Libro = row.Field<string>("LIBRO"),
                                         Acto = row.Field<string>("ACTO"),
                                         NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION"),
                                         FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                         Tomo = row.Field<int>("TOMO"),
                                         FolioInicial = row.Field<int>("FOLIO_INICIAL"),
                                         FolioFinal = row.Field<int>("FOLIO_FIN"),
                                         CodLibMov = row.Field<int>("COD_LIBRO_MOV"),
                                     })
                                     .Distinct();
                    var OrigenRepertorio = dtDatosAnotacion.AsEnumerable()
                                     .Where(n => n["NUM_INSCRIPCION"].ToString() == reper.NumeroInscripcion.ToString()
                                            && n["FECHA_INSCRIPCION"].ToString() == reper.FechaInscripcion.ToString()
                                            && n["COD_LIBRO_MOV"].ToString() == reper.CodLibMov.ToString())
                                     .Select(row => new
                                     {
                                         Canton = row.Field<string>("CANTON"),
                                         Origen = row.Field<string>("ORIGEN"),
                                         EscProtRes = row.Field<string>("ESC_PROT_RES"),
                                         NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION"),
                                         FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                         CodLibMov = row.Field<int>("COD_LIBRO_MOV"),
                                     })
                                     .Distinct();


                    var Porpietarios = dtPropietarios.AsEnumerable()
                                     .Where(n => n["NUM_INSCRIPCION"].ToString() == reper.NumeroInscripcion.ToString()
                                            && n["FECHA_INSCRIPCION"].ToString() == reper.FechaInscripcion.ToString()
                                            && n["COD_LIBRO_MOV"].ToString() == reper.CodLibMov.ToString())
                                    .Select(row => new
                                    {
                                        Partes = row.Field<string>("PAPEL"),
                                        Identificacion = row.Field<string>("CEDULA"),
                                        Nombre = row.Field<string>("CLIENTE"),
                                        EstadoCivil = row.Field<string>("ESTADO_CIVIL"),
                                        NumeroInscripcion = row.Field<int>("NUM_INSCRIPCION"),
                                        FechaInscripcion = row.Field<DateTime>("FECHA_INSCRIPCION"),
                                        CodLibMov = row.Field<int>("COD_LIBRO_MOV"),
                                    })
                                    .Distinct();

                    sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>Datos de Anotación</u></b></p>");
                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead ><tr>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Libro de Anotación</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Fecha de</br>Anotación</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Número de</br>Anotaciones</th>");
                    sb.AppendLine($"</tr></thead>");
                    sb.AppendLine($"<tbody>");
                    foreach (var anotacion in AnotacionRepertorio)
                    {
                        sb.AppendLine($"<tr><td><p class=\"parrafospequeno_sj\">{anotacion.LibroAnotacion}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{anotacion.FechaAnotacion.ToString("dd/MM/yyyy")}</td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{anotacion.NumeroAnotacion}</p></td>");
                    }
                    sb.AppendLine($"</tbody></table></br>");

                    switch (tipoCertificado)
                    {
                        case 21:
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>Datos de Asiento Registral Repuesto</u></b></p>");
                            break;
                        case 22:
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>Datos de Asiento Registral Marginado</u></b></p>");
                            break;
                        case 23:
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>Datos de Asiento Registral Afectado</u></b></p>");
                            break;
                    }

                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead ><tr>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Repertorio</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Acto</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Fecha de</br>Inscripción</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Libro</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Número de</br>Inscripción</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Tomo</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Folio</br>Inicial</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Folio</br>Final</th>");
                    sb.AppendLine($"</tr></thead>");
                    sb.AppendLine($"<tbody><tr>");
                    foreach (var repDet in RepertorioDetalle)
                    {
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.Reperotorio}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.Acto}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.FechaInscripcion.ToString("dd/MM/yyyy")}</td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.Libro}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.NumeroInscripcion}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.Tomo}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.FolioInicial}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{repDet.FolioFinal}</p></td>");
                    }
                    sb.AppendLine($"</tbody></table></br>");


                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead ><tr>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Cantón</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Origen</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Esc/Prot/Res</th>");
                    sb.AppendLine($"</tr></thead>");
                    sb.AppendLine($"<tbody><tr>");
                    foreach (var origen in OrigenRepertorio)
                    {
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{origen.Canton}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{origen.Origen}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{origen.EscProtRes}</p></td>");
                    }
                    sb.AppendLine($"</tbody></table></br>");


                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead ><tr>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Partes</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Doc. de Identidad</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Nombre</th>");
                    sb.AppendLine($"<th class=\"parrafospequeno_sj\">Estado Civil</th>");
                    sb.AppendLine($"</tr></thead>");
                    sb.AppendLine($"<tbody><tr>");
                    foreach (var propDet in Porpietarios)
                    {
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{propDet.Partes}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{propDet.Identificacion}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{propDet.Nombre}</p></td>");
                        sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{propDet.EstadoCivil}</p></td>");
                    }
                    sb.AppendLine($"</tbody></table></br>");


                }
                #endregion

                #region Predios
                sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                sb.AppendLine($"<thead ><tr>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Código</br>Catastral</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Descripción del Inmueble</th>");
                sb.AppendLine($"<th class=\"parrafospequeno_sj\">Estado</th>");
                sb.AppendLine($"</tr></thead>");
                sb.AppendLine($"<tbody><tr>");
                foreach (var item in predios)
                {
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{item.CodigoCatastral}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{item.Descirpcion}</p></td>");
                    sb.AppendLine($"<td><p class=\"parrafospequeno_sj\">{item.Estado}</p></td>");

                }
                sb.AppendLine($"</tbody></table></br>");
                #endregion


                sb.AppendLine($"<p class=\"parrafosmediano\">{texto2.Replace("\\n", "</br></br>")}</p>");

                sb.AppendLine($"<p class=\"parrafosmediano\" >Guayaquil, {fecEmision.ToString("D")}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombresFirmante} {apellidosFirmante}</b></td></tr>");
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

                var dataFooter = new DataFooter
                {
                    Solicitante = $"CED-{cedula}-{nombreSolicitante}",
                    FechaEmision = Convert.ToDateTime(dtDataCertificado.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = fechaVigencia,
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(dtDataCertificado.Rows[0]["VALOR"]).ToString("N2")}"
                };
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
                return true;
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
            string SpEjecutar = "";
            switch (tipoDocumento)
            {
                case 21: //REPOSICION DE FOLIO
                    SpEjecutar = "dbo.SP_ReposicionFolio";
                    break;

                case 22: //ANOTACION MARGINAL EXTRANJERIA

                    SpEjecutar = "dbo.SP_AnotacionMarginalExtranjeria";
                    break;

                case 23: //ANOTACION DE CESION DE DERECHO
                    SpEjecutar = "dbo.SP_AnotacionCesionDerechos";
                    break;
            }


            DataSet ds = ObtenerDatos(_cadenaConexion, SpEjecutar, parametros);

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
