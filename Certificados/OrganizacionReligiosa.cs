using SAC.CertificadosLibraryI.Auxiliares;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SAC.CertificadosLibraryI.Certificados
{
    public class OrganizacionReligiosa : Certificado
    {
        #region"Variables Globales"
        private string _cadenaConexion = string.Empty;
        private string _workPath = string.Empty;
        private string _reportPath = string.Empty;
        private string _cssPath = string.Empty;
        private string _imagesPath = string.Empty;
        private string _signedDocPath = string.Empty;
        private string _urlDescarga = string.Empty;
        private string _hostOpenKm = string.Empty;
        private string _userOpenKm = string.Empty;
        private string _passwordOpenKm = string.Empty;
        #endregion

        /// <summary>
        /// Constructor para genear el PDF a partir de una fuente de datos XML
        /// </summary>
        /// <param name="workPath"></param>
        /// <param name="reportPath"></param>
        /// <param name="cssPath"></param>
        /// <param name="imagesPath"></param>
        public OrganizacionReligiosa(string workPath, string reportPath, string cssPath, string imagesPath)
        {
            _workPath = workPath;
            _reportPath = reportPath;
            _cssPath = cssPath;
            _imagesPath = imagesPath;
        }

        /// <summary>
        /// Constructor para generar el PDF a partir de una consulta a la base de datos
        /// </summary>
        /// <param name="cadenaConexion"></param>
        /// <param name="workPath"></param>
        /// <param name="reportPath"></param>
        /// <param name="cssPath"></param>
        /// <param name="imagesPath"></param>
        public OrganizacionReligiosa(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

        #region"************* Métodos Principales *********** "
        protected string CargaCss(string rutaArchivo) // override async Task OnInitializedAsync()
        {


            string contenidoCSS = File.ReadAllText(rutaArchivo + "/reporte.css");
            return contenidoCSS;
        }

        /// <summary>
        /// Genera el documento pdf a partir del documento xml recibido en el dataSet
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        public bool GenerarPdfXml(DataSet ds)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataTable dtSol = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtSol.Rows[0]["DataXml"].ToString()));
                //************* Datos del cuerpo del reporte ********************************
                string cedula = doc.DocumentElement.SelectSingleNode(@"/row/CEDULA").InnerText;
                string tipo = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;

                string uso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string valor = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                if (valor == "" || valor == null) 
                { 
                    valor = "0.00"; 
                }
                // string TypeAXIS_COD = doc.DocumentElement.SelectSingleNode(@"/row/AXIS_COD").InnerText.Trim();
                string banco = doc.DocumentElement.SelectSingleNode(@"/row/BANCO").InnerText;
                string tramite = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;
                string cmpbtePago = null;
                if (cmpbtePago == null || cmpbtePago.ToString() == "")
                { 
                    cmpbtePago = "NO REGISTRADO"; 
                }

                string canal = doc.DocumentElement.SelectSingleNode(@"/row/CANAL").InnerText;
                string nombre = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE").InnerText;
                string tipoDocumento = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_TIPO").InnerText;
                string cedulaSolicitante = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_CEDULA").InnerText;
                string nombresSolicitante = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_NOMBRES").InnerText;
                //string TypeFECHAGeneral = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText.Trim();
                string fecha = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText;

                string fechaFin = doc.DocumentElement.SelectSingleNode(@"/row/FECHA_FIN").InnerText;
                string codigo = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;

                string texto = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO").InnerText;
                string texto2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;
                //  .....................  Datos de la empresa .................................
                string empresa = doc.DocumentElement.SelectSingleNode(@"/row/EMPRESA").InnerText;
                string notaAclaratoria = doc.DocumentElement.SelectSingleNode(@"/row/NOTA_ACLARATORIA").InnerText.Trim();
                string nombreQuienFirma = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE_QUIENFIRMA").InnerText;
                string cargoQuienFirma = doc.DocumentElement.SelectSingleNode(@"/row/CARGO_QUIENFIRMA").InnerText;

                var pathFirma = _reportPath + $"\\{codigo}.png";


                //*************  valido la firma venga llena *************************
                bool estaLleno = EstaLleno(dtSol, 0, "ImagenFirma");
                if (estaLleno)
                {
                    byte[] firma = (byte[])dtSol.Rows[0]["ImagenFirma"] == null ? null : (byte[])dtSol.Rows[0]["ImagenFirma"];
                    Comun.Upload(firma, pathFirma);
                }
                else
                {
#pragma warning disable CS0219 // La variable está asignada pero su valor es NULL
                    byte[] firmaQr = null;
#pragma warning restore CS0219 // La variable está asignada pero su valor es NULL
                }
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                // sb.AppendLine("<style> html { position: relative; min-height: 100%;}body { margin: 0px; font-family: Consolas; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ }qr {border:solid 1px blue; height:80px; width:80px;}@page {margin-top: 120pt;margin-bottom: 120pt;}.centrar {text-align: center;}.justificar {text-align: justify;} .imagen {width: 130px;height: 70px;} th.fondo {background-color: #E1E0DE; padding: 10px;} .parrafosgrande {text-align: center;font-family: 'Consolas', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Consolas', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Consolas', monospace; font-size: 8pt; text-align: justify;} .firmamediana {text-align: center; font-family: 'Consolas', monospace; font-size: 8pt;} </style>");
                sb.AppendLine("<style>" + CargaCss(_cssPath) + "</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p class=\"parrafosgrande\"><b>CERTIFICADO DE ORGANIZACIÓN RELIGIOSA</b></p>");
                //sb.AppendLine($"<hr><p style=\"text-align: center; font-family: 'Consolas', font-size: 16pt;\"><b>No. {TypeCODIGO}</b></p>");
                sb.AppendLine($"<hr><p class=\"parrafosgrande\" ><b>No. {codigo}</b></p>");
                sb.AppendLine($"<p class=\"parrafospequeno\" >{texto.Replace("\n", "</br>")}</p>");
                XmlNodeList nodeRegInscripcionList = doc.SelectNodes(@"/row/movimiento/inscripciones/inscripcion/iNSCIPT");
                if (nodeRegInscripcionList != null) { sb.AppendLine($"<u><p style=\"font-family: Helvetica; font-size: 12px;\"><b>REGISTRO DE INSCRIPCIÓN</b></p></u>"); }
                foreach (XmlNode dtNode in nodeRegInscripcionList) //
                {
                    if (dtNode != null)
                    {
                        var inscrip = dtNode.InnerText; // iNSCIPT // && dtNode.ChildNodes.Count>1

                        sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{inscrip}</p>");


                    }
                }

                XmlNodeList nodeRegObsList = doc.SelectNodes(@"/row/movimiento/reformas/reforma/OBSERVACION");
                if (nodeRegObsList != null)
                {
                    sb.AppendLine($"<u><p class=\"parrafospequeno\" ><b>REGISTRO DE REFORMAS</b></u></p>");
                }
                int i = 0;
                foreach (XmlNode dtNodeObs in nodeRegObsList) //
                {
                    if (dtNodeObs != null)
                    {
                        var reformas = dtNodeObs.InnerText; // OBSERVACION
                        i = dtNodeObs.ChildNodes.Count == 0 ? -1 : dtNodeObs.ChildNodes.Count;
                        if (i > -1) //
                        {
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{reformas.Replace("\n", "</br>")}</p>");
                            i++;
                        }
                    }
                }
                i = 0;
                #endregion


                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">{texto2.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">Guayaquil, {fecha}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{codigo}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombreQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargoQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<br/> <p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>NOTAS ACLARATORIAS:</b></p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 9px;\" >{notaAclaratoria.Replace("\n", "</br>")}</p><br/>");

                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\{codigo}.html";
                var pathToPdf = $"{_workPath}\\{codigo}.pdf";

                //  Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());

                if (estaLleno)
                {
                    byte[] firmaQr = (Byte[])dtSol.Rows[0]["ImagenFirma"];
                }
                else
                {
#pragma warning disable CS0219 // La variable está asignada pero su valor puede ser NULL
                    byte[] firmaQr = null;
#pragma warning restore CS0219 // La variable está asignada pero su valor puede ser NULL
                }


                byte[] codigoQrb = Comun.GenerarQr("Certificado " + codigo +
                  "\nNo. de Solicitud: " + tramite.ToString() +
                  "\nSolicitante: " + cedula.Trim() + " " + nombre.Trim() +
                  "\nFecha de Emisión: " + Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm") +
                  "\nFecha de Vencimiento: " + Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm") +
                  $"\nDescargue su certificado en {_urlDescarga}");


                var dataFooter = new DataFooter
                {
                    Solicitante = "CED-" + " " + cedula + " " + nombre,
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RPG GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = cmpbtePago,
                    Valor = valor
                };

                //********* Genera el pdf del documento a partir del reporte en html *************
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _workPath, logoRpg, logoAlcaldia, codigoQrb, dataFooter);

            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al generar el certificado.\n" + ex.Message, ex);
            }

            return procesoOk;
        }

        /// <summary>
        /// Genera el documento pdf a partir de datos consultados de base de datos de SISPRA
        /// </summary>
        /// <param name="firmaRegistrador"></param>
        /// <param name="tramite"></param>
        /// <param name="secuencial"></param>
        /// <param name="anioOrden"></param>
        /// <param name="numeroOrden"></param>
        /// <param name="tipoServicio"></param>
        /// <param name="generarFileHtml"></param>
        /// <returns></returns>
        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int secuencial, int tipoServicio, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataSet ds = ObtenerDatos(tramite, secuencial);

                DataTable dtData = ds.Tables[0];
                DataTable dtInscripciones = ds.Tables[1];
                DataTable dtAclaratorias = ds.Tables[2];

                string certificado = dtData.Rows[0]["CERTIFICADO"].ToString().Trim();
                string cedula = dtData.Rows[0]["IDEN_SOLICITANTE"].ToString().Trim();
                string nombreSolicitante = dtData.Rows[0]["NOMBRE_SOLICITANTE"].ToString().Trim();
                string nombresFirmante = dtData.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim();
                string apellidosFirmante = dtData.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = dtData.Rows[0]["CARGO_QUIENFIRMA"].ToString().Trim();
                string empresa = dtData.Rows[0]["EMPRESA"].ToString().Trim();
                string uso = dtData.Rows[0]["USO"].ToString().Trim();
                string texto = dtData.Rows[0]["TEXTO"].ToString().Trim();
                string texto2 = dtData.Rows[0]["TEXTO2"].ToString().Trim();
                //string notaAclaratoria = dtData.Rows[0]["NOTA_ACLARATORIA"].ToString().Trim();

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

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16px; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica'; font-size: 12px; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8px; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8px;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE ORGANIZACIÓN RELIGIOSA</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtData.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "")}</p>");

                if (dtInscripciones.Rows.Count > 0)
                {
                    sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>REGISTRO DE INSCRIPCIÓN</u></b></p>");

                    foreach (DataRow inscripcion in dtInscripciones.Rows)
                    {
                        var dataInscripcion = inscripcion["INSCRIPCION"].ToString();
                        var dataReforma = inscripcion["REFORMA"].ToString();

                        sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\"><b>Inscripción:</b></p>");
                        sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{dataInscripcion}</p>");
                        if (!string.IsNullOrEmpty(dataReforma))
                        {
                            var test = dataReforma.Replace("\n", "<br>");
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\"><b>Reforma:</b></p>");
                            sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{test}</p>");
                        }
                    }
                }

                #region firma del registrador
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto2.Replace("\\n", "</br></br>")}</p>");

                sb.AppendLine($"<p class=\"parrafosmediano\" >Guayaquil, {fecEmision.ToString("D")}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                //sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombresFirmante} {apellidosFirmante}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargoQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");
                #endregion

                #region notas aclaratorias
                sb.AppendLine($"<br/> <p class=\"parrafosmediano\" ><b>NOTAS ACLARATORIAS:</b></p>");

                foreach (DataRow nota in dtAclaratorias.Rows)
                {
                    sb.AppendLine($"<p class=\"subrayado\" style=\"font-family: 'Helvetica'; font-size: 9px;\"><b>{nota["TITULO_NOTA"].ToString()}</b></p>");
                    sb.AppendLine($"<p class=\"justificar\" style=\"font-family: 'Helvetica'; font-size: 9px;\">{nota["NOTA"].ToString()}</p>");
                }
                #endregion

                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\{certificado}.html";
                var pathToPdf = $"{_reportPath}\\{certificado}.pdf";

                if (generarFileHtml)
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                var pathFirma = $"{_reportPath}\\{certificado}.png";

                #region generación de pdf, firma e importacion a OpenKm
                var dataFooter = new DataFooter
                {
                    Solicitante = $"CED-{cedula}-{nombreSolicitante}",
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(dtData.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = fechaVigencia,
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(dtData.Rows[0]["VALOR"]).ToString("N2")}"
                };

                //genera el pdf del documento a partir del reporte en html
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

               //Upload(reporte, _workPath + $"\\{tramite}_{secuencial}.pdf");
                string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, fecEmision.Year, fecEmision.Month, tipoServicio).Result.ToString();

                Dictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "okp:sispra.tramite", tramite.ToString() },
                    { "okp:sispra.secuencia", secuencial.ToString() },
                    { "okp:sispra.tipo_documento", "" },
                    { "okp:sispra.tipo_servicio", tipoServicio.ToString() },
                    { "okp:sispra.anio", fecEmision.Year.ToString() },
                    { "okp:sispra.etapa", "6" },
                    { "okp:sispra.notificado", "NO" }, 
                    { "okp:sispra.nombre_documento", certificado }
                };

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

        private DataSet ObtenerDatos(int tramite, int secuencial)
        {
            List<SqlParameter> parametros = new List<SqlParameter>
            {
                new SqlParameter("@SOLINUMTRA", tramite),
                new SqlParameter("@SOLISEC", secuencial),
            };

            var resul = Comun.ActualizarVigencia(_cadenaConexion, tramite, secuencial);

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertOrganizacionReligiosa", parametros);

            return ds;
        }

        /// <summary>
        /// Función para verificar si un campo en una fila del datatable está lleno o vacío
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="filaIndice"></param>
        /// <param name="nombreCampo"></param>
        /// <returns></returns>
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
        //*************** Ordenar los propietarios de los predios con los nodos del Xml: BD ********

        // Función para agregar valores a la lista
        void ListaAgregarDatos(string tipoCed, string identificacion, string nombres)
        {    // Creó una instancia de Propietario, para invocarlo
            Propietario nuevoPropietario = new Propietario
            {
                TipoCed = tipoCed,
                Identificacion = identificacion,
                Nombre = nombres
            };
        }
    }
}
