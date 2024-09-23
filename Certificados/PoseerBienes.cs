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
    public class PoseerBienes : Certificado
    {
        private string _cadenaConexion;
        private string _workPath;
        private string _reportPath;
        private string _cssPath;
        private string _imagesPath;
        private string _signedDocPath;
        private string _urlDescarga;
        private string _hostOpenKm;
        private string _userOpenKm;
        private string _passwordOpenKm;

        public PoseerBienes(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

        public PoseerBienes(string workPath, string reportPath, string cssPath, string imagesPath)
        {
            _workPath = workPath;
            _reportPath = reportPath;
            _cssPath = cssPath;

            _imagesPath = imagesPath;
        }

        public bool GenerarPdfXml(DataSet ds, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\Logo Alcaldia Guayaquil 2.png");

                DataTable dtSol = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtSol.Rows[0]["DataXml"].ToString()));

                string cedula = doc.DocumentElement.SelectSingleNode(@"/row/CEDULA").InnerText;
                string tipo = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;
                string planHabita = doc.DocumentElement.SelectSingleNode(@"/row/PLAN_HABITACIONAL").InnerText;
                string uso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string valor = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                if (valor == "" || valor == null) { valor = "0.00"; }
                string codigoAxis = doc.DocumentElement.SelectSingleNode(@"/row/AXIS_COD").InnerText.Trim();
                string banco = doc.DocumentElement.SelectSingleNode(@"/row/BANCO").InnerText;
                string tramite = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;
                string cmpbtePago = doc.DocumentElement.SelectSingleNode(@"/row/COMPROBANTE_PAGO").InnerText;
                if (string.IsNullOrEmpty(cmpbtePago))
                { 
                    cmpbtePago = "NO REGISTRADO"; 
                }
                string codRecaudacion = doc.DocumentElement.SelectSingleNode(@"/row/RECAUDACION_COD").InnerText;
                string factura = doc.DocumentElement.SelectSingleNode(@"/row/FACTURA").InnerText;
                string canal = doc.DocumentElement.SelectSingleNode(@"/row/CANAL").InnerText;
                string nombre = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE").InnerText;
                string tipoIdentificacion = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_TIPO").InnerText;
                string cedulaSolicitante = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_CEDULA").InnerText;
                string nombreSolicitante = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_NOMBRES").InnerText;
                //string TypeFECHAGeneral = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText.Trim();
                string fecha = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText;

                string fechaFin = doc.DocumentElement.SelectSingleNode(@"/row/FECHA_FIN").InnerText;
                string codigo = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;

                string texto = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO").InnerText;
                string item = doc.DocumentElement.SelectSingleNode(@"/row/item").InnerText; //<item><detalle1>
                if (item == "" || item == null) { item = null; }

                string texto2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;

                string empresa = doc.DocumentElement.SelectSingleNode(@"/row/EMPRESA").InnerText;
                string notaAclaratoria = doc.DocumentElement.SelectSingleNode(@"/row/NOTA_ACLARATORIA").InnerText;
                string nombreQuienFirma = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE_QUIENFIRMA").InnerText;
                string cargoQuienFirma = doc.DocumentElement.SelectSingleNode(@"/row/CARGO_QUIENFIRMA").InnerText;

                var pathFirma = _reportPath + $"\\{codigo}.png";
                //*************  valido la firma venga llena *************************
                bool estaLleno = EstaLleno(dtSol, 0, "ImagenFirma");
                if (estaLleno) // dtSol.Rows[0]["ImagenFirma"] == null   
                {
                    byte[] firma = (byte[])dtSol.Rows[0]["ImagenFirma"] == null ? null : (byte[])dtSol.Rows[0]["ImagenFirma"];
                    Comun.Upload(firma, pathFirma);
                }
                else
                {
                    byte[] firma = null;
                }

                string dtablexml;
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> html { position: relative; min-height: 100%;}body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ }qr {border:solid 1px blue; height:80px; width:80px;}@page {margin-top: 120pt;margin-bottom: 120pt;}.centrar {text-align: center;}.justificar {text-align: justify;} .imagen {width: 130px;height: 70px;} th.fondo {background-color: #E1E0DE; padding: 10px;} .parrafosgrande {text-align: center;font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .firmamediana {text-align: center; font-family: 'Helvetica', monospace; font-size: 8pt;} </style>");
                sb.AppendLine("</head><body>");
                //sb.AppendLine($"<p style=\"text-align: center; font-family: 'Helvetica', font-size: 16pt;\"><b>CERTIFICADO DE POSEER BIENES</b></p>");
                sb.AppendLine($"<p class=\"parrafosgrande\" >CERTIFICADO DE POSEER BIENES</p>");
                //sb.AppendLine($"<hr><p style=\"text-align: center; font-family: 'Helvetica', font-size: 16pt;\"><b>No. {TypeCODIGO}</b></p>");
                sb.AppendLine($"<hr><p class=\"parrafosgrande\" >No. {codigo}</p>");
                sb.AppendLine($"<p class=\"parrafosmediano\" >{texto.Replace("\n", "</br>")}</p>");
                if (item != null) // ***** SI hay registros.. lleno la tabla con sus nodos *****
                {
                    #region"Validar campos para tabla"
                    // Select nodes using XPath
                    XmlNodeList nodeList = doc.SelectNodes(@"/row/item/detalle1");
                    int c = 0; dtablexml = "";
                    // Iterate through each detalle1 node
                    foreach (XmlNode detalleNode in nodeList)
                    {
                        // Accedo a los nodos del hijo
                        XmlNode matriculaNode = detalleNode.SelectSingleNode("MATRICULA_INMOBILIARIA");
                        XmlNode registroCatastralNode = detalleNode.SelectSingleNode("REGISTRO_CATASTRAL");
                        XmlNode ubicacionNode = detalleNode.SelectSingleNode("UBICACION");

                        if (c == 0 && nodeList.Count == 1) // una sola fila de datos
                        {
                            dtablexml = "<table border=0.5 class=\"justificar\"> <thead class=\"fondo\"><th class=\"fondo\">Código Catastral / Inscripciones</th><th class=\"fondo\">Linderos</th></tr></thead><tbody>";
                            dtablexml += "<tr><td><center><p class=\"parrafosmediano\">" + registroCatastralNode.InnerText + "</p></center></td>";
                            dtablexml += "<td><p class=\"parrafosmediano\" >" + ubicacionNode.InnerText.Replace("\n", "</br>") + "</p></td></tbody></table>";

                        }
                        if (c == 0 && nodeList.Count > 1) // si es mas de una fila
                        {
                            dtablexml = "<table border=0.5 class=\"justificar\"> <thead><th class=\"fondo\">Código Catastral / Inscripciones</th><th class=\"fondo\">Linderos</th></tr></thead><tbody>";
                            dtablexml += "<tr><td><center><p class=\"parrafosmediano\" > " + registroCatastralNode.InnerText + "</p></center></td>";
                            dtablexml += "<td><p class=\"parrafosmediano\" > " + ubicacionNode.InnerText.Replace("\n", "</br>") + "</p></td>";
                        }
                        c++;
                        if (c > 1 && nodeList.Count > 1) //  <\br>  \r\n
                        {
                            dtablexml += "<tr><td><p class=\"parrafosmediano\">" + registroCatastralNode.InnerText + "</p></td>";
                            dtablexml += "<td><p class=\"parrafosmediano\" >" + ubicacionNode.InnerText.Replace("\n", "</br>") + "</td></tr></tbody></table>";
                        }
                    } // fin del foreach
                    #endregion
                    sb.AppendLine($"<div class=\"centrar\"> {dtablexml.Trim()} </div>");
                }

                sb.AppendLine($"<p class=\"parrafosmediano\">{texto2.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p class=\"parrafosmediano\" >Guayaquil, {fecha}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{codigo}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombreQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargoQuienFirma.Trim()}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<br/> <p class=\"parrafosmediano\" ><b>NOTAS ACLARATORIAS:</b></p>");
                sb.AppendLine($"<p class=\"parrafospequeno\" >{notaAclaratoria.Replace("\n", "</br>")}</p><br/>");

                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\{codigo}.html";
                var pathToPdf = $"{_workPath}\\{codigo}.pdf";

                if (generarFileHtml)
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                if (estaLleno)
                {
                    byte[] firmaQr = (Byte[])dtSol.Rows[0]["ImagenFirma"];
                }
                else
                {
#pragma warning disable CS0219 // La variable está asignada pero nunca se usa su valor
                    byte[] firmaQr = null;
#pragma warning restore CS0219 // La variable está asignada pero nunca se usa su valor
                }


                byte[] codigoQrb = Comun.GenerarQr("Certificado " + codigo +
                  "\nNo. de Solicitud: " + tramite +
                  "\nSolicitante: " + cedula.Trim() + " " + nombre.Trim() +
                  "\nFecha de Emisión: " + Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm") +
                  "\nFecha de Vencimiento: " + Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm") +
                  $"\nDescargue su certificado en {_urlDescarga}");


                var dataFooter = new DataFooter
                {
                    Solicitante = "CED-" + cedula + " " + nombre,
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RPG GUAYAQUIL",
                    NumeroSolicitud = tramite,
                    ComprobantePago = cmpbtePago,
                    Valor = valor // Convert.ToDecimal(dtSol.Rows[0]["VALOR"]).ToString("N2")
                };

                //********* Genera el pdf del documento a partir del reporte en html *************
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _cssPath, logoRpg, logoAlcaldia, codigoQrb, dataFooter);

            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al generar el certificado.\n" + ex.Message, ex);
            }

            return procesoOk;
        }

        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int secuencial, int tipoServicio, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataSet ds = ObtenerDatos(tramite, secuencial);

                DataTable dtData = ds.Tables[0];
                DataTable dtInmuebles = ds.Tables[1];

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
                string notaAclaratoria = dtData.Rows[0]["NOTA_ACLARATORIA"].ToString().Trim();

                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombresFirmante + " " + apellidosFirmante +
                    "\nRAZON: CERTIFICADO " + certificado +
                    "\nLOCALIZACION: GUAYAQUIL" +
                    "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                    "\nNo. de Solicitud: " + dtData.Rows[0]["TRAMITE"].ToString() +
                    "\nSolicitante: " + cedula + " " + nombreSolicitante +
                    "\nFecha de Emisión: " + Convert.ToDateTime(dtData.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + Convert.ToDateTime(dtData.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm") +
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

                //var estadoPredioDesc = (dtInfoReg.Rows[0]["ESTADO_PREDIO"].ToString() == "AC") ? "ACTIVO" : "INACTIVO";

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE POSEER BIENES</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtData.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "</br>")}</p>");

                if (dtInmuebles.Rows.Count > 0)
                {
                    sb.AppendLine("<table class=\"tablaConBorde justificar\"> <thead><th class=\"fondo\">Código Catastral / Inscripciones</th><th class=\"fondo\">Linderos</th></tr></thead><tbody>");
                    foreach (DataRow inmueble in dtInmuebles.Rows)
                    {
                        var matricula = inmueble["MATRICULA_INMOBILIARIA"].ToString();
                        var registro = inmueble["REGISTRO_CATASTRAL"].ToString();
                        var ubicacion = inmueble["UBICACION"].ToString();
                        var usoInmueble = inmueble["USO"].ToString();
                        var estado = inmueble["ESTADO"].ToString();

                        sb.AppendLine($"<tr><td>{registro}</td><td class=\"justificar\">{ubicacion}</td></tr>");
                    }
                    sb.AppendLine($"</tbody></table>");
                }

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
                sb.AppendLine($"<p class=\"parrafospequeno\" >{notaAclaratoria.Replace("\\n", "</br>")}</p><br/>");

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
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(dtData.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(dtData.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = dtData.Rows[0]["TRAMITE"].ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(dtData.Rows[0]["VALOR"]).ToString("N2")}"
                };

                //genera el pdf del documento a partir del reporte en html
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

                string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, fecEmision.Year, fecEmision.Month, tipoServicio).Result.ToString();

                Dictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "okp:sispra.tramite", dtData.Rows[0]["TRAMITE"].ToString() },
                    { "okp:sispra.secuencia", secuencial.ToString() },
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

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertBienes", parametros);

            return ds;
        }

        #region"Validar campos llenos de los DataTables"
        // Función para verificar si un campo en una fila específica está lleno o vacío
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
        #endregion
    }
}
