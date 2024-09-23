using GestorDocumentalLibrary;
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
    public class Pronunciamiento : Certificado
    {
        private string _cadenaConexion = string.Empty;
        private string _workPath = string.Empty;
        private string _reportPath = string.Empty;
        private string _cssPath = string.Empty;
        private string _imagesPath = string.Empty;
        private string _signedDocPath = string.Empty;
        string _cssContent = string.Empty;
        private string _urlDescarga = string.Empty;
        private string _hostOpenKm = string.Empty;
        private string _userOpenKm = string.Empty;
        private string _passwordOpenKm = string.Empty;

        public Pronunciamiento(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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
        public Pronunciamiento(string workPath, string reportPath, string cssPath, string imagesPath) // by Acueva 2024-02 itext pdf Xml
        {
            _workPath = workPath;
            _reportPath = reportPath;
            _cssPath = cssPath;
            _imagesPath = imagesPath;
        }

        #region"GenerarPdfXml .... Métodos Principales ..."
        protected string CargaCss(string rutaArchivo) // override async Task OnInitializedAsync()
        {
            string contenidoCSS = File.ReadAllText(rutaArchivo + "/reporte.css");
            return contenidoCSS;
        }

        public bool GenerarPdfXml(DataSet ds, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataTable dtSol = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtSol.Rows[0]["DataXml"].ToString()));

                #region " *************** Cuerpo del Reporte ************ "
                string cedula = doc.DocumentElement.SelectSingleNode(@"/row/CEDULA").InnerText;
                string tipo = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;

                string uso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string valor = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                if (valor == "" || valor == null) { valor = "0.00"; }

                string banco = doc.DocumentElement.SelectSingleNode(@"/row/BANCO").InnerText;

                string cmpbtePago = "NO REGISTRADO";

                string tramite = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;

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
                string texto2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;
                string revisor = doc.DocumentElement.SelectSingleNode(@"/row/revisor/ITEM").InnerText;

                #endregion

                #region" Datos de la Empresa"
                //  .....................  Datos de la empresa .................................
                string empresa = doc.DocumentElement.SelectSingleNode(@"/row/EMPRESA").InnerText;
                string notaAclaratoria = doc.DocumentElement.SelectSingleNode(@"/row/NOTA_ACLARATORIA").InnerText;
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
#pragma warning disable CS0219 // La variable está asignada pero su valor es nulo
                    byte[] firmaQr = null;
#pragma warning restore CS0219 // La variable está asignada pero su valor es nulo
                }
                #endregion

                string dtablexml;
                _cssContent = CargaCss(_cssPath); // carga el archivo para trabajar los css de los reportes.
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                // sb.AppendLine("<style> html { position: relative; min-height: 100%;}body { margin: 0px; font-family: Consolas; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ }qr {border:solid 1px blue; height:80px; width:80px;}@page {margin-top: 120pt;margin-bottom: 120pt;}.centrar {text-align: center;}.justificar {text-align: justify;} .imagen {width: 130px;height: 70px;} th.fondo {background-color: #E1E0DE; padding: 10px;} .parrafosgrande {text-align: center;font-family: 'Consolas', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Consolas', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Consolas', monospace; font-size: 8pt; text-align: justify;} .firmamediana {text-align: center; font-family: 'Consolas', monospace; font-size: 8pt;} </style>");
                sb.AppendLine("<style>" + _cssContent + "</style>");
                sb.AppendLine("</head><body>");
                //sb.AppendLine($"<p style=\"text-align: center; font-family: 'Consolas', font-size: 16pt;\"><b>CERTIFICADO DE POSEER BIENES</b></p>");
                sb.AppendLine($"<p class=\"parrafosgrande\" >CERTIFICADO REGISTRAL</p>");
                //sb.AppendLine($"<hr><p style=\"text-align: center; font-family: 'Consolas', font-size: 16pt;\"><b>No. {TypeCODIGO}</b></p>");
                sb.AppendLine($"<hr><p class=\"parrafosgrande\" >No. {codigo}</p>");
                sb.AppendLine($"<p class=\"parrafosmediano\" >{texto.Replace("\n", "</br>")}</p>");
                if (revisor != null) // ***** SI hay registros.. lleno la tabla con sus nodos *****
                {
                    #region"Validar campos para tabla"
                    // Select nodes using XPath
                    XmlNodeList nodeObsList = doc.SelectNodes(@"/row/devolutiva/OBSERVACION"); // OBSERVACION
                    XmlNodeList nodeRevisorList = doc.SelectNodes(@"/row/devolutiva/revisor"); // revisor

                    dtablexml = "";
                    // Iterate through each detalle1 node
                    foreach (XmlNode detalleNode in nodeObsList)
                    {
                        // Accedo a los nodos del hijo
                        XmlNode obsNode = detalleNode.SelectSingleNode("OBSERVACION");
                        dtablexml += "<center><p class=\"parrafospequeno\" >" + obsNode.InnerText.Replace("\n", "</br>") + "</p></center>";

                    } // fin del foreach
                    int i = 0;
                    foreach (XmlNode dtNode in nodeRevisorList) // revisor / ITEM
                    {
                        if (dtNode != null)
                        {
                            if (i == 0 && dtNode.ChildNodes.Count <= 1) // siempre es un registro en bd
                            {
                                dtablexml += "</br> <center><table><tr>";
                                XmlNode itemNode = dtNode.SelectSingleNode("ITEM");
                                dtablexml += "<td><p class=\"parrafospequeno\" >" + itemNode.InnerText.Replace("\n", "</br>") + "</p></td></tr></table></center>";

                            }

                            i++;
                            if (i > 0 && dtNode.ChildNodes.Count > 1)
                            {
                                XmlNode itemNode = dtNode.SelectSingleNode("ITEM");
                                dtablexml += "<center><p class=\"parrafospequeno\" >" + itemNode.InnerText.Replace("\n", "</br>") + "</p></center>";
                            }
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
                    Solicitante = "CED-" + cedula + "" + nombre,
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RPG GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
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
        #endregion


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
        #endregion

        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int secuencial, int tipoServicio)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\Logo Alcaldia Guayaquil 2.png");

                List<SqlParameter> parametros = new List<SqlParameter>
                {
                    new SqlParameter("@SOLINUMTRA", tramite),
                    new SqlParameter("@SOLISEC", secuencial),
                };

                var resul = Comun.ActualizarVigencia(_cadenaConexion, tramite, secuencial);

                DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertPronunciamiento", parametros);

                DataTable dtSol = ds.Tables[0];
                DataTable dtData = ds.Tables[1];

                string nombreFirmante = dtSol.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim() + " " + dtSol.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = dtSol.Rows[0]["CARGO_QUIENFIRMA"].ToString();
                string empresa = dtSol.Rows[0]["EMPRESA"].ToString();

                string certificado = dtData.Rows[0]["Codigo"].ToString();
                var body = dtData.Rows[0]["Comentario"].ToString();

                var pathToHtml = $"{_reportPath}\\pronunciamiento.html";
                var pathToPdf = $"{_workPath}\\{tramite}_{secuencial}.pdf";

                GenerarHtml(pathToHtml, "CERTIFICACIÓN REGISTRAL", certificado, body);

                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombreFirmante +
                    "\nRAZON: CERTIFICADO " + certificado +
                    "\nLOCALIZACION: GUAYAQUIL" +
                    "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                    "\nNo. de Solicitud: " + dtSol.Rows[0]["TRAMITE"].ToString() +
                    "\nSolicitante: " + dtSol.Rows[0]["IDEN_SOLICITANTE"].ToString() + " " + dtSol.Rows[0]["NOMBRE_SOLICITANTE"].ToString() +
                    "\nFecha de Emisión: " + Convert.ToDateTime(dtSol.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + Convert.ToDateTime(dtSol.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nDescargue su certificado en " + dtSol.Rows[0]["URL_CONSULTA"].ToString());

                var imgFirmaQr = Comun.ImageFromByteArray(firmaQr);

                DateTime fecEmision = Convert.ToDateTime(dtSol.Rows[0]["FECHA"]);
                var nombresFirmante = dtSol.Rows[0]["NOMBRES_QUIENFIRMA"].ToString();
                var apellidosFirmante = dtSol.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString();

                Comun.DibujarFirma(new ImagenFirma
                {
                    ColorTexto = System.Drawing.Color.Black,
                    ColorBackground = System.Drawing.Color.White,
                    Fuente = "Calibri",
                    TamanoFuente = 12,
                    FirmaQr = imgFirmaQr,
                    NombresFirmante = nombresFirmante,
                    ApellidosFirmante = apellidosFirmante,
                    Cargo = cargoQuienFirma,
                    Empresa = empresa,
                    Alto = 185,
                    Ancho = 490,
                    RutaImagen = _workPath + "\\firma.png"
                });

                var dataFooter = new DataFooter
                {
                    Solicitante = dtSol.Rows[0]["IDEN_SOLICITANTE"].ToString() + " " + dtSol.Rows[0]["NOMBRE_SOLICITANTE"].ToString(),
                    Uso = dtSol.Rows[0]["USO"].ToString(),
                    FechaEmision = Convert.ToDateTime(dtSol.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(dtSol.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RPG GUAYAQUIL",
                    NumeroSolicitud = tramite.ToString(),
                    ComprobantePago = dtSol.Rows[0]["CMPBTE_PAGO"].ToString(),
                    Valor = Convert.ToDecimal(dtSol.Rows[0]["VALOR"]).ToString("N2")
                };

                //genera el pdf del documento a partir del reporte en html
                Comun.ManipulatePdf(pathToHtml, pathToPdf, _cssPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

                #region firma del documento y carga a OpenKm
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
                        Nombre = $"{tramite}_{secuencial}",
                        PathSinFirmar = _workPath,
                        PathFirmado = _signedDocPath,
                        CargarEnGestor = true,
                        RepositorioOkm = repositorioOkm,
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
            catch (GestorException ex)
            {
                throw new CertificadosLibraryException("Error al importar el documento al Gestor Documental.\n" + ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al generar el certificado.\n" + ex.Message, ex);
            }

            return procesoOk;
        }

        public static void GenerarHtml(string pathHtml, string nombreCertificado, string codigoCertificado, string body)
        {
            FileStream fs = null;

            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<link rel=\"stylesheet\" type=\"text/css\" href=\"../resources/css/estilos.css\" />");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-size: 18.0pt;\"><b>{nombreCertificado}</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-size: 16.0pt;\"><b>{codigoCertificado}</b></p>");
                sb.AppendLine(body);
                sb.AppendLine("<br/><br/><br/>");
                sb.AppendLine("<div class=\"centrar\"><img src=\"../temporal/firma.png\"></div><P style=\"text-align: center;\"></P>");
                sb.AppendLine("</body></html>");

                fs = new FileStream(pathHtml, FileMode.Create);
                using (StreamWriter writer = new StreamWriter(fs))
                {
                    writer.Write(sb.ToString());
                }
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }
        }


    }
}
