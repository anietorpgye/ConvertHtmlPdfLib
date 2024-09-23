using SAC.CertificadosLibraryI.Auxiliares;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Xml;

namespace SAC.CertificadosLibraryI.Certificados
{
    public class NotaInscripcion : Certificado
    {
        #region"**** Variables Globales ****"
        private string _cadenaConexion = string.Empty;
        private string _workPath = string.Empty;
        private string _pdfPath = string.Empty;
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
        public NotaInscripcion(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

        /// <summary>
        /// Constructor para genear el PDF a partir de una fuente de datos XML
        /// </summary>
        /// <param name="dataXml"></param>
        /// <param name="pdfPath"></param>
        /// <param name="reportPath"></param>
        /// <param name="imagesPath"></param>
        public NotaInscripcion(string dataXml, string pdfPath, string reportPath, string imagesPath)

        {
            _workPath = dataXml;
            _pdfPath = pdfPath;
            _reportPath = reportPath;
            _imagesPath = imagesPath;
        }

        /// <summary>
        /// Genera el PDF del certificado a partir de una fuente de datos XML.
        /// </summary>
        /// <param name="ds">Dataset que contiene el xml a leer para generar el pdf.</param>
        /// <param name="generarFileHtml">Flag para generar el archivo html que sirve de base para generar el pdf</param>
        /// <returns>Booleano Verdadero si generó el PDF, Falso si hubo algun error.</returns>
        public bool GenerarPdfXml(DataSet ds, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\Logo Alcaldia Guayaquil 2.png");
                DataTable dtXml = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtXml.Rows[0]["DataXml"].ToString()));

                #region "Registro de Actos"
                string tipoCer = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;
                // string uso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string valor = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                string tramite = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;
                string tipoIdenSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_TIPO").InnerText;
                string cedulaSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_CEDULA").InnerText;
                string nombreSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_NOMBRES").InnerText;
                string fecha = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText;
                // string fechaFin = doc.DocumentElement.SelectSingleNode(@"/row/FECHA_FIN").InnerText;
                string codCert = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;

                #endregion

                #region "tabla DE LA Empresa"
                string texto = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO").InnerText;
                string texto2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;
                string empresa = doc.DocumentElement.SelectSingleNode(@"/row/EMPRESA").InnerText;
                string nombreFirma = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE_QUIENFIRMA").InnerText;
                string cargo = doc.DocumentElement.SelectSingleNode(@"/row/CARGO_QUIENFIRMA").InnerText;
                string certificado = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;
                string notaAclaratoria = doc.DocumentElement.SelectSingleNode(@"/row/NOTA_ACLARATORIA").InnerText;

                var data = nombreFirma.Split(' ');
                string apellidosFirma = $"{data[data.Length - 1]} {data[data.Length - 2]}";

                if (data.Length.Equals(4))
                {
                    nombreFirma = $"{data[0]} {data[1]}";
                }
                else if (data.Length.Equals(5))
                {
                    nombreFirma = $"{data[0]} {data[1]} {data[2]}";
                }
                #endregion

                #region tabla imagenes
                byte[] codigoQr = Comun.GenerarQr("Certificado " + codCert +
                    "\nNo. de Solicitud: " + tramite.ToString() +
                    "\nSolicitante: " + cedulaSol + " " + nombreSol +
                    "\nFecha de Emisión: " + Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm") +
                    // "\nFecha de Vencimiento: " + Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm") +
                    $"\nDescargue su certificado en {_urlDescarga}");
                #endregion

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                //sb.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: E1E0DE;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 140px;height: 70px;} th,td {border: 1px solid black; padding: 2px 5px 2px 5px ; text-align: center; font-family: Helvetica; font-size: 12px; } .contenedorFirma {  display: grid;  grid-template-columns: 1fr;  justify-content:center; border: none;  grid-gap: 0; }.contenedorFirma div{ text-align: center; height: 50%;}@media (max-width:480px){.contenedorFirma { grid-template-columns: 100%; margin-bottom: 0;}}</style>");
                sb.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: E1E0DE;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 140px;height: 70px;} th,td {border: 1px solid black; padding: 2px 5px 2px 5px ; text-align: center; font-family: Helvetica; font-size: 12px;} /* Estilos para centrar la sección */ .contenedor {  display: flex;  justify-content: center;  align-items: center;  height: 100vh; /* Ajusta la altura según sea necesario */}/* Estilos para la sección */.mi-seccion {  text-align: center; font-family: 'Helvetica', sans-serif; font-size: 9px;}/* Estilos para la imagen */.mi-seccion img {  max-width: 70%; /* Ajusta el ancho máximo de la imagen */  margin-bottom: 1px; /* Ajusta el margen inferior según sea necesario */}/* Estilos para los párrafos */.mi-seccion p {  margin-bottom: 1px; /* Ajusta el margen inferior según sea necesario */} }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE NOTA DE INSCRIPCIÓN</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {codCert}</b></p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{texto.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>REGISTRO DE ACTOS</u></b></p>");
                sb.AppendLine($"<table style=\"width: 100%; border-collapse: collapse; border-spacing=1;\">");
                sb.AppendLine($"<thead ><tr style=\"background-color: #E1E0DE;\"><th>DOCUMENTO</th><th>FECHA</th><th>OFICINA</th></tr></thead>");
                sb.AppendLine($"<tbody>");
                var Actos = doc.DocumentElement.SelectNodes(@"/row/DATOS");
                foreach (XmlNode hijo in Actos)
                {
                    var hdocument = hijo.SelectSingleNode(@"DOCUMENTO").InnerText;
                    var hfecha = hijo.SelectSingleNode(@"FECHA").InnerText;
                    var hoficina = hijo.SelectSingleNode(@"OFICINA").InnerText;
                    if (hoficina != "" || hfecha != "" || hdocument != "")
                    {
                        sb.AppendLine($"<tr class=\"centrar\"><td><p style=\"font-family: Helvetica; font-size: 10px;\">{((hdocument == "") ? "-----" : hdocument.Trim())}</p></td><td><p style=\"font-family: Helvetica; font-size: 10px;\">{hfecha}</p></td><td><p style=\"font-family: Helvetica; font-size: 10px;\">{hoficina}</p></td></tr>");
                    }
                }
                sb.AppendLine($"</tbody></table></br>");


                sb.AppendLine($"<table style=\"width: 100%; border-collapse: collapse; border-spacing=1;\">");
                sb.AppendLine($"<thead><tr><th>Repertorio</th><th>Acto</th><th>Fecha de</br>Inscripción</th><th>Libro</th><th>Número de Inscripción</th><th>Tomo</th><th>Folio<br/>Inicial</th><th>Folio<br/>Final</th></tr></thead>");
                sb.AppendLine($"<tbody>");
                var folio = doc.DocumentElement.SelectNodes(@"/row/folio/detalle");
                foreach (XmlNode hijos in folio)
                {
                    var htipo = hijos.SelectSingleNode(@"TIPO").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"TIPO").InnerText.Trim();
                    var hrepertorio = hijos.SelectSingleNode(@"REPERTORIO").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"REPERTORIO").InnerText.Trim();
                    var hlibro = hijos.SelectSingleNode(@"LIBRO").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"LIBRO").InnerText.Trim();
                    var hacto = hijos.SelectSingleNode(@"ACTO").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"ACTO").InnerText.Trim();
                    var hnuminscripcion = hijos.SelectSingleNode(@"NUM_INSCRIPCION").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"NUM_INSCRIPCION").InnerText.Trim();
                    var hfechainscripcion = hijos.SelectSingleNode(@"FECHA_INSCRIPCION").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"FECHA_INSCRIPCION").InnerText.Trim();
                    var htomo = hijos.SelectSingleNode(@"TOMO").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"TOMO").InnerText.Trim();
                    var hfolioini = hijos.SelectSingleNode(@"FOLIO_INICIAL").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"FOLIO_INICIAL").InnerText.Trim();
                    var hfoliofin = hijos.SelectSingleNode(@"FOLIO_FIN").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"FOLIO_FIN").InnerText.Trim();
                    var hsecuencia = hijos.SelectSingleNode(@"SECUENCIA").InnerText == "" ? string.Empty : hijos.SelectSingleNode(@"SECUENCIA").InnerText.Trim();

                    sb.AppendLine($"<tr><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hrepertorio}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hacto}</td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hfechainscripcion}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hlibro}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hnuminscripcion}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{htomo}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hfolioini}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hfoliofin}</p></td></tr>");
                }
                sb.AppendLine($"</tbody></table></br>");

                var matriculaProp = doc.DocumentElement.SelectNodes(@"/row/folio/detalle/matriculas/matricula");
                sb.AppendLine($"<table style=\"width: 100%; border-collapse: collapse; border-spacing=1;\">");
                sb.AppendLine($"<thead ><tr><th>Código Catastral</th><th>Matrícula INM./</br>Ficha Registral</th><th>Descripción del Inmueble</th><th>Estado</th></tr></thead>");
                sb.AppendLine($"<tbody>");
                foreach (XmlNode h in matriculaProp)
                {
                    var hcodigo = h.SelectSingleNode(@"CODIGO_CATASTRAL").InnerText == "" ? string.Empty : h.SelectSingleNode(@"CODIGO_CATASTRAL").InnerText.Trim();
                    var hmatricula = h.SelectSingleNode(@"MATRICULA").InnerText == "" ? string.Empty : h.SelectSingleNode(@"MATRICULA").InnerText.Trim();
                    var hdescrip = h.SelectSingleNode(@"DESCRIPCION").InnerText == "" ? string.Empty : h.SelectSingleNode(@"DESCRIPCION").InnerText.Trim();
                    var hestado = h.SelectSingleNode(@"ESTADO").InnerText == "" ? string.Empty : h.SelectSingleNode(@"ESTADO").InnerText.Trim();
                    sb.AppendLine($"<tr><td style=\"width: 15%;\" ><p style=\"font-family: Helvetica; font-size: 8px;\">{hcodigo}</p></td><td style=\"width: 20%;\"><p style=\"font-family: Helvetica; font-size: 8px;\">{hmatricula}</p></td><td class=\"justificar\" style=\"width: 55%; \"><p style=\"font-family: Helvetica; font-size: 8px;\">{hdescrip}</p></td><td style=\"width: 10%;\"><p style=\"font-family: Helvetica; font-size: 8px;\">{hestado}</p></td></tr>");
                }
                sb.AppendLine($"</tbody></table></br>");

                var propietario = doc.DocumentElement.SelectNodes(@"/row/folio/detalle/propietarios/propietario");

                foreach (XmlNode hj in propietario)
                {
                    if (hj.HasChildNodes == true)
                    {
                        _contador++;
                        if (_contador == 1)// la primera vez, crea la cabecera de la tabla de Propietarios.
                        {
                            sb.AppendLine($"<table style=\"width: 100%; border-collapse: collapse; border-spacing=0;\">");
                            sb.AppendLine($"<thead class=\"centrar\"><tr><th style=\"width: 20%;\">Partes</th><th>Doc.de Identidad</th><th>Nombre</th><th>Estado Civil</th></tr></thead>");
                            sb.AppendLine($"<tbody>");
                        }
                    }
                    var hpartes = hj.SelectSingleNode(@"PAPEL").InnerText == "" ? string.Empty : hj.SelectSingleNode(@"PAPEL").InnerText.Trim();
                    var hdocid = hj.SelectSingleNode(@"CEDULA").InnerText == "" ? string.Empty : hj.SelectSingleNode(@"CEDULA").InnerText.Trim();
                    var hnombre = hj.SelectSingleNode(@"CLIENTE").InnerText == "" ? string.Empty : hj.SelectSingleNode(@"CLIENTE").InnerText.Trim();
                    var hestadocivil = hj.SelectSingleNode(@"ESTADO_CIVIL");
                    _estadocivil = "";
                    if (hestadocivil == null || hestadocivil.ToString() == string.Empty)
                    {
                        _estadocivil = "-----";
                    }
                    else
                    {
                        _estadocivil = Comun.GetEstadoCivil(hestadocivil.InnerText.Trim());
                    }


                    sb.AppendLine($"<tr class=\"centrar\"><td>{hpartes}</td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hdocid}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\">{hnombre}</p></td><td><p style=\"font-family: Helvetica; font-size: 8px;\"> {_estadocivil}</p> </td></tr>");
                } // ******** Fin del For de propietarios **************
                sb.AppendLine($"</tbody></table></br>");
                _contador = 0;

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">{texto2}</p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">Guayaquil, {fecha}</p>");

                //sb.AppendLine($"<table  border: none; style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                //sb.AppendLine($"<tr border: none;><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                //sb.AppendLine($"<tr><td></td></tr>");
                //sb.AppendLine($"<tr  border: none;><td class=\"centrar\"><b>{nombreFirma} {apellidosFirma}</b></td></tr>");
                //sb.AppendLine($"<tr  border: none;><td class=\"centrar\"><b>{cargo}</b></td></tr>");
                //sb.AppendLine($"<tr  border: none;><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                //sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<div class=\"contenedor\"><section class=\"mi-seccion\"><img src=\"temporal/{certificado}.png\" class=\"imagen\"><p><b>{nombreFirma} {apellidosFirma}</b></p><p><b>{cargo}</b></p><p><b>{empresa}</b></p></section></div>");
                sb.AppendLine($"<br/> <p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>NOTAS ACLARATORIAS:</b></p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 9px;\" class=\"justificar\">{notaAclaratoria.Replace("\n", "</br>")}</p><br/>");

                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\{certificado}.html";
                var pathToPdf = $"{_reportPath}\\{certificado}.pdf";

                if (generarFileHtml)
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                var pathFirma = $"{_reportPath}\\{certificado}.png";
                //*************************************************  Valido la firma venga llena *************************
                bool estaLleno = EstaLleno(dtXml, 0, "ImagenFirma");
                if (estaLleno)
                {
                    byte[] firmaReg = (Byte[])dtXml.Rows[0]["ImagenFirma"];
                    Comun.Upload(firmaReg, pathFirma);
                }
                else
                {
#pragma warning disable CS0219 // La variable está asignada para valores NULOS
                    byte[] firmaReg = null;
#pragma warning restore CS0219 //  La variable está asignada para valores NULOS
                }

                var dataFooter = new DataFooter
                {
                    Solicitante = tipoIdenSol + " " + cedulaSol.Trim() + " " + nombreSol.Trim(),
                    ShowUso = false,
                    // Uso = uso.ToUpper(),
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    ShowFechaVencimiento = false,
                    // FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "Matríz - Rpg Guayaquil",
                    NumeroSolicitud = tramite,
                    ComprobantePago = "NO REGISTRADO",//TypeCOMPROBANTE_PAGO,
                    Valor = valor == "" ? "0,00" : valor.Trim() // Convert.ToDecimal(XmlDto.Rows[0]["VALOR"]).ToString("N2")
                };

                //***************** Genera el pdf del documento a partir del reporte en html ****************************
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                procesoOk = false;
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
                DataTable dtNotas = ds.Tables[1];
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
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16pt; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica', monospace; font-size: 10pt; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8pt; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE NOTA DE INSCRIPCIÓN</b></p>");
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
                    sb.AppendLine($"<tr><td><p class=\"parrafospequeno_sj\">{reper.Repertorio}</p></td><td><p class=\"parrafospequeno_sj\">{reper.Acto}</td><td><p class=\"parrafospequeno_sj\">{reper.FechaInscripcion.ToString("dd/MM/yyyy")}</p></td><td><p class=\"parrafospequeno_sj\">{reper.Libro}</p></td><td><p class=\"parrafospequeno_sj\">{reper.NumeroInscripcion}</p></td><td><p class=\"parrafospequeno_sj\">{reper.Tomo}</p></td><td><p class=\"parrafospequeno_sj\">{reper.FolioInicial}</p></td><td><p class=\"parrafospequeno_sj\">{reper.FolioFinal}</p></td></tr>");
                    sb.AppendLine($"</tbody></table></br>");
                    #endregion

                    #region tabla inmueble
                    sb.AppendLine($"<table class=\"tablaConBorde justificar\">");
                    sb.AppendLine($"<thead><tr><th class=\"parrafospequeno_sj\">Código Catastral</th><th class=\"parrafospequeno_sj\">Matrícula INM./</br>Ficha Registral</th><th class=\"parrafospequeno_sj\">Descripción del Inmueble</th><th class=\"parrafospequeno_sj\">Estado</th></tr></thead>");
                    sb.AppendLine($"<tbody>");
                    sb.AppendLine($"<tr><td style=\"width: 15%;\" ><p class=\"parrafospequeno_sj\">{reper.Catastral}</p></td><td style=\"width: 20%;\"><p class=\"parrafospequeno_sj\">{reper.Matricula}</p></td><td class=\"justificar\" style=\"width: 55%; \"><p class=\"parrafospequeno\">{reper.Descripcion}</p></td><td style=\"width: 10%;\"><p class=\"parrafospequeno_sj\">{reper.Estado}</p></td></tr>");
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

                //ManipulatePdf(pathToHtml, pathToPdf, _cssPath, logoRpg, logoAlcaldia, codigoQrb, dataFooter);
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

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertNotaInscripcion", parametros);

            return ds;
        }

        /// <summary>
        /// Función para verificar si un campo en una fila del datatable está lleno o vacío
        /// </summary>
        /// <param name="dataTable">DataTable con los datos</param>
        /// <param name="filaIndice">Indice de la fila</param>
        /// <param name="nombreCampo">Nombre del campor a validar</param>
        /// <returns>Verdadero si el campo tiene datos, Falso si está vacio.</returns>
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
