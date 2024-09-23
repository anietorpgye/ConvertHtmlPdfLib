using SAC.CertificadosLibraryI.Auxiliares;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Xml;

namespace SAC.CertificadosLibraryI.Certificados
{
    public class HistoriaDominio : Certificado
    {
        private string _cadenaConexion = string.Empty;
        private string _data = string.Empty;
        private string _pdfPath = string.Empty;
        private string _reportPath = string.Empty;
        private string _imagesPath = string.Empty;
        private string _signedDocPath = string.Empty;
        private string _urlDescarga = string.Empty;
        private string _hostOpenKm = string.Empty;
        private string _userOpenKm = string.Empty;
        private string _passwordOpenKm = string.Empty;

        public HistoriaDominio(string dataXml, string pdfPath, string reportPath, string imagesPath)
        {
            _data = dataXml;
            _pdfPath = pdfPath;
            _reportPath = reportPath;
            _imagesPath = imagesPath;
        }

        public HistoriaDominio(string cadenaConexion, string workPath, string reportPath, string imagesPath, string signedDocPath, string urlDescarga)
        {
            _cadenaConexion = cadenaConexion;
            _pdfPath = workPath;
            _reportPath = reportPath;
            _imagesPath = imagesPath;
            _signedDocPath = signedDocPath;
            _urlDescarga = urlDescarga;
            _hostOpenKm = ObtenerParametro(_cadenaConexion, PARAM_HOST_OPENKM).Result;
            _userOpenKm = ObtenerParametro(_cadenaConexion, PARAM_USER_OPENKM).Result;
            _passwordOpenKm = ObtenerParametro(_cadenaConexion, PARAM_PWD_OPENKM).Result;
        }

        public bool GenerarPdfXml(DataSet ds)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\Logo Alcaldia Guayaquil 2.png");

                DataTable dtXml = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtXml.Rows[0]["DataXml"].ToString()));

                #region tabla solicitud
                string tipoCer = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;
                string planHab = doc.DocumentElement.SelectSingleNode(@"/row/PLAN_HABITACIONAL").InnerText;
                string uso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string valor = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                string tramite = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;
                string tipoIdenSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_TIPO").InnerText;
                string cedulaSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_CEDULA").InnerText;
                string nombreSol = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_NOMBRES").InnerText;
                string fecha = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText;
                string fechaFin = doc.DocumentElement.SelectSingleNode(@"/row/FECHA_FIN").InnerText;
                string codCert = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;
                #endregion

                #region tabla empresa
                string texto = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO").InnerText;
                string texto2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;
                string catastral = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO_CATASTRAL").InnerText;
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

                #region tabla propietarios
                var propietarios = doc.DocumentElement.SelectNodes(@"/row/propietarios/propietario");
                foreach (XmlNode propietario in propietarios)
                {
                    var tipoCedProp = propietario.SelectSingleNode(@"TIPO_CED").InnerText;
                    var nombreProp = propietario.SelectSingleNode(@"NOMBRE").InnerText;
                    var identiProp = propietario.SelectSingleNode(@"IDENTIFICACION").InnerText;
                }
                #endregion

                #region tabla info registral
                string ubicacion = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/UBICACION").InnerText;
                string estadoMuni = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/ESTADO_MUNICIPAL").InnerText;
                string estadoDepu = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/ESTADO_DEPURACION").InnerText;
                string estadoPredio = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/ESTADO_PREDIO").InnerText;
                string norte = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/NORTE").InnerText;
                string sur = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/SUR").InnerText;
                string este = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/ESTE").InnerText;
                string oeste = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/OESTE").InnerText;
                string areaEscr = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/AREAESCRITURA").InnerText;
                string fondoEscr = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FONDOESCRITURA").InnerText;
                string areaLevan = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/AREALEVANTAMIENTO").InnerText;
                string fondoLevan = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FONDOLEVANTAMINTO").InnerText;
                string frenteEscr = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FRENTEESCRITURA").InnerText;
                string frente1 = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FRENTE1").InnerText;
                string frente2 = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FRENTE2").InnerText;
                string frente3 = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FRENTE3").InnerText;
                string frente4 = doc.DocumentElement.SelectSingleNode(@"/row/catastro/forma/FRENTE4").InnerText;
                string estadoEdif = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/ESTADO_EDIFICACION").InnerText;
                string usoEdif = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/USO_EDIFICACION").InnerText;
                string alumbrado = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/ALUMBRADO").InnerText;
                string alcantarillado = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/ALCANTARILLADO").InnerText;
                string agua = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/AGUA").InnerText;
                string redTelef = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/REDTELEFONICA").InnerText;
                string pavimentacion = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/PAVIMENTACION").InnerText;
                string bordillo = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/BORDILLO").InnerText;
                string acera = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/ACERA").InnerText;
                string esquinero = doc.DocumentElement.SelectSingleNode(@"/row/catastro/complementario/ESQUINEROUSOMERIDIONAL").InnerText;
                #endregion

                #region tabla movimientos
                var movimientos = doc.DocumentElement.SelectNodes(@"/row/registros/registro");
                foreach (XmlNode movi in movimientos)
                {
                    var tipoReg = movi.SelectSingleNode(@"TIPO_REG").InnerText;
                    var conteo = movi.SelectSingleNode(@"CONTEO").InnerText;
                    var acto = movi.SelectSingleNode(@"ACTO").InnerText;
                    var fechaInsMov = movi.SelectSingleNode(@"FECHA_INSC").InnerText;
                    var tomo = movi.SelectSingleNode(@"TOMO").InnerText;
                    var numInsc = movi.SelectSingleNode(@"NUMERO_INSC").InnerText;
                    var numReper = movi.SelectSingleNode(@"NUMERO_REPER").InnerText;
                    var oficina = movi.SelectSingleNode(@"OFICINA_ORIGINAL").InnerText;
                    var canton = movi.SelectSingleNode(@"CANTON").InnerText;
                    var estProvResol = movi.SelectSingleNode(@"EST_PROV_RESOL")!=null? movi.SelectSingleNode(@"EST_PROV_RESOL").InnerText:"";
                    var oficio = movi.SelectSingleNode(@"OFICIO_TELEX").InnerText;
                    var observacion = (movi.SelectSingleNode(@"OBSERVACION") == null) ? string.Empty : movi.SelectSingleNode(@"OBSERVACION").InnerText;
                }
                #endregion

                #region tabla imagenes
                byte[] codigoQr = Comun.GenerarQr("Certificado " + codCert +
                    "\nNo. de Solicitud: " + tramite.ToString() +
                    "\nSolicitante: " + cedulaSol + " " + nombreSol +
                    "\nFecha de Emisión: " + Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm") +
                    $"\nDescargue su certificado en {_urlDescarga}");
                #endregion

                var estadoPredioDesc = (estadoPredio == "AC") ? "ACTIVO" : "INACTIVO";

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 140px;height: 70px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 18px;\"><b>CERTIFICADO DE HISTORIA DE DOMINIO</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 16px;\"><b>No. {codCert}</b></p>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 16px;\" >CÓDIGO CATASTRAL No. {catastral}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{texto.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>1. INFORMACIÓN DEL PREDIO</u></b></p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{ubicacion}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\"><b>Titular(es) Registrado(s):</b></p>");
                sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                foreach (XmlNode propietario in propietarios)
                {
                    var tipoCedProp = propietario.SelectSingleNode(@"TIPO_CED").InnerText;
                    var nombreProp = propietario.SelectSingleNode(@"NOMBRE").InnerText;
                    var identiProp = propietario.SelectSingleNode(@"IDENTIFICACION").InnerText;

                    sb.AppendLine($"<tr><td>{tipoCedProp}</td><td>{identiProp}</td><td>{nombreProp}</td></tr>");
                }
                sb.AppendLine($"</tbody></table>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>Estado:</b> {estadoPredioDesc}</p>");

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>2. INFORMACIÓN DEL CATASTRO MUNICIPAL</u></b></p>");
                sb.AppendLine($"<div class=\"row\"><table style=\"width: 100%;\"><tbody>");
                sb.AppendLine($"<tr><td style=\"width: 10%;\"><b>Norte:</b></td><td style=\"width: 90;\">{norte}</td></tr>");
                sb.AppendLine($"<tr><td><b>Sur:</b></td><td>{sur}</td></tr>");
                sb.AppendLine($"<tr><td><b>Este:</b></td><td>{este}</td></tr>");
                sb.AppendLine($"<tr><td><b>Oeste:</b>   </td><td>{oeste}</td></tr>");
                sb.AppendLine($"</tbody></table></div>");
                sb.AppendLine($"<div class=\"row\"><div class=\"column\"><table><tbody>");
                sb.AppendLine($"<tr><td><b>Área en Escrituras:</b></td><td>{areaEscr}</td></tr>");
                sb.AppendLine($"<tr><td><b>Fondo en Escrituras:</b></td><td>{fondoEscr}</td></tr>");
                sb.AppendLine($"<tr><td><b>Área en Levantamiento:</b></td><td>{areaLevan}</td></tr>");
                sb.AppendLine($"<tr><td><b>Fondo en Levantamiento:</b></td><td>{fondoLevan}</td></tr>");
                sb.AppendLine($"<tr><td><b>Tipo de Terreno:</b></td><td>{esquinero}</td></tr>");
                sb.AppendLine($"<tr><td><b>Estado:</b></td><td>{estadoEdif}</td></tr>");
                sb.AppendLine($"<tr><td><b>Uso:</b></td><td>{usoEdif}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente en Escritura:</b></td><td>{frenteEscr}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 1:</b></td><td>{frente1}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 2:</b></td><td>{frente2}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 3:</b></td><td>{frente3}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 4:</b></td><td>{frente4}</td></tr>");
                sb.AppendLine($"</tbody></table></div>");
                sb.AppendLine($"<div class=\"column\"><table><tbody>");
                sb.AppendLine($"<tr><td><b>Alumbrado:</b></td><td>{alumbrado}</td></tr>");
                sb.AppendLine($"<tr><td><b>Alcantarillado:</b></td><td>{alcantarillado}</td></tr>");
                sb.AppendLine($"<tr><td><b>Agua Potable:</b></td><td>{agua}</td></tr>");
                sb.AppendLine($"<tr><td><b>Red Telefónica:</b></td><td>{redTelef}</td></tr>");
                sb.AppendLine($"<tr><td><b>Pavimentación:</b></td><td>{pavimentacion}</td></tr>");
                sb.AppendLine($"<tr><td><b>Bordillo:</b></td><td>{bordillo}</td></tr>");
                sb.AppendLine($"<tr><td><b>Acera:</b></td><td>{acera}</td></tr>");
                sb.AppendLine($"</tbody></table></div></div>");

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>3. MOVIMIENTOS DEL PREDIO EN ORDEN CRONOLÓGICO</u></b></p>");
                foreach (XmlNode movi in movimientos)
                {
                    var tipoReg = movi.SelectSingleNode(@"TIPO_REG").InnerText;
                    var conteo = movi.SelectSingleNode(@"CONTEO").InnerText;
                    var acto = movi.SelectSingleNode(@"ACTO").InnerText;
                    var fechaInsMov = movi.SelectSingleNode(@"FECHA_INSC").InnerText;
                    var tomo = movi.SelectSingleNode(@"TOMO").InnerText;
                    var numInsc = movi.SelectSingleNode(@"NUMERO_INSC").InnerText;
                    var numReper = movi.SelectSingleNode(@"NUMERO_REPER").InnerText;
                    var oficina = movi.SelectSingleNode(@"OFICINA_ORIGINAL").InnerText;
                    var canton = movi.SelectSingleNode(@"CANTON").InnerText;
                    var estProvResol = movi.SelectSingleNode(@"EST_PROV_RESOL").InnerText;
                    var oficio = movi.SelectSingleNode(@"OFICIO_TELEX").InnerText;
                    var observacion = (movi.SelectSingleNode(@"OBSERVACION") == null) ? string.Empty : movi.SelectSingleNode(@"OBSERVACION").InnerText;
                    var partes = movi.SelectNodes(@"./partes/parte");

                    sb.AppendLine($"<p><table style=\"width: 100%; border-collapse: collapse; border-spacing=0;\"><tbody>");
                    sb.AppendLine($"<tr style=\"background-color: #D5DBDB;\"><td><b>No. {conteo}</b></td><td><b>{fechaInsMov}</b></td><td><b>{tipoReg}</b></td><td><b>{acto}</b></td></tr>");
                    sb.AppendLine($"</tbody></table>");
                    sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Tomo:</b></td><td style=\"width: 80%;\">{tomo}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>No. Inscripción:</b></td><td>{numInsc}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>No. Repertorio:</b></td><td>{numReper}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Cantón:</b></td><td>{canton}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Origen:</b></td><td>{oficina}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Oficio:</b></td><td>{oficio}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Esc./Pro./Res.:</b></td><td>{estProvResol}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%; vertical-align: text-top;\"><b>Observación:</b></td><td style=\"text-align: justify;\">{observacion}</td></tr>");
                    sb.AppendLine($"<tr><td style=\"width: 20%; vertical-align: text-top;\"><b>Partes:</b></td><td>");

                    sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                    foreach (XmlNode parte in partes)
                    {
                        var papel = parte.SelectSingleNode(@"PAPEL").InnerText;
                        var idenParte = parte.SelectSingleNode(@"IDENTIFICACION").InnerText;
                        var nomParte = parte.SelectSingleNode(@"NOMBRE").InnerText;
                        var estCivil = (parte.SelectSingleNode(@"ESTADO_CIVIL") == null) ? string.Empty : parte.SelectSingleNode(@"ESTADO_CIVIL").InnerText;
                        var domiParte = parte.SelectSingleNode(@"DOMICILIO").InnerText;
                        var fecInsParte = parte.SelectSingleNode(@"FECHA_INSC").InnerText;
                        var numInsParte = parte.SelectSingleNode(@"NUM_INSC").InnerText;
                        var libroParte = parte.SelectSingleNode(@"LIBRO").InnerText;

                        sb.AppendLine($"<tr><td>{papel}</td><td>{idenParte}</td><td>{nomParte}</td><td>{Comun.GetEstadoCivil(estCivil)}</td><td>{domiParte}</td></tr>");
                    }
                    sb.AppendLine($"</tbody></table>");

                    sb.AppendLine("</td></tr>");
                    sb.AppendLine($"</tbody></table></p>");

                }

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">{texto2}</p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">Guayaquil, {fecha}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombreFirma} {apellidosFirma}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargo}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<br/> <p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>NOTAS ACLARATORIAS:</b></p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 9px;\" class=\"justificar\">{notaAclaratoria.Replace("\n", "</br>")}</p><br/>");

                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\{certificado}.html";
                var pathToPdf = $"{_reportPath}\\{certificado}.pdf";

                //GenerarHtml(pathToHtml, sb.ToString());

                var pathFirma = $"{_reportPath}\\{certificado}.png";
                byte[] firmaReg = (Byte[])dtXml.Rows[0]["ImagenFirma"];

                //var imgFirmaManual = Comun.ImageFromByteArray(firmaReg);
                Comun.Upload(firmaReg, pathFirma);

                var dataFooter = new DataFooter
                {
                    Solicitante = "CED-" + tipoIdenSol + "" + nombreSol,
                    Uso = uso,
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RPG GUAYAQUIL",
                    NumeroSolicitud = tramite,
                    ComprobantePago = "",
                    Valor = valor // Convert.ToDecimal(dtSol.Rows[0]["VALOR"]).ToString("N2")
                };

                //genera el pdf del documento a partir del reporte en html

                //ManipulatePdf(pathToHtml, pathToPdf, _cssPath, logoRpg, logoAlcaldia, codigoQrb, dataFooter);
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

            }
            catch (Exception ex)
            {
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

                DataTable dt = ds.Tables[0];
                DataTable dtPropietarios = ds.Tables[1];
                DataTable dtEmpr = ds.Tables[2];
                DataTable dtNota = ds.Tables[3];
                DataTable dtInfoReg = ds.Tables[4];
                DataTable dtMovs = ds.Tables[5];

                string certificado = dtEmpr.Rows[0]["CERTIFICADO"].ToString();
                string nombreFirmante = dtEmpr.Rows[0]["NOMBRES_QUIENFIRMA"].ToString().Trim() + " " + dtEmpr.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString().Trim();
                string cargoQuienFirma = dtEmpr.Rows[0]["CARGO_QUIENFIRMA"].ToString();
                string empresa = dtEmpr.Rows[0]["EMPRESA"].ToString();

                byte[] firmaQr = Comun.GenerarQr("FIRMADO POR: " + nombreFirmante +
                    "\nRAZON: CERTIFICADO " + certificado +
                    "\nLOCALIZACION: GUAYAQUIL" +
                    "\nFECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"));

                byte[] codigoQr = Comun.GenerarQr("Certificado " + certificado +
                    "\nNo. de Solicitud: " + dt.Rows[0]["TRAMITE"].ToString() +
                    "\nSolicitante: " + dt.Rows[0]["IDEN_SOLICITANTE"].ToString() + " " + dt.Rows[0]["NOMBRE_SOLICITANTE"].ToString() +
                    "\nFecha de Emisión: " + Convert.ToDateTime(dt.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm") +
                    "\nFecha de Vencimiento: " + Convert.ToDateTime(dt.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm") +
                    $"\nDescargue su certificado en {_urlDescarga}");

                DateTime fecEmision = Convert.ToDateTime(dt.Rows[0]["FECHA"]);

                var imgFirmaQr = Comun.ImageFromByteArray(firmaQr);

                var nombresFirmante = dtEmpr.Rows[0]["NOMBRES_QUIENFIRMA"].ToString();
                var apellidosFirmante = dtEmpr.Rows[0]["APELLIDOS_QUIENFIRMA"].ToString();

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
                    RutaImagen = pathImagenFirma//@"C:\Users\avizuete.RPG\Pictures\firma chb.jpg"
                });

                var estadoPredioDesc = (dtInfoReg.Rows[0]["ESTADO_PREDIO"].ToString() == "AC") ? "ACTIVO" : "INACTIVO";

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE HISTORIA DE DOMINIO</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtEmpr.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 16px;\" >CÓDIGO CATASTRAL No. {dtEmpr.Rows[0]["CODIGO_CATASTRAL"].ToString()}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{dtEmpr.Rows[0]["TEXTO"].ToString().Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>1. INFORMACIÓN DEL PREDIO</u></b></p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{dtInfoReg.Rows[0]["UBICACION"].ToString()}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\"><b>Titular(es) Registrado(s):</b></p>");
                sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                foreach (DataRow propietario in dtPropietarios.Rows)
                {
                    var tipoCedProp = propietario["TIPO_CED"].ToString();
                    var nombreProp = propietario["NOMBRE"].ToString();
                    var identiProp = propietario["IDENTIFICACION"].ToString();

                    sb.AppendLine($"<tr><td>{tipoCedProp}</td><td>{identiProp}</td><td>{nombreProp}</td></tr>");
                }
                sb.AppendLine($"</tbody></table>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>Estado:</b> {estadoPredioDesc}</p>");

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>2. INFORMACIÓN DEL CATASTRO MUNICIPAL</u></b></p>");
                sb.AppendLine($"<div class=\"row\"><table style=\"width: 100%;\"><tbody>");
                sb.AppendLine($"<tr><td style=\"width: 10%;\"><b>Norte:</b></td><td style=\"width: 90;\">{dtInfoReg.Rows[0]["NORTE"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Sur:</b></td><td>{dtInfoReg.Rows[0]["SUR"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Este:</b></td><td>{dtInfoReg.Rows[0]["ESTE"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Oeste:</b>   </td><td>{dtInfoReg.Rows[0]["OESTE"].ToString()}</td></tr>");
                sb.AppendLine($"</tbody></table></div>");
                sb.AppendLine($"<div class=\"row\"><div class=\"column\"><table><tbody>");
                sb.AppendLine($"<tr><td><b>Área en Escrituras:</b></td><td>{dtInfoReg.Rows[0]["AREAESCRITURA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Fondo en Escrituras:</b></td><td>{dtInfoReg.Rows[0]["FONDOESCRITURA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Área en Levantamiento:</b></td><td>{dtInfoReg.Rows[0]["AREALEVANTAMIENTO"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Fondo en Levantamiento:</b></td><td>{dtInfoReg.Rows[0]["FONDOLEVANTAMINTO"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Tipo de Terreno:</b></td><td>{dtInfoReg.Rows[0]["ESQUINEROUSOMERIDIONA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Estado:</b></td><td>{dtInfoReg.Rows[0]["ESTADO_EDIFICACION"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Uso:</b></td><td>{dtInfoReg.Rows[0]["USO_EDIFICACION"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente en Escritura:</b></td><td>{dtInfoReg.Rows[0]["FRENTEESCRITURA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 1:</b></td><td>{dtInfoReg.Rows[0]["FRENTE1"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 2:</b></td><td>{dtInfoReg.Rows[0]["FRENTE2"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 3:</b></td><td>{dtInfoReg.Rows[0]["FRENTE3"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Frente 4:</b></td><td>{dtInfoReg.Rows[0]["FRENTE4"].ToString()}</td></tr>");
                sb.AppendLine($"</tbody></table></div>");
                sb.AppendLine($"<div class=\"column\"><table><tbody>");
                sb.AppendLine($"<tr><td><b>Alumbrado:</b></td><td>{dtInfoReg.Rows[0]["ALUMBRADO"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Alcantarillado:</b></td><td>{dtInfoReg.Rows[0]["ALCANTARILLADO"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Agua Potable:</b></td><td>{dtInfoReg.Rows[0]["AGUA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Red Telefónica:</b></td><td>{dtInfoReg.Rows[0]["REDTELEFONICA"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Pavimentación:</b></td><td>{dtInfoReg.Rows[0]["PAVIMENTACION"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Bordillo:</b></td><td>{dtInfoReg.Rows[0]["BORDILLO"].ToString()}</td></tr>");
                sb.AppendLine($"<tr><td><b>Acera:</b></td><td>{dtInfoReg.Rows[0]["ACERA"].ToString()}</td></tr>");
                sb.AppendLine($"</tbody></table></div></div>");

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>3. MOVIMIENTOS DEL PREDIO EN ORDEN CRONOLÓGICO</u></b></p>");

                int conteoPrevio = 0; bool hayTablaAbierta = false;
                foreach (DataRow movi in dtMovs.Rows)
                {
                    var tipoReg = movi["TIPO_REG"].ToString();
                    var conteo = Convert.ToInt32(movi["CONTEO"]);
                    var acto = movi["ACTO"].ToString();
                    var fechaInsMov = movi["FECHA_INSC"].ToString();
                    var tomo = movi["TOMO"].ToString();
                    var numInsc = movi["NUMERO_INSC"].ToString();
                    var numReper = movi[@"NUMERO_REPER"].ToString();
                    var oficina = movi["OFICINA_ORIGINAL"].ToString();
                    var canton = movi["CANTON"].ToString();
                    var estProvResol = movi["EST_PROV_RESOL"].ToString();
                    var oficio = movi["OFICIO_TELEX"].ToString();
                    var observacion = (movi["OBSERVACION"] == null) ? string.Empty : movi["OBSERVACION"].ToString();

                    var papel = movi["PAPEL"].ToString();
                    var idenParte = movi["IDENTIFICACION"].ToString();
                    var nomParte = movi["NOMBRE"].ToString();
                    var estCivil = (movi["ESTADO_CIVIL"] == null) ? string.Empty : movi["ESTADO_CIVIL"].ToString();
                    var domiParte = movi["DOMICILIO"].ToString();
                    var fecInsParte = movi["FECHA_INSC"].ToString();
                    var numInsParte = movi["NUM_INSC"].ToString();
                    var libroParte = movi["LIBRO"].ToString();

                    if (conteo != conteoPrevio)
                    {
                        if (hayTablaAbierta)
                        {
                            sb.AppendLine($"</tbody></table>"); //cierra la tabla de partes
                            sb.AppendLine("</td></tr>");
                            sb.AppendLine($"</tbody></table></p>"); //cierra la tabla del movimiento

                            hayTablaAbierta = false;
                        }
                        sb.AppendLine($"<p><table style=\"width: 100%; border-collapse: collapse; border-spacing=0;\"><tbody>");
                        sb.AppendLine($"<tr style=\"background-color: #D5DBDB;\"><td><b>No. {conteo}</b></td><td><b>{fechaInsMov}</b></td><td><b>{tipoReg}</b></td><td><b>{acto}</b></td></tr>");
                        sb.AppendLine($"</tbody></table>");
                        sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Tomo:</b></td><td style=\"width: 80%;\">{tomo}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>No. Inscripción:</b></td><td>{numInsc}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>No. Repertorio:</b></td><td>{numReper}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Cantón:</b></td><td>{canton}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Origen:</b></td><td>{oficina}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Oficio:</b></td><td>{oficio}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%;\"><b>Esc./Pro./Res.:</b></td><td>{estProvResol}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%; vertical-align: text-top;\"><b>Observación:</b></td><td style=\"text-align: justify;\">{observacion}</td></tr>");
                        sb.AppendLine($"<tr><td style=\"width: 20%; vertical-align: text-top;\"><b>Partes:</b></td><td>");
                        sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                        sb.AppendLine($"<tr><td>{papel}</td><td>{idenParte}</td><td>{nomParte}</td><td>{Comun.GetEstadoCivil(estCivil)}</td><td>{domiParte}</td></tr>");

                        hayTablaAbierta = true;
                    }
                    else
                    {
                        sb.AppendLine($"<tr><td>{papel}</td><td>{idenParte}</td><td>{nomParte}</td><td>{Comun.GetEstadoCivil(estCivil)}</td><td>{domiParte}</td></tr>");
                    }

                    conteoPrevio = conteo;
                }
                if (hayTablaAbierta)
                {
                    sb.AppendLine($"</tbody></table>"); //cierra la tabla de partes
                    sb.AppendLine("</td></tr>");
                    sb.AppendLine($"</tbody></table></p>"); //cierra la tabla del movimiento

                    hayTablaAbierta = false;
                }

                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">{dtEmpr.Rows[0]["TEXTO2"].ToString()}</p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\" class=\"justificar\">Guayaquil, {fecEmision.ToString("dd/MM/yyyy HH:mm")}</p>");

                sb.AppendLine($"<table style=\"width: 100%; page-break-inside: avoid;\"><tbody>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b><img src=\"temporal/{certificado}.png\" class=\"imagen\"></b></td></tr>");
                sb.AppendLine($"<tr><td></td></tr>");
                //sb.AppendLine($"<tr><td class=\"centrar\"><b>{nombreFirmante} {apellidosFirmante}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{cargoQuienFirma}</b></td></tr>");
                sb.AppendLine($"<tr><td class=\"centrar\"><b>{empresa}</b></td></tr>");
                sb.AppendLine($"</tbody></table>");

                sb.AppendLine($"<br/> <p class=\"subrayado\" style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>NOTAS ACLARATORIAS:</b></p>");
                foreach (DataRow nota in dtNota.Rows)
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
                    Solicitante = "CED-" + dt.Rows[0]["IDEN_SOLICITANTE"].ToString() + "-" + dt.Rows[0]["NOMBRE_SOLICITANTE"].ToString(),
                    Uso = dt.Rows[0]["USO"].ToString(),
                    FechaEmision = Convert.ToDateTime(dt.Rows[0]["FECHA"]).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(dt.Rows[0]["FECHA_FIN"]).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "MATRIZ - RP GUAYAQUIL",
                    NumeroSolicitud = dt.Rows[0]["TRAMITE"].ToString(),
                    ComprobantePago = "",
                    Valor = $"$ {Convert.ToDecimal(dt.Rows[0]["VALOR"]).ToString("N2")}"
                };

                //genera el pdf del documento a partir del reporte en html

                //ManipulatePdf(pathToHtml, pathToPdf, _cssPath, logoRpg, logoAlcaldia, codigoQrb, dataFooter);
                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia, codigoQr, dataFooter);

                //Upload(reporte, _workPath + $"\\{tramite}_{secuencial}.pdf");
                string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, fecEmision.Year, fecEmision.Month, tipoServicio).Result.ToString();

                Dictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "okp:sispra.tramite", dt.Rows[0]["TRAMITE"].ToString() },
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
                        Localizacion = "EPMRPG",
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

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertHistoriaDominio", parametros);

            return ds;
        }
    }
}
