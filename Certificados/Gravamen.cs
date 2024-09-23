using iText.Html2pdf;
using iText.IO.Image;
using iText.Kernel.Events;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using SAC.CertificadosLibraryI.Auxiliares;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Xml;

namespace SAC.CertificadosLibraryI.Certificados
{
    public class Gravamen : Certificado
    {
        #region"Variables Globales"
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
        List<Persona> listaDatoPropietarios = new List<Persona>();
        #endregion

        /// <summary>
        /// Constructor para generar el PDF a partir de una consulta a la base de datos
        /// </summary>
        /// <param name="cadenaConexion"></param>
        /// <param name="workPath"></param>
        /// <param name="reportPath"></param>
        /// <param name="cssPath"></param>
        /// <param name="imagesPath"></param>
        /// <param name="signedDocPath"></param>
        /// <param name="urlDescarga"></param>
        public Gravamen(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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
        public Gravamen(string dataXml, string pdfPath, string reportPath, string imagesPath)
        {
            _workPath = dataXml;
            _pdfPath = pdfPath;
            _reportPath = reportPath;
            _imagesPath = imagesPath;
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

                DataTable dtXml = ds.Tables[0];

                XmlDocument doc = new XmlDocument();
                doc.Load(new StringReader(dtXml.Rows[0]["DataXml"].ToString()));

                #region tabla propietarios
                var propietarios = doc.DocumentElement.SelectNodes(@"/row/propietarios/propietario");
                foreach (XmlNode propietario in propietarios)
                {
                    var tipoCedProp = propietario.SelectSingleNode(@"TIPO_CED").InnerText;
                    var nombreProp = propietario.SelectSingleNode(@"NOMBRE").InnerText;
                    var identiProp = propietario.SelectSingleNode(@"IDENTIFICACION").InnerText;
                }
                #endregion

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

                #region"Tabla info registral / info_juridica"

                string TypeCedula = doc.DocumentElement.SelectSingleNode(@"/row/CEDULA").InnerText;
                string TypeTipo = doc.DocumentElement.SelectSingleNode(@"/row/TIPO").InnerText;
                string TypePlanHabita = doc.DocumentElement.SelectSingleNode(@"/row/PLAN_HABITACIONAL").InnerText;
                string TypeUso = doc.DocumentElement.SelectSingleNode(@"/row/USO").InnerText;
                string TypeVALOR = doc.DocumentElement.SelectSingleNode(@"/row/VALOR").InnerText;
                if (TypeVALOR == "" || TypeVALOR == null) { TypeVALOR = "0.00"; }
                string TypeAXIS_COD = doc.DocumentElement.SelectSingleNode(@"/row/AXIS_COD").InnerText.Trim();
                string TypeBANCO = doc.DocumentElement.SelectSingleNode(@"/row/BANCO").InnerText;
                string TypeTRAMITE = doc.DocumentElement.SelectSingleNode(@"/row/TRAMITE").InnerText;
                string TypeCOMPROBANTE_PAGO = doc.DocumentElement.SelectSingleNode(@"/row/COMPROBANTE_PAGO").InnerText;
                if (TypeCOMPROBANTE_PAGO == null || TypeCOMPROBANTE_PAGO.ToString() == "")
                {
                    TypeCOMPROBANTE_PAGO = "NO REGISTRADO";
                }
                else
                {
                    TypeCOMPROBANTE_PAGO = TypeCOMPROBANTE_PAGO.Trim();
                }
                string TypeRECAUDACION_COD = doc.DocumentElement.SelectSingleNode(@"/row/RECAUDACION_COD").InnerText;
                string TypeFACTURA = doc.DocumentElement.SelectSingleNode(@"/row/FACTURA").InnerText;
                string TypeCANAL = doc.DocumentElement.SelectSingleNode(@"/row/CANAL").InnerText;
                string TypeNOMBRE = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE").InnerText;
                string TypeSOLICITANTE_TIPO = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_TIPO").InnerText;
                string TypeSOLICITANTE_CEDULA = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_CEDULA").InnerText;
                string TypeSOLICITANTE_NOMBRES = doc.DocumentElement.SelectSingleNode(@"/row/SOLICITANTE_NOMBRES").InnerText;
                string TypeFECHAGeneral = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText.Trim();
                string TypeFECHA = doc.DocumentElement.SelectSingleNode(@"/row/FECHA").InnerText;

                string TypeFECHA_FIN = doc.DocumentElement.SelectSingleNode(@"/row/FECHA_FIN").InnerText;
                string TypeCODIGO = doc.DocumentElement.SelectSingleNode(@"/row/CODIGO").InnerText;

                string TypeTEXTO = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO").InnerText;
                string Typeiteminfo_juridica = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica").InnerText; //<info_juridica><lindero>
                if (Typeiteminfo_juridica == "" || Typeiteminfo_juridica == null) { Typeiteminfo_juridica = null; }
                string Typeiteminfo_registros = doc.DocumentElement.SelectSingleNode(@"/row/registros").InnerText; //<registros><registro>
                if (Typeiteminfo_registros == null) { Typeiteminfo_registros = null; }

                string TypeTEXTO2 = doc.DocumentElement.SelectSingleNode(@"/row/TEXTO2").InnerText;

                string TypeEMPRESA = doc.DocumentElement.SelectSingleNode(@"/row/EMPRESA").InnerText;
                string TypeNOTA_ACLARATORIA = doc.DocumentElement.SelectSingleNode(@"/row/NOTA_ACLARATORIA").InnerText;
                string TypeNOMBRE_QUIENFIRMA = doc.DocumentElement.SelectSingleNode(@"/row/NOMBRE_QUIENFIRMA").InnerText;
                string TypeCARGO_QUIENFIRMA = doc.DocumentElement.SelectSingleNode(@"/row/CARGO_QUIENFIRMA").InnerText;

                //***************** INFORMACIÓN DEL PREDIO ************************************
                string norte = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/NORTE").InnerText;
                string sur = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/SUR").InnerText;
                string este = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/ESTE").InnerText;
                string oeste = doc.DocumentElement.SelectSingleNode(@"/row/catastro/lindero/OESTE").InnerText;

                string ubicacion = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/UBICACION").InnerText;
                string estadoPredio = doc.DocumentElement.SelectSingleNode(@"/row/info_juridica/lindero/ESTADO_PREDIO").InnerText;

                #endregion

                #region "tabla movimientos"
                var movimientos = doc.DocumentElement.SelectNodes(@"/row/registros/registro");

                foreach (XmlNode movi in movimientos) // tablas
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
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 18px;\"><b>CERTIFICADO DE GRAVAMEN</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 16px;\"><b>No. {codCert}</b></p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{texto.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>1.INFORMACIÓN DEL PREDIO</u></b></p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{ubicacion.Replace("\n", "</br>")}</p>");
                sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\"><b>Titular(es) Registrado(s):</b></p>");
                sb.AppendLine($"<table style=\"width: 100%;\"><tbody>");
                foreach (XmlNode propietario in propietarios)
                {
                    var tipoCedProp = propietario.SelectSingleNode(@"TIPO_CED").InnerText;
                    var identiProp = propietario.SelectSingleNode(@"IDENTIFICACION").InnerText;
                    var nombreProp = propietario.SelectSingleNode(@"NOMBRE").InnerText;
                    // Agrego a la lista.
                    if (identiProp != "")
                    {
                        ListaAgregarDatos(tipoCedProp.Trim(), identiProp.Trim(), nombreProp.Trim());
                    }
                    // sb.AppendLine($"<tr><td>{tipoCedProp}</td><td>{identiProp}</td><td>{nombreProp}</td></tr>");
                }
                // Ordenar la lista alfabéticamente por el nombre
                listaDatoPropietarios.Sort((p1, p2) => string.Compare(p1.Nombre, p2.Nombre));
                if (listaDatoPropietarios != null)
                {
                    foreach (var persona in listaDatoPropietarios)
                    {
                        sb.AppendLine($"<tr><td>{persona.TipoCed}</td><td>{persona.Identificacion}</td><td>{persona.Nombre}</td></tr>");
                    }
                }


                sb.AppendLine($"</tbody></table>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 12px;\"><b>Estado:</b> {estadoPredioDesc}</p>");
                sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>2.MOVIMIENTOS DEL PREDIO EN ORDEN CRONOLÓGICO</u></b></p>");
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
                    sb.AppendLine($"<tr><td style=\"width: 20%; vertical-align: text-top;\"><b>Observación:</b></td><td style=\"text-align: justify;\">{((observacion == string.Empty ? "-----" : observacion.Trim()))}</td></tr>");
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

                } // fin del recorrer las tablas y sub tablas: Predios > Partes

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

                //Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());

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
                //******************** Lleno un html para ver como queda el código **********optional*****
                // GenerarHtml(pathToHtml, sb.ToString());
                //****************************************************************************************
                var dataFooter = new DataFooter
                {
                    Solicitante = "CED-" + tipoIdenSol + "" + nombreSol,
                    Uso = uso.ToUpper(),
                    FechaEmision = Convert.ToDateTime(fecha).ToString("dd/MM/yyyy HH:mm"),
                    FechaVencimiento = Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy HH:mm"),
                    LugarCanal = "Matríz - Rpg Guayaquil",
                    NumeroSolicitud = tramite,
                    ComprobantePago = TypeCOMPROBANTE_PAGO,
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

        public bool GenerarPdf(Registrador firmaRegistrador, int tramite, int secuencial, int anioOrden, int numeroOrden, int tipoServicio, bool generarFileHtml = false)
        {
            bool procesoOk = true;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\logo_alcaldia.png");

                DataSet ds = ObtenerDatos(tramite, secuencial, anioOrden, numeroOrden);

                DataTable dtPropietarios = ds.Tables[0];
                DataTable dtData = ds.Tables[1];
                DataTable dtPredio = ds.Tables[2];
                DataTable dtMovs = ds.Tables[3]; 
                DataTable dtAclaratorias = ds.Tables[4];

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

                var estadoPredioDesc = (dtPredio.Rows[0]["ESTADO_PREDIO"].ToString() == "AC") ? "ACTIVO" : "INACTIVO";

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .tablaConBorde, .tablaConBorde th, .tablaConBorde tr, .tablaConBorde td {border: 1px solid black; border-collapse: collapse; margin-left: auto; margin-right: auto; text-align: center;} .tablaConBorde th {padding-top: 5px; padding-bottom: 5px; text-align: center; background-color: #F7F7F7;} .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .parrafosgrande {text-align: center; font-family: 'Helvetica', monospace; font-size: 16px; font-weight: bold;} .parrafosmediano {font-family: 'Helvetica'; font-size: 12px; text-align: justify;} .parrafospequeno {font-family: 'Helvetica', monospace; font-size: 8px; text-align: justify;} .parrafospequeno_sj {font-family: 'Helvetica', monospace; font-size: 8px;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine($"<p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>CERTIFICADO DE GRAVAMEN</b></p>");
                sb.AppendLine($"<hr><p style=\"text-align: center; font-family: Helvetica; font-size: 20px;\"><b>No. {dtData.Rows[0]["CERTIFICADO"].ToString()}</b></p>");
                sb.AppendLine($"<p class=\"parrafosmediano\">{texto.Replace("\n", "")}</p>");

                if (dtMovs.Rows.Count > 0)
                {
                    sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 16px;\" ><b><u>1. INFORMACIÓN DEL PREDIO</u></b></p>");
                    sb.AppendLine($"<p style=\"font-family: Helvetica; font-size: 12px;\" class=\"justificar\">{dtPredio.Rows[0]["UBICACION"].ToString()}</p>");
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

                    #region movimientos

                    sb.AppendLine($"<p style=\"font-family: 'Helvetica'; font-size: 16px;\" ><b><u>2. MOVIMIENTOS DEL PREDIO EN ORDEN CRONOLÓGICO</u></b></p>");

                    int conteoPrevio = 0; bool hayTablaAbierta = false;
                    foreach (DataRow movi in dtMovs.Rows)
                    {
                        var tipoReg = movi["TIPO_REG"].ToString();
                        var conteo = Convert.ToInt32(movi["CONTEO"]);
                        var acto = movi["ACTO"].ToString();
                        var fechaInsMov = movi["FECHA_INSC"].ToString();
                        var tomo = movi["TOMO"].ToString();
                        var numInsc = movi["NUMERO_INSC"].ToString();
                        var numReper = movi["NUMERO_REPER"].ToString();
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
                    #endregion
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

        private DataSet ObtenerDatos(int tramite, int secuencial, int anioOrden, int numeroOrden)
        {
            List<SqlParameter> parametros = new List<SqlParameter>
            {
                new SqlParameter("@SOLINUMTRA", tramite),
                new SqlParameter("@SOLISEC", secuencial),
                new SqlParameter("@NUMANO", anioOrden),
                new SqlParameter("@CODOTT", numeroOrden),
            };

            var resul = Comun.ActualizarVigencia(_cadenaConexion, tramite, secuencial);

            DataSet ds = ObtenerDatos(_cadenaConexion, "dbo.SP_CertGravamenes", parametros);

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

        //*************** Ordenar los propietarios de los predios con los nodos del Xml: Bd ********
        class Persona
        {
            public string TipoCed { get; set; }
            public string Identificacion { get; set; }
            public string Nombre { get; set; }
        }

        // Función para agregar valores a la lista
        void ListaAgregarDatos(string tipoCed, string identificacion, string nombre)
        {
            listaDatoPropietarios.Add(new Persona { TipoCed = tipoCed, Identificacion = identificacion, Nombre = nombre });
        }
    }
}
