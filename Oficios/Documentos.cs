using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAC.CertificadosLibraryI.Auxiliares;
using SISPRA.Entidades.Modelo.Produccion;
using SISPRA.Entidades.Modelo.Enums;
using SISPRA.Entidades.ModeloUtil;
using SISPRA.Entidades.Modelo.Catalago;
using SISPRA.Entidades.ModeloDTO.Administracion;
using Org.BouncyCastle.Tls;
using SISPRA.Entidades.Modelo.Administracion;
using GestorDocumentalLibrary;

namespace SAC.CertificadosLibraryI.Oficios
{
    public class Documentos : Certificado
    {

        #region"Variables Globales"
        private string _cadenaConexion = string.Empty;
        private string _workPath = string.Empty;
        private string _reportPath = string.Empty;
        private string _imagesPath = string.Empty;
        private string _cssPath = string.Empty;
        private string _signedDocPath = string.Empty;
        private string _urlDescarga = string.Empty;
        private string _hostOpenKm;
        private string _userOpenKm;
        private string _passwordOpenKm;
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
        public Documentos(string cadenaConexion, string workPath, string reportPath, string cssPath, string imagesPath, string signedDocPath, string urlDescarga)
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

        public Documentos()
        {
            _hostOpenKm = ObtenerParametro(_cadenaConexion, PARAM_HOST_OPENKM).Result;
            _userOpenKm = ObtenerParametro(_cadenaConexion, PARAM_USER_OPENKM).Result;
            _passwordOpenKm = ObtenerParametro(_cadenaConexion, PARAM_PWD_OPENKM).Result;
        }


        public byte[] GenerarPdf(TramiteDocumento documento, UsuarioRegistradorDTO usuarioRegistrador = null, bool generarFileHtml = false, bool siFirma = false, byte[] oficioRazon = null)
        {
            string pathToPdf = "";
            string docPdf = "";
            string footerPdf = "";
            byte[] codigoQr = null;

            try
            {
                string status = documento.TIPOABRDOC;
                if (documento.SubCategoriaList == null)
                {
                    documento.SubCategoriaList = new List<TramiteDocumentoSubCategoria>();
                }

                if (status.CompareTo(TipoDocumentoEnum.Negativa) == 0)
                {

                    docPdf = PlantillasDocumentos.Negativa;
                    docPdf = docPdf.Replace("#NEGATIVA", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);

                    if (documento.SubCategoriaList.Count > 0)
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "SUSTENTO LEGAL");
                        StringBuilder sb = new StringBuilder("<table border=\"1\" height=\"100%\" Style=\" width:100%\"> <tr> <th>Motivo</th> <th>Base Legal</th> <th>Descripcion Base Legal</th> </tr>");


                        foreach (TramiteDocumentoSubCategoria item in documento.SubCategoriaList)
                        {
                            sb.Append("<tr>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.CADONOMCAT}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARNOMART}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARDETART}</label>");
                            sb.Append("</td>");
                            sb.Append("</tr>");
                        }

                        sb.Append("</table>");
                        docPdf = docPdf.Replace("#ARTICULOS", sb.ToString());
                    }
                    else
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "");
                        docPdf = docPdf.Replace("#ARTICULOS", "");
                    }

                }
                else if (status.CompareTo(TipoDocumentoEnum.NegativaDefinitiva) == 0)
                {
                    docPdf = PlantillasDocumentos.Negativa;
                    docPdf = docPdf.Replace("#NEGATIVA", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                    if (documento.SubCategoriaList.Count > 0)
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "SUSTENTO LEGAL");
                        StringBuilder sb = new StringBuilder("<table border=\"1\" height=\"100%\"> <tr> <th>Motivo</th> <th>Base Legal</th> <th>Descripcion Base Legal</th> </tr>");

                        foreach (TramiteDocumentoSubCategoria item in documento.SubCategoriaList)
                        {
                            sb.Append("<tr>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.CADONOMCAT}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARNOMART}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARDETART}</label>");
                            sb.Append("</td>");
                            sb.Append("</tr>");
                        }

                        sb.Append("</table>");
                        docPdf = docPdf.Replace("#ARTICULOS", sb.ToString());
                    }
                    else
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "");
                        docPdf = docPdf.Replace("#ARTICULOS", "");
                    }

                }
                else if (status.CompareTo(TipoDocumentoEnum.OficioComplementario) == 0)
                {
                    docPdf = PlantillasDocumentos.OficioComplementario;
                    docPdf = docPdf.Replace("#OFICIO", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#CARGO", documento.DOORCARGO);
                    docPdf = docPdf.Replace("#DESTINATARIO", documento.DOORDESTIN);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                    docPdf = docPdf.Replace("#TITULO", documento.CARONOMCAR);
                    docPdf = docPdf.Replace("#CIUDAD", documento.CIUDNOMBRE);

                }
                else if (status.CompareTo(TipoDocumentoEnum.Pronunciamiento) == 0)
                {
                    docPdf = PlantillasDocumentos.Pronunciamiento;
                    docPdf = docPdf.Replace("#PRONUNCIAMIENTO", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                    if (documento.SubCategoriaList.Count > 0)
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "SUSTENTO LEGAL");
                        StringBuilder sb = new StringBuilder("<table border=\"1\" height=\"100%\"> <tr> <th>Motivo</th> <th>Base Legal</th> <th>Descripcion Base Legal</th> </tr>");

                        foreach (TramiteDocumentoSubCategoria item in documento.SubCategoriaList)
                        {
                            sb.Append("<tr>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.CADONOMCAT}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARNOMART}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARDETART}</label>");
                            sb.Append("</td>");
                            sb.Append("</tr>");
                        }

                        sb.Append("</table>");
                        docPdf = docPdf.Replace("#ARTICULOS", sb.ToString());
                    }
                    else
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "");
                        docPdf = docPdf.Replace("#ARTICULOS", "");
                    }
                }
                else if (status.CompareTo(TipoDocumentoEnum.PronunciamietoRectificacion) == 0)
                {
                    docPdf = PlantillasDocumentos.PronunciamientoRectificacion;
                    docPdf = docPdf.Replace("#PRONUNCIAMIENTO", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                    if (documento.SubCategoriaList.Count > 0)
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "SUSTENTO LEGAL");
                        StringBuilder sb = new StringBuilder("<table border=\"1\" height=\"100%\"> <tr> <th>Motivo</th> <th>Base Legal</th> <th>Descripcion Base Legal</th> </tr>");

                        foreach (TramiteDocumentoSubCategoria item in documento.SubCategoriaList)
                        {
                            sb.Append("<tr>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.CADONOMCAT}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARNOMART}</label>");
                            sb.Append("</td>");
                            sb.Append("<td>");
                            sb.Append($"<label Style=\"text-align:left\">{item.COARDETART}</label>");
                            sb.Append("</td>");
                            sb.Append("</tr>");
                        }

                        sb.Append("</table>");
                        docPdf = docPdf.Replace("#ARTICULOS", sb.ToString());
                    }
                    else
                    {
                        docPdf = docPdf.Replace("#SUSTENTOLEGAL", "");
                        docPdf = docPdf.Replace("#ARTICULOS", "");
                    }
                }
                else if (status.CompareTo(TipoDocumentoEnum.OficioNegativa) == 0)
                {
                    docPdf = PlantillasDocumentos.OficioNegativa;
                    docPdf = docPdf.Replace("#OFICIO", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#CARGO", documento.DOORCARGO);
                    docPdf = docPdf.Replace("#DESTINATARIO", documento.DOORDESTIN);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                    docPdf = docPdf.Replace("#TITULO", documento.CARONOMCAR);
                    docPdf = docPdf.Replace("#CIUDAD", documento.CIUDNOMBRE);
                }
                else if (status.CompareTo(TipoDocumentoEnum.OficioRazon) == 0)
                {
                    docPdf = "";
                }
                else if (status.CompareTo(TipoDocumentoEnum.ObservacionInscripcion) == 0)
                {
                    docPdf = PlantillasDocumentos.ObservacionInscripcion;
                    docPdf = docPdf.Replace("#NOBSERVACION", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                }
                else if (status.CompareTo(TipoDocumentoEnum.ObservacionRectificacion) == 0)
                {
                    docPdf = PlantillasDocumentos.ObservacionRectificacion;
                    docPdf = docPdf.Replace("#NOBSERVACION", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                }
                else if (status.CompareTo(TipoDocumentoEnum.TextoActa) == 0)
                {
                    docPdf = PlantillasDocumentos.TextoActa;
                    docPdf = docPdf.Replace("#NOBSERVACION", documento.DOORNUMDOC);
                    docPdf = docPdf.Replace("#OBSERVACION", documento.DOORDETHTML2);
                }
                if (docPdf.CompareTo("") != 0)
                {
                    documento.DOORDETHTML1 = docPdf;
                    //Genera QR
                    // Info QR
                    StringBuilder sbQr = new StringBuilder();
                    if (documento.SERVNOMBRE == null)
                    {
                        if (documento.DOORTIPSER == 1)
                        {
                            documento.SERVNOMBRE = "Inscripciones";
                        }
                    }
                    sbQr.Append($"Tipo de Servicio: {documento.SERVNOMBRE}");
                    if (documento.SERVNOMBRE.CompareTo("Inscripciones") == 0)
                    {
                        sbQr.Append("\n");
                        sbQr.Append($"Año de Repositorio: {documento.DOORNUMANO}");
                        sbQr.Append("\n");
                        sbQr.Append($"Numero de Repositorio: {documento.DOORCODOTT}");
                    }
                    else
                    {
                        sbQr.Append("\n");
                        sbQr.Append($"Año de Orden: {documento.DOORNUMANO}");
                        sbQr.Append("\n");
                        sbQr.Append($"Numero de Orden: {documento.DOORCODOTT}");
                    }

                    sbQr.Append("\n");
                    sbQr.Append($"Tipo de Documento: {documento.SERVNOMBRE}");
                    sbQr.Append("\n");
                    sbQr.Append(usuarioRegistrador.NOMBREREGISTRADOR);
                    sbQr.Append("\n");
                    sbQr.Append(usuarioRegistrador.CARGOREGISTRADOR);


                    codigoQr = Comun.GenerarQr(sbQr.ToString());
                    //string qrB64 = string.Format("data:{0};base64,{1}", "image/jpg", Convert.ToBase64String(codigoQr));
                    //footerPdf = PlantillasDocumentos.PiePagina;
                    //footerPdf = footerPdf.Replace("#QR", qrB64);

                    ////Firma de Registrador 
                    string firma = PlantillasDocumentos.FirmaRegistrador;
                    //=firma.Replace("#FECHAFIRMAREGISTRADOR", "Guayaquil,"+DateTime.Now.ToString("dd")+" de "+DateTime.Now.ToString("MMMM")+" de "+DateTime.Now.ToString("yyyy HH:mm"));
                    firma = firma.Replace("#NOMBREREGISTRADOR", usuarioRegistrador.NOMBREREGISTRADOR);
                    firma = firma.Replace("#CARGOREGISTRADOR", usuarioRegistrador.CARGOREGISTRADOR);
                    firma = firma.Replace("#RAZONSOCIAL", usuarioRegistrador.RAZONSOCIAL);
                    //firma=firma.Replace("#SCANFIRMA", "data:image/png;base64,"+PlantillasDocumentos.FirmaRegistradorEjemplo);

                    //StringBuilder sbQrFirma = new StringBuilder();
                    //sbQrFirma.Append($"FIRMADO POR: {usuarioRegistrador.NOMBREREGISTRADOR}");
                    //sbQr.Append("\n");
                    //sbQrFirma.Append("RAZON:");
                    //sbQr.Append("\n");
                    //sbQrFirma.Append("LOCALIZACION:");
                    //sbQr.Append("\n");
                    //sbQrFirma.Append("FECHA: "+DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    //sbQr.Append("\n");
                    //sbQrFirma.Append("VALIDAR CON: www.firmadigital.gob.ec");
                    //string qrFirmaB64 = await UtilitiesRepository.PrintQr(sbQrFirma.ToString());
                    //firma=firma.Replace("#QRFIRMA", qrFirmaB64);

                    StringBuilder confIni = new StringBuilder();

                    confIni.AppendLine("<html>");
                    confIni.AppendLine("<head>");
                    confIni.AppendLine("<meta charset=\"UTF-8\">");
                    confIni.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 140px;height: 70px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; }</style>");
                    confIni.AppendLine("</head><body>");
                    documento.DOORDETHTML1 = confIni.ToString() + documento.DOORDETHTML1 + firma + "</body></html>";// + "[FOOTER]" + footerPdf;

                    //documento.DOORDETHTML1 =documento.DOORDETHTML1+"[FOOTER]"+ footerPdf;
                }
                else
                {
                    documento.DOORDETHTML1 = "";
                }
                // Generar pdf en base del
                pathToPdf = $"{_reportPath}\\{documento.DOORNUMDOC}.pdf";
                if (status.CompareTo(TipoDocumentoEnum.OficioRazon) == 0)
                {
                    // Sube Oficio Razon al path
                    Comun.Upload(oficioRazon, pathToPdf);
                }
                else
                {
                    var footer = new DataFooter()
                    {
                        ShowComprobantePago = false,
                        ShowFechaEmision = false,
                        ShowFechaVencimiento = false,
                        ShowLugarCanal = false,
                        ShowNumeroSolicitud = false,
                        ShowSolicitante = false,
                        ShowUso = false,
                        ShowValor = false
                    };
                    byte[] registro = Convert.FromBase64String(PlantillasDocumentos.LogoRegistro);
                    byte[] municipio = Convert.FromBase64String(PlantillasDocumentos.LogoMunicipio);
                    Comun.ManipulatePdf(documento.DOORDETHTML1, pathToPdf, _reportPath, registro, municipio, codigoQr, footer);

                    var pathToHtml = $"{_reportPath}\\{documento.DOORNUMDOC}.html";

                    if (generarFileHtml)
                    {
                        Comun.GenerarArchivoHtml(pathToHtml, documento.DOORDETHTML1);
                    }
                }
                if (siFirma)
                {
                    // obtine UID OKM
                    string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, documento.DOORFECDOC.Value.Year, documento.DOORFECDOC.Value.Month, documento.DOORTIPSER).Result.ToString();

                    Dictionary<string, string> metadata = new Dictionary<string, string>
                    {
                        { "okp:sispra.nombre_documento", documento.DOORNUMDOC },
                        { "okp:sispra.orden_tramite", documento.DOORCODOTT.ToString() },
                        { "okp:sispra.tipo_documento", documento.DOORCODDOC.ToString() },
                        { "okp:sispra.tipo_servicio", documento.DOORTIPSER.ToString() },
                        { "okp:sispra.anio", documento.DOORNUMANO.ToString() },
                    };
                    // Frima Documento
                    var pathKeyStore = $"{_workPath}\\registrador.p12";
                    //genera el archivo de la firma p12
                    Comun.Upload(usuarioRegistrador.ARCHIVOFIRMA, pathKeyStore);
                    //genera el archivo del documento sin firmar
                    //Comun.Upload(pdfSinFirma, $"{_pdfPath}\\{tramite}_{secuencial}.pdf");

                    //firma el documento
                    //procesoOk = FirmarPdf(firmaRegistrador.Firma, firmaRegistrador.Password, pdfSinFirma, _workPath, repositorioOkm, metadata, $"{tramite}_{secuencial}");
                    bool procesoOk = Firma.FirmarDocumentoTramite(
                        new KeyStore
                        {
                            Password = usuarioRegistrador.CLAVEFIRMA,
                            Firma = usuarioRegistrador.ARCHIVOFIRMA,
                            Path = pathKeyStore
                        },
                        new Documento
                        {
                            Nombre = $"{documento.DOORNUMDOC}",
                            PathSinFirmar = _reportPath,
                            PathFirmado = _signedDocPath,
                            CargarEnGestor = true,
                            RepositorioOkm = repositorioOkm,
                            Razon = "Firma de Documento Tramite",
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


            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al generar el pdf.\n" + ex.Message, ex);
            }

            return Comun.FileToBytes(pathToPdf);
        }

    }
}
