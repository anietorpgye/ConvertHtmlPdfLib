using GestorDocumentalLibrary;
using SAC.CertificadosLibraryI.Auxiliares;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI.Inscripciones
{
    public class Inconformidad : Certificado
    {
        private string _cadenaConexion;
        private string _reportPath;
        private string _imagesPath;
        private string _hostOpenkm; 
        private string _userOpenkm;
        private string _passwordOpenkm;

        public Inconformidad(string cadenaConexion, string reportPath, string imagesPath)
        {
            _cadenaConexion = cadenaConexion;
            _reportPath = reportPath;
            _imagesPath = imagesPath;
            _hostOpenkm = ObtenerParametro(_cadenaConexion, PARAM_HOST_OPENKM).Result;
            _userOpenkm = ObtenerParametro(_cadenaConexion, PARAM_USER_OPENKM).Result;
            _passwordOpenkm = ObtenerParametro(_cadenaConexion, PARAM_PWD_OPENKM).Result;
        }

        /// <summary>
        /// Genera el pdf de la inconformidad y lo carga a OpenKm
        /// </summary>
        /// <param name="tramite"></param>
        /// <param name="secuencial"></param>
        /// <param name="mensaje"></param>
        /// <param name="fechaTramite"></param>
        /// <param name="tipoServicio"></param>
        /// <returns></returns>
        public bool GenerarPdf(int tramite, int secuencial, string mensaje, DateTime fechaTramite, int tipoServicio, bool generarFileHtml = false)
        {
            bool procesoOk = false;

            try
            {
                var logoRpg = Comun.FileToBytes(_imagesPath + "\\logo_rpg.png");
                var logoAlcaldia = Comun.FileToBytes(_imagesPath + "\\Logo Alcaldia Guayaquil 2.png");

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("<meta charset=\"UTF-8\">");
                sb.AppendLine("<style> * { box-sizing: border-box; } .row { margin-left:-5px; margin-right:-5px; } .column { float:left; width:50%; } .row::after {content:\"\"; clear:both;display:table;} html { position: relative; min-height: 100%;} body { margin: 0px; font-family: Helvetica; /* bottom = footer height */ padding: 25px; /*background-color: lightyellow;*/ } qr {border:solid 1px blue; height:80px; width:80px;} @page {margin-top: 100pt;margin-bottom: 120pt;} .centrar {text-align: center;} .justificar {text-align: justify;} .imagen {width: 370px;height: 90px;} th,td { padding: 2px 5px 2px 5px ; text-align: left; font-family: Helvetica; font-size: 12px; } .subrayado { text-decoration: underline; }</style>");
                sb.AppendLine("</head><body>");
                sb.AppendLine(mensaje);
                sb.AppendLine("</body></html>");

                var pathToHtml = $"{_reportPath}\\Inconformidad_{tramite}.html";
                var pathToPdf = $"{_reportPath}\\Inconformidad_{tramite}.pdf";

                if (generarFileHtml) 
                {
                    Comun.GenerarArchivoHtml(pathToHtml, sb.ToString());
                }

                Comun.ManipulatePdf(sb.ToString(), pathToPdf, _reportPath, logoRpg, logoAlcaldia);

                //Upload(reporte, _workPath + $"\\{tramite}_{secuencial}.pdf");
                string repositorioOkm = ObtenerRepositorioOpenKm(_cadenaConexion, fechaTramite.Year, fechaTramite.Month, tipoServicio).Result.ToString();

                Dictionary<string, string> metadata = new Dictionary<string, string>
                {
                    { "okp:sispra.tramite", tramite.ToString() },
                    { "okp:sispra.secuencia", secuencial.ToString() },
                    { "okp:sispra.tipo_documento", "" },
                    { "okp:sispra.tipo_servicio", tipoServicio.ToString() },
                    { "okp:sispra.anio", fechaTramite.Year.ToString() },
                    { "okp:sispra.etapa", "6" },
                    { "okp:sispra.notificado", "NO" },
                };

                Gestor gestor = new Gestor()
                {
                    HostGestor = _hostOpenkm,
                    UserGestor = _userOpenkm,
                    Password = _passwordOpenkm
                };
                var rc = gestor.ImportarIndividual(new Archivo
                {
                    Estado = false,
                    Path = pathToPdf,
                    FolderUuid = repositorioOkm,
                    UsarArchivoOkm = false,
                    GroupName = "okg:sispra",
                    Metadata = metadata
                });

                gestor = null;

                procesoOk = true;

                return procesoOk;
            }
            catch (Exception ex)
            {
                procesoOk = false;
            }

            return procesoOk;
        }

    }
}
