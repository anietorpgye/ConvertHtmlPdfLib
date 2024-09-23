using iText.Layout.Element;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class DataFooter
    {
        public string Solicitante { get; set; } = string.Empty;
        public string Uso { get; set; } = string.Empty;
        public string FechaEmision { get; set; } = string.Empty;
        public string FechaVencimiento { get; set; } = string.Empty;
        public string LugarCanal { get; set; } = string.Empty;
        public string NumeroSolicitud { get; set; } = string.Empty;
        public string ComprobantePago { get; set; } = string.Empty;
        public string Valor { get; set; } = string.Empty;
        public Image Qr { get; set; }
        public bool ShowSolicitante { get; set; } = true;
        public bool ShowUso { get; set; } = true;
        public bool ShowFechaEmision { get; set; } = true;
        public bool ShowFechaVencimiento { get; set; } = true;
        public bool ShowLugarCanal { get; set; } = true;
        public bool ShowNumeroSolicitud { get; set; } = true;
        public bool ShowComprobantePago { get; set; } = true;
        public bool ShowValor { get; set; } = true;
        public bool ShowQr { get; set; } = true;
    }
}
