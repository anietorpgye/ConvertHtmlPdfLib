using System.Drawing;

namespace SAC.CertificadosLibraryI
{
    public class ImagenFirma
    {
        public Color ColorTexto { get; set; }
        public Color ColorBackground { get; set; }
        public string Fuente { get; set; } = string.Empty;
        public int TamanoFuente { get; set; }
        public Image FirmaQr { get; set; }
        public string NombresFirmante { get; set; } = string.Empty;
        public string ApellidosFirmante { get; set; } = string.Empty;
        public string Cargo { get; set; } = string.Empty;
        public string Empresa { get; set; } = string.Empty;
        public int Alto { get; set; }
        public int Ancho { get; set; }
        public string RutaImagen { get; set; } = string.Empty;
    }
}
