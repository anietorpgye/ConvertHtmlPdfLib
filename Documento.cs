using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI
{
    public class Documento
    {
        /// <summary>
        /// Nombre del documento pdf SIN extensión
        /// </summary>
        public string Nombre { get; set; } = string.Empty;

        /// <summary>
        /// Path al documento sin firmar, sin nombre del archivo
        /// </summary>
        public string PathSinFirmar { get; set; } = string.Empty;

        /// <summary>
        /// path al documento firmado, sin nombre del archivo
        /// </summary>
        public string PathFirmado { get; set; } = string.Empty;

        /// <summary>
        /// Flag si se desea estampar la firma con codigo qr en el documento
        /// </summary>
        public bool EstamparFirma { get; set; }
        public string Razon { get; set; } = string.Empty;
        public string Localizacion { get; set; } = string.Empty;
        public byte[] Qr { get; set; }

        /// <summary>
        /// Flag si se desea cargar el documento firmado en OpenKm
        /// </summary>
        public bool CargarEnGestor { get; set; } = false;

        /// <summary>
        /// Uuid de la carpeta donde se cargará el documento en OpenKm
        /// </summary>
        public string RepositorioOkm { get; set; } = string.Empty;

        /// <summary>
        /// Nombre del grupo de metadatos
        /// </summary>
        public string GrupoOkm { get; set; } = string.Empty;

        public Dictionary<string, string> Metadata { get; set; }
    }
}
