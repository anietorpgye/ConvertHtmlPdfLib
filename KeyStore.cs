using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI
{
    public class KeyStore
    {
        /// <summary>
        /// Ruta completa a la firma digital incluido el nombre del archivo
        /// </summary>
        public string Path { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public byte[] Firma { get; set; }
    }
}
