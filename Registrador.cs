using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAC.CertificadosLibraryI
{
    public class Registrador
    {
        public byte[] Firma { get; set; }
        public string Password { get; set; } = string.Empty;
        public string PathKeyStore { get; set; } = string.Empty;
    }
}
