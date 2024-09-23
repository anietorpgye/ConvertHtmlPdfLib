namespace SAC.CertificadosLibraryI
{
    public class CertificadosLibraryException : ApplicationException
    {
        public CertificadosLibraryException(string mensaje, Exception orignal) : base(mensaje, orignal)
        {
        }

        public CertificadosLibraryException(string mensaje) : base(mensaje)
        {
        }
    }
}