using iText.Html2pdf.Attach;
using iText.Html2pdf.Attach.Impl;
using iText.StyledXmlParser.Node;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class QrCodeTagWorkerFactory : DefaultTagWorkerFactory
    {
        public override ITagWorker? GetCustomTagWorker(IElementNode tag, ProcessorContext context)
        {
            if (tag.Name().Equals("qr", StringComparison.OrdinalIgnoreCase))
            {
                return new QrCodeTagWorker(tag, context);
            }
            return null;
        }
    }
}
