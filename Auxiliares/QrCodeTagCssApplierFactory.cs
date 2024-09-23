using iText.Html2pdf.Css.Apply;
using iText.Html2pdf.Css.Apply.Impl;
using iText.StyledXmlParser.Node;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class QrCodeTagCssApplierFactory : DefaultCssApplierFactory
    {
        public override ICssApplier GetCustomCssApplier(IElementNode tag)
        {
            if (tag.Name().Equals("qr", StringComparison.OrdinalIgnoreCase))
            {
                return new BlockCssApplier();
            }

            return null;
        }
    }
}
