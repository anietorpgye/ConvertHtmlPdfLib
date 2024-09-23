using iText.Barcodes.Qrcode;
using iText.Barcodes;
using iText.Html2pdf.Attach;
using iText.Layout;
using iText.Layout.Element;
using iText.StyledXmlParser.Node;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class QrCodeTagWorker : ITagWorker
    {
        private static string[] allowedErrorCorrection = { "L", "M", "Q", "H" };
        private static string[] allowedCharSets = { "Cp437", "Shift_JIS", "ISO-8859-1", "ISO-8859-16" };

        private BarcodeQRCode? qrCode;
        private Image? qrCodeAsImage;

        public QrCodeTagWorker(IElementNode tag, ProcessorContext context)
        {
            IDictionary<EncodeHintType, object> hints = new Dictionary<EncodeHintType, object>();
            string charset = tag.GetAttribute("charset");
            if (CheckCharSet(charset))
            {
                hints.Add(EncodeHintType.CHARACTER_SET, charset);
            }

            string errorCorrection = tag.GetAttribute("errorcorrection");
            if (CheckErrorCorrectionAllowed(errorCorrection))
            {
                ErrorCorrectionLevel errorCorrectionLevel = getErrorCorrectionLevel(errorCorrection);
                hints.Add(EncodeHintType.ERROR_CORRECTION, errorCorrectionLevel);
            }
        }

        public void ProcessEnd(IElementNode element, ProcessorContext context)
        {
            qrCodeAsImage = new Image(qrCode.CreateFormXObject(context.GetPdfDocument()));
        }

        public bool ProcessContent(string content, ProcessorContext context)
        {
            qrCode = new BarcodeQRCode();
            qrCode.SetCode(content);
            return true;
        }

        public bool ProcessTagChild(ITagWorker childTagWorker, ProcessorContext context)
        {
            return false;
        }

        public IPropertyContainer GetElementResult()
        {
            return qrCodeAsImage;
        }

        private static bool CheckCharSet(string toValidate)
        {
            foreach (string charset in allowedCharSets)
            {
                if (toValidate.Equals(charset, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static ErrorCorrectionLevel getErrorCorrectionLevel(string lvl)
        {
            switch (lvl)
            {
                case "L":
                    return ErrorCorrectionLevel.L;
                case "M":
                    return ErrorCorrectionLevel.M;
                case "Q":
                    return ErrorCorrectionLevel.Q;
                case "H":
                    return ErrorCorrectionLevel.H;
            }

            return null;
        }

        private bool CheckErrorCorrectionAllowed(string toValidate)
        {
            foreach (string errorCorrection in allowedErrorCorrection)
            {
                if (toValidate.Equals(errorCorrection, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
