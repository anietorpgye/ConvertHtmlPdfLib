using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Layout.Properties;
using iText.Layout.Element;
using iText.Layout;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class Header : IEventHandler
    {
        private string _header;
        private Image _logoRpg;
        private Image _logoAlcaldia;

        public Header(string header, Image logoRpg, Image logoAlcaldia)
        {
            _header = header;
            _logoRpg = logoRpg;
            _logoAlcaldia = logoAlcaldia;
        }

        public virtual void HandleEvent(Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfDocument pdf = docEvent.GetDocument();

            PdfPage page = docEvent.GetPage();
            Rectangle pageSize = page.GetPageSize();

            Canvas canvas = new Canvas(new PdfCanvas(page), pageSize);
            canvas.SetFontSize(18);
            canvas.Add(_logoRpg);
            canvas.Add(_logoAlcaldia);

            // Write text at position
            canvas.ShowTextAligned(_header,
                pageSize.GetWidth() / 2,
                pageSize.GetTop() - 30, TextAlignment.CENTER);
            canvas.Close();
        }
    }
}
