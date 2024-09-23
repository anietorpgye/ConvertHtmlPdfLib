using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout;

namespace SAC.CertificadosLibraryI.Auxiliares
{
    public class Footer : IEventHandler
    {
        protected PdfFormXObject placeholder;
        protected float side = 20;
        protected float x = 467;//300;
        protected float y = 20;//25;
        protected float space = 4.5f;
        protected float descent = 3;
        protected float sizeText = 7f;

        private DataFooter _data;

        public Footer(DataFooter data)
        {
            _data = data;
            placeholder = new PdfFormXObject(new Rectangle(0, 0, side, side));
        }

        public virtual void HandleEvent(Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfDocument pdf = docEvent.GetDocument();
            PdfPage page = docEvent.GetPage();
            int pageNumber = pdf.GetPageNumber(page);
            Rectangle pageSize = page.GetPageSize();

            // Creates drawing canvas
            PdfCanvas pdfCanvas = new PdfCanvas(page);

            Canvas canvas = new Canvas(pdfCanvas, pageSize);

            Color blackColor = new DeviceCmyk(0f, 0f, 0f, 100f);
            Color grayColor = new DeviceCmyk(0f, 0f, 0f, 0.7f);
            pdfCanvas
                .SetStrokeColor(blackColor)
                .MoveTo(10, 90)
                .LineTo(490, 90)
                .ClosePathStroke();

            canvas.Add(_data.Qr);

            if (_data.ShowSolicitante)
            {
                Text tituloSolicitante = new Text("Solicitante:");
                tituloSolicitante.SetFontSize(sizeText);
                tituloSolicitante.SetFontColor(grayColor);

                Text dataSolicitante = new Text(_data.Solicitante);
                dataSolicitante.SetFontSize(sizeText);

                Paragraph pTitSol = new Paragraph(tituloSolicitante);
                Paragraph pDataSolicitante = new Paragraph(dataSolicitante);

                canvas.ShowTextAligned(pTitSol, 15, 75, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataSolicitante, 15, 65, TextAlignment.LEFT);
            }

            if (_data.ShowUso)
            {
                Text tituloUso = new Text("Uso:");
                tituloUso.SetFontSize(sizeText);
                tituloUso.SetFontColor(grayColor);

                Text dataUso = new Text(_data.Uso);
                dataUso.SetFontSize(sizeText);

                Paragraph pTitUso = new Paragraph(tituloUso);
                Paragraph pDataUso = new Paragraph(dataUso);

                canvas.ShowTextAligned(pTitUso, 350, 75, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataUso, 350, 65, TextAlignment.LEFT);
            }

            if (_data.ShowFechaEmision)
            {
                Text tituloFecEmision = new Text("Fecha y Hora de Emisión:");
                tituloFecEmision.SetFontSize(sizeText);
                tituloFecEmision.SetFontColor(grayColor);

                Text dataFecEmision = new Text(_data.FechaEmision);
                dataFecEmision.SetFontSize(sizeText);

                Paragraph pTitFecEmi = new Paragraph(tituloFecEmision);
                Paragraph pDataFecEmi = new Paragraph(dataFecEmision);

                canvas.ShowTextAligned(pTitFecEmi, 15, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataFecEmi, 15, 40, TextAlignment.LEFT);
            }

            if (_data.ShowFechaVencimiento)
            {
                Text tituloFecVencimiento = new Text("Fecha de Vencimiento:");
                tituloFecVencimiento.SetFontSize(sizeText);
                tituloFecVencimiento.SetFontColor(grayColor);

                Text dataFecVencimiento = new Text(_data.FechaVencimiento);
                dataFecVencimiento.SetFontSize(sizeText);

                Paragraph pTitVenc = new Paragraph(tituloFecVencimiento);
                Paragraph pDataFecVenc = new Paragraph(dataFecVencimiento);

                canvas.ShowTextAligned(pTitVenc, 100, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataFecVenc, 100, 40, TextAlignment.LEFT);
            }

            if (_data.ShowLugarCanal)
            {
                Text tituloLugar = new Text("Lugar/Canal de Emisión:");
                tituloLugar.SetFontSize(sizeText);
                tituloLugar.SetFontColor(grayColor);

                Text dataLugar = new Text(_data.LugarCanal);
                dataLugar.SetFontSize(sizeText);

                Paragraph pTitLugar = new Paragraph(tituloLugar);
                Paragraph pDataLugar = new Paragraph(dataLugar);

                canvas.ShowTextAligned(pTitLugar, 180, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataLugar, 180, 40, TextAlignment.LEFT);
            }

            if (_data.ShowNumeroSolicitud)
            {
                Text tituloNumSolicitud = new Text("No. de Solicitud:");
                tituloNumSolicitud.SetFontSize(sizeText);
                tituloNumSolicitud.SetFontColor(grayColor);

                Text dataNumSolicitud = new Text(_data.NumeroSolicitud);
                dataNumSolicitud.SetFontSize(sizeText);

                Paragraph pTitSolicitud = new Paragraph(tituloNumSolicitud);
                Paragraph pDataNumSol = new Paragraph(dataNumSolicitud);

                canvas.ShowTextAligned(pTitSolicitud, 280, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataNumSol, 280, 40, TextAlignment.LEFT);
            }

            if (_data.ShowComprobantePago)
            {
                Text tituloCmpbtePago = new Text("No. Comprobante de Pago:");
                tituloCmpbtePago.SetFontSize(sizeText);
                tituloCmpbtePago.SetFontColor(grayColor);

                Text dataCmpbtePago = new Text(_data.ComprobantePago);
                dataCmpbtePago.SetFontSize(sizeText);

                Paragraph pTitCmpbtePago = new Paragraph(tituloCmpbtePago);
                Paragraph pDataCmpbtePago = new Paragraph(dataCmpbtePago);

                canvas.ShowTextAligned(pTitCmpbtePago, 340, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataCmpbtePago, 340, 40, TextAlignment.LEFT);
            }

            if (_data.ShowValor)
            {
                Text tituloValor = new Text("Valor:");
                tituloValor.SetFontSize(sizeText);
                tituloValor.SetFontColor(grayColor);

                Text dataValor = new Text(_data.Valor);
                dataValor.SetFontSize(sizeText);

                Paragraph pTitValor = new Paragraph(tituloValor);
                Paragraph pDataValor = new Paragraph(dataValor);

                canvas.ShowTextAligned(pTitValor, 435, 50, TextAlignment.LEFT);
                canvas.ShowTextAligned(pDataValor, 435, 40, TextAlignment.LEFT);
            }

            Text paginas = new Text($"Pag.: {pageNumber} de");
            paginas.SetFontSize(sizeText);
            Paragraph p = new Paragraph(paginas);
            //Paragraph p = new Paragraph()
            //    .Add("Pág.: ")
            //    .Add(pageNumber.ToString())
            //    .Add(" de");

            canvas.ShowTextAligned(p, x, y, TextAlignment.RIGHT);
            canvas.Close();

            // Create placeholder object to write number of pages
            pdfCanvas.AddXObjectAt(placeholder, x + space, y - descent);
            pdfCanvas.Release();
        }

        public void WriteTotal(PdfDocument pdf)
        {
            Canvas canvas = new Canvas(placeholder, pdf);
            Text totalPaginas = new Text(pdf.GetNumberOfPages().ToString());
            totalPaginas.SetFontSize(sizeText);
            Paragraph pTotalPages = new Paragraph(totalPaginas);
            canvas.ShowTextAligned(pTotalPages, 0, descent, TextAlignment.LEFT);
            canvas.Close();
        }
    }
}
