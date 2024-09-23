using iText.Html2pdf;
using iText.IO.Image;
using iText.Kernel.Events;
using iText.Kernel.Pdf;
using MessagingToolkit.QRCode.Codec;
using SAC.CertificadosLibraryI.Auxiliares;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data.SqlClient;
using iText.Kernel.Geom;
using iText.Layout;

namespace SAC.CertificadosLibraryI
{
    public static class Comun
    {
        public static byte[] CargarImagen(string pathToImage)
        {
            // Load file meta data with FileInfo
            FileInfo fileInfo = new FileInfo(pathToImage);

            // The byte[] to save the data in
            byte[] data = new byte[fileInfo.Length];

            // Load a filestream and put its content into the byte[]
            using (FileStream fs = fileInfo.OpenRead())
            {
                fs.Read(data, 0, data.Length);
                fs.Close();
            }

            return data;
        }

        public static void Upload(Stream item, string path)
        {
            FileStream destFile = new FileStream(path, FileMode.OpenOrCreate);
            item.CopyTo(destFile);
            destFile.Close();
            item.Close();
        }

        public static void Upload(byte[] file, string path)
        {
            File.WriteAllBytes(path, file);
        }

        public static Image ImageFromByteArray(byte[] data)
        {
            MemoryStream ms = new MemoryStream(data);
            return Image.FromStream(ms);
        }

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

        public static byte[] FileToBytes(string ruta)
        {
            FileStream file = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.Delete);

            byte[] arreglo = new byte[file.Length];
            BinaryReader reader = new BinaryReader(file);
            arreglo = reader.ReadBytes(Convert.ToInt32(file.Length));
            file.Close();
            return arreglo;
        }

        public static byte[] GenerarQr(string texto)
        {
            byte[] imagenQR;
            QRCodeEncoder encoder = new QRCodeEncoder();
            Bitmap img = encoder.Encode(texto);
            Image QR = (Image)img;

            using (MemoryStream ms = new MemoryStream())
            {
                QR.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] imageBytes = ms.ToArray();

                imagenQR = imageBytes;
            }
            return imagenQR;
        }

        public static bool DibujarFirma(ImagenFirma data)
        {
            bool rc = true;
            try
            {
                Bitmap img = new Bitmap(data.Ancho, data.Alto);
                Graphics Gimg = Graphics.FromImage(img);
                Font imgFont = new Font(data.Fuente, data.TamanoFuente, FontStyle.Bold);
                SolidBrush bForeColor = new SolidBrush(data.ColorTexto);
                SolidBrush bBackColor = new SolidBrush(data.ColorBackground);
                SolidBrush bForeColorSign = new SolidBrush(Color.Gray);

                Gimg.FillRectangle(bBackColor, 0, 0, data.Ancho, data.Alto);

                //Gimg.FillRectangle(bBackColorSign, 290, 20, 250, 110);

                //Gimg.DrawImage(firmaQr, new PointF(290, 20));
                //Gimg.DrawImage(data.FirmaQr, 240, 20, 100, 100);
                Gimg.DrawImage(data.FirmaQr, 70, 25, 80, 80);

                Font imgFontFirma1 = new Font("Arial", 8);
                //PointF imgPoint = new PointF(345, 40);
                PointF imgPoint = new PointF(160, 40);
                Gimg.DrawString("Firmado electrónicamente por:", imgFontFirma1, bForeColorSign, imgPoint);

                Font imgFontFirma2 = new Font("Arial", 12, FontStyle.Bold);
                //imgPoint = new PointF(345, 60);
                imgPoint = new PointF(160, 60);
                Gimg.DrawString($"{data.NombresFirmante}", imgFontFirma2, bForeColor, imgPoint);

                //imgPoint = new PointF(345, 80);
                imgPoint = new PointF(160, 80);
                Gimg.DrawString($"{data.ApellidosFirmante}", imgFontFirma2, bForeColor, imgPoint);

                //imgPoint = new PointF(185, 128);
                imgPoint = new PointF(65, 128);
                Gimg.DrawString(data.Cargo, imgFont, bForeColor, imgPoint);

                //imgPoint = new PointF(135, 145);
                imgPoint = new PointF(15, 145);
                Gimg.DrawString(data.Empresa, imgFont, bForeColor, imgPoint);

                //var imgx = UniformNoise(img);

                img.Save(data.RutaImagen, ImageFormat.Png);
            }
            catch
            {
                rc = false;
            }

            return rc;
        }

        public static bool DibujarFirma2(ImagenFirma data)
        {
            bool rc = true;
            try
            {
                Bitmap img = new Bitmap(data.Ancho, data.Alto);
                Graphics Gimg = Graphics.FromImage(img);
                Font imgFont = new Font(data.Fuente, data.TamanoFuente, FontStyle.Bold);
                SolidBrush bForeColor = new SolidBrush(data.ColorTexto);
                SolidBrush bBackColor = new SolidBrush(data.ColorBackground);
                SolidBrush bForeColorSign = new SolidBrush(Color.Gray);

                Gimg.FillRectangle(bBackColor, 0, 0, data.Ancho, data.Alto);

                //Gimg.FillRectangle(bBackColorSign, 290, 20, 250, 110);

                //Gimg.DrawImage(firmaQr, new PointF(290, 20));
                //Gimg.DrawImage(data.FirmaQr, 240, 20, 100, 100);
                Gimg.DrawImage(data.FirmaQr, 5, 5, 80, 80);

                Font imgFontFirma1 = new Font("Arial", 8);
                //PointF imgPoint = new PointF(345, 40);
                PointF imgPoint = new PointF(90, 20);
                Gimg.DrawString("Firmado electrónicamente por:", imgFontFirma1, bForeColorSign, imgPoint);

                Font imgFontFirma2 = new Font("Arial", 12, FontStyle.Bold);
                //imgPoint = new PointF(345, 60);
                imgPoint = new PointF(90, 40);
                Gimg.DrawString($"{data.NombresFirmante}", imgFontFirma2, bForeColor, imgPoint);

                //imgPoint = new PointF(345, 80);
                imgPoint = new PointF(90, 60);
                Gimg.DrawString($"{data.ApellidosFirmante}", imgFontFirma2, bForeColor, imgPoint);

                img.Save(data.RutaImagen, ImageFormat.Png);
            }
            catch
            {
                rc = false;
            }

            return rc;
        }

        public static bool DibujarFirmaManual(ImagenFirma data)
        {
            bool rc = true;
            try
            {
                Bitmap img = new Bitmap(data.Ancho, data.Alto);
                Graphics Gimg = Graphics.FromImage(img);
                Font imgFont = new Font(data.Fuente, data.TamanoFuente, FontStyle.Regular);
                SolidBrush bForeColor = new SolidBrush(data.ColorTexto);
                SolidBrush bBackColor = new SolidBrush(data.ColorBackground);
                SolidBrush bForeColorSign = new SolidBrush(Color.Gray);

                //Gimg.FillRectangle(bBackColor, 0, 0, data.Ancho, data.Alto);

                Gimg.DrawImage(data.FirmaQr, 200, 25, 140, 80);

                Font imgFontFirma2 = new Font("Helvetica", 12, FontStyle.Bold);

                PointF imgPoint = new PointF(190, 110);
                Gimg.DrawString($"{data.NombresFirmante} {data.ApellidosFirmante}", imgFontFirma2, bForeColor, imgPoint);

                imgPoint = new PointF(65, 128);
                Gimg.DrawString(data.Cargo, imgFont, bForeColor, imgPoint);

                imgPoint = new PointF(15, 145);
                Gimg.DrawString(data.Empresa, imgFont, bForeColor, imgPoint);

                //var imgx = UniformNoise(img);

                img.Save(data.RutaImagen, ImageFormat.Png);
            }
            catch
            {
                rc = false;
            }

            return rc;
        }

        public static async Task<int> ActualizarVigencia(string cadenaConexion, int tramite, int secuencial)
        {
            Dal dal = new Dal(cadenaConexion)
            {
                Procedimiento = "dbo.SP_SolicitudDetalle",
                Parametros = new List<SqlParameter>
                {
                    new SqlParameter("@opcion", 6),
                    new SqlParameter("@SOLICODSOLI", tramite),
                    new SqlParameter("@SOLISEC", secuencial)
                }
            };

            var result = await dal.EjecutarNoQuery();

            return result;
        }

        public static string GetEstadoCivil(string codigo)
        {
            string estadoCivil = string.Empty;
            switch (codigo)
            {
                case "CA":
                    estadoCivil = "CASADO";
                    break;
                case "CE":
                    estadoCivil = "CASADO (*)";
                    break;
                case "DI":
                    estadoCivil = "DIVORCIADO";
                    break;
                case "ND":
                    estadoCivil = "NO DETERMINADO";
                    break;
                case "SO":
                    estadoCivil = "SOLTERO";
                    break;
                case "UE":
                    estadoCivil = "UNION DE HECHO (*)";
                    break;
                case "UH":
                    estadoCivil = "UNION DE HECHO";
                    break;
                case "VI":
                    estadoCivil = "VIUDO";
                    break;
            }

            return estadoCivil;
        }

        public static void ManipulatePdf(string htmlSource, string pdfDest, string src, byte[] imgRpg, byte[] imgAlcaldia, byte[] imgQr = null, DataFooter dataFooter = null)
        {
            iText.Layout.Element.Image rpg = new iText.Layout.Element.Image(ImageDataFactory.Create(imgRpg)).SetFixedPosition(50, 730);
            //rpg.ScaleToFit(184, 100);
            rpg.SetMaxWidth(138);
            rpg.SetMaxHeight(75);
            iText.Layout.Element.Image alcaldia = new iText.Layout.Element.Image(ImageDataFactory.Create(imgAlcaldia)).SetFixedPosition(340, 730);
            alcaldia.SetMaxHeight(90);
            alcaldia.SetMaxWidth(225);
            //alcaldia.ScaleToFit(158, 112);

            if (imgQr != null)
            {
                iText.Layout.Element.Image qr = new iText.Layout.Element.Image(ImageDataFactory.Create(imgQr)).SetFixedPosition(500, 20);
                qr.SetMaxHeight(80);
                qr.SetMaxWidth(80);

                //dataFooter.Qr = qr.ScaleToFit(80, 80);
                dataFooter.Qr = qr;
            }

            var pdfDocument = new PdfDocument(new PdfWriter(pdfDest));
            
            string header = "";
            //Logo logoHandler = new Logo(rpg, alcaldia);

            Header headerHandler = new Header(header, rpg, alcaldia);
            pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, headerHandler);

            Footer footerHandler = new Footer(dataFooter);
            if (dataFooter != null)
            {
                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, footerHandler);
            }

            //pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, logoHandler);


            // Base URI is required to resolve the path to source files
            ConverterProperties converterProperties = new ConverterProperties();
            converterProperties.SetBaseUri(src);
            converterProperties.SetCssApplierFactory(new QrCodeTagCssApplierFactory())
                .SetTagWorkerFactory(new QrCodeTagWorkerFactory());

            //FileStream fs = new FileStream(htmlSource, FileMode.Open);
            //HtmlConverter.ConvertToDocument(fs, pdfDocument, converterProperties);
            HtmlConverter.ConvertToDocument(htmlSource, pdfDocument, converterProperties);


            //if (fs != null)
            //{
            //    fs.Close();
            //    fs.Dispose();
            //}

            // Write the total number of pages to the placeholder
            if (dataFooter != null)
            {
                footerHandler.WriteTotal(pdfDocument);
            }

            pdfDocument.Close();
        }

        public static void GenerarArchivoHtml(string pathHtml, string dataHtml)
        {
            FileStream fs = null;

            try
            {
                fs = new FileStream(pathHtml, FileMode.Create);
                using (StreamWriter writer = new StreamWriter(fs))
                {
                    writer.Write(dataHtml);
                }
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }
        }
    }
}
