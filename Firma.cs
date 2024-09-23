using GestorDocumentalLibrary;
using iText.Bouncycastle.Crypto;
using iText.Bouncycastle.X509;
using iText.Commons.Bouncycastle.Cert;
using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Signatures;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.X509;
using SAC.CertificadosLibraryI.Auxiliares;

namespace SAC.CertificadosLibraryI
{
    public static class Firma
    {
        public static bool FirmarDocumento(KeyStore keyStore, Documento documento, GestorDocumental gestorDocumental)
        {
            bool firmaOk = false;
            FileStream writer = null;
            FileStream fsKs = null;
            try
            {
                var rutaDocumentoSinFirma = documento.PathSinFirmar + $"\\{documento.Nombre}.pdf";
                var rutaDocumentoFirmado = documento.PathFirmado + $"\\{documento.Nombre}_signed.pdf";

                var reader = new PdfReader(rutaDocumentoSinFirma);
                writer = new FileStream(rutaDocumentoFirmado, FileMode.Create, FileAccess.Write);
                var signer = new PdfSigner(reader, writer, new StampingProperties());

                if (documento.EstamparFirma)
                {
                    PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                    appearance
                        .SetRenderingMode(PdfSignatureAppearance.RenderingMode.GRAPHIC_AND_DESCRIPTION)
                        .SetReason(documento.Razon)
                        .SetLocation(documento.Localizacion)

                        // Specify if the appearance before field is signed will be used
                        // as a background for the signed field. The "false" value is the default value.
                        .SetReuseAppearance(false)
                        //.SetPageRect(rect)
                        //.SetPageNumber(3)
                        .SetSignatureGraphic(ImageDataFactory.Create(documento.Qr));

                    signer.SetFieldName("signField");
                }

                #region carga certificado
                Pkcs12Store pk12 = new Pkcs12StoreBuilder().Build();
                fsKs = new FileStream(keyStore.Path, FileMode.Open, FileAccess.Read);
                pk12.Load(fsKs, keyStore.Password.ToCharArray());
                string alias = null;
                foreach (var a in pk12.Aliases)
                {
                    alias = ((string)a);
                    if (pk12.IsKeyEntry(alias))
                        break;
                }

                ICipherParameters pk = pk12.GetKey(alias).Key;
                X509CertificateEntry[] ce = pk12.GetCertificateChain(alias);
                X509Certificate[] chain = new X509Certificate[ce.Length];
                for (int k = 0; k < ce.Length; ++k)
                {
                    chain[k] = ce[k].Certificate;
                }
                #endregion

                IExternalSignature pks = new PrivateKeySignature(new PrivateKeyBC(pk), DigestAlgorithms.SHA256);

                IX509Certificate[] certificateWrappers = new IX509Certificate[chain.Length];
                for (int i = 0; i < certificateWrappers.Length; ++i)
                {
                    certificateWrappers[i] = new X509CertificateBC(chain[i]);
                }
                // Sign the document using the detached mode, CMS or CAdES equivalent.
                signer.SignDetached(pks, certificateWrappers, null, null, null, 0, PdfSigner.CryptoStandard.CADES);

                //result = Comun.FileToBytes(rutaDocumentoFirmado);


                //Elimina Archivos temporales
                //await Delete(rutaDocumentoSinFirma);
                //await Delete(rutaDocumentoFirmado);
                //await Delete(rutaCompletaDelP12);

                if (documento.CargarEnGestor)
                {
                    //carga el documento en OPENKM
                    Gestor gestor = new Gestor()
                    {
                        HostGestor = gestorDocumental.Host,
                        UserGestor = gestorDocumental.Usuario,
                        Password = gestorDocumental.Contrasena
                    };
                    var rc = gestor.ImportarIndividual(new Archivo
                    {
                        Estado = false,
                        Path = rutaDocumentoFirmado,
                        FolderUuid = documento.RepositorioOkm,
                        UsarArchivoOkm = false,
                        GroupName = documento.GrupoOkm,
                        Metadata = documento.Metadata
                    });

                    gestor = null;
                }

                firmaOk = true;
            }
            catch (GestorException ex)
            {
                throw new CertificadosLibraryException("Error al importar el documento al Gestor Documental.\n" + ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al firmar el documento.\n" + ex.Message, ex);
            }
            finally
            {
                if (fsKs != null)
                {
                    fsKs.Close();
                    fsKs.Dispose();
                }
                if (writer != null)
                {
                    writer.Close();
                    writer.Dispose();
                }
            }

            return firmaOk;
        }


        public static bool FirmarDocumentoTramite(KeyStore keyStore, Documento documento, GestorDocumental gestorDocumental)
        {
            bool firmaOk = false;
            FileStream writer = null;
            FileStream fsKs = null;
            try
            {
                var rutaDocumentoSinFirma = documento.PathSinFirmar + $"\\{documento.Nombre}.pdf";
                var rutaDocumentoFirmado = documento.PathFirmado + $"\\{documento.Nombre}_signed.pdf";

                var reader = new PdfReader(rutaDocumentoSinFirma);
                writer = new FileStream(rutaDocumentoFirmado, FileMode.Create, FileAccess.Write);
                var signer = new PdfSigner(reader, writer, new StampingProperties());

                if (documento.EstamparFirma)
                {
                    PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                    appearance
                        .SetRenderingMode(PdfSignatureAppearance.RenderingMode.GRAPHIC_AND_DESCRIPTION)
                        .SetReason(documento.Razon)
                        .SetLocation(documento.Localizacion)

                        // Specify if the appearance before field is signed will be used
                        // as a background for the signed field. The "false" value is the default value.
                        .SetReuseAppearance(false)
                        //.SetPageRect(rect)
                        //.SetPageNumber(3)
                        .SetSignatureGraphic(ImageDataFactory.Create(documento.Qr));

                    signer.SetFieldName("signField");
                }

                #region carga certificado
                Pkcs12Store pk12 = new Pkcs12StoreBuilder().Build();
                fsKs = new FileStream(keyStore.Path, FileMode.Open, FileAccess.Read);
                pk12.Load(fsKs, keyStore.Password.ToCharArray());
                string alias = null;
                foreach (var a in pk12.Aliases)
                {
                    alias = ((string)a);
                    if (pk12.IsKeyEntry(alias))
                        break;
                }

                ICipherParameters pk = pk12.GetKey(alias).Key;
                X509CertificateEntry[] ce = pk12.GetCertificateChain(alias);
                X509Certificate[] chain = new X509Certificate[ce.Length];
                for (int k = 0; k < ce.Length; ++k)
                {
                    chain[k] = ce[k].Certificate;
                }
                #endregion

                IExternalSignature pks = new PrivateKeySignature(new PrivateKeyBC(pk), DigestAlgorithms.SHA256);

                IX509Certificate[] certificateWrappers = new IX509Certificate[chain.Length];
                for (int i = 0; i < certificateWrappers.Length; ++i)
                {
                    certificateWrappers[i] = new X509CertificateBC(chain[i]);
                }
                // Sign the document using the detached mode, CMS or CAdES equivalent.
                signer.SignDetached(pks, certificateWrappers, null, null, null, 0, PdfSigner.CryptoStandard.CADES);

                if (documento.CargarEnGestor)
                {
                    //carga el documento en OPENKM
                    Gestor gestor = new Gestor()
                    {
                        HostGestor = gestorDocumental.Host,
                        UserGestor = gestorDocumental.Usuario,
                        Password = gestorDocumental.Contrasena
                    };
                    var rc = gestor.ImportarSobreescribirDocumentoTramite(new Archivo
                    {
                        Estado = false,
                        Path = rutaDocumentoFirmado,
                        FolderUuid = documento.RepositorioOkm.Trim(),
                        UsarArchivoOkm = false,
                        GroupName = documento.GrupoOkm,
                        Metadata = documento.Metadata
                    });

                    gestor = null;
                }

                firmaOk = true;
            }
            catch (GestorException ex)
            {
                throw new CertificadosLibraryException("Error al importar el documento al Gestor Documental.\n" + ex.Message, ex);
            }
            catch (Exception ex)
            {
                throw new CertificadosLibraryException("Error al firmar el documento.\n" + ex.Message, ex);
            }
            finally
            {
                if (fsKs != null)
                {
                    fsKs.Close();
                    fsKs.Dispose();
                }
                if (writer != null)
                {
                    writer.Close();
                    writer.Dispose();
                }
            }

            return firmaOk;
        }

    }
}
