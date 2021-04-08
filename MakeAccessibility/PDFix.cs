using System;
using PDFixSDK.Pdfix;

namespace PDFix.App.Module
{
    class MakeAccessible
    {
        public static void Run(
            String openPath,                            // source PDF document
            String savePath,                            // output PDF document
            bool preflight,                             // preflight page before tagging
            String language,                            // document reading language
            String title,                               // document title
            String configPath,
            String email,
            String licenseKey
            // configuration file
            )
        {
            Pdfix pdfix = new Pdfix();
            if (pdfix == null)
                throw new Exception("Pdfix initialization fail");

            PdfDoc doc = pdfix.OpenDoc(openPath, "");
            if (doc == null)
                throw new Exception(pdfix.GetError());

            if (licenseKey.Length > 0)
            {
                if (email.Length > 0)
                {
                    // Authorization using an account name/key
                    var account_auth = pdfix.GetAccountAuthorization();
                    if (account_auth.Authorize(email, licenseKey) == false)
                    {
                        throw new Exception("PDFix SDK Account Authorization failed");
                    }
                }
                else
                {
                    // Authorization using the activation key
                    var standard_auth = pdfix.GetStandardAuthorization();
                    if (!standard_auth.IsAuthorized() && !standard_auth.Activate(licenseKey))
                    {
                        throw new Exception("PDFix SDK Standard Authorization failed");
                    }
                }
            }

            var doc_template = doc.GetTemplate();

            // convert to PDF/UA
            PdfAccessibleParams accParams = new PdfAccessibleParams();
            if (!doc.MakeAccessible(accParams, null, IntPtr.Zero))
                throw new Exception(pdfix.GetError());

            if (!doc.Save(savePath, Pdfix.kSaveFull))
                throw new Exception(pdfix.GetError());

            doc.Close();
            pdfix.Destroy();
        }
    }
}