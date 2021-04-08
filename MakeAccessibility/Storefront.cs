using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pageflex.Interfaces.Storefront;
using PDFix.App.Module;
using PDFixSDK.Pdfix;

namespace MakeAccessibility
{
    class Storefront : StorefrontExtension
    {
        public override string DisplayName
        {
            get
            {
                return "MakeAccessibility";
            }
        }

        public override string UniqueName
        {
            get
            {
                return "com.brandgate.MakeAccessibility";
            }
        }

        public override int DocOutput_After(string documentID)
        {
            string output = Storefront.GetValue("DocumentProperty", "FinalOutputLocation", documentID);
            string outputfile = "";
            string fileName = "";

            System.Threading.Thread.Sleep(10000);

            DirectoryInfo d = new DirectoryInfo(output);//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files

            //We only need the latest version
            fileName = Files[Files.Length - 1].Name;
            outputfile += output + "\\" + fileName;

            if (Storefront.GetValue("PrintingField", "PDFAccessible", documentID) == "1")
            {
                Storefront.LogMessage("Making PDF: " + outputfile + " accessible: " + output, null, documentID, 0, false);

                string config = "{\"template\": {\"add_tags\" : [{\"parse_form\": false}]}}";
                string accessiblePDFname = fileName.Substring(0, 11);
                //Call TAG function
                MakeAccessible.Run( outputfile, output + "\\" + accessiblePDFname + "Accessible.pdf",false, "no", fileName, config, "email", "key");
            }

            return eSuccess;
        }
    }
}
