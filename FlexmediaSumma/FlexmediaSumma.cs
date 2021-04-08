using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;
using Pageflex.Interfaces.Storefront;
using PageflexServices;

namespace FlexmediaSumma
{

    /*
2015-05-26 12:15:30
     */

    /// <summary>
    /// Klass som exporterar orderinformation till en XML-fil som ordersystemet
    /// Summa sedan läser in.
    /// </summary>
    public class FlexmediaSumma : StorefrontExtension
    {
        const string ModuleFieldOutputDirectory = "MD_FLEXSUMMA";
        const string ModuleFieldCopyDirectory = "MD_FLEXSUMMA_COPY";
        const string ModuleFieldCustomerID = "MD_FLEXSUMMACUSTOMERID";
        const string MetaFieldDisableXML = "MD_FLEXSUMMA_DISABLE";
        const string DocumentFieldHasExportedXML = "DF_FLEXSUMMA_HASEXPORTEDXML";

        //new constants from nyaljus summa pris
        public static readonly string ExtraFieldPrefix = "SummaField_";
        public static readonly string EfterbehandlingValPrefix = "EfterbehandlingVal_";
        public static readonly string EfterbehandlingParameterPrefix = "EfterbehandlingParameter[";
        public static readonly string EfterbehandlingKommentarPrefix = "EfterbehandlingKommentar[";
        public static readonly string TillbehorKommentarPrefix = "TillbehorKommentar[";
        public static readonly string TillbehorValPrefix = "TillbehorVal_";
        public static readonly string TillbehorParameterPrefix = "TillbehorParameter[";
        public static readonly string MetaDataFieldProductUsesSUMMAForPrice = "FlexPriceSumma_UseSummaForPrice";

#if DEBUG
        public static bool IsDebug = true;
#else
        public static bool IsDebug = false;
#endif

        public static readonly string MetaDataFieldProductNumner = "Artnr";

        public override string UniqueName
        {
            get
            {
                return "summa.flexmedia.no";
            }
        }

        public override string DisplayName
        {
            get
            {
                return "Flexmedia: Summa";
            }
        }

        public override int DocChangeStatus_After(string docId, string oldStatus, string newStatus, bool isViaSXI)
        {
            string productId = Storefront.GetValue("DocumentProperty", "ProductID", docId);
            string disable = Storefront.GetValue("ProductField", MetaFieldDisableXML, productId);

            if ("YES".Equals(disable))
                return eSuccess;

            // Get status whether info has been exported
            string sHasExported = Storefront.GetValue("PrintingField", DocumentFieldHasExportedXML, docId);

            if (sHasExported != "YES")
            {
                string ProductType = Storefront.GetValue("ProductProperty", "ProductType", productId);
                bool isPageflex = "PAGEFLEX".Equals(ProductType);
                // Storefront.LogMessage("Has not exported yet.", "", docId, 1, false);

                //if (newStatus == "DataPending")
                //if (newStatus == "Rendered")
                if ((newStatus == "Submitted" && isPageflex) || (!isPageflex && "Rendered".Equals(newStatus)))
                {
                    Storefront.LogMessage("Starting export of document information.", "", docId, 1, false);

                    OutputXMLFile(docId);

                    Storefront.LogMessage("Finished export of document information.", "", docId, 1, false);
                    Storefront.SetValue("PrintingField", DocumentFieldHasExportedXML, docId, "YES");
                }
            }
            else
            {
                // Storefront.LogMessage("Has already exported.", "", docId, 1, false);
            }

            return eSuccess;
        }

        public static string internal_villa = "RAS_internal_villa";
        public static string internal_farmhouse = "RAS_internal_farmHouse";
        public static string internal_ownedApartment = "RAS_internal_ownedApartment";
        public static string internal_rentedApartment = "RAS_internal_rentedApartment";
        public static string internal_unknown = "RAS_internal_unknown";

        private int OutputXMLFile(string docId)
        {
            if (IsDebug) Storefront.LogMessage("After outputing document " + Storefront.GetValue("DocumentProperty", "ExternalID", docId) + ", generating document xml.", "", docId, 3, false);

            // See if this is a test order and hence no xml should be generated
            if (Storefront.GetValue("PrintingField", "Instructions", docId) == "FLEXTEST")
            {
                Storefront.LogMessage("This is a test document, no xml for Summa is generated.", "", docId, 1, false);
                return eSuccess;
            }

            // Get order
            string orderId = Storefront.GetValue("DocumentProperty", "OrderGroupID", docId);
            if (IsDebug) Storefront.LogMessage("orderId: " + orderId, "", docId, 3, false);
            if (orderId == null) // Make sure the document is associated with an order.
            {
                Storefront.LogMessage("The generated document is not part of an order, so no xml file is generated.", orderId, docId, 1, false);
            }

            // Get the product ID
            string sProductID = Storefront.GetValue("DocumentProperty", "ProductID", docId);
            if (IsDebug) Storefront.LogMessage("sProductID: " + sProductID, "", docId, 3, false);
            // Read extension properties
            string sOutputDirectory = Storefront.GetValue("ProductField", ModuleFieldOutputDirectory, sProductID);

            if(String.IsNullOrEmpty(sOutputDirectory) || !hasWriteAccessToFolder(sOutputDirectory))
                sOutputDirectory = Storefront.GetValue("ModuleField", ModuleFieldOutputDirectory, UniqueName);


            string sCopyDirectory = Storefront.GetValue("ProductField", ModuleFieldCopyDirectory, sProductID);

            if (String.IsNullOrEmpty(sCopyDirectory) || !hasWriteAccessToFolder(sCopyDirectory))
                sCopyDirectory = Storefront.GetValue("ModuleField", ModuleFieldCopyDirectory, UniqueName);
            
            if (IsDebug) Storefront.LogMessage("sOutputDirectory: " + sOutputDirectory, "", docId, 3, false);
            string sDeploymentCode = Storefront.GetValue("ModuleField", ModuleFieldCustomerID, UniqueName);
            if (IsDebug) Storefront.LogMessage("sDeploymentCode: " + sDeploymentCode, "", docId, 3, false);
            string sExternalID = Storefront.GetValue("DocumentProperty", "ExternalID", docId);
            if (IsDebug) Storefront.LogMessage("sExternalID: " + sExternalID, "", docId, 3, false);

            // Make sure this document is not an 'annons'
            // First, we nee need to check that we have at least one order which is not an 'annons'
            string sPrintingType = Storefront.GetValue("ProductField", "PrintingType", sProductID);
            if (IsDebug) Storefront.LogMessage("sPrintingType: " + sPrintingType, "", docId, 3, false);
            sPrintingType = (sPrintingType == null) ? "" : sPrintingType.ToLower(); // Make sure we actually have a value
            if (sPrintingType == "annons")
            {
                Storefront.LogMessage("Document '" + docId + "' is an 'annons', no xml file generated", orderId, docId, 1, false);
                return eSuccess;
            }
            XmlTextWriter xmlWriter = null;
            //XmlTextWriter xmlWriter = new XmlTextWriter(sOutputDirectory + "\\" + sDeploymentCode + "-" + sExternalID + ".xml", null);

            string typ = Storefront.GetValue("PrintingField", "typ", docId);

            string outputFilePath = sOutputDirectory + "\\" + sDeploymentCode + "-" + sExternalID + ".xml";

            string RAS_internal_onlyphone = Storefront.GetValue("VariableValue", "RAS_internal_onlyphone", docId);
            string RAS_internal_fritidshus = Storefront.GetValue("VariableValue", "RAS_internal_fritidshus", docId);

           

            bool onlyPhoneSelection = false, fritidshus = false, is_villa_only = false, is_villa = false, is_farmHouse=false, is_rentedApartment = false, is_ownedApartment = false, is_unknown = false;
            if (!String.IsNullOrEmpty(RAS_internal_onlyphone) && "true".Equals(RAS_internal_onlyphone))
            {
                onlyPhoneSelection = true;
            }

            if (!String.IsNullOrEmpty(RAS_internal_fritidshus) && "true".Equals(RAS_internal_fritidshus))
            {
                fritidshus = true;
            }
            else {

                string RAS_villa = Storefront.GetValue("VariableValue", internal_villa, docId);
                
                is_villa = (!String.IsNullOrEmpty(RAS_villa) && "true".Equals(RAS_villa));
                if(is_villa){
                    string RAS_unknown = Storefront.GetValue("VariableValue", internal_unknown, docId);
                    string RAS_rented = Storefront.GetValue("VariableValue", internal_rentedApartment, docId);
                    string RAS_owned = Storefront.GetValue("VariableValue", internal_ownedApartment, docId);
                    string RAS_farmhouse = Storefront.GetValue("VariableValue", internal_farmhouse, docId);

                    is_unknown = (!String.IsNullOrEmpty(RAS_unknown) && "true".Equals(RAS_unknown));
                    is_rentedApartment = (!String.IsNullOrEmpty(RAS_rented) && "true".Equals(RAS_rented));
                    is_ownedApartment = (!String.IsNullOrEmpty(RAS_owned) && "true".Equals(RAS_owned));
                    is_farmHouse = (!String.IsNullOrEmpty(RAS_farmhouse) && "true".Equals(RAS_farmhouse));



                    is_villa_only = is_villa && !is_unknown && !is_rentedApartment && !is_ownedApartment;
                }
            }

            string antal = Storefront.GetValue("PrintingField", "PrintingQuantity", docId);

            try
            {
                string antalAlt = Storefront.GetValue("PrintingField", "PrintingQuantity2", docId);
                if (!String.IsNullOrEmpty(antalAlt))
                    antal = antalAlt;
            }
            catch { }

            try
            {
                string timedate = Storefront.GetValue("OrderProperty", "DateTimePlaced", orderId);
                timedate = timedate.Replace("'","").Replace("-","").Replace(":","");
                string[] timedatear = timedate.Split('T');
                string productName = Storefront.GetValue("DocumentProperty", "ProductName", docId);
                string PrintingFieldcustomer = Storefront.GetValue("PrintingField", "customer", docId);
                string ProductID = Storefront.GetValue("DocumentProperty", "ProductID", docId);
                string TimecutODR = Storefront.GetValue("ProductField", "TimecutODR", ProductID);
                if (String.IsNullOrEmpty(PrintingFieldcustomer) || "Fox".Equals(PrintingFieldcustomer))
                    PrintingFieldcustomer = Storefront.GetValue("UserProperty", "LogonName", Storefront.GetValue("DocumentProperty", "OwnerID", docId));


                if (!"yes".Equals(TimecutODR))
                {
                    if ("Endast tryck".Equals(typ))
                    {
                        productName = "Utskick_" + Storefront.GetValue("UserProperty", "LogonName", Storefront.GetValue("DocumentProperty", "OwnerID", docId));
                    }
                }

                if (onlyPhoneSelection)
                {
                    productName = productName.Replace("adress", "") + "_UF";

                    //productName += "_UF";
                }
                outputFilePath = sDeploymentCode + "_" +
                    PrintingFieldcustomer + "_" +productName;

                

                outputFilePath += "_" + timedatear[0] + "_" + timedatear[1].Substring(0, 4) + "_" +
                    sExternalID + "_" + antal + "ex.xml";

                if ("Endast tryck".Equals(typ)) {
                    string sdr = Storefront.GetValue("PrintingField", "sdr", docId);
                    if(!String.IsNullOrEmpty(sdr))
                        outputFilePath = sdr + outputFilePath;
                }

                outputFilePath = sOutputDirectory + "\\" + outputFilePath;

                xmlWriter = new XmlTextWriter(outputFilePath, null);
            }
            catch (Exception e)
            {
                Storefront.LogMessage("Error creating xml file: " + outputFilePath+ "," + e.Message, orderId, docId, 1, false);
                xmlWriter = new XmlTextWriter(outputFilePath, null);
            }



            try
            {
                xmlWriter.Formatting = Formatting.Indented;

                xmlWriter.WriteStartDocument();

                xmlWriter.WriteStartElement("order");

                // Write customer info
                string sUserID = Storefront.GetValue("DocumentProperty", "OwnerID", docId);
                if (IsDebug) Storefront.LogMessage("sUserID: " + sUserID, "", docId, 3, false);
                string sCustomCustomerID = Storefront.GetValue("UserField", "UserProfileCustomerID", sUserID);
                if (IsDebug) Storefront.LogMessage("sCustomCustomerID: " + sCustomCustomerID, "", docId, 3, false);
                if (sCustomCustomerID != null && sCustomCustomerID != "")
                {
                    XmlWriterWrite(xmlWriter, "CustomerID", sCustomCustomerID);
                }
                else
                {
                    XmlWriterWrite(xmlWriter, "CustomerID", Storefront.GetValue("ModuleField", ModuleFieldCustomerID, UniqueName));
                }

                // Write currency info
                XmlWriterWrite(xmlWriter, "Currency", Storefront.GetValue("SystemProperty", "IsoCurrencyCode", null));

                // Write order info
                XmlWriterWrite(xmlWriter, "OrderID", Storefront.GetValue("OrderProperty", "ExternalID", orderId));

                //////// Write payment and shipping info
                XmlWriterWrite(xmlWriter, "DeliveryFirstName", Storefront.GetValue("OrderField", "DeliveryFirstName", orderId));

                string DeliveryLastName = Storefront.GetValue("OrderField", "DeliveryLastName", orderId);
                string DeliveryAddress1 = Storefront.GetValue("OrderField", "DeliveryAddress1", orderId);
                string DeliveryPostalCode = Storefront.GetValue("OrderField", "DeliveryPostalCode", orderId);
                string DeliveryCity = Storefront.GetValue("OrderField", "DeliveryCity", orderId);

                if ("Endast tryck".Equals(typ))
                {
                    string DeliveryCompanyMW = Storefront.GetValue("PrintingField", "DeliveryCompanyMW", docId);
                    string DeliveryAdressMW = Storefront.GetValue("PrintingField", "DeliveryAdressMW", docId);
                    string DeliveryZipMW = Storefront.GetValue("PrintingField", "DeliveryZipMW", docId);
                    string DeliveryCityMW = Storefront.GetValue("PrintingField", "DeliveryCityMW", docId);

                    if (!String.IsNullOrEmpty(DeliveryCompanyMW))
                        DeliveryLastName = DeliveryCompanyMW;
                    if (!String.IsNullOrEmpty(DeliveryAdressMW))
                        DeliveryAddress1 = DeliveryAdressMW;
                    if (!String.IsNullOrEmpty(DeliveryZipMW))
                        DeliveryPostalCode = DeliveryZipMW;
                    if (!String.IsNullOrEmpty(DeliveryCityMW))
                        DeliveryCity = DeliveryCityMW;
                }

                XmlWriterWrite(xmlWriter, "DeliveryLastName", DeliveryLastName);

                try
                {
                    string VarCustomerID = Storefront.GetValue("VariableValue", "CustomerID", docId);
                    string PrintingCustomerID = Storefront.GetValue("PrintingField", "CustomerID", docId);
                    string OfficeCode = Storefront.GetValue("PrintingField", "OfficeCode", docId);

                    if("yes".Equals(Storefront.GetValue("ProductField", "UseOfficeCode", sProductID)))
                        XmlWriterWrite(xmlWriter, "DeliveryCustomerID", OfficeCode);
                    else if (!(String.IsNullOrEmpty(VarCustomerID)))
                        XmlWriterWrite(xmlWriter, "DeliveryCustomerID", VarCustomerID);
                    else if(!(String.IsNullOrEmpty(PrintingCustomerID)))
                        XmlWriterWrite(xmlWriter, "DeliveryCustomerID", PrintingCustomerID);
                    else
                        XmlWriterWrite(xmlWriter, "DeliveryCustomerID", Storefront.GetValue("UserField", "UserProfileLastName", sUserID));
                    
                }
                catch { }

                string sDeliveryFullAddress = Storefront.GetValue("OrderField", "DeliveryFullAddressBulow", orderId);
                // if (sDeliveryFullAddress == null || sDeliveryFullAddress == "")
                //   {
                XmlWriterWrite(xmlWriter, "DeliveryAddress1", DeliveryAddress1);
                XmlWriterWrite(xmlWriter, "DeliveryPostalCode", DeliveryPostalCode);
                XmlWriterWrite(xmlWriter, "DeliveryCity", DeliveryCity);
                // XmlWriterWrite(xmlWriter, "DeliveryCountry", Storefront.GetValue("OrderField", "DeliveryCountry", orderId));
                //    }
                //    else
                //    {
                //XmlWriterWrite(xmlWriter, "DeliveryFullAddress", sDeliveryFullAddress);

                XmlWriterWrite(xmlWriter, "DeliveryFullAddress", sDeliveryFullAddress);

                string DeliveryAddressInfo = Storefront.GetValue("OrderField", "DeliveryAddressInfo", orderId);


                XmlWriterWrite(xmlWriter, "DeliveryAddressInfo", DeliveryAddressInfo);

                XmlWriterWrite(xmlWriter, "information", Storefront.GetValue("UserField", "UserProfileLevInfo", sUserID));

                //    }

                string sShowPayment = Storefront.GetValue("OrderField", "show_payment", orderId);
                if (IsDebug) Storefront.LogMessage("sShowPayment: " + sShowPayment, "", docId, 3, false);
                if (sShowPayment != null && sShowPayment != "")
                {
                    string sPaymentFullAddress = Storefront.GetValue("OrderField", "PaymentFullAddress", orderId);
                    if (IsDebug) Storefront.LogMessage("sPaymentFullAddress: " + sPaymentFullAddress, "", docId, 3, false);
                    if (sPaymentFullAddress == null || sPaymentFullAddress == "")
                    {
                        //XmlWriterWrite(xmlWriter, "PaymentAddress1", Storefront.GetValue("OrderField", "PaymentAddress1", orderId));
                        //XmlWriterWrite(xmlWriter, "PaymentPostalCode", Storefront.GetValue("OrderField", "PaymentPostalCode", orderId));
                        //XmlWriterWrite(xmlWriter, "PaymentCity", Storefront.GetValue("OrderField", "PaymentCity", orderId));
                        //XmlWriterWrite(xmlWriter, "PaymentCountry", Storefront.GetValue("OrderField", "PaymentCountry", orderId));
                    }
                    else
                    {
                        XmlWriterWrite(xmlWriter, "PaymentFullAddress", sPaymentFullAddress);
                    }
                }

                // Contact information
                XmlWriterWrite(xmlWriter, "ContactFirstName", Storefront.GetValue("OrderField", "ContactFirstName", orderId));
                XmlWriterWrite(xmlWriter, "ContactLastName", Storefront.GetValue("OrderField", "ContactLastName", orderId));

                XmlWriterWrite(xmlWriter, "Hanteringskostnad", Storefront.GetValue("OrderProperty", "HandlingCharge", orderId));


                XmlWriterWrite(xmlWriter, "OrderDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // Extra fields
                for (int i = 1; i <= 5; i++)
                {
                    string sExtraField = Storefront.GetValue("OrderField", "SummaExtraField" + i.ToString(), orderId);
                    string sExtraFieldLabel = Storefront.GetValue("OrderField", "SummaExtraField" + i.ToString() + "Label", orderId);
                    bool bHasExtraField = (sExtraField != null);
                    bool bHasExtraFieldLabel = (sExtraFieldLabel != null && sExtraFieldLabel != "");
                    if (bHasExtraField)
                    {
                        if (bHasExtraFieldLabel)
                        {
                            XmlWriterWrite(xmlWriter, sExtraFieldLabel, sExtraField);
                        }
                        else
                        {
                            XmlWriterWrite(xmlWriter, "SummaExtraField" + i.ToString(), sExtraField);
                        }
                    }
                }
                if (IsDebug) Storefront.LogMessage("--AFTER EXTRA FIELDS: ", "", docId, 3, false);
                ///// Find Price
                // Check if it's a 'plock'-product
                string sIsPlock = Storefront.GetValue("ProductField", "Plock", sProductID);
                string sThePrice = "";
                if (sIsPlock != null && sIsPlock == "X") // If 'plock' that should be free, set price to zero.
                {
                    sThePrice = "0,00";
                }
                else
                {
                    // Check if we have a price table
                    string sPriceTable = Storefront.GetValue("DocumentProperty", "PerDocumentPriceTable", docId);
                    string sLeafTable = Storefront.GetValue("DocumentProperty", "PerLeafPriceTable", docId);
                    if (sPriceTable != null || sLeafTable != null || !String.IsNullOrEmpty(Storefront.GetValue("ProductField", MetaDataFieldProductUsesSUMMAForPrice, sProductID)))
                    {
                        sThePrice = Storefront.GetValue("DocumentProperty", "Price", docId);
                        if (sThePrice != null)
                            sThePrice = sThePrice.Replace(".", ",");
                    }
                    else // No price
                    {
                        sThePrice = "NA";
                    }
                }

                // Write info
                xmlWriter.WriteStartElement("document");

                
                if ("yes".Equals(Storefront.GetValue("PrintingField", "UseMW", docId)) || "yes".Equals(Storefront.GetValue("PrintingField", "UseTimecut", docId)))
                {
                    string ordernr = Storefront.GetValue("PrintingField", "ordernr", docId);
                    //mindworking file name
                    if (!String.IsNullOrEmpty(ordernr))
                        XmlWriterWrite(xmlWriter, "DocumentID", ordernr);
                    else
                        XmlWriterWrite(xmlWriter, "DocumentID", sExternalID);
                }
                else
                {
                    XmlWriterWrite(xmlWriter, "DocumentID", sExternalID);
                }
                string extra_bilder_choose = Storefront.GetValue("VariableValue", "extra_bilder_choose", docId);
                if (extra_bilder_choose == "" || extra_bilder_choose == null)
                {
                    extra_bilder_choose = Storefront.GetValue("PrintingField", "extra_bilder_choose", docId);
                }
                string sArtno = "";

                if (!String.IsNullOrEmpty(extra_bilder_choose))
                {
                    if (IsDebug) Storefront.LogMessage("extra_bilder_choose: " + extra_bilder_choose, "", docId, 3, false);

                    string current_chili_doc = Storefront.GetValue("VariableValue", "CHILI_PRODUCT_ID", docId);

                    string PagePlannerParams = Storefront.GetValue("ProductField", "PP_METADATA_FIELD_PARAMS", sProductID);
                    if (String.IsNullOrEmpty(current_chili_doc) || String.IsNullOrEmpty(PagePlannerParams))
                    {
                        sArtno = extra_bilder_choose;
                    }
                    else { 
                        //PP product
                        sArtno = Storefront.GetValue("ProductField", MetaDataFieldProductNumner, sProductID) + extra_bilder_choose;
                    }
                }
                else
                    sArtno = Storefront.GetValue("ProductField", MetaDataFieldProductNumner, sProductID);

                if (IsDebug) Storefront.LogMessage("sArtno: " + sArtno, "", docId, 3, false);

                string ProductNumber = "";

                if (onlyPhoneSelection)
                {
                    //XmlWriterWrite(xmlWriter, "ProductNumber", "7055");
                    ProductNumber = "7055";
                }
                else if (fritidshus)
                {
                    //XmlWriterWrite(xmlWriter, "ProductNumber", "3"+sArtno.Substring(1));
                    ProductNumber = "3"+sArtno.Substring(1);
                }
                else if (is_farmHouse && is_villa_only)
                {
                    //XmlWriterWrite(xmlWriter, "ProductNumber", "31" + sArtno.Substring(2));
                    ProductNumber = "31" + sArtno.Substring(2);
                }
                else
                {
                    if (sArtno != null && sArtno != "")
                    {
                        if ("Endast tryck".Equals(typ))
                        {
                            Storefront.LogMessage("Correct xml.", orderId, docId, 1, false);
                            //XmlWriterWrite(xmlWriter, "ProductNumber", Storefront.GetValue("ProductField", "ArtnrET", sProductID));
                            ProductNumber = Storefront.GetValue("ProductField", "ArtnrET", sProductID) + extra_bilder_choose;
                        }
                        else
                        {
                            //XmlWriterWrite(xmlWriter, "ProductNumber", sArtno);
                            ProductNumber = sArtno;
                        }
                    }
                    else
                    {
                        //XmlWriterWrite(xmlWriter, "ProductNumber", "");
                        ProductNumber = "";
                    }
                }


                int? numBringAddresses = 0;
                int? numOtherAddresses = 0;
                int? numBringNormalAddresses = 0;
                bool isRommFilter = false;
                bool isMovingFilter = false;

                bool isRacasseProduct = false;

                try
                {
                    string internal_response = "RAS_internalresponse";
                    string addressCountData = Storefront.GetValue("VariableValue", internal_response, docId);
                    Storefront.LogMessage("RAS_internalresponse:" + docId + " " + addressCountData, orderId, docId, 1, false);

                    if (!String.IsNullOrEmpty(addressCountData))
                    {

                        COLoackSearchResult lr = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<COLoackSearchResult>(addressCountData);
                        numBringAddresses = lr.numBringAddresses + lr.numRefexBringAddresses;
                        numOtherAddresses = lr.numOtherAddresses + lr.numRefexOtherAddresses;

                        numBringNormalAddresses = lr.numBringNormalAddresses + lr.numRefexBringNormalAddresses;

                        isRommFilter = lr.roomsOwnedApartment.Length > 0 || lr.roomsRentedApartment.Length > 0;
                        isMovingFilter = lr.lastMove.Length > 0;
                        Storefront.LogMessage("isRommFilter: " + isRommFilter + " hasmoved: "+ isMovingFilter + lr.lastMove.Length, orderId, docId, 1, false);

                        isRacasseProduct = true;
                    }


                }
                catch (Exception ex)
                {

                    Storefront.LogMessage("Exception parse RAS_internalresponse:" + ex.Message, orderId, docId, 1, false);

                }

                //if (!isRacasseProduct || (numBringAddresses != null && numBringAddresses > 0))
                //{
                    XmlWriterWrite(xmlWriter, "ProductNumber", ProductNumber);
                //}

                if (numBringNormalAddresses != null && numBringNormalAddresses > 0)
                {
                    if (!String.IsNullOrEmpty(ProductNumber))
                    {
                        StringBuilder sb = new StringBuilder(ProductNumber);


                        if (fritidshus)
                        {
                            sb[0] = '3';
                            sb[1] = '3';
                        }
                        else if (is_farmHouse)
                        {
                            sb[0] = '3';
                            sb[1] = '4';
                        }
                        else
                        {
                            sb[1] = '3';
                        }
                        string ProductNumberNormal = sb.ToString();
                        XmlWriterWrite(xmlWriter, "ProductNumberNormal", ProductNumberNormal);
                    }
                }

                string AntalBringNormal = Storefront.GetValue("PrintingField", "PrintingQuantity2BringNormal", docId);
                if (!String.IsNullOrEmpty(AntalBringNormal) && !AntalBringNormal.Equals("0"))
                {

                    if (!String.IsNullOrEmpty(ProductNumber))
                    {
                        StringBuilder sb = new StringBuilder(ProductNumber);
                        sb[1] = '3';
                        string ProductNumberNormal = sb.ToString();

                        XmlWriterWrite(xmlWriter, "ProductNumberNormal", ProductNumberNormal);
                    }
                }

                string PrintingQuantity2Posten = Storefront.GetValue("PrintingField", "PrintingQuantity2Posten", docId);

                if (!String.IsNullOrEmpty(PrintingQuantity2Posten) && !PrintingQuantity2Posten.Equals("0"))
                {
                    string ProductNumberPosten = "";
                    if (ProductNumber.Length > 1)
                        ProductNumberPosten = "2" + ProductNumber.Substring(1);

                    XmlWriterWrite(xmlWriter, "ProductNumberPosten", ProductNumberPosten);
                }

               

               

                string fbcount = Storefront.GetValue("VariableValue", "RAS_internal_fbcount", docId);

                if (numOtherAddresses!=null && numOtherAddresses > 0)
                {
                    //Och värdet för den ska hämtas från metadatafältet Artnr för produkten, men den första siffran i det värdet skall bytas ut mot en 2:a.
                    string artnr = Storefront.GetValue("ProductField", MetaDataFieldProductNumner, sProductID);
                    if (!String.IsNullOrEmpty(artnr))
                    {
                        if(fritidshus)
                            XmlWriterWrite(xmlWriter, "ProductNumberPosten", "4" + artnr.Substring(1));
                        else if (is_farmHouse)
                            XmlWriterWrite(xmlWriter, "ProductNumberPosten", "41" + artnr.Substring(2));

                        else XmlWriterWrite(xmlWriter, "ProductNumberPosten", "2" + artnr.Substring(1));
                    }
                }
                /*<ProductNumberAntalRum>7070</ProductNumberAntalRum>

should be added in case at least one of roomsOwnedApartment or roomsRentedApartment is added?
                 */
                if(isRommFilter){
                    if (onlyPhoneSelection)
                    {
                        XmlWriterWrite(xmlWriter, "ProductNumberAntalRum", "7075");
                    }
                    else
                    {
                        XmlWriterWrite(xmlWriter, "ProductNumberAntalRum", "7070");
                    }
                }

                if (!String.IsNullOrEmpty(fbcount))
                {
                    if (fbcount.Equals("1000"))
                        XmlWriterWrite(xmlWriter, "ProductNumberFacebook", "7060");
                    else if (fbcount.Equals("5000"))
                        XmlWriterWrite(xmlWriter, "ProductNumberFacebook", "7065");
                }

                //If lastMove is not used (empty), tag <ProductNumberFlytt> should not be added.
                if (isMovingFilter)
                {
                    XmlWriterWrite(xmlWriter, "ProductNumberFlytt", "7080");
                }

                //if (!"Endast tryck".Equals(typ))
                    XmlWriterWrite(xmlWriter, "Productname", Storefront.GetValue("ProductProperty", "DisplayName", sProductID));
      

                if (IsDebug) Storefront.LogMessage("--Productname", "", docId, 3, false);

                //XmlWriterWrite(xmlWriter, "PdfName", Storefront.GetValue("PrintingField", "RenamedOuputFilename", docId));
                //ProductField.ItemDescriptionTemplate
                //example:
                //<ProductField:Kund>_<DocumentProperty:ProductName>_v<PrintingField:Week>_<Order:OrderDate>_<Order:OrderTime>_<DocumentProperty:ExternalID>_<PrintingField:PrintingQuantity>ex
                if ("yes".Equals(Storefront.GetValue("PrintingField", "UseMW", docId)))
                {
                    XmlWriterWrite(xmlWriter, "PdfName", Storefront.GetValue("PrintingField", "ordernr", docId) + ".pdf");//filename
                }
                else
                {
                    try
                    {
                        string outputFilneName = new FieldParser()
                        {
                            DocumentID = docId,
                            OrderID = orderId,
                            UserID = this.Storefront.GetValue("SystemProperty", "LoggedOnUserID", (string)null)
                        }.Parse(Storefront.GetValue("ProductField", "ItemDescriptionTemplate", sProductID));

                        XmlWriterWrite(xmlWriter, "PdfName", outputFilneName + ".pdf");
                    }
                    catch (Exception ex)
                    {
                        Storefront.LogMessage("Exception while parsing:" + ex.Message, "", docId, 3, false);
                    }
                }
                string ProductID = Storefront.GetValue("DocumentProperty", "ProductID", docId);
                string TimecutODR = Storefront.GetValue("ProductField", "TimecutODR", ProductID);
                string TimecutPDFName = Storefront.GetValue("ProductField", "TimecutPDFName", ProductID);                
                if ("yes".Equals(TimecutODR) && !(String.IsNullOrEmpty(TimecutPDFName)))
                {
                    string ObjectStreetAddress = Storefront.GetValue("PrintingField", "ObjectStreetAddress", docId);
                    if(!String.IsNullOrEmpty(ObjectStreetAddress))
                    {
                        XmlWriterWrite(xmlWriter, "Description", ObjectStreetAddress);
                    }
                    else
                    {
                        XmlWriterWrite(xmlWriter, "Description", Storefront.GetValue("DocumentProperty", "Description", docId));
                    }
                }
                else
                {
                    //XmlWriterWrite(xmlWriter, "ItemDescriptionTemplate", Storefront.GetValue("DocumentProperty", "FinalOutputLocation", docId));
                    XmlWriterWrite(xmlWriter, "Description", Storefront.GetValue("DocumentProperty", "Description", docId));
                    if (IsDebug) Storefront.LogMessage("--Description", "", docId, 3, false);
                    // XmlWriterWrite(xmlWriter, "Quantity", Storefront.GetValue("PrintingField", "PrintingQuantity", docId));
                }

                if (onlyPhoneSelection)
                {
                    antal = "0";
                }
                XmlWriterWrite(xmlWriter, "Antal", antal);
                if (IsDebug) Storefront.LogMessage("--Antal", "", docId, 3, false);

                bool antalBringSet = false,
                    antalPostenSet = false;

                if (!"Endast tryck".Equals(typ))
                {
                    if (numBringAddresses != null && numBringAddresses > 0)
                    {
                        XmlWriterWrite(xmlWriter, "AntalBring", numBringAddresses.ToString());
                        antalBringSet = true;
                    }
                }

                string PrintingQuantity2Bring = Storefront.GetValue("PrintingField", "PrintingQuantity2Bring", docId);
                if (!antalBringSet && !String.IsNullOrEmpty(PrintingQuantity2Bring) && !PrintingQuantity2Bring.Equals("0"))
                {
                    XmlWriterWrite(xmlWriter, "AntalBring", PrintingQuantity2Bring);
                }

                if (numBringNormalAddresses != null && numBringNormalAddresses > 0)
                {
                    XmlWriterWrite(xmlWriter, "AntalBringNormal", numBringNormalAddresses.ToString());

                   
                }

                //string AntalBringNormal = Storefront.GetValue("PrintingField", "PrintingQuantity2BringNormal", docId);
                if (!String.IsNullOrEmpty(AntalBringNormal) && !AntalBringNormal.Equals("0"))
                {
                    XmlWriterWrite(xmlWriter, "AntalBringNormal", AntalBringNormal);
                }

                if (!"Endast tryck".Equals(typ))
                {
                    if (numOtherAddresses != null && numOtherAddresses > 0)
                    {
                        XmlWriterWrite(xmlWriter, "AntalPosten", numOtherAddresses.ToString());
                        antalPostenSet = true;
                    }
                }
                
                


                


                if (!antalPostenSet && !String.IsNullOrEmpty(PrintingQuantity2Bring) && !PrintingQuantity2Posten.Equals("0"))
                {
                    XmlWriterWrite(xmlWriter, "AntalPosten", PrintingQuantity2Posten);
                }

                

                /*string AntalTel = "";
                try
                {
                    string internal_response = "RAS_internalresponse";
                    string addressCountData = Storefront.GetValue("VariableValue", internal_response, docId);

                    if (!String.IsNullOrEmpty(addressCountData))
                    {

                        COLoackSearchResult lr = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<COLoackSearchResult>(addressCountData);
                        if (lr.numWithPhone != null)
                           AntalTel = lr.numWithPhone.ToString();
                    }
                    

                }
                catch (Exception ex)
                {

                    Storefront.LogMessage("Exception AntalTel:" + ex.Message, orderId, docId, 1, false);

                }*/

                XmlWriterWrite(xmlWriter, "AntalTel", Storefront.GetValue("PrintingField", "telephonecount", docId));

                if (!String.IsNullOrEmpty(fbcount))
                    XmlWriterWrite(xmlWriter, "AntalFacebook", fbcount);

                XmlWriterWrite(xmlWriter, "Instructions", Storefront.GetValue("PrintingField", "Instructions", docId) + Storefront.GetValue("VariableValue", "Instructions", docId));
                if (IsDebug) Storefront.LogMessage("--Instructions", "", docId, 3, false);
                //XmlWriterWrite(xmlWriter, "ProductNumber", Storefront.GetValue("ProductField", MetaDataFieldProductNumner, sProductID));

                //new fieldssda
                XmlWriterWrite(xmlWriter, "RefFirstName", Storefront.GetValue("PrintingField", "refex_name", docId));
                XmlWriterWrite(xmlWriter, "RefLastName", Storefront.GetValue("PrintingField", "refex_lastname", docId));
                XmlWriterWrite(xmlWriter, "RefAddress1", Storefront.GetValue("PrintingField", "refex_street", docId));
                XmlWriterWrite(xmlWriter, "RefPostalCode", Storefront.GetValue("PrintingField", "refex_zip", docId));
                XmlWriterWrite(xmlWriter, "RefCity", Storefront.GetValue("PrintingField", "refex_city", docId));

                XmlWriterWrite(xmlWriter, "Product", Storefront.GetValue("ProductField", "Product", sProductID));
                XmlWriterWrite(xmlWriter, "Format", Storefront.GetValue("ProductField", "Format", sProductID));
                XmlWriterWrite(xmlWriter, "Papper", Storefront.GetValue("ProductField", "Papper", sProductID));
                XmlWriterWrite(xmlWriter, "Distributor", Storefront.GetValue("ProductField", "Distributor", sProductID));


                //UserProfileCustomerDelivery_Date
                //
                string Delivery_Date = Storefront.GetValue("UserField", "UserProfileCustomerDelivery_Date", sUserID);
                if (String.IsNullOrEmpty(Delivery_Date))
                    Delivery_Date = Storefront.GetValue("ProductField", "Delivery_Date", sProductID);
                XmlWriterWrite(xmlWriter, "Delivery_Date", Delivery_Date);

               

                // XmlWriterWrite(xmlWriter, "DeliveryCity", Storefront.GetValue("OrderField", "DeliveryFullAddressBulow", sProductID));


                string bestalld_adress = Storefront.GetValue("ProductField", "bestalld_adress", sProductID);

                if ("Endast tryck".Equals(typ))               
                    XmlWriterWrite(xmlWriter, "Adr", "");
                else
                    if ("Beställd adress>".Equals(bestalld_adress))
                    XmlWriterWrite(xmlWriter, "Adr", "Y");

                try
                {
                    //Storefront.LogMessage("Adr_typ:" + Storefront.GetValue("ProductField", "Adr_typ", sProductID) + ", sProductID:" + sProductID, null, docId, 3, false);
                    //if ("Endast tryck".Equals(typ))                    
                    //    XmlWriterWrite(xmlWriter, "Adr_typ","");
                   // else 
                        XmlWriterWrite(xmlWriter, "Adr_typ", Storefront.GetValue("ProductField", "Adr_typ", sProductID));
                }
                catch (Exception ex) {
                    Storefront.LogMessage("EXCEPTION Adr_typ:"+ex.Message, null, docId, 3, false);
                }

                try
                {
                    string antalsidor = Storefront.GetValue("ProductField", "AntalSidor", sProductID);//AntalSidor
                    if (IsDebug) Storefront.LogMessage("antalsidor:" + antalsidor, null, docId, 3, false);

                    
                    if (IsDebug) Storefront.LogMessage("extra_bilder_choose:" + extra_bilder_choose, null, docId, 3, false);
                    string antalsid;
                    //string antalsid = (!String.IsNullOrEmpty(antalsidor))?antalsidor:extra_bilder_choose;
                    if (!String.IsNullOrEmpty(antalsidor))
                    {
                        antalsid = antalsidor;
                    }
                    else
                    {
                        antalsid = extra_bilder_choose;
                        if (extra_bilder_choose.Length > 2)
                            antalsid = extra_bilder_choose.Substring(extra_bilder_choose.Length - 2);
                    }

                    if (IsDebug) Storefront.LogMessage("antalsid:" + antalsid, null, docId, 3, false);

                    XmlWriterWrite(xmlWriter, "AntalSidor", antalsid);
                }
                catch {
                   
                }

                /*XmlWriterWrite(xmlWriter, "Format", Storefront.GetValue("ProductField", "Format", sProductID));

                

                XmlWriterWrite(xmlWriter, "Efterbehandling", Storefront.GetValue("ProductField", "Efterbehandling", sProductID));

                */
                //if (IsDebug) Storefront.LogMessage("--ProductNumber", "", docId, 3, false);
                // Add Efterbehandling and Tillbehor
                // Note that Efterbehandling is separated by newlines (\r), which we'll need to change to commas
                string sEfterbehandling = Storefront.GetValue("PrintingField", "Efterbehandling", docId);
                if (sEfterbehandling != null) sEfterbehandling = sEfterbehandling.Replace("\n", ",").Replace(" ", "");
                if (IsDebug) Storefront.LogMessage("sEfterbehandling: " + sEfterbehandling, "", docId, 3, false);
                XmlWriterWrite(xmlWriter, "Efterbehandling", sEfterbehandling);

                /*  string sTillbehor = Storefront.GetValue("PrintingField", "Tillbehor", docId);
                  if (!String.IsNullOrEmpty(sTillbehor)) sTillbehor = sTillbehor.Replace("\n", ", ");

                  if (IsDebug) Storefront.LogMessage("Tillbehor: " + sTillbehor, "", docId, 3, false);
                  XmlWriterWrite(xmlWriter, "Tillbehor" , sTillbehor );*/

                // Loop through all fields and look for custom fields

                FieldValue[] fvvFieldValues = Storefront.GetAllValues("PrintingField", docId);
                foreach (FieldValue fvFieldValue in fvvFieldValues)
                {
                    // Ignore empty fields
                    if (String.IsNullOrEmpty(fvFieldValue.fieldValue)) continue;

                    // Is it a 'EfterbehandlingVal' field?
                    if (fvFieldValue.fieldName.StartsWith(EfterbehandlingValPrefix))
                    {
                        XmlWriterWrite(xmlWriter, "EfterbehandlingVal", fvFieldValue.fieldValue);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(EfterbehandlingParameterPrefix)) // Is it a efterbehandling paramater?
                    {
                        // Get the ID
                        string sTempString = fvFieldValue.fieldName.Substring(EfterbehandlingParameterPrefix.Length);
                        string ID = sTempString.Substring(0, sTempString.IndexOf("]"));
                        XmlWriterWrite(xmlWriter, "EfterbehandlingParameter", fvFieldValue.fieldValue, "ID", ID);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(EfterbehandlingKommentarPrefix)) // Is it a efterbehandling komentar?
                    {
                        // Get the ID
                        string sTempString = fvFieldValue.fieldName.Substring(EfterbehandlingKommentarPrefix.Length);
                        string ID = sTempString.Substring(0, sTempString.IndexOf("]"));
                        XmlWriterWrite(xmlWriter, "EfterbehandlingKommentar", fvFieldValue.fieldValue, "ID", ID);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(TillbehorValPrefix))
                    {
                        XmlWriterWrite(xmlWriter, "TillbehorVal", fvFieldValue.fieldValue);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(TillbehorParameterPrefix)) // Is it a tilbehor paramater?
                    {
                        // Get the ID
                        string sTempString = fvFieldValue.fieldName.Substring(TillbehorParameterPrefix.Length);
                        string ID = sTempString.Substring(0, sTempString.IndexOf("]"));
                        XmlWriterWrite(xmlWriter, "TillbehorParameter", fvFieldValue.fieldValue, "ID", ID);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(TillbehorKommentarPrefix)) // Is it a tilbehor komentar?
                    {
                        // Get the ID
                        string sTempString = fvFieldValue.fieldName.Substring(TillbehorKommentarPrefix.Length);
                        string ID = sTempString.Substring(0, sTempString.IndexOf("]"));
                        XmlWriterWrite(xmlWriter, "TillbehorKommentar", fvFieldValue.fieldValue, "ID", ID);
                    }
                    else if (fvFieldValue.fieldName.StartsWith(ExtraFieldPrefix)) // Is it a 'custom' field?
                    {

                        XmlWriterWrite(xmlWriter, fvFieldValue.fieldName.Substring(ExtraFieldPrefix.Length), fvFieldValue.fieldValue);
                    }
                }

                if (IsDebug) Storefront.LogMessage("--AFTER CUSTOM FIELDS", "", docId, 3, false);

                XmlWriterWrite(xmlWriter, "Tryck", Storefront.GetValue("ProductField", "Tryck", sProductID));
                XmlWriterWrite(xmlWriter, "Price", sThePrice);
                // XmlWriterWrite(xmlWriter, "Plock", Storefront.GetValue("ProductField", "Plock", sProductID));

                // Write delivery options, if available
                string sShowDelivery = Storefront.GetValue("PrintingField", "show_delivery", docId);
                sShowDelivery = (sShowDelivery == null) ? "" : sShowDelivery.ToLower(); // Make sure we actually have a value
                if (sShowDelivery == "yes" || sShowDelivery == "ja") // Do we have custom delivery fields for this document?
                {
                    XmlWriterWrite(xmlWriter, "PODeliveryFirstName", Storefront.GetValue("PrintingField", "PODeliveryFirstName", docId));
                    XmlWriterWrite(xmlWriter, "PODeliveryLastName", Storefront.GetValue("PrintingField", "PODeliveryLastName", docId));
                    string sPODeliveryFullAddress = Storefront.GetValue("PrintingField", "PODeliveryFullAddress", docId);
                    if (sPODeliveryFullAddress != null && sPODeliveryFullAddress != "")
                    {
                        XmlWriterWrite(xmlWriter, "PODeliveryAddress1", Storefront.GetValue("PrintingField", "PODeliveryAddress1", docId));
                        XmlWriterWrite(xmlWriter, "PODeliveryPostalCode", Storefront.GetValue("PrintingField", "PODeliveryPostalCode", docId));
                        XmlWriterWrite(xmlWriter, "PODeliveryCity", Storefront.GetValue("PrintingField", "PODeliveryCity", docId));
                        XmlWriterWrite(xmlWriter, "PODeliveryCountry", Storefront.GetValue("PrintingField", "PODeliveryCountry", docId));

                    }
                    else
                    {
                        XmlWriterWrite(xmlWriter, "PODeliveryFullAddress", sPODeliveryFullAddress);
                    }
                }

                //  XmlWriterWrite(xmlWriter, "File_confirm_adress", Storefront.GetValue("VariableValue", "confirm_adress", docId));
                //  XmlWriterWrite(xmlWriter, "File_AntalSidor", Storefront.GetValue("VariableValue", "AntalSidor", docId));
                //   XmlWriterWrite(xmlWriter, "File_Address", Storefront.GetValue("PrintingField", "Address", docId));

                string Stop_Date = Storefront.GetValue("UserField", "UserProfileCustomerStop_Date", sUserID);
                if (String.IsNullOrEmpty(Stop_Date))
                    Stop_Date = Storefront.GetValue("ProductField", "Stop_Date", sProductID);

                XmlWriterWrite(xmlWriter, "Stop_Date", Stop_Date);


                xmlWriter.WriteEndElement(); // End 'document'

                xmlWriter.WriteEndElement(); // End 'order'

                xmlWriter.WriteEndDocument();


                if (IsDebug) Storefront.LogMessage("--DONE", "", docId, 3, false);
            }
            catch (Exception e)
            {
                Storefront.LogMessage("Unable to generate order XML. Error: " + e.Message.ToString()+", "+e.StackTrace.ToString(), orderId, "", 3, true);
            }
            xmlWriter.Flush();
            xmlWriter.Close();

            try
            {
                FileInfo fi = new FileInfo(outputFilePath);

                File.Copy(outputFilePath, sCopyDirectory + "\\" + fi.Name);
            }
            catch (Exception ex) {
                Storefront.LogMessage("Exception copying file:" + ex.Message, null, null, 1, false);
            }
            

            return eSuccess;
        }

        public override int DocClone_After(string docId, string oldDocId)
        {
            string sHasExported = Storefront.GetValue("PrintingField", DocumentFieldHasExportedXML, docId);

            // Make sure the copy is not marked as exported
            if (sHasExported == "YES")
            {
                Storefront.SetValue("PrintingField", DocumentFieldHasExportedXML, docId, "");
            }

            return eSuccess;
        }

        /// <summary>
        /// Helper function that makes sure that we don't write any empty strings to the XML-file.
        /// </summary>
        /// <param name="xmlWriter"></param>
        /// <param name="tag"></param>
        /// <param name="value"></param>
        private void XmlWriterWrite(XmlTextWriter xmlWriter, string tag, string value)
        {
            //if (value != null && value.Replace(" ", "") != "") 
            if (value == null)
                value = "";
            xmlWriter.WriteElementString(tag, value);
        }

        /// <summary>
        /// Helper function that makes sure that we don't write any empty strings to the XML-file.
        /// </summary>
        /// <param name="xmlWriter"></param>
        /// <param name="tag"></param>
        /// <param name="value"></param>
        private void XmlWriterWrite(XmlTextWriter xmlWriter, string tag, string value, string param, string paramValue)
        {
            if ((value != null && value.Replace(" ", "") != "") || (!String.IsNullOrEmpty(paramValue)))
            {
                xmlWriter.WriteStartElement(tag);
                xmlWriter.WriteAttributeString(param, paramValue);
                xmlWriter.WriteString(value);
                xmlWriter.WriteEndElement();
            }

        }

        /// <summary>
        /// Helper function which checks whether we have permission to write to the folder
        /// </summary>
        /// <param name="folderPath">The folder to check.</param>
        /// <returns>True if we have permission, false otherwize.</returns>
        private bool hasWriteAccessToFolder(string folderPath)
        {
            string sFileName = System.Guid.NewGuid().ToString() + ".txt";
            string sFullFileName = folderPath + "\\" + sFileName;

            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(sFullFileName);
                file.Close();

                // Apparently successfull, delete the file
                File.Delete(sFullFileName);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Helper function which tests if the customer ID can be used to generate file names.
        /// </summary>
        /// <param name="folderPath">Path for output folder.</param>
        /// <param name="customerID">Customer ID to check</param>
        /// <returns></returns>
        private bool canUseCustomerID(string folderPath, string customerID)
        {
            string sFullFileName = folderPath + "\\" + System.Guid.NewGuid().ToString() + customerID + ".txt";

            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(sFullFileName);
                file.Close();

                // Apparently successfull, delete the file
                File.Delete(sFullFileName);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public override int GetConfigurationHtml(KeyValuePair[] parameters, out string ret)
        {
            ret = null;
            ConfigurationHtmlBuilder config = new ConfigurationHtmlBuilder();

            if (parameters == null)
            {
                config.AddHeader();
                config.AddSectionHeader(DisplayName + " - " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
                string sOutputDirectory = Storefront.GetValue("ModuleField", ModuleFieldOutputDirectory, UniqueName);
                config.AddTextField("Absolute pat to output directory, to override add Metadata field <i>" + ModuleFieldOutputDirectory + "</i>", ModuleFieldOutputDirectory, sOutputDirectory, true, true, "Must be non-empty");

                string sCopyDirectory = Storefront.GetValue("ModuleField", ModuleFieldCopyDirectory, UniqueName);
                config.AddTextField("Absolute pat to output copy directory, to override add Metadata field <i>" + ModuleFieldCopyDirectory + "</i>", ModuleFieldCopyDirectory, sCopyDirectory, true, true, "Must be non-empty");

                string sCustomerID = Storefront.GetValue("ModuleField", ModuleFieldCustomerID, UniqueName);
                config.AddTextField("Customer ID", ModuleFieldCustomerID, sCustomerID, true, true, "Must be non-empty");

                config.AddMessage("To disable creating XML for some products add MetaData field <i>"+MetaFieldDisableXML+"</i> value <i>YES</i>");

                config.AddFooter();
                ret = config.html;
            }
            else
            {
                bool bHasErrors = false;
                bool bDirectoryExists = true;
                bool bOutputDirectoryMissing = false;
                bool bHasWriteAccess = true;
                bool bHasCustomerID = true;
                bool bLegalCustomerID = true;

                // First check folder ok
                foreach (KeyValuePair kv in parameters)
                {
                    if (kv.Name == ModuleFieldOutputDirectory)
                    {
                        if (kv.Value != "")
                        {
                            // Make sure directory exists:
                            if (!Directory.Exists(kv.Value))
                            {
                                bDirectoryExists = false;
                                bHasErrors = true;
                            }
                            else if (!hasWriteAccessToFolder(kv.Value))
                            {
                                bHasWriteAccess = false;
                                bHasErrors = true;
                            }
                            else
                            {
                                Storefront.SetValue("ModuleField", kv.Name, UniqueName, kv.Value);
                            }
                        }
                        else
                        {
                            bOutputDirectoryMissing = true;
                            bHasErrors = true;
                        }
                    }

                    if (kv.Name == ModuleFieldCopyDirectory)
                    {
                        if (kv.Value != "")
                        {
                            // Make sure directory exists:
                            if (!Directory.Exists(kv.Value))
                            {
                                bDirectoryExists = false;
                                bHasErrors = true;
                            }
                            else if (!hasWriteAccessToFolder(kv.Value))
                            {
                                bHasWriteAccess = false;
                                bHasErrors = true;
                            }
                            else
                            {
                                Storefront.SetValue("ModuleField", kv.Name, UniqueName, kv.Value);
                            }
                        }
                        else
                        {
                            bOutputDirectoryMissing = true;
                            bHasErrors = true;
                        }
                    }
                }
                // Now check customer ID
                foreach (KeyValuePair kv in parameters)
                {
                    if (kv.Name == ModuleFieldCustomerID)
                    {
                        if (!canUseCustomerID(Storefront.GetValue("ModuleField", ModuleFieldOutputDirectory, UniqueName), kv.Value))
                        {
                            bLegalCustomerID = false;
                            bHasErrors = true;
                        }
                        else if (kv.Value == "")
                        {
                            bHasCustomerID = false;
                            bHasErrors = true;
                        }
                        if (kv.Value != "")
                        {
                            Storefront.SetValue("ModuleField", kv.Name, UniqueName, kv.Value);
                        }
                    }
                }

                if (bHasErrors)
                {
                    config.AddHeader();
                    config.AddSectionHeader(DisplayName + " - " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
                    string sOutputDirectory = Storefront.GetValue("ModuleField", ModuleFieldOutputDirectory, UniqueName);
                    config.AddTextField("Supported file types", ModuleFieldOutputDirectory, sOutputDirectory, true, true, "Must be non-empty");
                    string sCustomerID = Storefront.GetValue("ModuleField", ModuleFieldCustomerID, UniqueName);
                    config.AddTextField("Customer ID", ModuleFieldCustomerID, sCustomerID, true, true, "Must be non-empty");
                    config.AddFooter();
                    ret = config.html;
                    if (bOutputDirectoryMissing)
                    {
                        ret += "<p style='color:red'>Must provide output directory.</p>";
                    }
                    else if (!bDirectoryExists)
                    {
                        ret += "<p style='color:red'>Must provide and existing output directory.</p>";
                    }
                    else if (!bHasWriteAccess)
                    {
                        ret += "<p style='color:red'>Write access denied for this directory.</p>";
                    }
                    if (!bHasCustomerID)
                    {
                        ret += "<p style='color:red'>Must provide a customer ID.</p>";
                    }
                    if (!bLegalCustomerID)
                    {
                        ret += "<p style='color:red'>Not a valid customer ID (must be usable as filename).</p>";
                    }
                }
            }

            return eSuccess;
        }
    }
}
