using System;
using System.Web;
using Pageflex.Interfaces.Storefront;
using System.Net;
using System.Collections.Specialized;
using System.Text;
using ImagebankInIframePages;
using System.Globalization;

namespace LogonGrandID
{
    public class LogonGrandID : StorefrontExtension
    {
        public const string WP_PASSWORD = "userg1obalpassw0rd";
        public const string CUMULUS_PASSWORD = "botkyrakg1obalpassw0rd";
        public const string GLOBAL_TOKEN = "yriBG2gMWjMYBBtG6H99";

        public override string UniqueName {
            get
            {
                return "com.brandgate.LogonGrandID";
            }
        }

        public override string DisplayName {
            get
            {
                return "LogonGrandID";
            }
        }

        public static string FindUser(string name)
        {
            return PFWeb.StorefrontAPI.Storefront.FindUserID(name);
        }

        public static void CreateNewUser(string firstname)
        {
            string uid = null;
            AddLogMessage("Creating Storefront User: " + firstname);
            int result = PFWeb.StorefrontAPI.Storefront.CreateUser(firstname, out uid);

            

            if (result == eSuccess)
            {
                PFWeb.StorefrontAPI.Storefront.SetValue("UserField", "UserProfileFirstName", uid, firstname);
                PFWeb.StorefrontAPI.Storefront.SetValue("UserField", "UserProfileLastName", uid, " ");
                PFWeb.StorefrontAPI.Storefront.SetValue("UserProperty", "IsActive", uid, "1");
            }
        }

        public static void CreateNewWPUser(string firstname)
        {
            try
            {
                WebClient wpClient = new WebClient();
                string url = PFWeb.StorefrontAPI.Storefront.GetValue("ModuleField", "WP_URL", "com.brandgate.newsletteriniframe");
                AddLogMessage("Creating WP User: " + firstname);
                string output = wpClient.DownloadString(string.Format("{0}/wp-admin/admin-ajax.php?action=my_action&username={1}&password={2}&token={3}", url, firstname, WP_PASSWORD, GLOBAL_TOKEN));
            }
            catch (Exception ex)
            {
                AddLogMessage("Error while creating WP User: " + ex);
                PFWeb.StorefrontAPI.Storefront.SendEmail("ivan@brandgate.no", "mindaugas@brandgate.no", "", "", "Attension required",
                    "User " + firstname + " failed to be created in wordpress", null, null, null);
            }

        }

        public static int CreateCumulusUser(string userId,string prefix)
        {
            PFWeb.StorefrontAPI.FieldValue[] userFieldValues = PFWeb.StorefrontAPI.Storefront.GetAllValues("UserField", userId);

            string name = null;

            foreach (PFWeb.StorefrontAPI.FieldValue fv in userFieldValues)
            {
                if (fv.fieldName.Equals("UserProfileFirstName"))
                    name = fv.fieldValue;
            }
            name = RemoveDiacritics(name);
            NameValueCollection values = new NameValueCollection();

            values.Add("{53c7a211-dac4-11d6-b6be-0050baeba6c7}", prefix + name);//username
            values.Add("{24b7790b-6646-4cfe-9feb-5879fd1b8da4}", CUMULUS_PASSWORD);//password
            values.Add("{7c437141-daa4-11d6-b6be-0050baeba6c7}", name);//first name
            values.Add("{7c437143-daa4-11d6-b6be-0050baeba6c7}", name);//last name
            values.Add("{7c43714f-daa4-11d6-b6be-0050baeba6c7}", name);//email
            values.Add("encoding", "UTF-8");
            string response = "";

            string address = Base.getConfigValue(Base.IMAGEBANK_URL) + "createaccountwindow.jspx";
            try
            {
                using (WebClient wc = new WebClient())
                {
                    //wc.Headers.Add();
                    wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    wc.Encoding = Encoding.UTF8;                    
                    //response = wc.UploadString(address, "POST", data);
                    response = Encoding.UTF8.GetString(wc.UploadValues(address, "POST", values));
                }
            }
            catch (Exception ex)
            {
                response = "Exception: " + ex.Message;
            }

            //check
            if (!response.Contains("userSuccessfullyCreatedIdentifier"))
            {
                PFWeb.StorefrontAPI.Storefront.LogMessage("User " + name + " creating failed, response:" + response, null, null, 0, false);
                PFWeb.StorefrontAPI.Storefront.SendEmail("ivan@brandgate.no", "mindaugas@brandgate.no", "", "", "Attension required",
                    "User " + name + " failed to be created in cumulus, response:" + response, null, null, null);
            }

            return eSuccess;
        }

        public override int PageLoad(string pageBaseName, string eventName)
        {
            if (pageBaseName.Equals("login_aspx"))
            {
                string logoutEvent = HttpContext.Current.Request.QueryString["logout"];
                string sID = HttpContext.Current.Request.QueryString["grandidsession"];
                string ticket = HttpContext.Current.Request.QueryString["ticket"];
                string successURL = HttpContext.Current.Request.QueryString["SuccessUrl"];

                if (!string.IsNullOrEmpty(logoutEvent))
                {
                    WebClient webClient = new WebClient();
                    string str = webClient.DownloadString(string.Format("{0}/Logout?apiKey={1}&authenticateServiceKey={2}&sessionid={3}", Getvalue("service_url"), Getvalue("apiKey"), Getvalue("authenticateServiceKey"), HttpContext.Current.Session["userID"]));
                    AddLogMessage("Logout event: " + str);
                }
                if (!string.IsNullOrEmpty(ticket))
                {
                    
                }
                else if (String.IsNullOrEmpty(sID))
                {                    
                    AddLogMessage("Session id not found: " + sID);
                    // Step 1. and 2. FederatedLogin request - returns seesionID and redirectUrl
                    HttpContext.Current.Session["url"] = "&successURL=" + successURL;
                    LoginData user = LoginRequest(successURL);

                    // Step 3. and 4. Redirect browser to GrandID 
                    HttpContext.Current.Response.Redirect(user.redirectUrl);
                }
                else
                {
                    AddLogMessage("Session id found: " + sID);
                    HttpContext.Current.Session["userID"] = sID;
                    // Step 5. and 6. GetSession - returns user data
                    LoginData user = GetUserProfile(sID);

                    if (string.IsNullOrEmpty(user.userAttributes.givenname))
                    {
                        CWPLoginToken str = LoginCWP(Getvalue("cwpApiUrl"));
                        CWPUserInfo cwpInfo = UserInfo(str.access_token, user.userAttributes.contactid, Getvalue("cwpApiUrl"));
                        AddLogMessage("CWP info: " + cwpInfo);
                        user.userAttributes.givenname = cwpInfo.results[0].firstname;
                        //user.userAttributes.adress = cwpInfo.results[0].address1_line1;
                    }

                    // Step 7. Check if user exists in Storefront
                    if (PFWeb.StorefrontAPI.Storefront.FindUserID(user.userAttributes.givenname) == null)
                    {
                        // Step 8. Create new user in Storefront
                        CreateNewUser(user.userAttributes.givenname);
                        //CreateNewWPUser(user.userAttributes.givenname);
                        CreateCumulusUser(PFWeb.StorefrontAPI.Storefront.FindUserID(user.userAttributes.givenname),Getvalue("prefix"));
                    }
                    // Step 9. Create a ticket for the user
                    ticket = PFWeb.StorefrontAPI.Storefront.GetTicketForUserLogin(user.userAttributes.givenname);

                    // Step 10. Redirect with the ticket
                    HttpContext.Current.Response.Redirect("Login.aspx?ticket=" + ticket + "&logonName=" + user.userAttributes.givenname + HttpContext.Current.Session["url"]);
                }

                
                return eSuccess;
            }
            else return eDoNotCall;
        }

        public static CWPLoginToken LoginCWP(string CWPapiUrl)
        {
            WebClient webClient = new WebClient();
            var reqparm = new NameValueCollection();
            reqparm.Add("grant_type", "password");
            reqparm.Add("username", "brand");
            reqparm.Add("password", "sommar");
            byte[] responsebytes = webClient.UploadValues(string.Format("{0}/token", CWPapiUrl), "POST", reqparm);

            string responseStr = Encoding.Default.GetString(responsebytes);

            return new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<CWPLoginToken>(responseStr);
        }

        public static CWPUserInfo UserInfo(string token,string contact_id,string CWPapiUrl)
        {
            WebClient webClient = new WebClient();
            webClient.Headers.Add("Content-Type", "application/json");
            webClient.Headers.Add("Authorization", "Bearer "+ token);
            string str = webClient.DownloadString(string.Format("{0}/API/Feed/VIS_Visionse_GetUserInfoById?contactid={1}", CWPapiUrl, contact_id));
            return new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<CWPUserInfo>(str);
        }

        public static void AddLogMessage(string message)
        {
            PFWeb.StorefrontAPI.Storefront.LogMessage(message, null, null, 1, false);
        }

        public T ApiRequest<T>(string url)
        {
            try
            {
                WebClient wc = new WebClient();
                AddLogMessage("ApiRequest URL given: " + url);
                string str = wc.DownloadString(url);
                AddLogMessage("String downloaded: " + str);

                return new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<T>(str);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error:", ex);
            }
        }

        public LoginData LoginRequest(string successURL)
        {
            return ApiRequest<LoginData>(string.Format("{0}/FederatedLogin?apiKey={1}&authenticateServiceKey={2}&callbackUrl={3}", Getvalue("service_url"), Getvalue("apiKey"), Getvalue("authenticateServiceKey"), Getvalue("callbackUrl")+ successURL));
        }

        public LoginData GetUserProfile(string grandidsession)
        {
            return ApiRequest<LoginData>(string.Format("{0}/GetSession?apiKey={1}&authenticateServiceKey={2}&sessionid={3}", Getvalue("service_url"), Getvalue("apiKey"), Getvalue("authenticateServiceKey"), grandidsession));
        }

        public override int GetConfigurationHtml(KeyValuePair[] parameters,out string HTML_configString)
        {
            // to configure the DLL a single parameter called 'setup'
            // needs to be specified
            string config = "<br><br><b>MyExtension Configuration</b><br>" +
            "Service URL: <input type='text' name='service_url' value='" + Getvalue("service_url") + "'><br>" +
            "API key: <input type='text' name='apiKey' value='" + Getvalue("apiKey") + "'><br>" +
            "Authenticate Service key: <input type='text' name='authenticateServiceKey' value='" + Getvalue("authenticateServiceKey") + "'><br>" +
            "Callback URL: <input type='text' name='callbackUrl' value='" + Getvalue("callbackUrl") + "'><br>" +
            "CWP API URL: <input type='text' name='cwpApiUrl' value='" + Getvalue("cwpApiUrl") + "'><br>" +
            "Prefix: <input type='text' name='prefix' value='" + Getvalue("prefix") + "'>";

            HTML_configString = null;
            if (parameters == null)
            {
                HTML_configString = config;
            }
            else
            {
                // find the parameter we care about
                foreach (KeyValuePair p in parameters)
                {
                    if (p.Name.Equals("service_url") || p.Name.Equals("apiKey") || p.Name.Equals("authenticateServiceKey") || p.Name.Equals("callbackUrl") || p.Name.Equals("prefix") || p.Name.Equals("cwpApiUrl"))
                    {
                        Storefront.SetValue("ModuleField", p.Name, UniqueName, p.Value);
                    }
                }
            }
            return eSuccess;
        }

        public string Getvalue(string name)
        {
            return Storefront.GetValue("ModuleField", name, UniqueName);
        }

        public static String RemoveDiacritics(string s)
        {
            string normalizedString = s.Normalize(NormalizationForm.FormD);
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < normalizedString.Length; i++)
            {
                char c = normalizedString[i];
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }
            return stringBuilder.ToString();
        }
    }
}
