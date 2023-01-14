using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Intuit.Ipp;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using System.Net;
using System.IO;
using Newtonsoft.Json;

namespace Pen.Utilities.QuickBook
{
    public class clsQuickBook
    {
        public string _ClientID { get; set; }
        public string _ClientSecret { get; set; }
        public string _RedirectURL { get; set; }
        public string _Environment { get; set; }
        public string _BaseURL { get; set; }
        public string _RefreshToken { get; set; }

        public string _OuthBearerToken { get; set; }

        public string refresh_token { get; set; }

        public OAuth2Client auth2Client;
        public clsQuickBook()
        {
        }
        public clsQuickBook(string ClientID, string ClientSecret, string RedirectURL, string Environment)
        {
            _ClientID = ClientID;
            _ClientSecret = ClientSecret;
            _RedirectURL = RedirectURL;
            _Environment = Environment;
            auth2Client = new OAuth2Client(_ClientID, _ClientSecret, _RedirectURL, _Environment);
        }
        public clsQuickBook(string ClientID, string ClientSecret, string refresh_Token, string Environment, string OuthBearerTokenURL)
        {
            _ClientID = ClientID;
            _ClientSecret = ClientSecret;
            _RefreshToken = refresh_Token;
            _OuthBearerToken = OuthBearerTokenURL;
        }
        public clsQuickBook(string baseURL)
        {
            _BaseURL = baseURL;
        }
        public string GetAuthorizationURL()
        {
            List<OidcScopes> scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            return auth2Client.GetAuthorizationURL(scopes);
           
        }
        public string PerformRefreshToken()
        {
            string access_token = "";
            string cred = string.Format("{0}:{1}", _ClientID, _ClientSecret);
            string enc = Convert.ToBase64String(Encoding.ASCII.GetBytes(cred));
            string basicAuth = string.Format("{0} {1}", "Basic", enc);

            // build the  request
            string refreshtokenRequestBody = string.Format("grant_type=refresh_token&refresh_token={0}",
                _RefreshToken
                );

            // send the Refresh Token request
            HttpWebRequest refreshtokenRequest = (HttpWebRequest)WebRequest.Create(_OuthBearerToken);
            refreshtokenRequest.Method = "POST";
            refreshtokenRequest.ContentType = "application/x-www-form-urlencoded";
            refreshtokenRequest.Accept = "application/json";
            //Adding Authorization header
            refreshtokenRequest.Headers[HttpRequestHeader.Authorization] = basicAuth;

            byte[] _byteVersion = Encoding.ASCII.GetBytes(refreshtokenRequestBody);
            refreshtokenRequest.ContentLength = _byteVersion.Length;
            Stream stream = refreshtokenRequest.GetRequestStream();
            stream.Write(_byteVersion, 0, _byteVersion.Length);
            stream.Close();
            try
            {
                //get response
                HttpWebResponse refreshtokenResponse = (HttpWebResponse)refreshtokenRequest.GetResponse();
                using (var refreshTokenReader = new StreamReader(refreshtokenResponse.GetResponseStream()))
                {
                    //read response
                    string responseText = refreshTokenReader.ReadToEnd();

                    // decode response
                    Dictionary<string, string> refreshtokenEndpointDecoded = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseText);

                    if (refreshtokenEndpointDecoded.ContainsKey("error"))
                    {
                        // Check for errors.
                        if (refreshtokenEndpointDecoded["error"] != null)
                        {
                            return "";
                        }
                    }
                    else
                    {
                        //if no error
                        if (refreshtokenEndpointDecoded.ContainsKey("refresh_token"))
                        {
                            refresh_token = refreshtokenEndpointDecoded["refresh_token"];
                            if (refreshtokenEndpointDecoded.ContainsKey("access_token"))
                            {
                                //save both refresh token and new access token in permanent store
                                access_token = refreshtokenEndpointDecoded["access_token"];
                            }
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        var exceptionDetail = response.GetResponseHeader("WWW-Authenticate");
                        if (exceptionDetail != null && exceptionDetail != "")
                        {
                            access_token = "";
                        }
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // read response body
                            string responseText = reader.ReadToEnd();
                            if (responseText != null && responseText != "")
                            {
                                access_token = "";
                            }
                        }
                    }
                }
            }
            return access_token;
        }

        public Boolean IsCustomerExists(string fName, string lName, string emailID, string baseURL)
        {
            try
            {
                if (String.IsNullOrEmpty(baseURL))
                {
                    throw new Exception("Please provide base URL");
                }
                if (System.Web.HttpContext.Current.Session["access_token"] != null && System.Web.HttpContext.Current.Session["refresh_token"] != null &&
                    System.Web.HttpContext.Current.Session["realmId"] != null)
                {
                    OAuth2RequestValidator reqValidator = new OAuth2RequestValidator(System.Web.HttpContext.Current.Session["access_token"].ToString());
                    ServiceContext context = new ServiceContext(System.Web.HttpContext.Current.Session["realmId"].ToString(), IntuitServicesType.QBO, reqValidator);
                    context.IppConfiguration.MinorVersion.Qbo = "23";
                    context.IppConfiguration.BaseUrl.Qbo = baseURL;// "https://sandbox-quickbooks.api.intuit.com/";
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    QueryService<Customer> customerQueryService = new QueryService<Customer>(context);
                    List<Customer> cust = customerQueryService.ExecuteIdsQuery("SELECT * FROM Customer WHERE GivenName = '" + fName + "' and FamilyName='" + lName + "' STARTPOSITION  20").ToList();
                    if (cust.Count > 0)
                    {
                        if (emailID != "")
                        {
                            foreach (var C in cust)
                            {
                                if (C.PrimaryEmailAddr != null)
                                {
                                    string emailCustID = C.PrimaryEmailAddr.Address;
                                    if (emailCustID == emailID)
                                        return true;
                                }
                                return false;
                            }
                        }
                        return true;
                    }
                    return false;
                }
                else
                {
                    throw new Exception("Quickbook session expired, please authenticate again");
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public string AddCustomer(Customer cust, string baseURL)
        {
            try
            {
                if (String.IsNullOrEmpty(baseURL))
                {
                    throw new Exception("Please provide base URL");
                }
                if (System.Web.HttpContext.Current.Session["access_token"] != null && System.Web.HttpContext.Current.Session["refresh_token"] != null &&
                    System.Web.HttpContext.Current.Session["realmId"] != null)
                {
                    OAuth2RequestValidator reqValidator = new OAuth2RequestValidator(System.Web.HttpContext.Current.Session["access_token"].ToString());
                    ServiceContext context = new ServiceContext(System.Web.HttpContext.Current.Session["realmId"].ToString(), IntuitServicesType.QBO, reqValidator);
                    context.IppConfiguration.MinorVersion.Qbo = "23";
                    context.IppConfiguration.BaseUrl.Qbo = baseURL;
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    QueryService<Customer> customerQueryService = new QueryService<Customer>(context);
                    DataService dS = new DataService(context);
                    Customer CustomerAdd = dS.Add(cust);
                    return CustomerAdd.Id;
                }
                else
                {
                    throw new Exception("Quickbook session expired, please authenticate again");
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Boolean IsCustomerExists(string fName, string lName, string emailID, string baseURL, string access_token,string realmId)
        {
            try
            {
                if (String.IsNullOrEmpty(baseURL))
                {
                    throw new Exception("Please provide base URL");
                }
                if (access_token != null && realmId != null)
                {
                    OAuth2RequestValidator reqValidator = new OAuth2RequestValidator(access_token);
                    ServiceContext context = new ServiceContext(realmId, IntuitServicesType.QBO, reqValidator);
                    context.IppConfiguration.MinorVersion.Qbo = "23";
                    context.IppConfiguration.BaseUrl.Qbo = baseURL;
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    QueryService<Customer> customerQueryService = new QueryService<Customer>(context);
                    List<Customer> cust = customerQueryService.ExecuteIdsQuery("SELECT * FROM Customer WHERE GivenName = '" + fName + "' and FamilyName='" + lName + "' STARTPOSITION  20").ToList();
                    if (cust.Count > 0)
                    {
                        if (emailID != "")
                        {
                            foreach (var C in cust)
                            {
                                if (C.PrimaryEmailAddr != null)
                                {
                                    string emailCustID = C.PrimaryEmailAddr.Address;
                                    if (emailCustID == emailID)
                                        return true;
                                }
                                return false;
                            }
                        }
                        return true;
                    }
                    return false;
                }
                else
                {
                    throw new Exception("Quickbook session expired, please authenticate again");
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public string AddCustomer(Customer cust, string baseURL, string access_token, string realmId)
        {
            try
            {
                if (String.IsNullOrEmpty(baseURL))
                {
                    throw new Exception("Please provide base URL");
                }
                if (access_token != null && realmId != null)
                {
                    OAuth2RequestValidator reqValidator = new OAuth2RequestValidator(access_token);
                    ServiceContext context = new ServiceContext(realmId, IntuitServicesType.QBO, reqValidator);
                    context.IppConfiguration.MinorVersion.Qbo = "23";
                    context.IppConfiguration.BaseUrl.Qbo = baseURL;
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    QueryService<Customer> customerQueryService = new QueryService<Customer>(context);
                    DataService dS = new DataService(context);
                    Customer CustomerAdd = dS.Add(cust);
                    return CustomerAdd.Id;
                }
                else
                {
                    throw new Exception("Quickbook session expired, please authenticate again");
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
