using DocuSign.eSign.Api;
using DocuSign.eSign.Client;
using DocuSign.eSign.Model;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Pen.Utilities.Docusign
{
    public class DocuSignTemplate
    {
        public string templateId { get; set; }
        public string name { get; set; }
    }
    public class clsDocuSign
    {
        private string _IntegrationKey { get; set; }
        private string _secretKey { get; set; }
        private string _accountID { get; set; }
        private string _TokenURL { get; set; }

        public bool _isSSLEnabled = true;
        public clsDocuSign(string IntegrationKey, string secretKey, string tokenURL, string accountID, bool isSSLEnabled = true)
        {
            _IntegrationKey = IntegrationKey;
            _secretKey = secretKey;
            _TokenURL = tokenURL;
            _accountID = accountID;
        }
        public string GetUser(string code, string userProfileURL)
        {
            try
            {
                if (String.IsNullOrEmpty(userProfileURL))
                {
                    throw new Exception("Provide User Profile URL");
                }
                string token = getToken(_IntegrationKey + ":" + _secretKey, code);
                var client = new RestClient(userProfileURL);
                if (_isSSLEnabled)
                {
                    System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                }
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Authorization", "Bearer " + token);
                IRestResponse response = client.Execute(request);
                client = null;
                request = null;
                return response.Content;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetToken(string code)
        {
            return getToken(_IntegrationKey + ":" + _secretKey, code);
        }

        public string GetBase64String(string inputString)
        {
            string base64Decoded = inputString;
            string base64Encoded;
            byte[] data = System.Text.ASCIIEncoding.ASCII.GetBytes(base64Decoded);
            base64Encoded = System.Convert.ToBase64String(data);
            return base64Encoded;
        }

        public string getToken(string IntSecretKey, string code)
        {
            try
            {
                if (String.IsNullOrEmpty(_TokenURL))
                {
                    throw new Exception("Provide Token URL");
                }
                var client = new RestClient(_TokenURL);
                if (_isSSLEnabled)
                {
                    System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                }
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Basic " + GetBase64String(IntSecretKey));
                request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                request.AddParameter("grant_type", "authorization_code");
                request.AddParameter("code", code);
                IRestResponse response = client.Execute(request);
                if (response.Content != "")
                {
                    var jObject = JObject.Parse(response.Content);
                    client = null;
                    request = null;
                    return jObject.GetValue("access_token").ToString();
                }
                else
                {
                    client = null;
                    request = null;
                    if (response.ErrorException != null)
                    {
                        throw new Exception(response.ErrorException.Message.ToString());
                    }
                    return "";
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        public List<DocuSignTemplate> getTemplate(string TemplateURL, string Token)
        {
            try
            {
                if (String.IsNullOrEmpty(Token))
                {
                    throw new Exception("Provide Token ");
                }
                if (String.IsNullOrEmpty(TemplateURL))
                {
                    throw new Exception("Provide Template URL ");
                }
                if (String.IsNullOrEmpty(_accountID))
                {
                    throw new Exception("Provide Account ID");
                }
                List<DocuSignTemplate> template = new List<DocuSignTemplate>();
                var client = new RestClient(String.Format(TemplateURL, _accountID));
                if (_isSSLEnabled)
                {
                    System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                }
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Bearer " + Token);
                request.AddHeader("Content-Type", "application/json");
                IRestResponse response = client.Execute(request);
                if (response.Content != "")
                {
                    var jObject = JObject.Parse(response.Content);
                    client = null;
                    request = null;
                    for (int i = 0; i < jObject.GetValue("envelopeTemplates").Count(); i++)
                    {
                        template.Add(new DocuSignTemplate()
                        {
                            name = jObject.GetValue("envelopeTemplates").ElementAt(i).SelectToken("name").ToString(),
                            templateId = jObject.GetValue("envelopeTemplates").ElementAt(i).SelectToken("templateId").ToString()
                        });
                    }
                    return template;
                }
                else
                {
                    client = null;
                    request = null;
                    if (response.ErrorException != null)
                    {
                        throw new Exception(response.ErrorException.Message.ToString());
                    }
                    return template;
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public string Create(string addressMail, string signerEmail, string signerName, string ccEmail, string ccName, string AccessToken, string templateID, string baseURL)
        {
            var accessToken = AccessToken;
            var basePath = baseURL + "/restapi";
            var accountId =_accountID;
            var templateId = templateID;
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new Exception("Provide Token ");
            }
            string envelopeId = DoWork(addressMail,signerEmail, signerName, ccEmail,
               ccName, accessToken, basePath, accountId, templateId);
            return envelopeId;
        }

        public string DoWork(string addressMail, string signerEmail, string signerName, string ccEmail,
            string ccName, string accessToken, string basePath,
            string accountId, string templateId)
        {
            var apiclient = new ApiClient(basePath);
            apiclient.Configuration.AccessToken = accessToken;
            EnvelopesApi envelopesApi = new EnvelopesApi(apiclient);
            EnvelopeDefinition envelope = MakeEnvelope(addressMail,signerEmail, signerName, ccEmail, ccName, templateId);
            EnvelopeSummary result = envelopesApi.CreateEnvelope(accountId, envelope);

            return result.EnvelopeId;
        }

        private EnvelopeDefinition MakeEnvelope(string addressMail, string signerEmail, string signerName,
           string ccEmail, string ccName, string templateId)
        {
            EnvelopeDefinition env = new EnvelopeDefinition();
            env.TemplateId = templateId;


            FullName fullName = new FullName
            {
                TabLabel = "Full Name",
                Value = signerName
            };

            Email email = new Email
            {
                TabLabel = "Email",
                Value = signerEmail
            };

            DateSigned dateSigned = new DateSigned
            {
                TabLabel = "Date Signed",
                Value =DateTime.Now.ToShortDateString()
            };

            Company Cmp = new Company
            {
                TabLabel = "Company",
                Value = ""
            };
            Text txt = new Text
            {
                 TabLabel= "Property Address",
                 Value=""
            };
            Text txt1 = new Text
            {
                TabLabel = "Tax Parcel",
                Value = ""
            };

            Text txt2 = new Text
            {
                TabLabel = "County",
                Value = ""
            };
            Text txt3 = new Text
            {
                TabLabel = "Mailing Address",
                Value = addressMail
            };

            TemplateRole signer1 = new TemplateRole();
            signer1.Email = signerEmail;
            signer1.Name = signerName;
            signer1.RoleName = "Client";
            signer1.Tabs = new Tabs()
            {
                FullNameTabs = new List<FullName> { fullName },
                EmailTabs = new List<Email> { email },
                DateSignedTabs = new List<DateSigned> { dateSigned },
                CompanyTabs = new List<Company> { Cmp },
                TextTabs=new List<Text> { txt, txt1, txt2, txt3 }
            };

            env.TemplateRoles = new List<TemplateRole> { signer1 };
            env.Status = "sent";
            return env;
        }

    }    
}

