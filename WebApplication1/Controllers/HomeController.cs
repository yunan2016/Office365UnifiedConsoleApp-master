using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.Exchange.WebServices.Data;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            return View();
        }
        public ActionResult ShowInfo()
        {

          
            string token = (string)Session["access_token"];
            if (string.IsNullOrEmpty(token))
            {


                return View();
            }
            try
            {



                string sURL = "https://outlook.office.com/api/v2.0/me/messages/AAMkADQyZjE2NzY3LWEyNjEtNGI3Ny05YmE5LWM3Yjk1N2RiZjg2YQBGAAAAAACjfX4Ad-CaQJG0NYun281fBwC90J0GawtZTaVMfIdaQJ8wAAAAAAEMAAC90J0GawtZTaVMfIdaQJ8wAAExISUWAAA=";

                WebRequest request1 = WebRequest.Create(sURL);
                request1.Method = "DELETE";
                request1.Headers.Add("Authorization", "Bearer " + token);
                HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                if (response1.StatusCode == HttpStatusCode.OK)
                {
                    // some code
                }
               //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
               //service.HttpHeaders.Add("Authorization", "Bearer " + token);
               //service.PreAuthenticate = true;
               //service.SendClientLatencies = true;
               //service.EnableScpLookup = false;
               //service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
               //EmailMessage email = new EmailMessage(service);

               //email.ToRecipients.Add("dennis@O365E3W15.onmicrosoft.com");

               //email.Subject = "HelloWorld";
               //email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API_hahah");

               //email.Save(WellKnownFolderName.Inbox);

                using (var client = new HttpClient())
                {
                   
                    ////Enable signon and read users' profile
                    //using (var request = new HttpRequestMessage(HttpMethod.Put, "https://graph.microsoft.com/v1.0/me/contacts"))
                    //{
                    //    request.Headers.Add("Authorization", "Bearer " + token);
              
                    //    request.Headers.Add("Content-Type", "application/image");

                    //    //response = await client.PutAsync(gizmoUrl, gizmo);
                    //}


               using (var request = new HttpRequestMessage(HttpMethod.Get, "https://outlook.office.com/api/v1.0/me/messages?$filter=Subject eq 'test123'"))
               {
                   request.Headers.Add("Authorization", "Bearer " + token);
                   request.Headers.Add("Accept", "application/json;odata.metadata=minimal");

                   using (var response = client.SendAsync(request).Result)
                   {

                       if (response.StatusCode == HttpStatusCode.OK)
                       {
                           var model = JsonConvert.DeserializeObject<RootObject>(response.Content.ReadAsStringAsync().Result);
                           //foreach (var item in model.value)
                           //{
                           //    var receivedDateTime = item.receivedDateTime;
                           //}

                       }

                   }
               }
                   

                //    using (var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime gt 2016-02-29T05:19:58Z"))
                //    {
                //        request.Headers.Add("Authorization", "Bearer " + token);
                //        request.Headers.Add("Accept", "application/json;odata.metadata=minimal");

                //        using (var response = client.SendAsync(request).Result)
                //        {

                //            if (response.StatusCode == HttpStatusCode.OK)
                //            {
                //                //var model = JsonConvert.DeserializeObject<RootObject>(response.Content.ReadAsStringAsync().Result);
                //                //foreach (var item in model.value)
                //                //{
                //                //    var receivedDateTime = item.receivedDateTime;
                //                //}

                //            }

                //        }
                //    }
                }

                return View();
            }
            catch (AdalException ex)
            {
                return Content(string.Format("ERROR retrieving messages: {0}", ex.Message));
            }
        }

        public void sendEmail()
        {
            //SmtpClient client = new SmtpClient("smtp.office365.com", 587);
            //client.EnableSsl = true;
            //client.Credentials = new System.Net.NetworkCredential("fx@msdnofficedev.onmicrosoft.com", "asdf123(");
            //MailAddress from = new MailAddress("fx@msdnofficedev.onmicrosoft.com", String.Empty, System.Text.Encoding.UTF8);
            //MailAddress to = new MailAddress("fx@msdnofficedev.onmicrosoft.com");
            //MailMessage message = new MailMessage(from, to);
            //message.Body = "The message I want to send.";
            //message.BodyEncoding = System.Text.Encoding.UTF8;
            //message.Subject = "The subject of the email";
            //message.SubjectEncoding = System.Text.Encoding.UTF8;
       
            //client.SendAsync(message, message);

            //MailMessage msg = new MailMessage();
            //msg.To.Add(new MailAddress("v-nany@hotmail.com", "Nan YU"));
            //msg.From = new MailAddress("fx@msdnofficedev.onmicrosoft.com", "Fei Xue");
            //msg.Subject = "This is a Test Mail";
            //msg.Body = "This is a test message using Exchange OnLine";
            //msg.IsBodyHtml = true;

            //SmtpClient client = new SmtpClient();
            //client.UseDefaultCredentials = false;
            //client.Credentials = new System.Net.NetworkCredential("fx@msdnofficedev.onmicrosoft.com", "asdf123(");
            //client.Port = 587; // You can use Port 25 if 587 is blocked
            //client.Host = "smtp.office365.com";
            //client.DeliveryMethod = SmtpDeliveryMethod.Network;
            //client.EnableSsl = true;
            //try
            //{
            //    client.Send(msg);
            
            //}
            //catch (Exception ex)
            //{
               
            //}



            //var _mailServer = new SmtpClient();
            //_mailServer.UseDefaultCredentials = false;
            //_mailServer.Credentials = new NetworkCredential("fx@msdnofficedev.onmicrosoft.com", "asdf123(");
            //_mailServer.Host = "smtp.office365.com";
            //_mailServer.TargetName = "STARTTLS/smtp.office365.com"; // same behaviour if this lien is removed
            //_mailServer.Port = 587;
            //_mailServer.EnableSsl = true;

            //var eml = new MailMessage();
            //eml.Sender = new MailAddress("fx@msdnofficedev.onmicrosoft.com");
            //eml.From = eml.Sender;

            //eml.To.Add(new MailAddress("v-nany@hotmail.com", "v-nany@hotmail.com"));

            //eml.Subject = "Test message";
            //eml.Body = "Test message body";

            //_mailServer.Send(eml);
        }

        public class RootObject
        {

            public List<Value> value { get; set; }
        }

        public class Value
        {
            public string id { get; set; }
            public List<object> BusinessPhones { get; set; }
            public List<object> EmailAddresses { get; set; }
            //public string createdDateTime { get; set; }
            //public string lastModifiedDateTime { get; set; }
            //public string changeKey { get; set; }
            //public List<object> categories { get; set; }
            //public string receivedDateTime { get; set; }
            //public string sentDateTime { get; set; }
            //public bool hasAttachments { get; set; }
            //public string subject { get; set; }
            //public string bodyPreview { get; set; }
            //public string importance { get; set; }
            //public string parentFolderId { get; set; }
            //public List<object> ccRecipients { get; set; }
            //public List<object> bccRecipients { get; set; }
            //public List<object> replyTo { get; set; }
            //public string conversationId { get; set; }
            //public bool? isDeliveryReceiptRequested { get; set; }
            //public bool isReadReceiptRequested { get; set; }
            //public bool isRead { get; set; }
            //public bool isDraft { get; set; }
            //public string webLink { get; set; }
        }

        public class EmailModel
        {
            public string id { get; set; }
            public string receivedDateTime { get; set; }
            public string sentDateTime { get; set; }
            public string subject { get; set; }
            public string from { get; set; }
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult SignIn()
        {


            //// crate client credentila
            //var authContext2 = new AuthenticationContext("https://login.microsoftonline.com/O365E3W15.onmicrosoft.com");
            //var token = authContext2.AcquireToken("https://graph.microsoft.com/", new ClientCredential("699ab0a8-118f-4de2-a0ba-10ae6f69edfb", "1eqFkV+wX3BpazfVcFfxGnaHKZx1g2KtejqhFn6ad4Y=")).AccessToken;

           // sendEmail();
            var authContext = new AuthenticationContext("https://login.microsoftonline.com/common");


            // The url in our app that Azure should redirect to after successful signin 
            string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);


            // Generate the parameterized URL for Azure signin 
            Uri authUri = authContext.GetAuthorizationRequestURL("https://outlook.office.com", ConfigurationManager.AppSettings["ClientID"],
                new Uri(redirectUri), UserIdentifier.AnyUser, null);

            // Redirect the browser to the Azure signin page 
            return Redirect(authUri.ToString());
        }
        public async Task<ActionResult> Authorize()
        {
            // Get the 'code' parameter from the Azure redirect
            string authCode = Request.Params["code"];

            AuthenticationContext authContext = new AuthenticationContext("https://login.microsoftonline.com/common");

            // The same url we specified in the auth code request
            string redirectUri = Url.Action("Authorize", "Home", null, Request.Url.Scheme);

            // Use client ID and secret to establish app identity
            ClientCredential credential = new ClientCredential(ConfigurationManager.AppSettings["ClientID"], ConfigurationManager.AppSettings["ClientSecret"]);

            try
            {
                // Get the token
                var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
                    authCode, new Uri(redirectUri), credential, "https://outlook.office.com");

                // Save the token in the session
                Session["access_token"] = authResult.AccessToken;


                return Redirect(Url.Action("ShowInfo", "Home", null, Request.Url.Scheme));
            }
            catch (AdalException ex)
            {
                return Content(string.Format("ERROR retrieving token: {0}", ex.Message));
            }
        }

    }
}