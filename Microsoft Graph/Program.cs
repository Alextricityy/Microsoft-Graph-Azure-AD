using Microsoft.Graph;
using Newtonsoft.Json.Linq;

//need these 2 packages

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

/*  Who     When        Version Change
 *  ==========================================================================================================================================
 *  Alex    19/02/2019  1.0.0  Completed first version of project using azure active directory information instead of user permission's to access office 365 information
 *
 */

namespace Microsoft_Graph
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var access_token = GetAccessToken();
            Console.WriteLine(access_token);
            usetoken(access_token).GetAwaiter().GetResult();

            Console.Read();
        }

        public static string GetAccessToken()
        {
            using (var wb = new WebClient())
            {
                var url = "https://login.microsoftonline.com/mcsa.co.uk/oauth2/token";
                var data = new NameValueCollection();
                data["username"] = ConfigurationManager.AppSettings["userLogin"].ToString();
                data["password"] = ConfigurationManager.AppSettings["userPassword"].ToString();
                data["grant_type"] = ConfigurationManager.AppSettings["grantType"].ToString();
                data["resource"] = ConfigurationManager.AppSettings["resource"].ToString();
                data["client_secret"] = ConfigurationManager.AppSettings["clientSecret"].ToString();
                data["client_id"] = ConfigurationManager.AppSettings["clientId"].ToString();

                var response = wb.UploadValues(url, "POST", data);
                string responseInString = Encoding.UTF8.GetString(response);

                var testuserjson = JObject.Parse(responseInString);
                return testuserjson["access_token"].ToString();
            }
        }

        public static async Task usetoken(string token)
        {
            //PublicClientApplication clientApp = new PublicClientApplication(ConfigurationManager.AppSettings["clientId"].ToString());
            // use access token to genereate service client
            //call other methods within this to authenticate
            GraphServiceClient graphclient = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                    }));
            var currentUser = await graphclient.Me.Request().GetAsync();
            sendMail(graphclient).GetAwaiter().GetResult();
        }

        private static async Task GetMailAsync(GraphServiceClient client)
        {
            var currentUser = await client.Me.Request().GetAsync();
            Console.WriteLine(currentUser.Messages.Count);
        }

        private static async Task sendMail(GraphServiceClient client)
        {
            IList<Recipient> messageToList = new List<Recipient>();
            User currentUser = client.Me.Request().GetAsync().Result;

            Recipient currentUserRecipient = new Recipient();
            EmailAddress currentUserEmailAdress = new EmailAddress();
            currentUserEmailAdress.Address = currentUser.UserPrincipalName;
            currentUserEmailAdress.Name = currentUser.DisplayName;
            currentUserRecipient.EmailAddress = currentUserEmailAdress;
            messageToList.Add(currentUserRecipient);

            // Send mail to signed in user and the recipient list
            Console.WriteLine();
            Console.WriteLine("Sending mail....");
            Console.WriteLine();
            try
            {
                ItemBody messageBody = new ItemBody();
                messageBody.Content = "Testing Mailbox to others";
                messageBody.ContentType = BodyType.Text;

                Message newMessage = new Message();
                newMessage.Subject = "\nCompleted test run from console app.";
                newMessage.ToRecipients = messageToList;
                newMessage.Body = messageBody;

                client.Me.SendMail(newMessage, true).Request().PostAsync();
                Console.WriteLine("\nMail sent to {0}", currentUser.DisplayName);
            }
            catch (Exception)
            {
                Console.WriteLine("\nUnexpected Error attempting to send an email");
                throw;
            }
        }

        private static async Task moveMail(string MailId, GraphServiceClient client)
        {
            //send specific id to folder autolog
            var currentUser = await client.Me.MailFolders.Request().GetAsync();
            IList<Recipient> messageToList = new List<Recipient>();

            Console.WriteLine();
            Console.WriteLine("Moving mail....");
            Console.WriteLine();
            var processedFolder = "";
            var needsAttentionFolder = "";

            foreach (MailFolder folder in currentUser)
            {
                //get specific id's of the folder's name
                if (folder.DisplayName == "Inbox")
                {
                    processedFolder = folder.Id;
                }
                if (folder.DisplayName == "Autolog")
                {
                    needsAttentionFolder = folder.Id;
                }
            }
            Message movedMsg = await client.Me.Messages[MailId].Move(needsAttentionFolder).Request().PostAsync();
        }

        private static async Task getMail(string Subject, GraphServiceClient client)
        {
            var currentUser = await client.Me.MailFolders.Inbox.Messages.Request().GetAsync();
            IList<Recipient> messageToList = new List<Recipient>();

            Console.WriteLine();
            Console.WriteLine("Getting mail....");
            Console.WriteLine();

            foreach (Message email in currentUser)
            {
                //get all mail
                Console.WriteLine(email.Id);
                Console.WriteLine(email.Subject);
                if (Subject == email.Subject)
                {
                    //get specific
                    Console.WriteLine("Got Specific Mail!");
                }
            }
        }
    }
}