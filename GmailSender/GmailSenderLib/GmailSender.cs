// TBD:
// - PGP sign messages

using MimeKit;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Collections;

namespace GmailSenderLib
{
    /// <summary>
    /// Convenience class to create System.Net.MailMessage for use in GmailSender.Send method
    /// </summary>
    public class SimpleMailMessage : System.Net.Mail.MailMessage
    {
        public SimpleMailMessage(string subject, string body, string to, string from = null, bool? isBodyHTML = null)
        {
            //if (from != null) base.From = new System.Net.Mail.MailAddress(from);
            base.Subject = subject;
            base.Body = body;
            base.IsBodyHtml = (isBodyHTML != null) ? ((bool)isBodyHTML) : (body.IndexOf("<html>", System.StringComparison.OrdinalIgnoreCase) != -1);
            base.To.Add(to);
        }
    }

    public class GmailSender
    {
        private GmailService service;

        public GmailSender(string ApplicationName)
        {
            UserCredential credential;

            string[] Scopes = { GmailService.Scope.GmailSend };

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("Google.Apis.Tokens", true)).Result;
            }

            // Create Gmail API service.
            this.service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public string Send(SimpleMailMessage mailMessage)
        {
            return this.service.Users.Messages.Send(
                new Google.Apis.Gmail.v1.Data.Message
                {
                    Raw = Encode(MimeKit.MimeMessage.CreateFromMailMessage(mailMessage).ToString())
                },
                "me"
            ).Execute().Id;
        }

        static private string Encode(string text)
        {
            return System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(text))
                .Replace('+', '-')
                .Replace('/', '_')
                .Replace("=", "");
        }
    }
}
