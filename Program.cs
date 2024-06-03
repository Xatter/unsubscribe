using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading;
using System.Collections.Concurrent;
using System.Net.Sockets;
using System.Text.Json;

namespace GmailAPIExample
{

    public class Program
    {
        const string EMAIL_ADDRESS = "jfwallac@gmail.com";
        static readonly string[] Scopes = { GmailService.Scope.MailGoogleCom };
        const string ApplicationName = "UnsubscribeMe";


        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define the folder name to search for.
            string folderName = "Crap";

            // Get the folder ID based on the folder name.
            string folderId = service.GetFolderId(folderName);

            if (string.IsNullOrEmpty(folderId))
            {
                Console.WriteLine($"Folder '{folderName}' not found.");
                return;
            }

            IList<Message> allMessages = new List<Message>();

            if (File.Exists("allMessages.json"))
            {
                foreach (var line in File.ReadAllLines("allMessages.json"))
                {
                    var record = JsonSerializer.Deserialize<Message>(line);
                    allMessages.Add(record);
                }
            }
            else
            {
                allMessages = service.FetchAllMessagesInFolder(folderId);
            }

            var toUnsubscribe = allMessages.Where(msg => msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase)) != null);
            Console.WriteLine($"Found {allMessages.Count()} messages in folder: {folderId}");

            var groupedByHeader = toUnsubscribe.GroupBy(msg =>
            {
                var fromHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("From", StringComparison.OrdinalIgnoreCase));
                var replyHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("Reply-To", StringComparison.OrdinalIgnoreCase));

                if (fromHeader != null) { return fromHeader.Value; }
                else if (replyHeader != null) { return replyHeader.Value; }
                else { return String.Empty; }
            }).OrderByDescending(g => g.Count());

            HashSet<string> seen = new HashSet<string>();
            if (File.Exists("seen.json"))
                seen = new HashSet<string>(File.ReadAllLines("seen.json"));

            foreach (var group in groupedByHeader)
            {
                if (seen.Contains(group.Key))
                {
                    Console.WriteLine($"Already processed {group.Key}, skipping ....... ");
                    service.BatchDelete(group.Select(m => m.Id).ToList());
                    continue;
                }

                var firstMessage = group.FirstOrDefault();
                var unsubscribeHeader = firstMessage?.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase));
                var values = unsubscribeHeader?.Value.Split(",").Select(x => ExtractUnsubscribeUrl(x));
                if (values != null)
                {
                    foreach (var unsubscribeUrl in values)
                    {
                        if (!string.IsNullOrEmpty(unsubscribeUrl))
                        {

                            if (unsubscribeUrl.StartsWith("mailto"))
                            {
                                SendUnsubscribeEmail(service, unsubscribeUrl);
                                Console.WriteLine($"Unsubscribe email sent for message with List-Unsubscribe: {group.Key}");
                            }
                            else
                            {
                                try
                                {
                                    HttpClient client = new HttpClient();
                                    client.GetAsync(unsubscribeUrl).Wait();
                                    Console.WriteLine($"Unsubscribe URL visited {unsubscribeUrl}");
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine($"Error while visiting URL: {unsubscribeUrl}: {e}");
                                }
                            }
                        }
                    }
                }

                service.BatchDelete(group.Select(m => m.Id).ToList());
                File.AppendAllLines("seen.json", new List<string>() { group.Key });
            }

            File.Delete("allMessages.json");
            File.Delete("seen.json");
        }



        static string ExtractUnsubscribeUrl(string headerValue)
        {
            // Extract the unsubscribe URL from the header value.
            // The header value format is typically: <mailto:unsubscribe@example.com> or <https://example.com/unsubscribe>
            // You may need to adjust this based on the specific format of the unsubscribe URLs you encounter.
            var urlStart = headerValue.IndexOf('<');
            var urlEnd = headerValue.IndexOf('>');

            if (urlStart >= 0 && urlEnd > urlStart)
            {
                return headerValue.Substring(urlStart + 1, urlEnd - urlStart - 1);
            }

            return null;
        }

        static void SendUnsubscribeEmail(GmailService service, string unsubscribeUrl)
        {
            var withoutMailto = unsubscribeUrl.Substring(unsubscribeUrl.IndexOf(":") + 1);
            var toAddress = withoutMailto;
            var subject = String.Empty;

            if (withoutMailto.Contains('?'))
            {
                toAddress = withoutMailto.Split("?")[0];
                var query = withoutMailto.Split("?")[1];
                subject = query.Substring(query.IndexOf("subject=") + ("subject=".Count()));
            }

            // Create the unsubscribe email message.
            var message = new Message
            {
                Raw = CreateMessageBody(EMAIL_ADDRESS, toAddress, subject, "")
            };

            // Send the unsubscribe email.
            Helpers.Retry(service.Users.Messages.Send(message, "me").Execute);
        }

        static string CreateMessageBody(string from, string to, string subject, string body)
        {
            var msgBody = $"From: {from}\r\nTo: {to}\r\nSubject: {subject}\r\n\r\n{body}";
            Console.WriteLine(msgBody);
            return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(msgBody));
        }
    }
}
