using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading;

namespace GmailAPIExample
{
    class Program
    {
        static string[] Scopes = { GmailService.Scope.MailGoogleCom };
        static string ApplicationName = "UnsubscribeMe";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
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
            string folderId = GetFolderId(service, folderName);

            if (string.IsNullOrEmpty(folderId))
            {
                Console.WriteLine($"Folder '{folderName}' not found.");
                return;
            }

            // List all messages in the specified folder.
            // ListMessagesResponse response = service.Users.Messages.List("me").Q($"in:{folderId}").Execute();
            var response = service.Users.Messages.List("me").Execute();
            IList<Message> messages = response.Messages;

            if (messages != null && messages.Count > 0)
            {
                foreach (var message in messages)
                {
                    // Get the message details.
                    Message msg = service.Users.Messages.Get("me", message.Id).Execute();

                    // Check if the "List-Unsubscribe" header exists.
                    var unsubscribeHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase));

                    if (unsubscribeHeader != null)
                    {
                        var values = unsubscribeHeader.Value.Split(",").Select(x => ExtractUnsubscribeUrl(x));
                        foreach (var unsubscribeUrl in values)
                        {
                            if (!string.IsNullOrEmpty(unsubscribeUrl))
                            {
                                if (unsubscribeUrl.StartsWith("mailto"))
                                {
                                    SendUnsubscribeEmail(service, unsubscribeUrl);
                                    Console.WriteLine($"Unsubscribe email sent for message with ID: {msg.Id}");
                                }
                                else
                                {
                                    HttpClient client = new HttpClient();
                                    client.GetAsync(unsubscribeUrl).Wait();
                                    Console.WriteLine($"Unsubscribe URL visited {unsubscribeUrl}");
                                }

                                try
                                {
                                    var deleteRequest = service.Users.Messages.Delete("me", msg.Id);
                                    deleteRequest.Execute();
                                }
                                catch(Google.GoogleApiException e)
                                {
                                    Console.WriteLine($"Could not delete {msg.Id}: {e}");
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No messages found in the specified folder.");
            }
        }

        static string GetFolderId(GmailService service, string folderName)
        {
            // List all labels in the user's mailbox.
            var labels = service.Users.Labels.List("me").Execute().Labels;

            // Find the label with the specified folder name.
            var folder = labels.FirstOrDefault(label => label.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

            return folder?.Id;
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
                // Raw = CreateUnsubscribeEmailBody(unsubscribeUrl)
                Raw = CreateMessageBody("jfwallac@gmail.com", toAddress, subject, "")
            };

            // Send the unsubscribe email.
            service.Users.Messages.Send(message, "me").Execute();
        }

        static string CreateUnsubscribeEmailBody(string unsubscribeUrl)
        {
            // Create the email body for the unsubscribe request.
            // This is a simple example, and you may need to adjust it based on the specific requirements of the unsubscribe process.
            var body = $"To: {unsubscribeUrl}\r\nSubject: Unsubscribe\r\n\r\nPlease unsubscribe me from this mailing list.";

            return Base64UrlEncode(body);
        }

        static string CreateMessageBody(string from, string to, string subject, string body)
        {
            var msgBody = $"From: {from}\r\nTo: {to}\r\nSubject: {subject}\r\n\r\n{body}";
            Console.WriteLine(msgBody);
            return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(msgBody));
        }

        static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }
    }
}
