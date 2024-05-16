﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading;

namespace GmailAPIExample
{
    class Program
    {
        const string EMAIL_ADDRESS = "jfwallac@gmail.com";
        static readonly string[] Scopes = { GmailService.Scope.MailGoogleCom };
        const string ApplicationName = "UnsubscribeMe";

        static T Safe<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e);
            }

            return default(T);
        }

        static T Retry<T>(Func<T> f)
        {
            while (true)
            {
                try
                {
                    return f();
                }
                catch (Google.GoogleApiException ex)
                {
                    if (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
                    {
                        // Extract the retry time from the exception message
                        Console.WriteLine(ex.Message);
                        string retryAfter = ex.Message.Split('(')[0].Trim().Split(' ').Last();
                        // Create a CultureInfo object with the "en-US" culture
                        CultureInfo culture = new CultureInfo("en-US");

                        // Define the custom date and time format
                        string format = "yyyy-MM-ddTHH:mm:ss.fffZ";

                        DateTime retryTime = DateTime.ParseExact(retryAfter, format, culture);

                        // Calculate the delay until the retry time
                        TimeSpan delay = retryTime - DateTime.Now;

                        if (delay.TotalSeconds > 0)
                        {
                            Console.WriteLine($"User-rate limit exceeded. Retrying after {delay.TotalSeconds} seconds...");

                            // Wait for the specified delay before retrying
                            System.Threading.Thread.Sleep(delay);

                            // Retry the operation
                        }
                        else
                        {
                            Console.WriteLine("User-rate limit exceeded, but the retry time has already passed.");
                        }
                    }
                    else
                    {
                        // Handle other types of GoogleApiException
                        Console.WriteLine($"GoogleApiException occurred: {ex.Message}");
                    }
                } catch (TaskCanceledException) {
                    Console.WriteLine($"Got TaskCanceledException: Retrying ...");
                }
            }
        }

        static IList<Message> FetchAllMessages(GmailService service, string folderId)
        {
            var result = new List<Message>();
            string? nextPageToken = String.Empty;

            do
            {
                var request = service.Users.Messages.List("me");
                request.LabelIds = new List<string>() { folderId };
                request.PageToken = nextPageToken;
                request.MaxResults = 500;

                var response = Safe(request.Execute);

                IList<Message>? messages = response?.Messages;
                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine($"Got {messages.Count} messages");
                    Console.WriteLine($"Total so far: {result.Count}");
                    foreach (var message in messages)
                    {
                        var msg = Retry(() => service.Users.Messages.Get("me", message.Id).Execute());
                        if (msg != null)
                        {
                            result.Add(msg);
                        }
                    }
                }

                nextPageToken = response?.NextPageToken;
            } while (!String.IsNullOrEmpty(nextPageToken));

            return result;
        }

        static void Main(string[] args)
        {
            var cache = new HashSet<string>();
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
            string folderId = GetFolderId(service, folderName);

            if (string.IsNullOrEmpty(folderId))
            {
                Console.WriteLine($"Folder '{folderName}' not found.");
                return;
            }

            var allMessages = FetchAllMessages(service, folderId);
            var toUnsubscribe = allMessages.Where(msg => msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase)) != null);
            var toDelete = allMessages.Where(msg => msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase)) == null);

            var batchDeleteRequest = new BatchDeleteMessagesRequest
            {
                Ids = toDelete.Select(m => m.Id).ToList()
            };
            service.Users.Messages.BatchDelete(batchDeleteRequest, "me");

            Console.WriteLine($"Found {allMessages.Count()} messages in folder: {folderId}");
            Console.WriteLine($"Found {toDelete.Count()} messages that do not have List-Unsubscribe header... deleting");

            var groupedByHeader = toUnsubscribe.GroupBy(msg =>
                msg.Payload.Headers.FirstOrDefault(header => header.Name.Equals("List-Unsubscribe", StringComparison.OrdinalIgnoreCase))?.Value
            );

            foreach (var group in groupedByHeader)
            {
                var values = group.Key.Split(",").Select(x => ExtractUnsubscribeUrl(x));
                foreach (var unsubscribeUrl in values)
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

                var deleteRequest = new BatchDeleteMessagesRequest
                {
                    Ids = group.Select(m => m.Id).ToList()
                };
                service.Users.Messages.BatchDelete(batchDeleteRequest, "me");
            }
        }

        static void DeleteMail(GmailService service, Message message)
        {
            try
            {
                Console.WriteLine($"Deleting Message {message.Id}");
                var deleteRequest = service.Users.Messages.Delete("me", message.Id);
                Safe(deleteRequest.Execute);
            }
            catch (Google.GoogleApiException e)
            {
                Console.WriteLine($"Could not delete {message.Id}: {e}");
            }
        }

        static string GetFolderId(GmailService service, string folderName)
        {
            // List all labels in the user's mailbox.
            var labels = Safe(service.Users.Labels.List("me").Execute)?.Labels;

            // Find the label with the specified folder name.
            var folder = labels?.FirstOrDefault(label => label.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

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
                Raw = CreateMessageBody(EMAIL_ADDRESS, toAddress, subject, "")
            };

            // Send the unsubscribe email.
            Retry(service.Users.Messages.Send(message, "me").Execute);
        }

        static string CreateMessageBody(string from, string to, string subject, string body)
        {
            var msgBody = $"From: {from}\r\nTo: {to}\r\nSubject: {subject}\r\n\r\n{body}";
            Console.WriteLine(msgBody);
            return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(msgBody));
        }
    }
}
