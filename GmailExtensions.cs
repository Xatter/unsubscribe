using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Requests;
using System.Text.Json;

namespace GmailAPIExample
{
    public static class GmailExtensions
    {
        public static void BatchDelete(this GmailService service, IList<string> ids)
        {
            // Gmail API limitation, can only delete 1000 messages at a time
            var batches = ids.Chunk(1000);
            foreach (var batch in batches)
            {
                var deleteRequest = new BatchDeleteMessagesRequest
                {
                    Ids = batch
                };

                // Maybe this needs to be retried?
                service.Users.Messages.BatchDelete(deleteRequest, "me").Execute();
            }
        }
        public static string? GetFolderId(this GmailService service, string folderName)
        {
            // List all labels in the user's mailbox.
            var labels = Helpers.Safe(service.Users.Labels.List("me").Execute)?.Labels;

            // Find the label with the specified folder name.
            var folder = labels?.FirstOrDefault(label => label.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

            return folder?.Id;
        }
        public static IList<Message> FetchAllMessagesInFolder(this GmailService service, string? folderId = null)
        {
            var result = new List<Message>();
            string? nextPageToken = String.Empty;
            BatchRequest batch = new BatchRequest(service);

            do
            {
                var request = service.Users.Messages.List("me");
                if (folderId != null)
                {
                    request.LabelIds = new List<string>() { folderId };
                }
                request.PageToken = nextPageToken;
                request.MaxResults = 1000;

                var response = Helpers.Safe(request.Execute);

                IList<Message>? messages = response?.Messages;
                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine($"Got {messages.Count} messages");
                    Console.WriteLine($"Total so far: {result.Count}");
                    foreach (var message in messages)
                    {
                        var getRequest = service.Users.Messages.Get("me", message.Id);
                        getRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Metadata;

                        var msg = Helpers.Retry(getRequest.Execute);

                        if (msg != null)
                        {
                            File.AppendAllLines("allMessages.json", new List<string> { JsonSerializer.Serialize(msg) });
                            result.Add(msg);
                        }
                    }
                }

                nextPageToken = response?.NextPageToken;
                batch = new BatchRequest(service);
            } while (!String.IsNullOrEmpty(nextPageToken));

            return result;
        }

        public static void SendUnsubscribeEmail(this GmailService service, string email, string unsubscribeUrl)
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
                Raw = CreateMessageBody(email, toAddress, subject, "")
            };

            // Send the unsubscribe email.
            try
            {
                Helpers.Retry(service.Users.Messages.Send(message, "me").Execute);
            }
            catch
            {
                Console.WriteLine(message.Raw);
            }
        }

        static string CreateMessageBody(string from, string to, string subject, string body)
        {
            var msgBody = $"From: {from}\r\nTo: {to}\r\nSubject: {subject}\r\n\r\n{body}";
            Console.WriteLine(msgBody);
            return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(msgBody));
        }
    }
}
