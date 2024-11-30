using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace readgooglemailssize
{
    class GmailAPIExample
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly }; // "https://www.googleapis.com/auth/gmail.readonly"
        static string ApplicationName = "Gmail API Example";

        private IDictionary<string, string> _labelMap;

        private IList<MailMessageInformation> _messagesInfo;

        static void Main(string[] args)
        {
            try
            {
                new GmailAPIExample().Run();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        void Run()
        {
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
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

            _labelMap = GetEmailLabels(service);
            //ListLargestEmails(service);
            _messagesInfo = ListLargestEmailsParallel(service);
        }

        IList<MailMessageInformation> ListLargestEmailsParallel(GmailService service)
        {
            var messageInfoBag = new ConcurrentBag<MailMessageInformation>();

            int maxSizeInBytes = 150000;
            maxSizeInBytes = 500000;
            maxSizeInBytes = 1000000;
            maxSizeInBytes = 0;
            maxSizeInBytes = 10000000;

            int maxMatchingMailsCount = 100;

            // Fetch messages
            var request = service.Users.Messages.List("me");
            request.MaxResults = 100; // Adjust for more results.
            request.MaxResults = 75;
            request.Q = ""; // Add Gmail query to filter if needed, e.g., has:attachment.
            request.Q = $"larger:{maxSizeInBytes}"; // Add Gmail query to filter if needed, e.g., has:attachment.

            int index = 0;
            int indexPage = 0;

            int totalNumberOfMails = 0;
            Int64 totalEstimatedSize = 0;

            bool first = true;
            ListMessagesResponse response = null;
            string nextPageToken = null;
            List<ListMessagesResponse> responseList = new List<ListMessagesResponse>();

            while (first || !string.IsNullOrEmpty(response.NextPageToken))
            {
                if (first)
                {
                    response = request.Execute();
                    first = false;
                }
                else
                {
                    request.PageToken = nextPageToken;
                    response = request.Execute();
                }
                nextPageToken = response?.NextPageToken;
                if (null != response)
                    responseList.Add(response);
            }

            var options = new ParallelOptions()
            {
                MaxDegreeOfParallelism = 20
            };

            object _incrementsLock = new object();
            int matchingMailsCount = 0;
            Parallel.ForEach(responseList, options, (r) =>
            {
                var tid = AppDomain.GetCurrentThreadId();
                foreach (var messageItem in r.Messages)
                {
                    lock (_incrementsLock)
                        if (matchingMailsCount >= maxMatchingMailsCount)
                            break;

                    lock (_incrementsLock)
                        ++totalNumberOfMails;
                    //Interlocked.Increment(ref totalNumberOfMails);

                    var messageRequest = service.Users.Messages.Get("me", messageItem.Id);
                    var message = messageRequest.Execute();

                    var from = GetHeader(message, "From");
                    var to = GetHeader(message, "To");
                    var subject = GetHeader(message, "Subject");
                    var labels = GetLabels(message);

                    if (message.SizeEstimate > maxSizeInBytes)
                    {
                        lock (_incrementsLock)
                        {
                            ++matchingMailsCount;
                            totalEstimatedSize += (Int64)message.SizeEstimate;
                            if (matchingMailsCount > maxMatchingMailsCount)
                                break;

                        //Interlocked.Add(ref totalEstimatedSize, (Int64)message.SizeEstimate);
                        //Interlocked.Increment(ref matchingMailsCount);
                        //if (matchingMailsCount >= maxMatchingMailsCount)
                        //    break;

                            WriteLine($"{matchingMailsCount}\t{tid}\t{message.SizeEstimate} bytes\t{DateTimeOffset.FromUnixTimeMilliseconds(message.InternalDate ?? 0L)}\t{from.SmartSubString(0, 30)}\t{to.SmartSubString(0, 30)}\t{subject.SmartSubString(0, 30)}\t{message.Snippet.ToString().SmartSubString(0, 80)}\t{string.Join(",", labels)}");
                            messageInfoBag.Add(new MailMessageInformation()
                            {
                                Subject = subject,
                                From = from,
                                To = to,
                                Date = DateTimeOffset.FromUnixTimeMilliseconds(message.InternalDate ?? 0L).DateTime,
                                Size = message.SizeEstimate ?? 0,
                                Snippet = message.Snippet,
                                Labels = string.Join(",", labels),
                            });
                        }
                    }
                }
            });
            Console.WriteLine($"matching mails: {Math.Min(maxMatchingMailsCount, matchingMailsCount)}\tEstimated size: {Microsoft.VisualBasic.Strings.FormatNumber(totalEstimatedSize, 0)}");
            return messageInfoBag.ToList();
        }

        void ListLargestEmails(GmailService service)
        {
            // Fetch messages
            Google.Apis.Gmail.v1.UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
            request.MaxResults = 100; // Adjust for more results.
            request.Q = "larger:150000"; // Add Gmail query to filter if needed, e.g., has:attachment.

            int index = 0;
            int totalNumberOfMails = 0;

            bool first = true;
            ListMessagesResponse response = null;
            string nextPageToken = null;
            while (first || !string.IsNullOrEmpty(response.NextPageToken))
            {
                if (first)
                {
                    response = request.Execute();
                    first = false;
                }
                else
                {
                    request.PageToken = nextPageToken;
                    response = request.Execute();
                }
                nextPageToken = response.NextPageToken;

                if (response.Messages != null)
                {
                    foreach (var messageItem in response.Messages)
                    {
                        ++totalNumberOfMails;
                        // Fetch each message details
                        var messageRequest = service.Users.Messages.Get("me", messageItem.Id);
                        var message = messageRequest.Execute();

                        var from = GetHeader(message, "From");
                        var subject = GetHeader(message, "Subject");
                        Console.Write($"{++index}\r");
                        if (message.SizeEstimate > 01 * 15000)
                            Console.WriteLine($"{message.SizeEstimate} bytes\t{DateTimeOffset.FromUnixTimeMilliseconds(message.InternalDate ?? 0L)}\t{from.SmartSubString(0, 30)}\t{subject.SmartSubString(0, 30)}\t{message.Snippet.ToString().SmartSubString(0, 80)}");
                    }
                }
            }
            Console.WriteLine($"{totalNumberOfMails} mails read");
        }

        string GetHeader(Message message, string headerName)
        {
            var ret = (message.Payload.Headers.Where(x => x.Name == headerName).Select(x => x.Value).FirstOrDefault()) ?? string.Empty;
            return ret;
        }

        IList<string> GetLabels(Message message)
        {
            IList<string> ret = new List<string>();
            foreach (var labelId in message.LabelIds)
            {
                string txt;
                if (!_labelMap.TryGetValue(labelId, out txt))
                    txt = labelId;
                ret.Add(txt);
            }
            return ret;
        }

        IDictionary<string, string> GetEmailLabels(GmailService service)//, string messageId)
        {
            //// Fetch the message by ID
            //var messageRequest = service.Users.Messages.Get("me", messageId);
            //var message = messageRequest.Execute();

            //// Display the labels associated with the message
            //Console.WriteLine($"Message ID: {message.Id}");
            //Console.WriteLine("Labels:");

            //if (message.LabelIds != null)
            //{
            //    foreach (var labelId in message.LabelIds)
            //    {
            //        Console.WriteLine($"- {labelId}");
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No labels found for this message.");
            //}

            // Fetch the list of all labels to map IDs to names
            var labelsRequest = service.Users.Labels.List("me");
            var labelsResponse = labelsRequest.Execute();

            // Create a dictionary to map label IDs to label names
            Dictionary<string, string> labelMap = new Dictionary<string, string>();
            foreach (var label in labelsResponse.Labels)
            {
                labelMap[label.Id] = label.Name;
            }
            return labelMap;

            // Display the label names for this message
            Console.WriteLine("Label Names:");
            //if (message.LabelIds != null)
            //{
            //    foreach (var labelId in message.LabelIds)
            //    {
            //        if (labelMap.ContainsKey(labelId))
            //        {
            //            Console.WriteLine($"- {labelMap[labelId]}");
            //        }
            //        else
            //        {
            //            Console.WriteLine($"- {labelId} (Unknown Label)");
            //        }
            //    }
            //}
        }

        private static object _lock = new object();
        static void WriteLine(string s)
        {
            lock (_lock)
                System.Console.WriteLine(s);
        }
    }
}