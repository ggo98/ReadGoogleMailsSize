using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

static class GmailAPIExample
{
    static string[] Scopes = { GmailService.Scope.GmailReadonly }; // "https://www.googleapis.com/auth/gmail.readonly"
    static string ApplicationName = "Gmail API Example";

    static void Main(string[] args)
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

        GetEmailLabels(service);
        //ListLargestEmails(service);
        ListLargestEmailsParallel(service);
    }

    static void ListLargestEmailsParallel(GmailService service)
    {
        int maxSizeInBytes = 150000;
        maxSizeInBytes = 500000;
        maxSizeInBytes = 1000000;

        // Fetch messages
        var request = service.Users.Messages.List("me");
        request.MaxResults = 100; // Adjust for more results.
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
            //MaxDegreeOfParallelism = 5
        };

        int matchingMailsCount = 0;
        Parallel.ForEach(responseList, options, (r) =>
        {
            var tid = AppDomain.GetCurrentThreadId();
            foreach (var messageItem in r.Messages)
            {
                ++totalNumberOfMails;

                var messageRequest = service.Users.Messages.Get("me", messageItem.Id);
                var message = messageRequest.Execute();

                var from = GetHeader(message, "From");
                var subject = GetHeader(message, "Subject");
                //Console.Write($"{++index}\r");
                //WriteLine($"{tid}\t{message.SizeEstimate} bytes");
                if (message.SizeEstimate > 01 * maxSizeInBytes)
                {
                    Interlocked.Add(ref totalEstimatedSize, (Int64)message.SizeEstimate);
                    Interlocked.Increment(ref matchingMailsCount);
                    WriteLine($"{tid}\t{message.SizeEstimate} bytes\t{DateTimeOffset.FromUnixTimeMilliseconds(message.InternalDate ?? 0L)}\t{from.SmartSubString(0, 30)}\t{subject.SmartSubString(0, 30)}\t{message.Snippet.ToString().SmartSubString(0, 80)}");
                }
            }
        });
        Console.WriteLine($"matching mails: {matchingMailsCount}\tEstimated size: {Microsoft.VisualBasic.Strings.FormatNumber(totalEstimatedSize, 0)}");
    }

    private static object _lock = new object();
    static void WriteLine(string s)
    {
        lock (_lock)
            System.Console.WriteLine(s);
    }

    static void ListLargestEmails(GmailService service)
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

    static string GetHeader(Message message, string headerName)
    {
        var ret = (message.Payload.Headers.Where(x => x.Name == headerName).Select(x => x.Value).FirstOrDefault()) ?? string.Empty;
        return ret;
    }

    static IDictionary<string, string> GetEmailLabels(GmailService service)//, string messageId)
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

    public static string SmartSubString(this string s, int startIndex, int length)
    {
        int slen = s.Length;
        if (startIndex > slen)
            return string.Empty;
        int len = slen - startIndex;
        if (len > length)
            len = length;
        if (len < 0)
            return string.Empty;
        string ret = s.Substring(startIndex, len);
        return ret;
    }

}