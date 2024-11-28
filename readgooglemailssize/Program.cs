using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Linq;
using System.Threading;

static class GmailAPIExample
{
    static string[] Scopes = { GmailService.Scope.GmailReadonly };
    static string ApplicationName = "Gmail API Example";

    static void Main(string[] args)
    {
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

        ListLargestEmails(service);
    }

    static void ListLargestEmails(GmailService service)
    {
        // Fetch messages
        var request = service.Users.Messages.List("me");
        request.MaxResults = 100; // Adjust for more results.
        request.Q = ""; // Add Gmail query to filter if needed, e.g., has:attachment.

        var response = request.Execute();
        int index = 0;

        while (!string.IsNullOrEmpty(response.NextPageToken))
        {
            if (response.Messages != null)
            {
                foreach (var messageItem in response.Messages)
                {
                    // Fetch each message details
                    var messageRequest = service.Users.Messages.Get("me", messageItem.Id);
                    var message = messageRequest.Execute();

                    Console.Write($"{++index}\r");
                    if (message.SizeEstimate > 150000)
                        Console.WriteLine($"{message.SizeEstimate} bytes\t{DateTimeOffset.FromUnixTimeMilliseconds(message.InternalDate ?? 0L)}\t{message.Snippet.ToString().SmartSubString(0, 100)}");

                    //Console.WriteLine($"Message ID: {message.Id}");
                    //Console.WriteLine($"Size Estimate: {message.SizeEstimate} bytes");
                    //Console.WriteLine($"Snippet: {message.Snippet}");
                    //Console.WriteLine(new string('-', 40));
                }
            }
            request.PageToken = response.NextPageToken;
            response = request.Execute();
        }
        //else
        //{
        //    Console.WriteLine("No messages found.");
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