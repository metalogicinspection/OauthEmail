using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using EASendMail;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            //localBrowserPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";

            localBrowserPath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
            if (!File.Exists(localBrowserPath))
            {
                localBrowserPath = "C:\\Program Files (x86)\\Microsoft\\Edge Beta\\Application\\msedge.exe";
            }
            //SendMailWithXOAUTH2("Jhe@metalogicinspection.com","eyJ0eXAiOiJKV1QiLCJub25jZSI6InVUbm81R0tMbWs4dVVsUWgzckMxQ0NjdnZRd09xbXVPbWlyTTctb3pYNzQiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2UyNzM0MTZkLWVkNzItNGVkNC04MDYxLWEzOGQ5MDllZDM1Mi8iLCJpYXQiOjE2MzA1MzUzOTMsIm5iZiI6MTYzMDUzNTM5MywiZXhwIjoxNjMwNTM5MjkzLCJhY2N0IjowLCJhY3IiOiIxIiwiYWlvIjoiQVRRQXkvOFRBQUFBK0pORWYxNGR5b0RjaHpOSE81NllhWnEyb0xQVVc1L094Mm9RR1pyWElXSENXOG1HMmxSVVVnZTVOZmJXNktvSSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiTWV0YWxvZ2ljIEVtYWlsIiwiYXBwaWQiOiIyYTU5MzhmNC00M2QyLTQ2NjgtYWVmMy01NTA2MDNiOTAxMDgiLCJhcHBpZGFjciI6IjAiLCJlbmZwb2xpZHMiOltdLCJmYW1pbHlfbmFtZSI6IkhlIiwiZ2l2ZW5fbmFtZSI6IkppYWh1aSIsImlwYWRkciI6IjIwNy4yMjguMTA1LjIwNSIsIm5hbWUiOiJKaWFodWkgSGUiLCJvaWQiOiJmYWQxMDUzZS1mNGY0LTQxYmMtODNkMi1mZTJkZWZkYWNmNDkiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMTA5MTA2NzUyMC0zMjQxMDM3MjEwLTI5MTg1MTk1My0xMjI5IiwicHVpZCI6IjEwMDM3RkZFOURDRThDNUMiLCJyaCI6IjAuQVJ3QWJVRno0bkx0MUU2QVlhT05rSjdUVXZRNFdTclNRMmhHcnZOVkJnTzVBUWdjQUd3LiIsInNjcCI6IlNNVFAuU2VuZCIsInNpZCI6IjI4ZTQ5MjYwLTI0MGEtNDFjOC1hNTdlLTMwZDBkZTc0MjdlMSIsInN1YiI6IjBTakxwanRZZXhzVmRBU3FIZzQyLVdoVzBvQ3BWN25kaHo0eFJ1MTh5bVUiLCJ0aWQiOiJlMjczNDE2ZC1lZDcyLTRlZDQtODA2MS1hMzhkOTA5ZWQzNTIiLCJ1bmlxdWVfbmFtZSI6IkpIZUBNZXRhbG9naWNJbnNwZWN0aW9uLmNvbSIsInVwbiI6IkpIZUBNZXRhbG9naWNJbnNwZWN0aW9uLmNvbSIsInV0aSI6IjJqNlNZX2c1MVVpSEFkaC1VVllRQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdfQ.MYjAogzsFUHn9x1LDvt3zVqUCaya3f3pgYkKN2gbXsLyMEHCPtVYH_rsthFXgeqfRW-XBtPR683CD8uk7M-oFc3_ZUWPKnvLTexkQF3OAw7TIdrKOe9LXe-KIFBb6I3anyOFLJy2RICU6FrK3eRL3rcnmlnjBtjySXjMmpSqpDIu8evRp47qQYMVqhtqiqcGRGfCramDgweY27b_jqvcU9KJ3BZLTnF3QVHjpMOj9gKwzOW9CVYUPchEthsa2TKubQyBIEGrYvKc_1DhyGqozXs6rZnwL_W1g6qNkq8IqbXLxLjj7LjL_cvFsHQRXwobqTrSUPQjUyQznwXVYeSUkw");
            //Console.WriteLine("+------------------------------------------------------------------+");
            //Console.WriteLine("  Sign in with MS OAuth                                             ");
            //Console.WriteLine("   If you got \"This app isn't verified\" information in Web Browser, ");
            //Console.WriteLine("   click \"Advanced\" -> Go to ... to continue test.");
            //Console.WriteLine("+------------------------------------------------------------------+");
            //Console.WriteLine("");
            //Console.WriteLine("Press any key to sign in...");
            //Console.ReadKey();

            try
            {
                Program p = new Program();
                p.DoOauth();
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.ToString());
            }
            Console.ReadKey();
        }

        static void SendMailWithXOAUTH2(string userEmail, string accessToken)
        {
            // Office365 server address
            SmtpServer oServer = new SmtpServer("outlook.office365.com");

            // Using 587 port, you can also use 465 port
            oServer.Port = 587;
            // enable SSL connection
            oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

            // use SMTP OAUTH 2.0 authentication
            oServer.AuthType = SmtpAuthType.XOAUTH2;
            // set user authentication
            oServer.User = userEmail;
            // use access token as password
            oServer.Password = accessToken;

            SmtpMail oMail = new SmtpMail("TryIt");
            // Your email address
            oMail.From = userEmail;

            // Please change recipient address to yours for test
            oMail.To = "bsnow@MetalogicInspection.com";

            oMail.Subject = "test email from Office365 account with OAUTH 2";
            oMail.TextBody = "this is a test email sent from c# project with EWS.";

            Console.WriteLine("start to send email using OAUTH 2.0 ...");

            SmtpClient oSmtp = new SmtpClient();
            oSmtp.SendMail(oServer, oMail);

            Console.WriteLine("The email has been submitted to server successfully!");
        }

        static string localBrowserPath = string.Empty;
        // client configuration
        // You should create your client id and client secret,
        // do not use the following client id in production environment, it is used for test purpose only.
        const string clientID = "2a5938f4-43d2-4668-aef3-550603b90108";
        const string clientSecret = "5dc832d2-bb26-4597-92d9-5e9b9d06e1b6";
        // set SMTP scope
        const string scope = "https://outlook.office.com/SMTP.Send%20offline_access%20email%20openid";
        const string authUri = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        const string tokenUri = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        // if your application is single tenant, please use tenant id instead of common in authUri and tokenUri
        // for example, your tenant is 669595d0-a4d7-47c5-8040-cf9970400e48, then
        // const string authUri = "https://login.microsoftonline.com/669595d0-a4d7-47c5-8040-cf9970400e48/oauth2/v2.0/authorize";
        // const string tokenUri = "https://login.microsoftonline.com/669595d0-a4d7-47c5-8040-cf9970400e48/oauth2/v2.0/token";

        static int GetRandomUnusedPort()
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }

        async void DoOauth()
        {
            // Creates a redirect URI using an available port on the loopback address.
            string redirectUri = string.Format("http://127.0.0.1:{0}/", "53977");
            Console.WriteLine("redirect URI: " + redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectUri);
            Console.WriteLine("Listening ...");
            http.Start();

            // Creates the OAuth 2.0 authorization request.
            var authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}",
            //var authorizationRequest = string.Format("{0}?response_type=code&grant_type=refresh_token&scope={1}&redirect_uri={2}&client_id={3}",
                authUri,
                scope,
                Uri.EscapeDataString(redirectUri),
                clientID
            );

            // Opens request in the browser.
            Process.Start(localBrowserPath, authorizationRequest);

            // Waits for the OAuth authorization response.
            var context = await http.GetContextAsync();

            // Brings the Console to Focus.
            BringConsoleToFront();

            // Sends an HTTP response to the browser.
            var response = context.Response;
            //string responseString = string.Format("<html><head></head><body>Please return to the app and close current window.</body></html>");
            //var buffer = Encoding.UTF8.GetBytes(responseString);
            var buffer = new byte[1024];
            response.ContentLength64 = buffer.Length;
            var responseOutput = response.OutputStream;
            response.OutputStream.Write(buffer, 0, buffer.Length);
            //Task responseTask = responseOutput.WriteAsync(buffer, 0, buffer.Length).ContinueWith((task) =>
            //{
            //    responseOutput.Close();
            //    http.Stop();
            //    Console.WriteLine("HTTP server stopped.");
            //});
            http.Stop();
            Console.WriteLine("HTTP server stopped.");
            var data = System.Text.Encoding.UTF8.GetString(buffer);

            // Checks for errors.
            if (context.Request.QueryString.Get("error") != null)
            {
                Console.WriteLine(string.Format("OAuth authorization error: {0}.", context.Request.QueryString.Get("error")));
                return;
            }

            if (context.Request.QueryString.Get("code") == null)
            {
                Console.WriteLine("Malformed authorization response. " + context.Request.QueryString);
                return;
            }

            // extracts the code
            var code = context.Request.QueryString.Get("code");
            Console.WriteLine("Authorization code: " + code);

            var responseText = await RequestAccessToken(code, redirectUri);
            Console.WriteLine(responseText);

            var parser = new OAuthResponseParser();
            parser.Load(responseText);

            var user = parser.EmailInIdToken;
            var accessToken = parser.AccessToken;

            Console.WriteLine("User: {0}", user);
            Console.WriteLine("AccessToken: {0}", accessToken);

            //SendMailWithXOAUTH2(user, accessToken);
        }




        async Task<string> RequestAccessToken(string code, string redirectUri)
        {
            Console.WriteLine("Exchanging code for tokens...");

            // builds the  request
            string tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&grant_type=authorization_code",
                code,
                Uri.EscapeDataString(redirectUri),
                clientID
                );

            // if you use it in web application, please add clientSecret parameter
            /*string tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&client_secret={3}&grant_type=authorization_code",
                code,
                Uri.EscapeDataString(redirectUri),
                clientID,
                clientSecret
                );
            */
            // sends the request
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(tokenUri);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

            byte[] _byteVersion = Encoding.ASCII.GetBytes(tokenRequestBody);
            tokenRequest.ContentLength = _byteVersion.Length;

            Stream stream = tokenRequest.GetRequestStream();
            await stream.WriteAsync(_byteVersion, 0, _byteVersion.Length);
            stream.Close();

            try
            {
                // gets the response
                WebResponse tokenResponse = await tokenRequest.GetResponseAsync();
                using (StreamReader reader = new StreamReader(tokenResponse.GetResponseStream()))
                {
                    // reads response body
                    return await reader.ReadToEndAsync();
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        Console.WriteLine("HTTP: " + response.StatusCode);
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // reads response body
                            string responseText = await reader.ReadToEndAsync();
                            Console.WriteLine(responseText);
                        }
                    }
                }

                throw ex;
            }
        }

        // Hack to bring the Console window to front.

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public void BringConsoleToFront()
        {
            SetForegroundWindow(GetConsoleWindow());
        }
    }
}