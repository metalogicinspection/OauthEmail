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

            SendMailWithXOAUTH2("software@metalogicinspection.com", "eyJ0eXAiOiJKV1QiLCJub25jZSI6IllQU3NJcnVRMlRDQi1FRTVsdzNFVS1JZzFSMG5fT2Q4Qm1fMlctWHA5aWMiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2UyNzM0MTZkLWVkNzItNGVkNC04MDYxLWEzOGQ5MDllZDM1Mi8iLCJpYXQiOjE2MzA2MDMwMjksIm5iZiI6MTYzMDYwMzAyOSwiZXhwIjoxNjMwNjA2OTI5LCJhY2N0IjowLCJhY3IiOiIxIiwiYWlvIjoiQVZRQXEvOFRBQUFBMGk1c0tnTHg0K2FkK29QSWlzY0pnRHNhOXBsOStMRWhPYTFpU1hYTEp3Z2g0cTl0eThxM2IyUFB5WWUvZ2IvQmpGOXNnbVZBT2ZiUFk3VitFNHNBL20xbFJpRUxpMlMxaVoySXgwUVJMams9IiwiYW1yIjpbInB3ZCIsIm1mYSJdLCJhcHBfZGlzcGxheW5hbWUiOiJNZXRhbG9naWMgRW1haWwiLCJhcHBpZCI6IjJhNTkzOGY0LTQzZDItNDY2OC1hZWYzLTU1MDYwM2I5MDEwOCIsImFwcGlkYWNyIjoiMCIsImVuZnBvbGlkcyI6W10sImlwYWRkciI6IjIwNy4yMjguMTA1LjIwNSIsIm5hbWUiOiJzb2Z0d2FyZSIsIm9pZCI6IjlmNzc0OWQ5LWExNTItNDJhYi04ZmQ3LTAxOTVhMmZjZTk0YiIsInB1aWQiOiIxMDAzMjAwMEJENkExMzFBIiwicHdkX2V4cCI6IjEyOTUyIiwicHdkX3VybCI6Imh0dHBzOi8vcG9ydGFsLm1pY3Jvc29mdG9ubGluZS5jb20vQ2hhbmdlUGFzc3dvcmQuYXNweCIsInJoIjoiMC5BUndBYlVGejRuTHQxRTZBWWFPTmtKN1RVdlE0V1NyU1EyaEdydk5WQmdPNUFRZ2NBT2suIiwic2NwIjoiU01UUC5TZW5kIiwic2lkIjoiNWM1NjAyNGItMDZhNy00MjA3LTg5MzgtMDlhNzg5MzQ5NGQyIiwic3ViIjoiX0gtVlpEcUZnajVLaWxaczg5ZmgxZkNYRlZHY3NramZMNWcwS3Uwb25GZyIsInRpZCI6ImUyNzM0MTZkLWVkNzItNGVkNC04MDYxLWEzOGQ5MDllZDM1MiIsInVuaXF1ZV9uYW1lIjoic29mdHdhcmVAbWV0YWxvZ2ljaW5zcGVjdGlvbi5jb20iLCJ1cG4iOiJzb2Z0d2FyZUBtZXRhbG9naWNpbnNwZWN0aW9uLmNvbSIsInV0aSI6IkplS09WbUdaM1VHZ3p0MWpKVnc0QUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdfQ.RljRQOXRouTRBU1T2vHtY_0i2Hudf8G91YI0zk1WnNoj0_l5hdm7G4JFI-yG7S_aseA4HjI9BV4Y-N-AVqbKpJcutyxewAlmSmXYBUsYGlysGY89hQQymsSetQK3vWtDpxaa8FQzDNbAV-jFAeBkl5Zad_fKOfRXyNkYtqaA5UAtJNe5QNqyc-z1dVmYeDoWZq2nfm1hWwAiADWKtY304BdsxqKEQisBUow94VqpV9bL4nJZvRq_ooM4SMySySWRh0SmCfn621UnJKlcnYpR8eLbgv1aELse4xtn9g5dAUzK7Xh9LtO8hrxlM3kBg5DvGY5QK162Nk2ngRon5jqtIw");
            //Console.WriteLine("+------------------------------------------------------------------+");
            //Console.WriteLine("  Sign in with MS OAuth                                             ");
            //Console.WriteLine("   If you got \"This app isn't verified\" information in Web Browser, ");
            //Console.WriteLine("   click \"Advanced\" -> Go to ... to continue test.");
            //Console.WriteLine("+------------------------------------------------------------------+");
            //Console.WriteLine("");
            //Console.WriteLine("Press any key to sign in...");
            //Console.ReadKey();

            //try
            //{
            //    Program p = new Program();
            //    p.DoOauthAndSendEmail();
            //}
            //catch (Exception ep)
            //{
            //    Console.WriteLine(ep.ToString());
            //}
            //Console.ReadKey();
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
            oMail.To = "jhe@MetalogicInspection.com";

            oMail.Subject = "test email from Office365 account with OAUTH 2";
            oMail.TextBody = "this is a test email sent from c# project with EWS.";

            Console.WriteLine("start to send email using OAUTH 2.0 ...");

            SmtpClient oSmtp = new SmtpClient();
            oSmtp.SendMail(oServer, oMail);

            Console.WriteLine("The email has been submitted to server successfully!");
        }

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

        async void DoOauthAndSendEmail()
        {
            // Creates a redirect URI using an available port on the loopback address.
            string redirectUri = string.Format("http://127.0.0.1:{0}/", GetRandomUnusedPort());
            Console.WriteLine("redirect URI: " + redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectUri);
            Console.WriteLine("Listening ...");
            http.Start();

            // Creates the OAuth 2.0 authorization request.
            string authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}&prompt=login",
                authUri,
                scope,
                Uri.EscapeDataString(redirectUri),
                clientID
            );
            
            // Opens request in the browser.
            Process.Start("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", authorizationRequest);

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

            string responseText = await RequestAccessToken(code, redirectUri);
            Console.WriteLine(responseText);

            OAuthResponseParser parser = new OAuthResponseParser();
            parser.Load(responseText);

            var user = parser.EmailInIdToken;
            var accessToken = parser.AccessToken;

            Console.WriteLine("User: {0}", user);
            Console.WriteLine("AccessToken: {0}", accessToken);

            SendMailWithXOAUTH2(user, accessToken);
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