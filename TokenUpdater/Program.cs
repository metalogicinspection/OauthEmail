using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using EASendMail;

namespace TokenUpdater
{
    class Program
    {
        // client configuration
        // You should create your client id and client secret,
        // do not use the following client id in production environment, it is used for test purpose only.
        const string clientID = "2a5938f4-43d2-4668-aef3-550603b90108";
        const string clientSecret = "5dc832d2-bb26-4597-92d9-5e9b9d06e1b6";
        const string localPortForListerningToken = "53977";

        /// <summary>
        /// the path to store the token in blob storage server
        /// </summary>
        const string tokenBlobPath = "LatestEmailToken.txt";

        private static readonly BlobContainerClient BlobContainer = new BlobContainerClient(
            "DefaultEndpointsProtocol=https;AccountName=metalogicreportingblob;AccountKey=3vWVwJcTe/2cmbKYnux7v+qk3cSkW1gbsyE1oVCKJ2kPk1uao4KiZMTxv65Sq/LK2UeDynvo4ZgvKOlHIC6wdA==;EndpointSuffix=core.windows.net",
            "reports");

        
        

        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now + "Start getting latest token");


            #region Delete old loacal stored tokens
            //the EASendMail dll will store the toknen in cur windows profile's download folder,
            //delete the ones from previous runs
            string download = Path.Combine(Environment.GetEnvironmentVariable("USERPROFILE") ,"Downloads");
            var files = Directory.GetFiles(download).Select(x=> new {Path = x, FileName = Path.GetFileNameWithoutExtension(x)}).Where(x => x.FileName.Equals("download") || x.FileName.StartsWith("download ")).ToList();
            foreach (var file in files)
            {
                File.Delete(file.Path);
            }

            #endregion


            localBrowserPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
            if (!File.Exists(localBrowserPath))
            {
                localBrowserPath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
            }
            var p = new Program();
            p.DoOauth().Wait();
           



        }
        


        static string localBrowserPath = string.Empty;
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

        async Task DoOauth()
        {
            try
            {
                // Creates a redirect URI using an available port on the loopback address.
                //string redirectUri = string.Format("http://127.0.0.1:{0}/", localPortForListerningToken);
                string redirectUri = string.Format("http://127.0.0.1:{0}/", localPortForListerningToken);
                Console.WriteLine("redirect URI: " + redirectUri);

                // Creates an HttpListener to listen for requests on that redirect URI.
                var http = new HttpListener();
                http.Prefixes.Add(redirectUri);
                Console.WriteLine("Listening ...");
                try
                {
                    http.Start();

                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

                // Creates the OAuth 2.0 authorization request.
                var authorizationRequest =
                    $"{authUri}?response_type=code&scope={scope}&redirect_uri={Uri.EscapeDataString(redirectUri)}&client_id={clientID}";

                // Opens request in the browser.
                var proc = Process.Start(localBrowserPath, authorizationRequest);

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
                http.Stop();
                Console.WriteLine("HTTP server stopped.");
                proc.Kill();
                Console.WriteLine("Blowser Closed.");
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

                //Console.WriteLine("User: {0}", user);
                //Console.WriteLine("AccessToken: {0}", accessToken);

                Console.WriteLine(DateTime.Now + " Got the token");


                var blobCurFileClient = BlobContainer.GetBlobClient(tokenBlobPath);
                var byteArray = Encoding.ASCII.GetBytes(accessToken);
                var stream = new MemoryStream(byteArray);
                 


                await blobCurFileClient.UploadAsync(stream, true, CancellationToken.None);
                Console.WriteLine(DateTime.Now + " Uploaded token to blob");
                
                Thread.Sleep(new TimeSpan(0,0,30,5));
                var localtion = typeof(Program).Assembly.Location.Replace(".dll", ".exe");
                Console.WriteLine(localtion);

                Process.Start(localtion);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                File.WriteAllText("error.txt", e.ToString());
                throw;
            }
           
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
