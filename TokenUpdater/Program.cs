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
using Azure.Storage;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
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

            try
            {
                var p = new Program();
                p.DoOauth().Wait();
            }
            catch (Exception e)
            {
                LogWriter.WriteLog("Wait 5 mins to restart due to error: " , e);
                File.WriteAllText("error.txt", e.ToString());
                Thread.Sleep(new TimeSpan(0,0,5));
                Main(args);
            }
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
            LogWriter.WriteLog("Start a new pass for email server");

            // Creates a redirect URI using an available port on the loopback address.
            //string redirectUri = string.Format("http://127.0.0.1:{0}/", localPortForListerningToken);
            var redirectUri = string.Format("http://127.0.0.1:{0}/", localPortForListerningToken);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectUri);
            LogWriter.WriteLog("Listening ... from redirect URI:" + redirectUri);
            try
            {
                http.Start();
            }
            catch (Exception e)
            {
                LogWriter.WriteLog("Failed to start http server. ", e);
                throw new Exception("failed to start");
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
            LogWriter.WriteLog("HTTP server stopped.");
            Thread.Sleep(new TimeSpan(0,0,0,3));
            proc.Kill();
            LogWriter.WriteLog("Blowser Closed.");
            var data = System.Text.Encoding.UTF8.GetString(buffer);

            // Checks for errors.
            if (context.Request.QueryString.Get("error") != null)
            {
                LogWriter.WriteLog(string.Format("OAuth authorization error: {0}.", context.Request.QueryString.Get("error")), EventLogEntryType.Error);
                throw new Exception("OAuth authorization error");
            }

            if (context.Request.QueryString.Get("code") == null)
            {
                LogWriter.WriteLog("Malformed authorization response. " + context.Request.QueryString);
                throw new Exception("Malformed authorization response");
            }

            // extracts the code
            var code = context.Request.QueryString.Get("code");
            var responseText = await RequestAccessToken(code, redirectUri);

            var parser = new OAuthResponseParser();
            parser.Load(responseText);

            var user = parser.EmailInIdToken;
            var accessToken = parser.AccessToken;


            var durationBeforeTokenExpire = new TimeSpan(0, 45,0);
            var intervalBetweenPulling = new TimeSpan(0, 0, 0, 3);
            LogWriter.WriteLog(DateTime.Now + " Got the token , starting pulling emails from blob for another  " + durationBeforeTokenExpire.TotalMinutes + " minutes in every " + intervalBetweenPulling.TotalSeconds + " seconds");

            for (var interIndex = 0; interIndex < durationBeforeTokenExpire/ intervalBetweenPulling; interIndex++)
            {
                var pendingEmailsList = BlobContainer.GetBlobs(prefix: "PendingOutEmails/").ToList();
                LogWriter.WriteLog("found " + pendingEmailsList.Count + " emails");
                foreach (var curPendingEmail in pendingEmailsList)
                {
                    var stream2 = new MemoryStream();
                    var cancelToken = new CancellationToken();
                    var storageTransferOptions = new StorageTransferOptions
                    {
                        //bytes * 1000000 = MB
                        InitialTransferSize = 1000000,
                        MaximumConcurrency = 1,
                        MaximumTransferSize = long.MaxValue
                    };

                    var blobRequestConditions = new BlobRequestConditions();


                    var blobCurFileClient = BlobContainer.GetBlobClient(curPendingEmail.Name);

                    var response2 = blobCurFileClient.DownloadTo(stream2, blobRequestConditions, storageTransferOptions, cancelToken);


                    var emailContent = Encoding.ASCII.GetString(stream2.ToArray());

                    var toEmail = string.Empty;
                    var subject = string.Empty;
                    var body = string.Empty;

                    var readEmailFailed = true;
                    try
                    {
                        var lines = emailContent.Split(Environment.NewLine);
                        toEmail = lines[0];
                        subject = lines[1];

                        var bodySb = new StringBuilder();
                        for (var i = 2; i < lines.Length; ++i)
                        {
                            bodySb.AppendLine(lines[i]);

                        }

                        body = bodySb.ToString();
                        readEmailFailed = false;
                    }
                    catch (Exception e)
                    {
                        LogWriter.WriteLog("failed to match email format! ", e);
                        readEmailFailed = true;
                    }

                    if (readEmailFailed)
                    {
                        blobCurFileClient.Delete();
                    }
                    else
                    {
                        LogWriter.WriteLog("Read Email " + curPendingEmail.Name);
                        SendMailWithXOAUTH2(user, accessToken, toEmail, subject, body);
                        blobCurFileClient.Delete();
                    }
                }

                Thread.Sleep(intervalBetweenPulling);
            }
           

            LogWriter.WriteLog("finished current loop");

            Thread.Sleep(new TimeSpan(0, 0, 0, 1));
            var localtion = typeof(Program).Assembly.Location.Replace(".dll", ".exe");

            try
            {
                Process.Start(localtion);
            }
            catch (Exception e)
            {
                LogWriter.WriteLog("failed to start next pass", e);
            }

        }


        static void SendMailWithXOAUTH2(string userEmail, string accessToken, string toEmail, string subject, string body)
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
            oMail.To = toEmail;

            oMail.Subject = subject;
            oMail.TextBody = body;
            
            LogWriter.WriteLog("start to send email using OAUTH 2.0 ...");

            SmtpClient oSmtp = new SmtpClient();
            oSmtp.SendMail(oServer, oMail);

            LogWriter.WriteLog("Sent email to " + toEmail);
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
