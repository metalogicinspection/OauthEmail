using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using Azure;
using Azure.Storage;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs.Specialized;
using EASendMail;

namespace EmailClientSending
{
    class Program
    {   
        /// <summary>
        /// the path to store the token in blob storage server
        /// </summary>
        const string tokenBlobPath = "LatestEmailToken.txt";

        private static readonly BlobContainerClient BlobContainer = new BlobContainerClient(
            "DefaultEndpointsProtocol=https;AccountName=metalogicreportingblob;AccountKey=3vWVwJcTe/2cmbKYnux7v+qk3cSkW1gbsyE1oVCKJ2kPk1uao4KiZMTxv65Sq/LK2UeDynvo4ZgvKOlHIC6wdA==;EndpointSuffix=core.windows.net",
            "reports");


        static void Main(string[] args)
        {
            var blobCurFileClient = BlobContainer.GetBlobClient(tokenBlobPath);
            var stream = new MemoryStream();
            var cancelToken = new CancellationToken();
            var storageTransferOptions = new StorageTransferOptions
            {
                //bytes * 1000000 = MB
                InitialTransferSize = 1000000,
                MaximumConcurrency = 1,
                MaximumTransferSize = long.MaxValue
            };

            var blobRequestConditions = new BlobRequestConditions();
            var response = blobCurFileClient.DownloadTo(stream, blobRequestConditions, storageTransferOptions, cancelToken);
            var value = Encoding.ASCII.GetString(stream.ToArray());
            Console.Write("Got token : \n" + value);


            SendMailWithXOAUTH2("software@metalogicinspection.com", value);
            //if (response.Status == Convert.ToInt32(HttpStatusCode.OK)
            //{
            //    using (StreamReader streamReader = new StreamReader(stream))
            //    {
            //        messageJson = streamReader.ReadToEnd();
            //    }
            //}


        }

        static void SendMailWithXOAUTH2(string userEmail, string accessToken)
        {
            // Office365 server address
            SmtpServer oServer = new SmtpServer("outlook.office365.com");

            // use EWS protocol
            oServer.Protocol = ServerProtocol.ExchangeEWS;
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
            oMail.To = "jhe@metalogicinspection.com";

            oMail.Subject = "test email from Office365 account with OAUTH 2";
            oMail.TextBody = "this is a test email sent from c# project with EWS.";

            Console.WriteLine("start to send email using OAUTH 2.0 ...");

            SmtpClient oSmtp = new SmtpClient();
            oSmtp.SendMail(oServer, oMail);

            Console.WriteLine("The email has been submitted to server successfully!");
        }
    }
}
