using System;
using System.IO;
using System.Text;
using Azure.Storage.Blobs;

namespace EmailClientSending
{
    class Program
    {
        private static readonly BlobContainerClient BlobContainer = new BlobContainerClient(
            "DefaultEndpointsProtocol=https;AccountName=metalogicreportingblob;AccountKey=3vWVwJcTe/2cmbKYnux7v+qk3cSkW1gbsyE1oVCKJ2kPk1uao4KiZMTxv65Sq/LK2UeDynvo4ZgvKOlHIC6wdA==;EndpointSuffix=core.windows.net",
            "reports");


        static void Main(string[] args)
        {
            SendMailToRelayServer("jhe@metalogicinspection.com", "adfibobfo", "tes\nsdf\n\nsd34f\nt");
        }

        static void SendMailToRelayServer(string toEmailAddress, string subject, string body)
        {
            var sb = new StringBuilder();
            sb.AppendLine(toEmailAddress);
            sb.AppendLine(subject);
            sb.Append(body);
            var stream = new MemoryStream(Encoding.ASCII.GetBytes(sb.ToString()));

            var fileName = Guid.NewGuid() + ".txt";
            var curEmailRemotePath = string.Concat("PendingOutEmails/", fileName);


            var blobCurFileClient = BlobContainer.GetBlobClient(curEmailRemotePath);
            blobCurFileClient.Upload(stream);
            
            Console.WriteLine(DateTime.Now + " Uploaded email task to blob");
        }
    }
}
