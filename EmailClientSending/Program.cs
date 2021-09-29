using System;
using System.Collections.Generic;
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


            //var blobCurFileClient = BlobContainer.GetBlobClient("PendingOutEmailAttachments/fbbee48e-1cc6-4686-aa75-b31e7783dfad.dat");

            //var stream2 = new MemoryStream();
            //blobCurFileClient.DownloadTo(stream2);
            //var bytes = stream2.ToArray();
            //File.WriteAllBytes("e:\\test.pdf", bytes);


            //var bytes1 = File.ReadAllBytes("1.pdf");
            //var bytes2 = File.ReadAllBytes("2.pdf");
            SendMailToRelayServer("bsnow@metalogicinspection.com", "asf324234", "4948948\nasdfafd\nadsfasfd");
        }

        static void SendMailToRelayServer(string toEmailAddress, string subject, string body, List<Tuple<string, byte[]>> attachments = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine(toEmailAddress);
            sb.AppendLine(subject);
            sb.Append(body);
            sb.AppendLine();


            if (attachments != null)
            {
                foreach (var attachment in attachments)
                {
                    var curAttachmentRemoteFileName = Guid.NewGuid() + ".dat";
                    var curAttachmentRemoteFilePath = string.Concat("PendingOutEmailAttachments/", curAttachmentRemoteFileName);

                    var curAttachmentBlobFileClient = BlobContainer.GetBlobClient(curAttachmentRemoteFilePath);
                    curAttachmentBlobFileClient.Upload(new MemoryStream(attachment.Item2));
                    sb.Append(curAttachmentRemoteFileName); 
                    sb.Append(" ");
                    sb.AppendLine(attachment.Item1);
                }
            }

        


            var stream = new MemoryStream(Encoding.ASCII.GetBytes(sb.ToString()));

            var fileName = Guid.NewGuid() + ".txt";

            var curEmailRemotePath = string.Concat("PendingOutEmails/", fileName);

            var blobCurFileClient = BlobContainer.GetBlobClient(curEmailRemotePath);
            blobCurFileClient.Upload(stream);
            
            Console.WriteLine(DateTime.Now + " Uploaded email task to blob");
        }
    }
}
