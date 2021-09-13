﻿using System;
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
            var bytes1 = File.ReadAllBytes("1.pdf");
            var bytes2 = File.ReadAllBytes("2.pdf");
            SendMailToRelayServer("jhe@metalogicinspection.com", "test5555555", "4948948", new List<Tuple<string, byte[]>>(){new Tuple<string, byte[]>("1.pdf", bytes1), new Tuple<string, byte[]>("2.pdf", bytes2)});
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
