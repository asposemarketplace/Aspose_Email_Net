using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System.Web.UI.WebControls;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Outlook.Pst;
using Aspose.Email.Outlook;


namespace Aspose.EmailProcessing.Library
{
    public class ExchangeHelper : MailHelper
    {
        public string Domain { get; set; }

        public ExchangeHelper(string su, string dom, string u, string p, MailTypeEnum mte)
        {
            ServerURL = su; Domain = dom; Username = u; Password = p; MailType = mte;
        }

        private IEWSClient client
        {
            get { return (IEWSClient)MailClient; }
        }

        public override bool VerfiyCredentials()
        {
            try
            {
                NetworkCredential credentials = new NetworkCredential(Username, Password, Domain);
                IEWSClient client = EWSClient.GetEWSClient(ServerURL, credentials);
                MailClient = client;
            }
            catch (Exception)
            {
                HttpContext.Current.Session.Remove(Constants.MailHelperSession);
                return false;
            }
            return true;
        }

        private void AddFolderToList(ref List<MailFolder> foldersList, ref ExchangeFolderInfoCollection folderInfoColl, string folderName)
        {
            try
            {
                var folList = (from obj in folderInfoColl where obj.DisplayName.ToLower().Contains(folderName) select obj);
                if (folList != null)
                {
                    var folder = folList.FirstOrDefault();
                    if (folder != null)
                    {
                        foldersList.Add(new MailFolder(folder.DisplayName, folder.Uri));
                        folderInfoColl.Remove(folder);
                    }
                }
            }
            catch (Exception) { }
        }

        public override void PopulateFoldersList(ref Repeater repater)
        {
            // Register callback method for SSL validation event
            ServicePointManager.ServerCertificateValidationCallback += RemoteCertificateValidationHandler;

            try
            {
                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                string rootUri = client.GetMailboxInfo().RootUri;
                // List all the folders from Exchange server
                ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(rootUri);

                List<MailFolder> foldersList = new List<MailFolder>();
                List<MailFolder> foldersList2 = new List<MailFolder>();

                AddFolderToList(ref foldersList, ref folderInfoCollection, "inbox");
                AddFolderToList(ref foldersList, ref folderInfoCollection, "drafts");
                AddFolderToList(ref foldersList, ref folderInfoCollection, "sent");
                AddFolderToList(ref foldersList, ref folderInfoCollection, "deleted");
                AddFolderToList(ref foldersList, ref folderInfoCollection, "trash");

                foldersList2 = (from list in folderInfoCollection
                                select new MailFolder()
                                {
                                    FolderName = list.DisplayName,
                                    FolderUri = list.Uri
                                }).ToList();

                foldersList.AddRange(foldersList2);

                repater.DataSource = foldersList;
                repater.DataBind();
            }
            catch (System.Exception) { }
        }

        private static bool RemoteCertificateValidationHandler(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true; //ignore the checks and go ahead
        }

        public override int ListMessagesInFolder(ref GridView gridView, string folderUri)
        {
            try
            {
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(folderUri);
                List<Message> messagesList = new List<Message>();

                messagesList = (from msg in msgCollection
                                orderby msg.Date descending
                                select new Message()
                                {
                                   UniqueUri = msg.UniqueUri,
                                   Date = msg.Date,
                                   Subject = msg.Subject,
                                   From = msg.From.DisplayName
                                }
                               ).ToList();

                gridView.DataSource = messagesList;
                gridView.DataBind();
                gridView.Visible = true;

                return msgCollection.Count;
            }
            catch (Exception) { }
            return 0;
        }

        public override void ExportFolderToPST(string folderUri)
        {
            if (!Directory.Exists(HttpContext.Current.Server.MapPath("~/output")))
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("~/output"));

            ExchangeFolderInfo info = client.GetFolderInfo(folderUri);
            ExchangeFolderInfoCollection fc = new ExchangeFolderInfoCollection();

            string outputFileName = System.Guid.NewGuid().ToString() + ".pst";

            fc.Add(info);
            client.Backup(fc, HttpContext.Current.Server.MapPath(string.Format("~/output/{0}", outputFileName)), Aspose.Email.Outlook.Pst.BackupOptions.Recursive);

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "application/octet-stream";
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=" + outputFileName);
            HttpContext.Current.Response.WriteFile(HttpContext.Current.Server.MapPath(string.Format("~/output/{0}", outputFileName)));
            HttpContext.Current.Response.End();
        }

        public override void ExportSelectedMessagesToPST(string folderName, List<string> messagesUrisList)
        {
            if (!Directory.Exists(HttpContext.Current.Server.MapPath("~/output")))
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("~/output"));

            //Create PST
            string outputFileName = System.Guid.NewGuid().ToString() + ".pst";
            PersonalStorage pst = PersonalStorage.Create(HttpContext.Current.Server.MapPath(string.Format("~/output/{0}", outputFileName)), FileFormatVersion.Unicode);
            pst.RootFolder.AddSubFolder(folderName);
            FolderInfo outfolder = pst.RootFolder.GetSubFolder(folderName);

            foreach (string mUri in messagesUrisList)
            {
                MailMessage mail = client.FetchMessage(mUri);
                outfolder.AddMessage(MapiMessage.FromMailMessage(mail));
            }

            pst.Dispose();

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "application/octet-stream";
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=" + outputFileName);
            HttpContext.Current.Response.WriteFile(HttpContext.Current.Server.MapPath(string.Format("~/output/{0}", outputFileName)));
            HttpContext.Current.Response.End();
        }

        public void GetFoldersList()
        {
            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);
            ExchangeFolderInfo info = client.GetFolderInfo(client.MailboxInfo.RootUri);

            foreach (ExchangeMessageInfo msgInfo in msgCollection)
            {
                MailMessage msg = client.FetchMessage(msgInfo.UniqueUri);
            }
        }
    }
}