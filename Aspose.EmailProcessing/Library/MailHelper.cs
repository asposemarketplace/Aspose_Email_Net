using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System.Web.UI.WebControls;
using System.IO;

namespace Aspose.EmailProcessing.Library
{
    public enum MailTypeEnum
    {
        SMTP = 1,
        POP3 = 2,
        IMAP = 3,
        ExchangeServer = 4
    }

    public class MailFolder
    {
        public MailFolder() { }
        public MailFolder(string fName, string fUri) { FolderName = fName; FolderUri = fUri; }

        public string FolderName { get; set; }
        public string FolderUri { get; set; }
    }

    public class Message
    {
        public Message() { }
        public Message(string u, DateTime d, string s, string f) { UniqueUri = u; Date = d; Subject = s; From = f; }

        public string UniqueUri { get; set; }
        public DateTime Date { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
    }

    public abstract class MailHelper
    {
        public MailHelper() { }

        public object MailClient { get; set; }
        public string ServerURL { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public MailTypeEnum MailType { get; set; }

        public static MailHelper Current
        {
            get
            {
                string licenseFile = HttpContext.Current.Server.MapPath("~/App_Data/Aspose.Total.lic");
                if (File.Exists(licenseFile))
                {
                    Aspose.Email.License license = new Aspose.Email.License();
                    license.SetLicense(licenseFile);
                }

                if (HttpContext.Current.Session[Constants.MailHelperSession] != null)
                {
                    return (MailHelper)HttpContext.Current.Session[Constants.MailHelperSession];
                }
                else
                {
                    HttpContext.Current.Response.Redirect("~/Default.aspx");
                }

                return null;
            }
        }

        public abstract bool VerfiyCredentials();

        public abstract void PopulateFoldersList(ref Repeater treeNode);

        public abstract int ListMessagesInFolder(ref GridView gridView, string folderUri);

        public abstract void ExportFolderToPST(string folderUri);

        public abstract void ExportSelectedMessagesToPST(string folderName, List<string> messagesUris);
    }
}