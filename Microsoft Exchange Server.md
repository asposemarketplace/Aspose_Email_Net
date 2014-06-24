<h3>Include Namespaces</h3>
<p>Please make sure to add the following namespaces before trying the code below</p>
<pre><code>using Aspose.Email.Outlook.Pst;
using Aspose.Email.Outlook;
</code></pre>
<h3><a name="connect-to-microsoft-exchange"></a>Connect to Microsoft Exchange Server using Aspose.Email</h3>
<p>Connecting to Microsoft Exchange server is very easy with just two lines of code</p>
<pre><code>NetworkCredential credentials = new NetworkCredential(Username, Password, Domain);
IEWSClient client = EWSClient.GetEWSClient(ServerURL, credentials);
</code></pre>
<h3><a name="get-folders-list-from-microsoft-exchange"></a>Get folders list from Microsoft Exchange server</h3>
<p>You can get list of Mailbox folders from Microsoft Exchange Server using the following lines of code</p>
<pre><code>ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
string rootUri = client.GetMailboxInfo().RootUri;
// List all the folders from Exchange server
ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(rootUri);
</code></pre>
<h3><a name="get-all-messages-from-microsoft-exchange-folder"></a>Get all messages from a Microsoft Exchange Folder</h3>
<p>Once you get list of all folders from a Microsoft Exchange Server Mailbox, each folder contains a Uri which enables you to fetch all messages from that folder, here is a single line of code which can do that for you</p>
<pre><code>ExchangeMessageInfoCollection msgCollection = client.ListMessages(folderUri);
</code></pre>
<h3><a name="get-single-message-from-microsoft-exchange"></a>Get single message from Microsoft Exchange Server</h3>
<p>If you have bound all the messages to a GridView or some other control, upon selecting a single message you will need to get that single message again for showing its details or using it for any other purpose. You can use this single line of code to do that</p>
<pre><code>MailMessage mail = client.FetchMessage(messageUri);
</code></pre>
<h3><a name="export-entire-microsoft-exchange-folder-to-pst-file"></a>Export entire Microsoft Exchange Server folder to a PST file</h3>
<p>Aspose.Email gives you a great option to export the entire Microsoft Exchange Server folder to a PST file with having to loop through all messages and putting them in PST file one by one. Please check the following lines of code to do that</p>
<pre><code>ExchangeFolderInfo info = client.GetFolderInfo(folderUri);
ExchangeFolderInfoCollection fc = new ExchangeFolderInfoCollection();
string outputFilePath = System.Guid.NewGuid().ToString() + ".pst";
fc.Add(info);
client.Backup(fc,outputFilePath, Aspose.Email.Outlook.Pst.BackupOptions.Recursive);
</code></pre>
<h3><a name="export-selected-messages-from-microsoft-exchange-to-pst-file"></a>Export selected messages from Microsoft Exchange Server to a PST file</h3>
<p>Aspose.Email also gives you the option to export selected messages from any folder to a PST file. Below is the code to do that</p>
<pre><code>string outputFilePath = System.Guid.NewGuid().ToString() + ".pst";
PersonalStorage pst = PersonalStorage.Create(outputFilePath, FileFormatVersion.Unicode);
pst.RootFolder.AddSubFolder(folderName);
FolderInfo outfolder = pst.RootFolder.GetSubFolder(folderName);
foreach (string mUri in messagesUrisList)
{
    MailMessage mail = client.FetchMessage(mUri);
    outfolder.AddMessage(MapiMessage.FromMailMessage(mail));
}
pst.Dispose();
</code></pre>