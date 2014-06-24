<h3>Include Namespaces</h3>
<p>Please make sure to add the following namespaces before trying the code below</p>
<pre><code>using Aspose.Email.Imap;
using Aspose.Email.Outlook.Pst;
using Aspose.Email.Outlook;
</code></pre>
<h3><a name="connect-to-imap"></a>Connect to IMAP Server using Aspose.Email</h3>
<p>You can connect to an IMAP server by using the following lines of code. If the server is SSL enabled then you will need to provide its SSL port otherwise default port i.e. 143 will be used for non SSL connection.</p>
<pre><code>ImapClient client = new ImapClient(ServerURL, (SSLEnabled ? SSLPort : 143), Username, Password);
if (SSLEnabled)
{
    client.EnableSsl = true;
    client.SecurityMode = Aspose.Email.Imap.ImapSslSecurityMode.Implicit; // set security mode
}
</code></pre>
<h3><a name="get-folders-list-from-imap"></a>Get folders list from IMAP Server</h3>
<p>You can get list of Mailbox folders from IMAP Server using this line of code</p>
<pre><code>ImapFolderInfoCollection folderInfoColl = client.ListFolders();
</code></pre>
<h3><a name="get-all-messages-from-imap-folder"></a>Get all messages from a Microsoft Exchange Folder</h3>
<p>Once you get list of all folders from an IMAP Server Mailbox, you can use the Folder name to first select that folder and then get list of all messages in that folder.</p>
<pre><code>client.SelectFolder(folderName);
ImapMessageInfoCollection msgInfoColl = client.ListMessages();
</code></pre>
<h3><a name="get-single-message-from-imap"></a>Get single message from IMAP Server</h3>
<p>If you have bound all the messages to a GridView or some other control, upon selecting a single message you will need to get that single message again for showing its details or using it for any other purpose. You can use this single line of code to do that</p>
<pre><code>MailMessage mail = client.FetchMessage(uniqueId);
</code></pre>
<h3><a name="export-entire-imap-folder-to-pst-file"></a>Export entire IMAP Server folder to a PST file</h3>
<p>Please check the following lines of code to export an entire folder to PST file</p>
<pre><code>string outputFileName = System.Guid.NewGuid().ToString() + ".pst";
PersonalStorage pst = PersonalStorage.Create(outputFileName, FileFormatVersion.Unicode);
pst.RootFolder.AddSubFolder(folderName);
FolderInfo outfolder = pst.RootFolder.GetSubFolder(folderName);

client.SelectFolder(folderName);
ImapMessageInfoCollection messageInfoCollection = client.ListMessages();

foreach (ImapMessageInfo msg in messageInfoCollection)
{
    MailMessage mail = client.FetchMessage(msg.UniqueId);
    outfolder.AddMessage(MapiMessage.FromMailMessage(mail));
}

pst.Dispose();
</code></pre>
<h3><a name="export-selected-messages-from-imap-to-pst-file"></a>Export selected messages from IMAP Server to a PST file</h3>
<p>You can export selected IMAP messages to a PST file. Below is the code to do that</p>
<pre><code>string outputFileName = System.Guid.NewGuid().ToString() + ".pst";
PersonalStorage pst = PersonalStorage.Create(outputFileName, FileFormatVersion.Unicode);
pst.RootFolder.AddSubFolder(folderName);
FolderInfo outfolder = pst.RootFolder.GetSubFolder(folderName);

client.SelectFolder(folderName);

foreach (string mUri in messagesUrisList)
{
    MailMessage mail = client.FetchMessage(mUri);
    outfolder.AddMessage(MapiMessage.FromMailMessage(mail));
}

pst.Dispose();
</code></pre>