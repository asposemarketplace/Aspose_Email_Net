using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.EmailProcessing.Library;
using System.Web.UI.HtmlControls;
using Aspose.Email.Imap;
using Aspose.Email.Outlook.Pst;
using Aspose.Email.Outlook;

namespace Aspose.EmailProcessing
{
    public partial class MailBox : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ErrorsDiv.Visible = false;

            if (!Page.IsPostBack)
            {
                try
                {
                    MailHelper.Current.PopulateFoldersList(ref MailFoldersRepeater);
                }
                catch (Exception ex)
                {
                    ErrorsDiv.Visible = true;
                    ErrorLiteral.Text = ex.ToString();
                }
            }
        }

        protected void MailFoldersRepeater_ItemCommand(Object Sender, RepeaterCommandEventArgs e)
        {
            try
            {
                if (e.CommandName == "ShowFolderDetails")
                {
                    foreach (RepeaterItem item in MailFoldersRepeater.Items)
                    {
                        if (item.ItemType == ListItemType.Item || item.ItemType == ListItemType.AlternatingItem)
                        {
                            HtmlGenericControl folderRowLi = item.FindControl("folderRowLi") as HtmlGenericControl;
                            folderRowLi.Attributes.Remove("class");
                        }
                    }

                    LinkButton FolderNameLinkButton = e.Item.FindControl("FolderNameLinkButton") as LinkButton;
                    ViewState["SelectedFolderName"] = FolderNameLinkButton.Text;

                    Label FolderUriLabel = e.Item.FindControl("FolderUriLabel") as Label;
                    ViewState["SelectedFolderUri"] = FolderUriLabel.Text;

                    HtmlGenericControl folderRowLi2 = e.Item.FindControl("folderRowLi") as HtmlGenericControl;
                    folderRowLi2.Attributes.Add("class", "active");

                    RenderSelectedMessages(FolderUriLabel.Text);
                }
            }
            catch (Exception ex)
            {
                ErrorsDiv.Visible = true;
                ErrorLiteral.Text = ex.ToString();
            }
        }

        private void RenderSelectedMessages(string folderUri)
        {
            int count = MailHelper.Current.ListMessagesInFolder(ref MessagesGridView, folderUri);
            PagingDiv.Visible = ExportButtonsDiv.Visible = (count > 0) ? true : false;

            InfoDiv.Visible = false;

            if (MessagesGridView.Rows.Count > 0)
            {
                MessagesGridView.UseAccessibleHeader = true;
                MessagesGridView.HeaderRow.TableSection = TableRowSection.TableHeader;
                //MessagesGridView.FooterRow.TableSection = TableRowSection.TableFooter;
            }
        }

        protected void PST_ExportSelectedLinkButton_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> messagesList = new List<string>();

                foreach (GridViewRow row in MessagesGridView.Rows)
                {
                    if (row.RowType == DataControlRowType.DataRow)
                    {
                        CheckBox chkRow = (row.Cells[0].FindControl("SelectedCheckBox") as CheckBox);
                        if (chkRow.Checked)
                        {
                            string UniqueUri = MessagesGridView.DataKeys[row.RowIndex].Value.ToString();
                            messagesList.Add(UniqueUri);
                        }
                    }
                }

                if (messagesList.Count() > 0)
                {
                    ErrorsDiv.Visible = false;
                    MailHelper.Current.ExportSelectedMessagesToPST(ViewState["SelectedFolderName"].ToString(), messagesList);
                }
                else
                {
                    // Binding again to make sure paging works
                    RenderSelectedMessages(ViewState["SelectedFolderUri"].ToString());
                    ErrorsDiv.Visible = true;
                    ErrorLiteral.Text = "Please select one or more messages to export";
                }
            }
            catch (Exception ex)
            {
                ErrorsDiv.Visible = true;
                ErrorLiteral.Text = ex.ToString();
            }
        }

        protected void PST_ExportAllLinkButton_Click(object sender, EventArgs e)
        {
            try
            {
                MailHelper.Current.ExportFolderToPST(ViewState["SelectedFolderUri"].ToString());
            }
            catch (Exception ex)
            {
                ErrorsDiv.Visible = true;
                ErrorLiteral.Text = ex.ToString();
            }
        }

        protected void LogoutButton_Click(object sender, EventArgs e)
        {
            if (HttpContext.Current.Session[Constants.MailHelperSession] != null)
            {
                HttpContext.Current.Session.Remove(Constants.MailHelperSession);
            }
            HttpContext.Current.Response.Redirect("~/Default.aspx");
        }
    }
}