<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Site.Master" CodeBehind="MailBox.aspx.cs" Inherits="Aspose.EmailProcessing.MailBox" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <script type="text/javascript">
        $(document).ready(function () {
            $('.selectAllCheckBox input[type="checkbox"]').click(function (event) {  //on click
                if (this.checked) { // check select status
                    $('.selectableCheckBox input[type="checkbox"]').each(function () { //loop through each checkbox
                        this.checked = true;  //select all checkboxes with class "checkbox1"              
                    });
                } else {
                    $('.selectableCheckBox input[type="checkbox"]').each(function () { //loop through each checkbox
                        this.checked = false; //deselect all checkboxes with class "checkbox1"                      
                    });
                }
            });

            $('.tablesorterPager img').click(function (event) {  //on click
                $('.selectAllCheckBox input[type="checkbox"]').each(function () {  //on click
                    this.checked = false;
                });
            });

            $('#MessagesGridView')
            .tablesorter({ widthFixed: true, widgets: ['zebra'], headers: { 0: { sorter: false } } })
            .tablesorterPager({ container: $('#pager'), size: 20, positionFixed: false });
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <br />
    <div class="container">
        <table class="mailBoxTable">
            <tr>
                <td class="coloredRow" colspan="2">
                    <div class="floatLeft">
                        <h4>Welcome</h4>
                    </div>
                    <div>
                    </div>
                    <div id="ExportButtonsDiv" runat="server" visible="false" class="btn-group floatRight">
                        <button data-toggle="dropdown" type="button" class="btn btn-default btn-sm dropdown-toggle">
                            PST Export &nbsp;&nbsp;&nbsp;<span class="caret"></span>
                        </button>
                        <ul role="menu" class="dropdown-menu">
                            <li>
                                <asp:LinkButton ID="PST_ExportSelectedLinkButton" OnClick="PST_ExportSelectedLinkButton_Click" runat="server">Export Selected Messages</asp:LinkButton>
                            </li>
                            <li>
                                <asp:LinkButton ID="PST_ExportAllLinkButton" OnClick="PST_ExportAllLinkButton_Click" runat="server">Export Selected Folder</asp:LinkButton>
                            </li>
                        </ul>
                    </div>
                </td>
            </tr>
            <tr>
                <td class="rowBorderTop rowBorderRight width20">
                    <asp:Repeater ID="MailFoldersRepeater" runat="server" OnItemCommand="MailFoldersRepeater_ItemCommand">
                        <HeaderTemplate>
                            <ul style="max-width: 260px;" class="nav nav-pills nav-stacked">
                        </HeaderTemplate>
                        <ItemTemplate>
                            <li id="folderRowLi" runat="server">
                                <asp:LinkButton ID="FolderNameLinkButton" runat="server" Text='<%# Eval("FolderName")%>' CommandName="ShowFolderDetails" />
                                <asp:Label ID="FolderUriLabel" runat="server" Text='<%# Eval("FolderUri")%>' Visible="false" />
                            </li>
                        </ItemTemplate>
                        <FooterTemplate>
                            </ul>
                        </FooterTemplate>
                    </asp:Repeater>
                </td>
                <td class="rowBorderTop width80">
                    <div class="alert alert-warning" id="ErrorsDiv" runat="server" visible="false">
                        <strong>Oops!</strong>
                        <asp:Literal ID="ErrorLiteral" Text="Please select one or more messages to export" runat="server"></asp:Literal>
                    </div>
                    <div id="InfoDiv" runat="server">
                        Please select folder from the left side to view messages.
                        <br />
                        <br />
                        This sample demostrates some of very cool features provided by Aspose.Email for .NET. The features in this version includes
                        <ul>
                            <li>Connect to Exchange and IMAP servers</li>
                            <li>Fetch folders list and messages in those folders</li>
                            <li>Export selected/all messages in a folder to PST</li>
                            <li>The generated PST is can be download immediately</li>
                        </ul>
                    </div>
                    <asp:GridView ID="MessagesGridView" EmptyDataText="There are no messages in this folder." EmptyDataRowStyle-CssClass="emptyClass"  GridLines="None" BorderWidth="0" AutoGenerateColumns="false" CssClass="table table-striped" DataKeyNames="UniqueUri" ClientIDMode="Static" runat="server">
                        <Columns>
                            <asp:TemplateField>
                                <HeaderTemplate>
                                    <asp:CheckBox ID="SelectAllCheckBox" CssClass="selectAllCheckBox" runat="server" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="SelectedCheckBox" CssClass="selectableCheckBox" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="From" HeaderText="From" />
                            <asp:BoundField DataField="Subject" HeaderText="Subject" />
                            <asp:BoundField DataField="Date" HeaderText="Date" />
                        </Columns>
                    </asp:GridView>

                    <div id="PagingDiv" runat="server" class="pagingOuterDiv " visible="false">
                        <asp:Panel Width="550px" ID="pager" ClientIDMode="Static" CssClass="tablesorterPager" runat="server">
                            <img src="/images/first.png" class="first" />
                            <img src="/images/prev.png" class="prev" />
                            <input type="text" class="pagedisplay" />
                            <img src="/images/next.png" class="next" />
                            <img src="/images/last.png" class="last" />
                            <div class="recordsToDisplay">
                                <label>Records to Display</label>
                                <select class="pagesize">
                                    <option value="5">5</option>
                                    <option value="10">10</option>
                                    <option selected="selected" value="30">20</option>
                                    <option value="40">30</option>
                                </select>
                            </div>
                        </asp:Panel>
                    </div>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
