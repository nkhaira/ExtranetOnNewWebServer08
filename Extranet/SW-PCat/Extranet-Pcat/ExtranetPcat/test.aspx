<%@ Page language="c#" Codebehind="test.aspx.cs" AutoEventWireup="false" Inherits="ExtranetPcat.test" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
    <HEAD>
        <META NAME="MainCategory" CONTENT="Products">
        <META NAME="PageType" CONTENT="Manuals">
        <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
        <meta content="C#" name="CODE_LANGUAGE">
        <meta content="JavaScript" name="vs_defaultClientScript">
        <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    </HEAD>
    <body>
        <form id="Form1" method="post" runat="server">
            <table cellSpacing="0" cellPadding="0" width="780" border="0">
                <tr>
                    <td>
                        <asp:datagrid id="dbAssets" runat="server" AutoGenerateColumns="False" PageSize="25" GridLines="None"
                            CellPadding="5" EnableViewState="False">
                            <AlternatingItemStyle CssClass="trAlternate"></AlternatingItemStyle>
                            <Columns>
                                <asp:TemplateColumn>
                                    <HeaderStyle CssClass="tableHead"></HeaderStyle>
                                    <ItemTemplate>
                                        <b>
                                            <asp:label id="lblDemoName" runat="server" text='<%# DataBinder.Eval(Container.DataItem, "AssetName") %>'/></b><br>
                                        <asp:label id="lblDescription" runat="server" text='<%# DataBinder.Eval(Container.DataItem, "Description") %>'/>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn>
                                    <HeaderStyle CssClass="tableHead"></HeaderStyle>
                                    <ItemStyle Width="10px"></ItemStyle>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn>
                                    <HeaderStyle CssClass="tableHead"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:HyperLink ID="lkDemo" Runat="server" Target="_blank" NavigateUrl='<%# DataBinder.Eval(Container.DataItem, "OracleID")%>'>
                                            <%# LinkTitle + " (" + DataBinder.Eval(Container.DataItem, "FileSize") + ")"%>
                                        </asp:HyperLink><br>
                                        <asp:label id="Label2" runat="server" text='' />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:datagrid>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</HTML>
