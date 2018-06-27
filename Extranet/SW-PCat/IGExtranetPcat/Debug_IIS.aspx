<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Debug_IIS.aspx.cs" Inherits="Fluke_Debug.Debug_IIS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
 <div>
        <table width="100%">
            <tr> 
                <td colspan="2" align="center" class="Text">
                    
                    This page will display List of All the Directories in IIS. 
                </td>
            </tr>
             <tr>
                <td colspan="2">
                    &nbsp;
                </td>
            </tr>
            <tr> 
                <td colspan="2">                    
                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
             <tr>
                <td style="width:50%" class="Text" align="right">
                    Server Name :&nbsp;
                </td>
                <td>
                    <asp:TextBox id="txtServerName" runat="server">dtmevtsvfn05</asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtServerName"
                        ErrorMessage="Please enter server name.">Please enter server name.</asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="Text" align="right">
                    Directory Name :&nbsp;
                </td>
                <td>
                    <asp:TextBox id="txtDirectoryName" runat="server">inetpub</asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <asp:Button ID="btnGetList" runat="server" Text="Get List of Directories" OnClick="btnGetList_Click" />
                </td>
                
            </tr>
            <tr>
                <td align="right" class="Text" style="width: 472px">
                    IIS Version :&nbsp;
                </td>
                <td>
                    <asp:RadioButton ID="rdbIISVersion5" runat="server" Checked="true" GroupName="IIS" Text="5.1 or previous" />
                    <asp:RadioButton ID="rdbIISVersion6" runat="server" GroupName="IIS" Text="6.0 or later" />
                </td>
            </tr>
            <tr> 
                <td colspan="2" class="Text">                    
                    <asp:Label ID="lblMessage" runat="server" ></asp:Label>
                </td>
            </tr>
             <tr>
                <td colspan="2">
                    &nbsp;
                </td>
            </tr>
             <tr>
                <td colspan="2">
                    <asp:GridView ID="gvListOfSites" runat="server" AutoGenerateColumns="false" Width="100%">
                        <Columns>   
                            <asp:BoundField HeaderText="Site Name" DataField="SiteName"   />
                            <asp:BoundField HeaderText="IIS Path" DataField="IISPath" HeaderStyle-Width="20%" />
                            <asp:BoundField HeaderText="Has Browse Access" DataField="HasBrowseAccess" />
                            <asp:BoundField HeaderText="Has Execute Access" DataField="HasExecuteAccess" />
                            <asp:BoundField HeaderText="Has Read Access" DataField="HasReadAccess" />
                            <asp:BoundField HeaderText="Has Write Access" DataField="HasWriteAccess" />
                            <asp:BoundField HeaderText="Is Anonymous Access Allow" DataField="IsAnonymousAccessAllow" />
                            <asp:BoundField HeaderText="Is Basic Authentication Set" DataField="IsBasicAuthenticationSet" />
                            <asp:BoundField HeaderText="Is NTLM Authentication Set" DataField="IsNTLMAuthenticationSet" />
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
        </table>
        
    </div>
    </form>
</body>
</html>
