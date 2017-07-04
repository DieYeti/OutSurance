<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1._Default" %>
<%@Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <table>
        <tr>
            <th>
                Upload File:
            </th>
        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="fuNewFile" runat="server" ClientIDMode="AutoID" />
            </td>
        </tr>
        <tr><td></td></tr>
        <tr>
            <td>
                <asp:Button ID="btnUpload" runat="server" Text="Generate Outputs" OnClick="btnUpload_Click" />
            </td>
        </tr>
        <tr><td></td></tr>
        <tr>
            <td>
                <asp:Label ID="lblMessage" runat="server" />
            </td>
        </tr>
        <tr><td></td></tr>
        <tr>
            <td>
                <asp:HyperLink ID="hlSurnames" runat="server" Text="Download Surnames File" Visible="false" />
            </td>
        </tr>
        <tr>
            <td>
                <asp:HyperLink ID="hlAddresses" runat="server" Text="Download Addresses File" Visible="false" />
            </td>
        </tr>
    </table>
</asp:Content>