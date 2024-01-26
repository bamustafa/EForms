<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SubmitExcel.aspx.cs" Inherits="SubmitExcel" EnableSessionState="True" Debug="true" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Import Namespace="Microsoft.SharePoint.WebControls" %> 
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C" %>
<%@ Import Namespace ="System.Net"  %>
<%@ Import Namespace ="System.IO"  %>

<asp:Content ID="Content4" ContentPlaceHolderId="PlaceHolderTitleBreadcrumb" runat="server">
    <asp:SiteMapPath SiteMapProvider="SPContentMapProvider" id="ContentMap" SkipLinkText="" runat="server"/>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderId="PlaceHolderPageTitleInTitleArea" runat="server">
<asp:Label runat="server" ID="Label1" Visible = "true">Submit Excel</asp:Label>

</asp:Content>

<asp:Content ID="Content1" ContentPlaceHolderId="PlaceHolderMain" runat="server">
   <asp:Label runat="server" ID="MsgLabel" Visible = "true"></asp:Label>
   <br/>
   <br/>
   <br/>
   <asp:Button runat="server" ID="LinkButton1" Text="Ok"  Width="112px" OnClick="LinkButton1_Click" />
   
</asp:Content>


