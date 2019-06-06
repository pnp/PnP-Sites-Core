<%@ Page Language="C#" MasterPageFile="~masterurl/default.master" Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Taxonomy" Namespace="Microsoft.SharePoint.Taxonomy" Assembly="Microsoft.SharePoint.Taxonomy, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content ID="Content1" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
   <SharePoint:projectproperty property="Title" runat="server" />
</asp:Content>
<asp:Content ContentPlaceHolderID="PlaceHolderPageImage" runat="server">
    <img src="/_layouts/15/images/blank.gif?rev=23" width='1' height='1' alt="" /></asp:Content>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    <SharePoint:ProjectProperty Property="Title" runat="server" />
</asp:Content>
<asp:Content ContentPlaceHolderID="PlaceHolderTitleAreaClass" runat="server">
    <SharePoint:ProjectProperty Property="Title" runat="server" />
</asp:Content>
<asp:Content ID="Content5" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server" />
<asp:Content ContentPlaceHolderID="PlaceHolderSearchArea" runat="server">
    <SharePoint:DelegateControl runat="server"
        ControlId="SmallSearchInputBox" />
</asp:Content>
<asp:Content ContentPlaceHolderID="PlaceHolderLeftActions" runat="server" />
<asp:Content ContentPlaceHolderID="PlaceHolderPageDescription" runat="server" />
<asp:Content ID="Content9" ContentPlaceHolderID="PlaceHolderBodyAreaClass" runat="server">
</asp:Content>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <table width="100%" cellpadding="0" cellspacing="0" id="project-home">
        <tr>
            <td valign="top" width="100%">
                <WebPartPages:WebPartZone runat="server" FrameType="TitleBarOnly" Title="<%$Resources:cms,WebPartZoneTitle_Right%>" ID="MainColumn" Orientation="Vertical" />
            </td>
            <td>&#160;</td>
        </tr>
    </table>
</asp:Content>