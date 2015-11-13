<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="PDFConvertor.ascx.cs" Inherits="ConvertToPdf.PDFConvertor.PDFConvertor" %>
<style>
    .buttons{
        margin-left:0px !important;
        margin-top:5px;
        width:150px;
        float:left;
        margin-right:5px;
    }
    .merge{
        width:305px;
    }
    .textbox{
        width:293px
    }
    .buttonBox {
        width: 319px;
    }
    .resultBox{
        float:left;
        clear:both;
    }
    
</style>
<div>
    <div><asp:Label ID="lblInstructions" runat="server" Text="Enter the name of the document library you wish to convert"></asp:Label></div>
    <div><asp:TextBox CssClass="textbox" ID="txtLibrary" runat="server"></asp:TextBox></div>
    <div class="buttonBox">
        <asp:Button ID="btnWordConvertToPDF" CssClass="buttons" runat="server" Text="Convert Word to PDF" OnClick="btnWordConvertToPDF_Click" />
        
    </div>
</div>
<div class="resultBox"><asp:Label ID="ltResult" Text="" runat="server" /></div>

