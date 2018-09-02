<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="DataProcessorVB._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>ASP.NET</h1>
        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS and JavaScript.</p>
        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
    </div>

    <div class="row">
        <div>        
            <asp:FileUpload ID="FileUpload1" runat="server" />          
            <asp:Button ID="btnImport" runat="server" Text="Import" OnClick="btnImport_Click"/>
            <br />          <asp:Label ID="Label1" runat="server" ForeColor="Green" />
            <br />          <asp:Label ID="Label2" runat="server" ForeColor="Red" />
            <br />          <asp:Label ID="lblError" runat="server" ForeColor="Red" />     

        </div>
    </div>

</asp:Content>
