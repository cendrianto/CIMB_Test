<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Dashboard.aspx.cs" Inherits="CIMB_trial.Dashboard" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div class="row">
            <asp:FileUpload id="Fileupload" runat="server" accept="application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></asp:FileUpload>
            
        </div>
        <div class="row">
            <asp:Button ID="BtnImportExcel" runat="server" Text="Import Excel" ToolTip="Import" OnClick="BtnImportExcel_Click" /><br/>
        </div>
        <div class="row">
            <asp:Button ID="BtnUpload" runat="server" Text="Upload File" ToolTip="Upload" OnClick="BtnUpload_Click" /><br/>
        </div>
        <div class="row">
            <asp:Button ID="btnLogout" runat="server" Text="Log Out" ToolTip ="Log out" OnClick="btnLogout_Click" />
        </div>
        <div class="row">
            <asp:Label ID="lbMsg" runat="server" ForeColor="Red"></asp:Label>
        </div>
    </form>
</body>
</html>
