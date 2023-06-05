<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Register.aspx.cs" Inherits="CIMB_trial.Register" %>

<!DOCTYPE html>


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server"><title>Login</title></head>
<body>
    <form id="LoginForm" runat="server">
        <div class="row">
            <asp:Label ID="lbUID" runat="server" Text="Username"></asp:Label>
            <asp:TextBox ID="tbUN" runat="server" ></asp:TextBox>      
        </div>
        <div class="row">
            <asp:Label ID="lbPwd" runat="server" Text="Password"></asp:Label>
            <asp:TextBox ID="tbPwd" runat="server" TextMode="Password" ></asp:TextBox>
        </div>
        <div class="row">
            <asp:Label ID="lbPwd2" runat="server" Text="Confim Password"></asp:Label>
            <asp:TextBox ID="tbPwd2" runat="server" TextMode="Password" ></asp:TextBox>
        </div>
        <div class="row">
            <asp:Button ID="btnConfirm"  Width="100px" runat="server" Text="Confirm" OnClick="BtnConfirm_Click"  ToolTip="Confirm" />
        </div>
        <div class="row">
            <asp:Button ID="BtnBackLogin"  Width="100px" runat="server" Text="Back To Login" OnClick="BtnBackLogin_Click" ToolTip="Back To Login" />
        </div>
        <div class="row">
            <asp:Label ID="lbMsg" runat="server" ForeColor="Red"></asp:Label>
        </div>

    </form>
</body>
</html>

