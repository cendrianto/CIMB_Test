<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="CIMB_trial.Login" %>

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
            <asp:TextBox ID="tbPwd" runat="server" Text="Password" TextMode="Password" ></asp:TextBox>
        </div>
        <div class="row">
            <asp:Button ID="btnLogin"  Width="100px" runat="server" Text="Login" OnClick="btnLogin_Click" ToolTip="Login" />
        </div>
        <div class="row">
            <asp:Button ID="btnRegister"  Width="100px" runat="server" Text="Register" OnClick="btnRegister_Click" ToolTip="Register" />
        </div>
        <div class="row">
            <asp:Label ID="lbMsg" runat="server" ForeColor="Red"></asp:Label>
        </div>

    </form>
</body>
</html>

