<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TokenManage.aspx.vb" Inherits="FacebookSEO.TokenManage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style ="max-width :90%;">
    longTimeToken到期时间：<asp:TextBox ID="txtLongTokenExpriedin" runat="server"></asp:TextBox>
    <br /><br /><br />
    新的shortTimeToken:<asp:TextBox ID="txtShortToken" runat="server" 
            Width="644px"></asp:TextBox>
            <br /><br />
        <asp:Button ID="btnGetNewLongtoken" runat="server" Text="获取新的longTimeToken" />
        <br /><br />
        <asp:Label ID="lbllongTimeToken" runat="server" Text=""></asp:Label>
    </div>
    </form>
</body>
</html>
