<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="posts.aspx.vb" Inherits="FacebookSEO.posts" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <%--head--%>
    <%=head%>
     <%--head--%>
    <link rel="stylesheet" type="text/css" href="/Styles/fancybox.css" />
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="Scripts/jquery-1.4.1.min.js"></script>
    <script type="text/javascript" src="/Scripts/jquery.fancybox-1.3.1.pack.js"></script>
    <script type="text/javascript">
        $(function () {
            $(".iframe").fancybox({
                'width': '50%',
                'height': '75%',
                'autoScale': true
            });
        });
    </script>
    <style type="text/css">
        .paginator { font: 17px Arial, Helvetica, sans-serif;padding:20px 40px 20px 0; margin: 0px;}
        .paginator a {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;margin-right: 2px;color:#d9946e }
        .paginator a:visited {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;}
        .paginator .cpb {padding: 6px 12px;font-weight: bold; font-size: 16px;border:none}
        .paginator a:hover {color: #fff; background: #d4d4d4;border-color:#d4d4d4;text-decoration: none;}
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <%--header--%>
    <%=header%>
    <%--header--%>

    <div>
         <asp:Repeater ID="ReItems" runat="server" ViewStateMode ="Disabled" >
    <HeaderTemplate>
      <table align="center"  cellpadding="0" border="0" cellspacing="0" width="700">
        <tr>
    </HeaderTemplate>
    <ItemTemplate>
        <td >
            <div  class ="Itemli">
                <div style ="  width: 240px;margin: 5px 5px 5px 5px;">
                    
                    <a class="iframe" href='<%# String .Format("/post.aspx?siteid={0}&id={1}",siteid,Eval("ProdouctID")) %>'>
                        <img src='<%# DataBinder.Eval(Container.DataItem, "PictureUrl")%>' style ="max-width :240px;height :auto;margin-left :auto ;margin-right :auto ;" />
                    </a>
                    <br /><br />
                    
                    <%# String.Format("{0}<br/>", Eval("PictureAlt"))%>
                    
                    <div style ="height: 20px;"></div>
                </div>
            </div>
        </td>
    </ItemTemplate>
    <FooterTemplate>
        </tr></table>
    </FooterTemplate>
    </asp:Repeater>
    </div>
    
    <%--footer--%>
    <%=footer %>
    <%--footer--%>
    </div>
    </form>
</body>
</html>
