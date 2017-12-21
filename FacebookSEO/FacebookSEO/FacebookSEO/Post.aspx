<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Post.aspx.vb" Inherits="FacebookSEO.Post1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta name="description" content="" /> 
    <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=2.0" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="/Styles/K11style.css" />
    <style type="text/css">
        .paginator { font: 17px Arial, Helvetica, sans-serif;padding:20px 40px 20px 0; margin: 0px;}
        .paginator a {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;margin-right: 2px;color:#d9946e }
        .paginator a:visited {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;}
        .paginator .cpb {padding: 6px 12px;font-weight: bold; font-size: 16px;border:none}
        .paginator a:hover {color: #fff; background: #d4d4d4;border-color:#d4d4d4;text-decoration: none;}
    </style>
    <!-- 新 Bootstrap 核心 CSS 文件 -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.4/css/bootstrap.min.css">
    <!-- 可选的Bootstrap主题文件（一般不用引入） -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.4/css/bootstrap-theme.min.css">
    <!-- jQuery文件。务必在bootstrap.min.js 之前引入 -->
    <script src="http://cdn.bootcss.com/jquery/1.11.3/jquery.min.js"></script>
    <!-- 最新的 Bootstrap 核心 JavaScript 文件 -->
    <script src="http://cdn.bootcss.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
    <script type="text/javascript" src="/Scripts/jquery-1.4.1.min.js"></script>
    <script type="text/javascript" src="/Scripts/jquery-1.10.2.js"></script>
    <script type ="text/javascript">
         // GA tracking code
        (function (i, s, o, g, r, a, m) {
            i['GoogleAnalyticsObject'] = r; i[r] = i[r] || function () {
                (i[r].q = i[r].q || []).push(arguments);
            }, i[r].l = 1 * new Date(); a = s.createElement(o),
            m = s.getElementsByTagName(o)[0]; a.async = 1; a.src = g;
            m.parentNode.insertBefore(a, m);
        })(window, document, 'script', '//www.google-analytics.com/analytics.js', 'ga');
        ga('create', 'UA-49545219-1', 'auto');
        ga('send', 'pageview');
    </script>
</head>
<body>
    <%--header--%>
    <%=header%>
    <%--header--%>
    <div style ="margin-left:50px;">
        <%=pageTitle%> &nbsp; &gt; 
        <%= product.PictureAlt%>&nbsp; &nbsp; 
    </div>

        <div style ="margin-left :50px;">
            <a href ='<%=prePostUrl  %>'>
                <img id="img1"src='/images/pre.png' />
            </a>
            <a href ="#">
                <img id="img" src='<%=IIf(String.IsNullOrEmpty (product.PictureUrl),"",product.PictureUrl)%>'
                alt ='<%=IIf(String.IsNullOrEmpty (product.PictureAlt),"",product.PictureAlt)%>'/>
            </a>  
            <a href ='<%=nextPostUrl %>'>
                <img id="img2"src='/images/next.png'/>
            </a>
        </div>
        <div style ="height:10px;width:1px;"></div>
        <div style ="margin-left :50px;">
            <%=String.Format("{0}<br/><br/>{1}", IIf(String.IsNullOrEmpty(product.PictureAlt), "", product.PictureAlt),
            IIf(String.IsNullOrEmpty(product.Description), "", product.Description))%>
        </div>
        <div style ="margin-left :50px;width :500px">
            <br />
            <%
                Dim writeStr As String = ""
                If (product.Currency.Trim = "WB") Then
                    writeStr = "发布日期：" & product.ExpiredDate & " &nbsp;&nbsp; 来源：新浪微博 &nbsp;&nbsp; "
                    writeStr = writeStr + "<img src='/images/sinaweibologo.JPG'/><br/>"
                    writeStr = writeStr + "点击查看香港groupbuyer的<a href ='" & shopsite.shopUrl & "' style='color:#0664c1;text-decoration:none;'>" & shopsite.SiteNameSc & "</a>店铺信息。"
                Else
                    writeStr = "發佈日期：" & product.ExpiredDate & " &nbsp;&nbsp; 來源：Facebook &nbsp;&nbsp;"
                    writeStr = writeStr & "<img src='/images/facebooklogo.JPG'/><br/> "
                    writeStr = writeStr + "點擊查看香港groupbuyer的<a href ='" & shopsite.shopUrl & "' style='color:#0664c1;text-decoration:none;'>" & shopsite.SiteName & "</a>店鋪信息。"
                End If
                Response.Write(writeStr)
                %>
        </div>
        <div style ="height:50px;width:1px;"></div>
        <%--footer--%>
        <%=footer  %>
    <%--footer--%>
</body>
</html>
