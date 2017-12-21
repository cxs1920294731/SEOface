<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="FacebookSEO._Default1" %>
<%@ Import Namespace="FacebookSEO" %>
<%@ Register TagPrefix="webdiyer" Namespace="Wuqi.Webdiyer" Assembly="AspNetPager, Version=7.4.5.0, Culture=neutral, PublicKeyToken=fb0a0fe055d40fd4" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>groupbuyer</title>
    <link rel="stylesheet" type="text/css" href="/Styles/K11style.css" />
    <meta name="description" content="Group Buyer 提供香港最新最Hit團購優惠，包括餐飲團購、美食團購、紅酒團購、戲票團購、優惠券團購、美容團購、化妝品團購、課程團購，團購力量全港最大！" /> 
    <meta name="keywords" content="Group Buy,團購,Group Buyer,團購家" /> 
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .paginator { font: 17px Arial, Helvetica, sans-serif;padding:20px 40px 20px 0; margin: 0px;}
        .paginator a {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;margin-right: 2px;color:#d9946e }
        .paginator a:visited {padding: 6px 12px; border: solid 1px #ddd; background: #fff; text-decoration: none;}
        .paginator .cpb {padding: 6px 12px;font-weight: bold; font-size: 16px;border:none}
        .paginator a:hover {color: #fff; background: #d4d4d4;border-color:#d4d4d4;text-decoration: none;}
    </style>
    <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=2.0" /><meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=2.0" />
    <!-- 新 Bootstrap 核心 CSS 文件 -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.4/css/bootstrap.min.css">
    <!-- 可选的Bootstrap主题文件（一般不用引入） -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.4/css/bootstrap-theme.min.css">
    <!-- jQuery文件。务必在bootstrap.min.js 之前引入 -->
    <script src="http://cdn.bootcss.com/jquery/1.11.3/jquery.min.js"></script>
    <!-- 最新的 Bootstrap 核心 JavaScript 文件 -->
    <script src="http://cdn.bootcss.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
    <link rel="stylesheet" type="text/css" href="/Styles/K11style.css" />
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="/Styles/fancybox.css" />
    <script type="text/javascript" src="/Scripts/jquery-1.10.2.min.js"></script>
    <script type ="text/javascript" src="/Scripts/jquery.wookmark.min.js"></script>
    <script type ="text/javascript" src="/Scripts/jquery.imagesloaded.js"></script>
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
        <script type="text/javascript">
            $(function () {
                var loadedImages = 0, // Counter for loaded images
                    handler = $('#tiles .Itemli'); // Get a reference to your grid items.
                
                // Prepare layout options.
                //container: $('#list'),
                //offset: 10,
                //itemWidth: 200
                var options = {
                    autoResize: true, // This will auto-update the layout when the browser window is resized.
                    container: $('#tiles'), // Optional, used for some extra CSS styling
                    offset: 20, // Optional, the distance between grid items
                    outerOffset: 3 ,// Optional, the distance to the containers border
                };
                $('#tiles').imagesLoaded(function () {
                    // Call the layout function.
                    handler.wookmark(options);
                }).progress(function (instance, image) {
                    // Update progress bar after each image load
                    loadedImages++;
                   // if (loadedImages == handler.length)
                        //$('.progress-bar').hide();
                   // else
                       // $('.progress-bar').width((loadedImages / handler.length * 100) + '%');
                });
            });
        </script>
    <div>
        <%--header--%>
        <%=header%>
        <%--header--%>
        <div style ="margin-left:50px;height:30px;">
            <%=pageTitle  %>
        </div>
        <div id="main" role="main">
            <div id="tiles">
            <asp:Repeater ID="ReItems" runat="server" ViewStateMode ="Disabled" >
                <HeaderTemplate>
                </HeaderTemplate>
                <ItemTemplate>
                    <div class ="Itemli">
                        <div style ="margin: 5px 5px 5px 5px;">
                            <a href='<%# UrlValid .getUrlValid (Convert.ToString (Eval ("k11siteurl")) ) %>'  style='color:#3e3e3e; text-decoration:none;font-weight :bold ;'>
                                <%#
                                    Eval("sitetitlename")
                                    %>
                            </a>
                            <br/>
                            <!--
                                <a class="iframe" href='<%# String.Format("Default.aspx?shopName={0}&siteid={2}", UrlValid.getUrlValid(Convert.ToString(Eval("sitename"))), UrlValid.getUrlValid(Convert.ToString(Eval("PictureAlt"))), Eval("ProdouctID")) %>' style ='text-decoration:none;'>
                                -->
                            <a class="iframe" href='<%# String.Format("/{0}/post/{2}.aspx", UrlValid.getUrlValid(Convert.ToString(Eval("sitename"))), UrlValid.getUrlValid(Convert.ToString(Eval("PictureAlt"))), Eval("ProdouctID")) %>' style ='text-decoration:none;'>
                                <img class ="img" src='<%# DataBinder.Eval(Container.DataItem, "Prodouct")%>' alt='<%# Eval("PictureAlt") %>'/>
                            </a>
                            <br /><br />
                            <a href='<%# String .Format("/{0}/post/{2}.aspx", UrlValid.getUrlValid(Convert.ToString(Eval("sitename"))), UrlValid.getUrlValid(Convert.ToString(Eval("PictureAlt"))), Eval("ProdouctID")) %>'  style='color:#3e3e3e; text-decoration:none;'>
                                <%# String.Format("{0}<br/>", Eval("PictureAlt"))%>
                            </a>
                            <div style ="height: 20px;"></div>
                        </div>
                    </div>
                </ItemTemplate>
                <FooterTemplate>
                </FooterTemplate>
            </asp:Repeater>
            </div>
        </div>
        <webdiyer:AspNetPager ID="AspNetPager1" runat="server" Width="100%" UrlPaging="true" pagesize="100" CssClass="paginator"
            HorizontalAlign="Center" EnableTheming="true" ShowPageIndexBox="Never" CurrentPageButtonClass="cpb" UrlPageIndexName="pageindex">
        </webdiyer:AspNetPager>
        <br/>
        <div style ="margin-left:50px;">
            <!--
                          <%=homePage%>
                -->
        </div>
        <div style="margin-left: 20px;">
            <span style="display: block; line-height: 25px; color: #333333;font-size: 14px;">
                <%
                    Dim keySiteid As Integer
                    Dim hottopic As String = ""
                    If (siteid <> "0") Then
                        If (Integer.TryParse(siteid, keySiteid)) Then
                            If (siteid = "-1") Then
                                keySiteid = msiteid
                                hottopic = ""
                            Else
                                hottopic = "本店"
                            End If
                            Response.Write("<div style='max-width:95%;'>")
                            Dim listKeyWords As New List(Of KeyWord)
                            If (prodType = "WB") Then
                                listKeyWords = (From key In entity.KeyWords
                                                Where key.Type = "tst" AndAlso key.Siteid = keySiteid
                                                Select key).ToList()
                                If (listKeyWords.Count > 0) Then
                                    Response.Write(hottopic & "热门话题：<br/>")
                                End If
                            Else
                                listKeyWords = (From key In entity.KeyWords
                                                Where key.Type <> "tst" AndAlso key.Siteid = keySiteid
                                                Select key).ToList()
                                If (listKeyWords.Count > 0) Then
                                    Response.Write(hottopic & "熱門話題：<br/>")
                                End If
                            End If

                            For Each itemKey As KeyWord In listKeyWords
                                Response.Write("<span><a href='" & itemKey.KeyUrl & "/" & itemKey.ID & "' style='color:#3e3e3e; text-decoration:none;font-size:13px;'> #" &
                                                            itemKey.KeyWord1 & "</a></span>")
                                Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
                            Next
                            Response.Write("</div>")

                        End If
                    End If
                    %>
            </span>
            <br/>
            <br />
            <span id="cao" style="display: block; line-height: 25px; color: #333333;font-size: 14px;">
                  <%
                      If (siteid = "0") Then '仅在首页显示分类tag
                          For Each cate As CateTag In parentCateTags
                              Dim cateid As Integer = cate.ID
                              Response.Write("<div style='float:left;width:30%;min-width:400px;'>")
                              If (prodType = "WB") Then
                                  Response.Write("<h1>" & cate.CateNameSC & "</h1><br/>")
                              Else
                                  Response.Write("<h1>" & cate.CateName & "</h1><br/>")
                              End If
                              chlidCateTags = (From c In entity.CateTags
                                               Where c.ParentID = cateid
                                               Select c).ToList()

                              For Each childcate As CateTag In chlidCateTags
                                  If (prodType = "WB") Then
                                      Response.Write("<span class='label label-warning'><a href='" & UrlValid.getUrlValid(childcate.CateUrlSc) & "' style='color:#3e3e3e; text-decoration:none;font-size:13px;'>" &
                                                      childcate.CateNameSC & "</a></span><br/>")
                                  Else
                                      Response.Write("<span class='label label-warning'><a href='" & UrlValid.getUrlValid(childcate.CateUrl) & "' style='color:#3e3e3e; text-decoration:none;font-size:13px;'>" &
                                                     childcate.CateName & "</a></span><br/>")
                                  End If

                                  Dim listshopid As String() = {}
                                  Common.LogText(childcate.ID.ToString())
                                  Common.LogText(childcate.HasShopID)
                                  If Not (childcate.HasShopID Is Nothing) Then
                                      listshopid = childcate.HasShopID.Split(",")
                                  End If

                                  For i As Integer = 0 To listshopid.Count - 1
                                      Common.LogText(listshopid(i))
                                      Dim mysiteid As Integer
                                      If (Integer.TryParse(listshopid(i), mysiteid)) Then
                                          Dim autosite As AutomationSite = (From a In entity.AutomationSites
                                                                            Where a.siteid = mysiteid
                                                                            Select a).FirstOrDefault()

                                          If Not (autosite Is Nothing) Then
                                              Dim writeString As String = ""
                                              If (prodType = "WB") Then
                                                  If (isAutoPlanExist(mysiteid, "HA")) Then '简体时仅展示出有微博的店铺
                                                      writeString = writeString & "<a href='" & UrlValid.getUrlValid(autosite.k11SiteUrlSC) & "' style='color:#3e3e3e; text-decoration:none;'>"
                                                      writeString = writeString & autosite.SiteNameSc
                                                      writeString = writeString & "</a>&nbsp; &nbsp; &nbsp;&nbsp;"
                                                  End If
                                              Else
                                                  If (isAutoPlanExist(mysiteid, "HO")) Then '繁体时仅展示出有fb的店铺
                                                      writeString = writeString & "<a href='" & UrlValid.getUrlValid(autosite.k11SiteUrl) & "' style='color:#3e3e3e; text-decoration:none;'>"
                                                      writeString = writeString & autosite.SiteName
                                                      writeString = writeString & "</a>&nbsp; &nbsp; &nbsp;&nbsp;"
                                                  End If
                                              End If
                                              Response.Write(writeString)
                                          End If
                                      End If
                                  Next
                                  Response.Write("</br>")
                              Next
                              Response.Write("</div>")
                              Response.Write("<div style='float:left;width:20px;'>&nbsp;&nbsp;</div>")
                          Next
                          Response.Write("<div style='clear:both;'></div>")
                      End If
                     %>
            </span>
            <br/>
            <br/>
            <h1>
                <!--
                <%=k11Description%>
                    -->
            </h1>
            <br/>
            <br/>
        </div>
        <%--footer--%>
        <%=footer %>
        <%--footer--%>
    </div>
</body>
</html>
