Imports System
Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Text
Imports System.Text.RegularExpressions

Public Class TmallTaobaoSite
    Inherits EBusinessSite
    Dim efhelper As New EFHelper()

    Public Sub New(ByVal myshopUrl As String)
        MyBase.ShopUrl = StandardlizeShopURl(myshopUrl)

        If (ShopUrl.Contains("tmall.com")) Then
            ShopType = "tmall"
        ElseIf (ShopUrl.Contains("taobao.com")) Then
            ShopType = "taobao"
        End If

        SortedType.Add("新品", "/search.htm?orderType=newOn_desc")
        SortedType.Add("热销", "/search.htm?orderType=hotsell_desc")
        SortedType.Add("低价", "/search.htm?orderType=price_asc")
        SortedType.Add("收藏", "/search.htm?orderType=hotkeep_desc")
        If (ShopType = "tmall") Then
            SortedType.Add("人气", "/search.htm?orderType=koubei")
        Else
            SortedType.Add("人气", "/search.htm?orderType=coefp_desc")
        End If
        SortedType.Add("往期热门产品", "/")
        DefaultSubject = "[FIRSTNAME] 你好," & ShopName & "为你带来： [FIRST_PRODUCT]"
        DefaultTriggerSubject = "[FIRSTNAME] 您的专属资讯（第[VOL_NUMBER]期）"
        TemplateType = "P"
        BannerImgRegex = "(?:data-ks-lazyload=""|url\()(/(?:/[\w\d-_!.-]+)+?(?:.bmp|.jpg|.jpeg|.png|.gif))"
    End Sub

    Public Overrides Function GetProductList(ByVal pageUrl As String, ByVal siteId As Integer) As List(Of Product)
        Dim productList As List(Of Product) = efhelper.GetAsynProductList(pageUrl, 0)
        Return productList
    End Function

    Public Overrides Function GetShopNameandLogo() As String()
        Dim document As New HtmlDocument()
        Dim resultString(2) As String
        Try
            document = efhelper.GetHtmlDocByUrlTmall(ShopUrl)
            Dim shopName As String
            Dim shopNameNode As HtmlNode = document.DocumentNode.SelectSingleNode("//div[@class='hd-shop-name']/a[1]")
            shopName = shopNameNode.InnerText.Trim
            resultString(0) = shopName
        Catch ex As Exception
            efhelper.Log(ex)
            resultString(0) = -1
        End Try

        Try
            Dim shopLogoNode As HtmlNode = document.DocumentNode.SelectSingleNode("//div[@id='hd']")
            Dim shopLogoHtml As String = shopLogoNode.InnerHtml.Trim
            Dim shopLogo As String = RegularExpressions.Regex.Match(shopLogoHtml, "<img.+?src=[\""'](.+?)[\""'].+?>").Groups(1).Value
            If (shopLogo = "http://a.tbcdn.cn/s.gif") Then 'tmall店中会出现真正的图片属性为data-ks-lazyload
                shopLogo = RegularExpressions.Regex.Match(shopLogoHtml, "<img.+?data-ks-lazyload=[\""'](.+?)[\""'].+?>").Groups(1).Value
            End If

            If (String.IsNullOrEmpty(shopLogo)) Then
                shopLogo = RegularExpressions.Regex.Match(shopLogoHtml, "url\([\w\W]+?\)").Value
                shopLogo = shopLogo.Replace("url(", "").Replace(")", "")
            End If
            shopLogo = EFHelper.AddHttpForAli(shopLogo)
            If shopLogo = "http://assets.alicdn.com/s.gif" Then '识别错误
                shopLogo = ""
            End If
            resultString(1) = shopLogo
        Catch ex As Exception
            efhelper.Log(ex)
            resultString(1) = "http://app.rspread.com/spreaderfiles/16577/171409/output/img/logo.jpg"
        End Try
        Return resultString
    End Function

    Public Overrides Function GetSortedTypeURl(ByVal cateName As String) As String
        cateName = cateName.Trim
        If (ShopUrl.EndsWith("/")) Then
            ShopUrl = ShopUrl.Substring(0, ShopUrl.Length - 1)
        End If

        Dim sortedcaterul As String
        If (SortedType.TryGetValue(cateName, sortedcaterul)) Then
            If Not (sortedcaterul.StartsWith("/")) Then
                sortedcaterul = "/" & sortedcaterul
            End If
            Return ShopUrl & sortedcaterul
        Else
            Return ""
        End If
        
    End Function

    ''' <summary>
    ''' 将店铺（目前淘宝/天猫）链接规范化
    ''' </summary>
    ''' <param name="shopUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function StandardlizeShopURl(ByVal shopUrl As String) As String
        shopUrl = shopUrl.ToLower()
        Dim index As Integer = shopUrl.IndexOf("?")
        If (index > 0) Then
            shopUrl = shopUrl.Substring(0, shopUrl.IndexOf("?"))
        End If

        shopUrl = shopUrl.Replace("/shop/view_shop.htm", "").Replace("/shop/view_shop.html", "")
        shopUrl = shopUrl.Replace("/index.htm", "").Replace("/index.html", "")

        If (shopUrl.EndsWith("/")) Then
            shopUrl = shopUrl.Substring(0, shopUrl.Length - 1)
        End If
        Return shopUrl
    End Function


End Class
