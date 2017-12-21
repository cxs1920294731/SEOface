
Imports System.Xml
Imports System.Linq
Imports System.Configuration


Public Class GetSiteMap
    Public Shared Sub GetRootUrl(ByVal path As String)
        Dim parentNodePath As String = "/urlset/url"
        Dim urlNode As String = "url"
        Dim locNode As String = "loc"
        Dim lastmodNode As String = "lastmod"
        Dim changefreq As String = "changefreq"
        Dim priority As String = "priority"
        XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
        'XMLHelper.CreateOrUpdateXmlNodeByXPath(padPath, "/urlset", "url", "")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, "http://tsimshatsui.k11.com/")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "1.0")
        XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, "http://tsimshatsui.k11.com/AllShops")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
        XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "1.0")
    End Sub

    Public Shared Sub GetShopUrl(ByVal path As String)
        Dim parentNodePath As String = "/urlset/url"
        Dim urlNode As String = "url"
        Dim locNode As String = "loc"
        Dim lastmodNode As String = "lastmod"
        Dim changefreq As String = "changefreq"
        Dim priority As String = "priority"


        Dim listShop As New List(Of AutomationSite)
        Dim EF As New FaceBookForSEOEntities
        listShop = EF.AutomationSites.ToList()
        '店铺URL形如：http://fb.hk.k11.com/sitename=chowtaifook&siteid=109
        For Each shop As AutomationSite In listShop
            If Not (String.IsNullOrEmpty(shop.SiteName)) Then
                Dim url As String = "http://tsimshatsui.k11.com/" & UrlValid.getUrlValid(shop.SiteName.Trim) & "/" & shop.siteid
                XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, url)
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.8")

                Dim urltst As String = "http://tsimshatsui.k11.com/tst/" & UrlValid.getUrlValid(shop.SiteNameSc.Trim) & "/" & shop.siteid
                XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, urltst)
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.8")
            End If
        Next

        Dim listCate As List(Of CateTag) = EF.CateTags.ToList()
        For Each cate As CateTag In listCate
            If Not (String.IsNullOrEmpty(cate.HasShopID)) Then
                Dim url As String = "http://tsimshatsui.k11.com" & cate.CateUrl
                XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, url)
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.5")

                Dim urltst As String = "http://tsimshatsui.k11.com" & cate.CateUrlSc
                XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, urltst)
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.5")
            End If
        Next

        Dim listKeyWord As List(Of KeyWord) = EF.KeyWords.ToList()
        For Each keyItem As KeyWord In listKeyWord
            If (keyItem.Siteid > 0) Then
                Dim url As String = "http://tsimshatsui.k11.com" & keyItem.KeyUrl
                XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, url)
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.5")
            End If
        Next
    End Sub

    Public Shared Sub GetPostUrl(ByVal path As String)
        Dim listCategory As New List(Of AutoProduct)
        Dim EF As New FaceBookForSEOEntities
        listCategory = EF.AutoProducts.ToList()
        '分类URL形如：http://12Mart.net/post.aspx?post={0}&id={1}
        For Each post As AutoProduct In listCategory
            If Not (String.IsNullOrEmpty(post.PictureAlt)) Then
                Try

                    Dim sitename As AutomationSite = EF.AutomationSites.FirstOrDefault(Function(a) a.siteid = post.SiteID)
                    Dim url As String = "http://tsimshatsui.k11.com/" & UrlValid.getUrlValid(sitename.SiteName) & "/post/" & post.ProdouctID & ".aspx"

                    Dim parentNodePath As String = "/urlset/url"
                    Dim urlNode As String = "url"
                    Dim locNode As String = "loc"
                    Dim lastmodNode As String = "lastmod"
                    Dim changefreq As String = "changefreq"
                    Dim priority As String = "priority"
                    XMLHelper.CreateXmlNodeByXPath(path, "/urlset", "url")
                    'XMLHelper.CreateOrUpdateXmlNodeByXPath(padPath, "/urlset", "url", "")
                    XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", locNode, url)
                    XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", lastmodNode, DateTime.Now.ToString("yyyy-MM-dd"))
                    XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", changefreq, "daily")
                    XMLHelper.CreateOrUpdateXmlNodeByXPath(path, parentNodePath, "/urlset/url", priority, "0.5")
                Catch ex As Exception
                    Common.LogText("sitemap error:" & post.ProdouctID)
                    Continue For
                End Try
            End If
        Next
    End Sub



    Public Shared Function SiteMap()

        '生成sitemap.xml文件
        Common.LogText("start to write sitemap.NowTime:")
        Dim filepath As String = ConfigurationManager.AppSettings("siteMapPath").ToString().Trim()
        XMLHelper.CreateXmlDocument(filepath, "urlset", "1.0", "utf-8", "yes")

        GetRootUrl(filepath)
        GetShopUrl(filepath)
        GetPostUrl(filepath)

        XMLHelper.CreateOrUpdateXmlAttributeByXPath(filepath, "urlset", "xmlns", "http://www.sitemaps.org/schemas/sitemap/0.9")
    End Function

End Class
