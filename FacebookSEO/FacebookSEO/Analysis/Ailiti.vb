
Imports HtmlAgilityPack

Public Class Ailiti

    Private efHelper As New EFHelper()

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
            ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        GetCategory(siteId) 'insert catagory information in table Catagories
        Dim categoryName As String = "ailiti"
        Dim pageUrl As String = "http://ailiti.tmall.com/p/rd221267.htm?spm=a1z10.1-b.w7519107-3274097132.17.TefRkl&scene=taobao_shop"
        GetBanner(siteId, IssueID, categoryName, planType, pageUrl)

        Dim allCategoryUrl As String = "http://ailiti.tmall.com/category.htm"
        Dim xinpinurlandSubject As List(Of String) = efHelper.GetXinpinUrlandSubject(allCategoryUrl)
        Dim xinpinUrl As String = xinpinurlandSubject.Item(0).Trim
        efHelper.LogText("xinpinUrl:" & xinpinUrl)
        Dim subject As String = xinpinurlandSubject.Item(1).Trim
        Dim section As String = "CA"
        Dim productCount As Integer = 8
        efHelper.FetchProduct(xinpinUrl, categoryName, section, planType, productCount, DateTime.Now, DateTime.Now, siteId, IssueID)

        efHelper.InsertIssueSubject(IssueID, "爱丽缇——" & subject)
    End Sub


    Public Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim myCategory As New Category
        myCategory.Category1 = "ailiti"
        myCategory.Description = "ailiti"
        myCategory.Url = ""
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String, ByVal pageUrl As String)
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-16).ToShortDateString()
        Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        ' Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='zui_area']/table/tr/td")  '/div/div/a/
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='zui_area']")  '2016.3.3hoyho根据网站修改此xpath
        Dim ad As New Ad()


        'Dim anode As HtmlNode = banners(1).SelectSingleNode("a")
        'If Not (anode Is Nothing) Then
        '    ad.Url = efHelper.AddHttpForAli(anode.GetAttributeValue("href", "").ToString.Trim())
        '    ad.PictureUrl = efHelper.AddHttpForAli(anode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim())
        'Else
        '    anode = banners(1).SelectSingleNode("img")
        '    If Not (anode Is Nothing) Then
        '        ad.Url = pageUrl
        '        ad.PictureUrl = efHelper.AddHttpForAli(anode.GetAttributeValue("data-ks-lazyload", "").ToString.Trim())
        '    End If
        'End If
        
        Try
            Dim anode As HtmlNode = banners(0).SelectSingleNode("a")

            ad.Url = efHelper.AddHttpForAli(anode.GetAttributeValue("href", "").ToString.Trim())
            ad.PictureUrl = efHelper.AddHttpForAli(anode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim())

        Catch ex As Exception
            Try
                Dim anode As HtmlNode = banners(0).SelectSingleNode("img")

                ad.Url = pageUrl
                ad.PictureUrl = efHelper.AddHttpForAli(anode.GetAttributeValue("data-ks-lazyload", "").ToString.Trim())
            Catch ex1 As Exception
                efHelper.Log(ex1)
            End Try
        End Try
        ' If Not (efHelper.isAdSent(siteId, ad.PictureUrl, startDate, endDate, planType, True)) AndAlso Not String.IsNullOrEmpty(ad.Url) Then '控制2周内获取的Ad不重复：
        ad.SiteID = siteId
        ad.Lastupdate = Now

        Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
        listAd = efHelper.GetListAd(siteId)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName) '从Catagories表中获取categoryId
        Dim adid As Integer = efHelper.InsertAd(ad, listAd, categoryId)
        efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)
        'End If
    End Sub
End Class
