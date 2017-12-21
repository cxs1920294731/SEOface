

Imports HtmlAgilityPack
Imports System.Text.RegularExpressions

Public Class Aotu

    Private efHelper As New EFHelper()
    Private Shared efContext As New EmailAlerterEntities()

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
            ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String,
            ByVal preSubject As String, ByVal productCount As Integer)
        'GetCategory(siteId)
        Dim cate As String = categories.Split("^")(0).Trim
        If (planType.Contains("HO") OrElse planType.Contains("HA")) Then

            Dim categoryName As String = "短袖T恤"
            Dim pageUrl As String = url
            'Dim bannerImgIndex As Integer = 4 '指定获取页面的第几张图片作为banner
            'GetBanner(siteId, IssueID, categoryName, planType, pageUrl, bannerImgIndex)
            efHelper.GetBanner(siteId, IssueID, categoryName, planType, "https://")

            Dim allCategoryUrl As String = "http://onttno.tmall.com/category-1043342838.htm"
            Dim xinpinurlandSubject As List(Of String) = efHelper.GetXinpinUrlandSubject(allCategoryUrl)
            Dim xinpinUrl As String = xinpinurlandSubject.Item(0).Trim
            Dim subject As String = xinpinurlandSubject.Item(1).Trim
            Dim section As String = "CA"

            efHelper.FetchProduct(xinpinUrl, "新品", section, planType, productCount, DateTime.Now, DateTime.Now, siteId, IssueID)

            efHelper.InsertIssueSubject(IssueID, "亲爱的[FIRSTNAME]，ONTTNO傲徒推荐 " & subject)
        ElseIf (planType.Contains("HP")) Then

            Dim myCate As Category = (From c In efContext.Categories
                                     Where c.Category1 = cate And c.SiteID = siteId
                                     Select c).FirstOrDefault()
            efHelper.FetchProduct(myCate.Url, cate, "CA", planType, productCount, DateTime.Now.AddDays(-30), DateTime.Now, siteId, IssueID)

            Dim contactlistCount As Integer
            Dim saveListName1 As String = cate & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate, saveListName1, 90, ChooseStrategy.Favorite, "draft", spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName1, "draft")
            End If
            
            efHelper.HandleSubject(IssueID, siteId, "ONTTNO傲徒", planType, preSubject, cate, "CA")
            'efHelper.InsertIssueSubject(IssueID, "亲爱的[FIRSTNAME]，ONTTNO傲徒全心推荐：" & cate & "（第" & (query.Count + 1).ToString.PadLeft(2, "0") & "）期")
        End If
    End Sub


    Public Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim myCategory As New Category
        myCategory.Category1 = "短袖T恤"
        myCategory.Description = "短袖T恤"
        myCategory.Url = "http://onttno.tmall.com/category-855953756.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory, siteId)
        '短袖T恤 + 短袖衬衫 + 休闲裤 + 西装会场
        Dim myCategory2 As New Category
        myCategory2.Category1 = "短袖衬衫"
        myCategory2.Description = "短袖衬衫"
        myCategory2.Url = "http://onttno.tmall.com/category-855953757.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "休闲裤"
        myCategory3.Description = "休闲裤"
        myCategory3.Url = "http://onttno.tmall.com/category-855953758.htm"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "西装会场"
        myCategory4.Description = "西装会场"
        myCategory4.Url = "http://onttno.tmall.com/category-1029239340.htm"
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory4, siteId)

        Dim myCategory5 As New Category
        myCategory5.Category1 = "新品"
        myCategory5.Description = "新品"
        myCategory5.Url = "http://onttno.tmall.com"
        myCategory5.SiteID = siteId
        myCategory5.LastUpdate = lastUpdate
        efHelper.InsertOrUpdateCate(myCategory5, siteId)

    End Sub

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String,
                         ByVal pageUrl As String, ByVal bannerImgIndex As Integer)
        'Dim endDate As String = Now.AddDays(1).ToShortDateString()
        'Dim startDate As String = Now.AddDays(-16).ToShortDateString()
        'Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        'Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='zui_area']/table/tr/td")  '/div/div/a/
        'Dim ad As New Ad()
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
        Dim ad As New Ad()
        Dim imgRegex As String = "<img.*?(?:>|\/>)"
        Dim txtDocHtml As String = efHelper.GetHtmlStringByUrlAli(pageUrl)
        Dim mCollection As MatchCollection = Regex.Matches(txtDocHtml, imgRegex)
        Dim matched As Match = mCollection(bannerImgIndex - 1)
        If Not (matched Is Nothing) Then
            Dim bannerImg As String = mCollection(bannerImgIndex - 1).Value
            Dim srcRegex As String = "data-ks-lazyload=[\'\""]?([^\'\""]*)[\'\""]?"

            Dim srcMatched As Match = Regex.Match(bannerImg, srcRegex)
            If Not (srcMatched Is Nothing) Then
                ad.PictureUrl = srcMatched.Groups(1).Value.Trim

                Dim ahrefReg As String = "<a.*?href=""(.*?)"".*?(?:" & ad.PictureUrl & ").*?(?:>|\/>)"
                Dim ahrefMatched As Match = Regex.Match(txtDocHtml, ahrefReg)
                If Not (ahrefMatched Is Nothing) Then
                    ad.Url = ahrefMatched.Groups(1).Value.Trim
                Else
                    ad.Url = pageUrl
                End If
            End If
        End If

        If Not String.IsNullOrEmpty(ad.PictureUrl) Then '控制2周内获取的Ad不重复：
            ad.Url = efHelper.AddHttpForAli(ad.Url)
            ad.PictureUrl = efHelper.AddHttpForAli(ad.PictureUrl)
            ad.SiteID = siteId
            ad.Lastupdate = Now

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteId)
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim adid As Integer = efHelper.InsertAd(ad, listAd, categoryId)
            efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)
        End If
    End Sub


End Class
