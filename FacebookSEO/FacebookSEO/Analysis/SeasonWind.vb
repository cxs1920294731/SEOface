
Imports HtmlAgilityPack
Imports System.Net
Imports System.IO

Public Class SeasonWind
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        If planType.Trim = "HA" Then
            GetCategory(siteId)
            Dim bannerCategoryName As String = "新品"
            'GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            Dim pageUrl As String = "http://seasonwind.tmall.com/"
            'efHelper.GetBanner(siteId, IssueID, bannerCategoryName, planType, pageUrl, 3)
            efHelper.GetBanner(siteId, IssueID, bannerCategoryName, planType, "http://")
            GetNewProducts(siteId, IssueID, planType)
            'GetCategoryProducts(siteId, IssueID, "new", planType)

            'efHelper.InsertContactList(IssueID, "B顾客", "draft")
            Dim querySubject = (From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct).FirstOrDefault()
            If Not (querySubject Is Nothing) Then
                efHelper.InsertIssueSubject(IssueID, "亲爱的[FIRSTNAME]，本周季候风为您带来：" & querySubject.ToString)
            Else
                efHelper.InsertIssueSubject(IssueID, "亲爱的[FIRSTNAME]，来季候风遇见多彩的你")
            End If
        Else '个性化邮件
            Dim contactlistCount As Integer = 0

            Dim cate As String() = categories.Split("^")

            Dim categoryUrlC As String
            Dim categoryUrlL As String
            Dim iProIssueCount As Integer = 6

            Dim queryURL = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case cate(0).Trim
                        categoryUrlL = q.Url.Trim '& "?orderType=hotsell_desc"
                    Case cate(1).Trim
                        categoryUrlC = q.Url.Trim '& "?orderType=hotsell_desc"
                End Select
            Next
            efHelper.FetchProduct(categoryUrlL, cate(0), "CA", planType, 10, DateTime.Now.AddDays(-30), DateTime.Now, siteId, IssueID)
            efHelper.FetchProduct(categoryUrlC, cate(1), "CA", planType, 10, DateTime.Now.AddDays(-30), DateTime.Now, siteId, IssueID)
            'GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            'GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            Dim sendingStatus = "draft"
            Dim saveListName As String = "SeasonWindSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 90, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
            Dim saveListName1 As String = "SeasonWindSpecial" & cate(1).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(1).Trim, saveListName1, 90, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName1, sendingStatus)
            End If

            Dim query = From i In efContext.Issues
                     Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
            efHelper.InsertIssueSubject(IssueID, "Hi，亲爱的[FIRSTNAME]，本周特别推荐：" & cate(0).Trim & "/" & cate(1).Trim & "（第" & (query.Count + 1).ToString & "期）" & "(AD)")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "新品"
        myCategory.Description = "新品"
        myCategory.Url = "http://seasonwind.tmall.com/category-972469795.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory1 As New Category
        myCategory1.Category1 = "连衣裙"
        myCategory1.Description = "连衣裙"
        myCategory1.Url = "http://seasonwind.tmall.com/category-860761397.htm"
        myCategory1.SiteID = siteId
        myCategory1.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory1, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "衬衫"
        myCategory2.Description = "衬衫"
        myCategory2.Url = "http://seasonwind.tmall.com/category-860769628.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "T恤"
        myCategory3.Description = "T恤"
        myCategory3.Url = "http://seasonwind.tmall.com/category-860785969.htm"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "半身裙"
        myCategory4.Description = "半身裙"
        myCategory4.Url = "http://seasonwind.tmall.com/category-728518655.htm"
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory4, siteId)

        Dim myCategory5 As New Category
        myCategory5.Category1 = "裤装"
        myCategory5.Description = "裤装"
        myCategory5.Url = "http://seasonwind.tmall.com/category-374742989.htm"
        myCategory5.SiteID = siteId
        myCategory5.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory5, siteId)

        Dim myCategory6 As New Category
        myCategory6.Category1 = "外套"
        myCategory6.Description = "外套"
        myCategory6.Url = "http://seasonwind.tmall.com/category-534602147.htm"
        myCategory6.SiteID = siteId
        myCategory6.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory6, siteId)

        Dim myCategory7 As New Category
        myCategory7.Category1 = "上装"
        myCategory7.Description = "上装"
        myCategory7.Url = "http://seasonwind.tmall.com/category-374742987.htm"
        myCategory7.SiteID = siteId
        myCategory7.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory7, siteId)

    End Sub

    Private Sub GetNewProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String)
        Dim xinpinUrls As String = "http://seasonwind.tmall.com/category-972469795.htm"
        Dim categoryDoc As HtmlDocument
        Try
            categoryDoc = efHelper.GetHtmlDocByUrlTmall(xinpinUrls)
        Catch ex As Exception
            categoryDoc = efHelper.GetHtmlDocByUrlTmall(xinpinUrls)
        End Try
        'Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")
        Dim xinpinLi As HtmlNode = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")(0)
        Dim xinpinUrl As String = xinpinLi.SelectSingleNode("h4/a").GetAttributeValue("href", "").Trim()
        If Not (xinpinUrl.StartsWith("http")) Then
            If (xinpinUrl.StartsWith("//")) Then
                xinpinUrl = "http:" & xinpinUrl
            Else
                xinpinUrl = "http://" & xinpinUrl
            End If
        End If
        efHelper.FetchProduct(xinpinUrl, "新品", "CA", planType, 15, DateTime.Now, DateTime.Now, siteId, issueId)
        'GetProducts(siteId, issueId, xinpinUrl, "新品", 15, Now, planType)
    End Sub

    Public Sub GetCategoryProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim xinpinUrl As String = ""

        Dim queryURL = From c In efContext.Categories
                        Where c.SiteID = siteId
                        Select c
        For Each q In queryURL
            Select Case q.Category1.Trim()
                Case "新品"
                    xinpinUrl = q.Url.Trim()
            End Select
        Next
        'Dim maxProductCount As Integer = 34
        Dim iProIssueCount As Integer
        Dim updateTime As DateTime = Now
        iProIssueCount = 9
        GetProducts(siteId, Issueid, xinpinUrl, "新品", iProIssueCount, updateTime, planType)
    End Sub

    Public Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product)
            listProducts = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)
            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)

            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next

            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    Private Function GetListProducts(ByVal categoryUrl As String, ByVal siteId As Integer, ByVal lastUpdate As DateTime, _
                                          ByVal categoryname As String, ByVal planType As String, ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper()
        Dim index As Integer
        Try
            Dim categoryDoc As HtmlDocument
            Try
                categoryDoc = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            Catch ex As Exception
                categoryDoc = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            End Try
            Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")

            For Each productLine In productLines
                If (productLine.GetAttributeValue("class", "").ToString.ToLower() = "pagination") Then
                    Exit For
                End If
                Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
                If Not (productDls Is Nothing) Then
                    For Each dl In productDls

                        Dim myProduct As New Product
                        Dim productName As String = dl.SelectSingleNode("dd[@class='detail']/a").InnerText
                        If productName.Contains("运费") Or productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
                            Continue For
                        End If
                        myProduct.Prodouct = productName
                        myProduct.Url = dl.SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                        If (myProduct.Url.Contains("&rn")) Then
                            index = myProduct.Url.IndexOf("&rn")
                            myProduct.Url = myProduct.Url.Remove(index)
                        End If
                        If (listProductUrl.Contains(myProduct.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(myProduct.Url)
                        End If
                        myProduct.Url = efHelper.AddHttpForAli(myProduct.Url)
                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now, planType)) Then
                            myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                            myProduct.PictureUrl = efHelper.AddHttpForAli(myProduct.PictureUrl)
                            myProduct.Description = productName
                            myProduct.LastUpdate = lastUpdate
                            myProduct.SiteID = siteId
                            listProducts.Add(myProduct)
                        End If
                    Next
                End If
                If (listProducts.Count >= iProCount * 2) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when categoryUrl:" & categoryUrl & " categoryName:" & categoryname & ex.StackTrace)
        End Try
        Return listProducts
    End Function

    '专门用于获取新品/p/***的页面
    Private Function GetListNewProducts(ByVal categoryUrl As String, ByVal siteId As Integer, ByVal lastUpdate As DateTime, _
                                          ByVal categoryname As String, ByVal planType As String, ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper()
        Dim index As Integer
        Try
            Dim categoryDoc As HtmlDocument
            Try
                categoryDoc = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            Catch ex As Exception
                categoryDoc = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            End Try
            Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='J_TModule']/div/div/span/table/tbody/tr[1]/td/table/tbody")

            For Each productDetails In productLines
                Dim tr As HtmlNodeCollection = productDetails.SelectNodes("tr")
                Dim myProduct As New Product
                Dim productName As String = tr(3).InnerText
                If productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
                    Continue For
                End If
                myProduct.Prodouct = productName
                myProduct.Url = tr(2).SelectSingleNode("td/div/a").GetAttributeValue("href", "").Trim().Replace("&amp;", "&")
                If (myProduct.Url.Contains("&rn")) Then
                    index = myProduct.Url.IndexOf("&rn")
                    myProduct.Url = myProduct.Url.Remove(index)
                End If
                If (listProductUrl.Contains(myProduct.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(myProduct.Url)
                End If

                If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now, planType)) Then
                    myProduct.Discount = tr(5).InnerText.ToString.ToLower().Replace("rmb", "").Trim
                    myProduct.PictureUrl = tr(2).SelectSingleNode("td/div/a/img").GetAttributeValue("data-ks-lazyload", "")
                    myProduct.Description = productName
                    myProduct.LastUpdate = lastUpdate
                    myProduct.SiteID = siteId
                    listProducts.Add(myProduct)
                End If
                If (listProducts.Count >= iProCount * 2) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when categoryUrl:" & categoryUrl & " categoryName:" & categoryname & ex.StackTrace)
        End Try
        Return listProducts
    End Function

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://seasonwind.tmall.com/"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

        Catch ex As Exception
            LogText("GetPromotionBanner Error:" & ex.Message.ToString)
        End Try
    End Sub


    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-10).ToShortDateString()
        Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Dim index1 As Integer

        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='yinc']/div[@class='main']/div/a")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            Dim ad As New Ad()
            ad.Url = banners(counter).GetAttributeValue("href", "").ToString.Trim()
            If Not (ad.Url.ToLower.Contains("http://")) Then
                If (ad.Url.Trim().StartsWith("/")) Then
                    ad.Url = "http://seasonwind.tmall.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                Else
                    ad.Url = "http://seasonwind.tmall.com/" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                End If
            End If

            ad.PictureUrl = banners(counter).GetAttributeValue("style", "").ToString.Trim()
            If (ad.PictureUrl.Contains("(")) Then
                index1 = ad.PictureUrl.IndexOf("(")
                ad.PictureUrl = ad.PictureUrl.Substring(index1 + 1)
            End If
            If (ad.PictureUrl.Contains(")")) Then
                index1 = ad.PictureUrl.IndexOf(")")
                ad.PictureUrl = ad.PictureUrl.Substring(0, index1)
            End If

            ad.PictureUrl = efHelper.AddHttpForAli(ad.PictureUrl)

            If (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
                Continue For
            End If

            ad.SiteID = siteid
            ad.Lastupdate = Now
            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteid)

            adid = efHelper.InsertAd(ad, listAd, categroyId)
            listAdIds.Add(adid)
            If (listAdIds.Count = 1) Then
                Return listAdIds
            End If
        Next

        If (listAdIds.Count = 0) Then
            banners = hd.DocumentNode.SelectNodes("//div[@class='sldbox']/div[@class='bd']/div[@class='J_TWidget carousel']/div[@class='content']/ul/li/a")  '/div/div/a/
            For counter As Integer = 0 To banners.Count - 1
                Dim ad As New Ad()
                ad.Url = banners(counter).GetAttributeValue("href", "").ToString.Trim()
                If Not (ad.Url.ToLower.Contains("http://")) Then
                    If (ad.Url.Trim().StartsWith("/")) Then
                        ad.Url = "http://seasonwind.tmall.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                    Else
                        ad.Url = "http://seasonwind.tmall.com/" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                    End If
                End If

                ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim()
                If Not (ad.PictureUrl.ToLower.Contains("http://")) Then
                    If (ad.PictureUrl.Trim.StartsWith("/")) Then
                        ad.PictureUrl = "http://seasonwind.tmall.com" & banners(counter).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim()
                    Else
                        ad.PictureUrl = "http://seasonwind.tmall.com/" & banners(counter).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim()
                    End If
                End If

                If (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
                    Continue For
                End If

                ad.SiteID = siteid
                ad.Lastupdate = Now
                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)

                adid = efHelper.InsertAd(ad, listAd, categroyId)
                listAdIds.Add(adid)
                If (listAdIds.Count = 1) Then
                    Return listAdIds
                End If
            Next
        End If

        Return listAdIds
    End Function

    Public Function GetListLittleAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='skin-box-bd clear-fix']/span/table/tbody/tr/td")
        Try
            For counter As Integer = 0 To banners.Count - 1
                If (banners(counter).InnerHtml <> "") Then
                    Dim Ad As Ad = New Ad()

                    Ad.Url = banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                    If Not (Ad.Url.ToLower.Contains("http://")) Then
                        If (Ad.Url.Trim.StartsWith("/")) Then
                            Ad.Url = "http://seasonwind.tmall.com" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        Else
                            Ad.Url = "http://seasonwind.tmall.com/" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        End If
                    End If
                    If (Ad.Url = "") Then
                        Continue For
                    End If

                    Ad.PictureUrl = banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
                    Try
                        If Not (Ad.PictureUrl.ToLower().Contains("http://")) Then
                            If (Ad.PictureUrl.Trim.StartsWith("/")) Then
                                Ad.PictureUrl = "http://seasonwind.tmall.com" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
                            Else
                                Ad.PictureUrl = "http://seasonwind.tmall.com/" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
                            End If
                        End If
                    Catch ex As Exception
                        'ignore
                    End Try
                    Ad.Type = "B"
                    Ad.SiteID = siteid
                    Ad.Lastupdate = nowTime

                    Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                    listAd = efHelper.GetListAd(siteid)
                    If (Ad.PictureUrl <> "") Then
                        adid = InsertAd(Ad, listAd, categroyId)
                        listAdIds.Add(adid)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
        Return listAdIds
    End Function

    Public Sub InsertAdsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal listAdId As List(Of Integer), ByVal iProIssueCount As Integer)
        Dim i As Integer = 0
        For Each li In listAdId
            If i < iProIssueCount Then
                Dim aIssue As New Ads_Issue
                aIssue.AdId = li
                aIssue.SiteId = siteId
                aIssue.IssueID = issueId
                efContext.AddToAds_Issue(aIssue)
                i = i + 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
                                    ByVal iProIssueCount As Integer)

        Dim queryProductId = (From pro In efContext.Products_Issue
                           Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId
                           Select pro.ProductId).ToList() '同一个产品分属于不同的类别，避免同一个issue获取到相同的产品
        Dim i As Integer = 0
        For Each li In listProductId
            If i < iProIssueCount AndAlso Not (queryProductId.Contains(li)) Then
                Dim pIssue As New Products_Issue
                pIssue.ProductId = li
                pIssue.SiteId = siteId
                pIssue.IssueID = issueId
                pIssue.SectionID = sectionId
                efContext.AddToProducts_Issue(pIssue)
                i = i + 1
            End If
            If (queryProductId.Contains(li)) Then
                i = i - 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub

    Public Shared Function GetHtmlString(ByVal pageUrl As String) As String
        Dim ressting As String
        Try
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            request.Timeout = 120000
            'request.Headers.Add("Accept-Language", "zh-CN")
            ' request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch")
            request.Headers.Add("Accept-Language", "zh-CN")
            'request.Headers.Add("Host", "octlegendast.tmall.com")
            'request.Headers.Add("Connection", "keep-alive")
            request.Referer = "http://chuyu.tmall.com/?spm=a1z10.1.w5001-3242313717.3.mfNGN4&scene=taobao_shop"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()

            Dim resStreamReader As StreamReader = New StreamReader(resStream, System.Text.Encoding.GetEncoding("gb2312"))
            ressting = resStreamReader.ReadToEnd()
            ressting = ressting.Replace("\r", "")
            ressting = ressting.Replace("\n", "")
            ressting = ressting.Replace("\t", "")
            ressting = ressting.Replace("\", "")
            ressting = ressting.Replace("&nbsp;", "")
        Catch ex As Exception
            Throw New Exception("Request time out or bad url")
        End Try
        Return ressting
    End Function
    '判断重复条件为pictureURL
    Public Function InsertAd(ByVal ad As Ad, ByVal listAd As List(Of Ad), ByVal categoryID As Integer) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryID Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim newAd As New Ad()
        If Not (String.IsNullOrEmpty(ad.Ad1)) Then
            newAd.Ad1 = ad.Ad1
        End If
        If Not (String.IsNullOrEmpty(ad.Description)) Then
            newAd.Description = ad.Description
        End If
        If Not (String.IsNullOrEmpty(ad.Type)) Then
            newAd.Type = ad.Type
        End If
        newAd.Url = ad.Url
        newAd.PictureUrl = ad.PictureUrl
        newAd.SizeHeight = ad.SizeHeight
        newAd.SizeWidth = ad.SizeWidth
        newAd.SiteID = ad.SiteID
        newAd.Lastupdate = ad.Lastupdate


        If (JudgeAd(newAd.PictureUrl, listAd)) Then
            newAd.Categories.Add(category)
            efContext.AddToAds(newAd)
            efContext.SaveChanges()
            Return newAd.AdID
        Else
            Dim updateAd = efContext.Ads.FirstOrDefault(Function(m) m.PictureUrl = ad.PictureUrl)
            If Not (String.IsNullOrEmpty(ad.Ad1)) Then
                updateAd.Ad1 = ad.Ad1
            End If
            If Not (String.IsNullOrEmpty(ad.Description)) Then
                updateAd.Description = ad.Description
            End If
            If Not (String.IsNullOrEmpty(ad.Type)) Then
                updateAd.Type = ad.Type
            End If
            updateAd.Url = ad.Url
            updateAd.PictureUrl = ad.PictureUrl
            updateAd.SizeHeight = ad.SizeHeight
            updateAd.SizeWidth = ad.SizeWidth
            updateAd.SiteID = ad.SiteID
            updateAd.Lastupdate = ad.Lastupdate

            '2014/2/24新增，防止一个Banner有多个adCategory关系
            Dim updateCategory = updateAd.Categories

            'Dim queryCate = From p In efContext.Products
            '                Where p.ProdouctID = updateProduct.ProdouctID
            '                Select p
            'Dim cate = queryCate.Single.Categories
            Dim counter As Integer = updateCategory.Count
            For i As Integer = 0 To counter - 1
                updateCategory(0).Ads.Remove(updateAd)
            Next
            efContext.SaveChanges()

            If Not updateCategory.Contains(category) Then
                category.Ads.Add(updateAd)
            End If
            efContext.SaveChanges()

            Return updateAd.AdID
        End If
    End Function
    '判断重复条件为pictureURL
    Private Function JudgeAd(ByVal url As String, ByVal list As List(Of Ad)) As Boolean
        For Each li In list
            If (li.PictureUrl.Trim() = url.Trim()) Then
                Return False
            End If
        Next
        Return True
    End Function

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class

