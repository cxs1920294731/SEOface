Imports HtmlAgilityPack
Imports System.Net
Imports System.IO

Public Class Nollmet

    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private domin As String = "http://nollmet.tmall.com/"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        GetCategory(siteId)
        Dim bannerCategoryName As String = "新品"
        GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
        GetNewProducts(siteId, IssueID, planType)
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "新品"
        myCategory.Description = "新品"
        myCategory.Url = "http://nollmet.tmall.com/category-313054830.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Private Sub GetNewProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String)
        Dim xinpinUrls As String = "http://nollmet.tmall.com/category-313054830.htm"
        Dim categoryDoc As HtmlDocument
        Try
            categoryDoc = efHelper.GetHtmlDocByUrlTmall(xinpinUrls)
        Catch ex As Exception
            categoryDoc = efHelper.GetHtmlDocByUrlTmall(xinpinUrls)
        End Try
        'Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")
        Dim xinpinLi As HtmlNode = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")(0)
        Dim xinpinUrl As String = xinpinLi.SelectSingleNode("h4/a").GetAttributeValue("href", "").Trim()
        Dim subject As String = xinpinLi.SelectSingleNode("h4/a").InnerText().Trim()
        AddIssueSubject(issueId, siteId, subject)
        xinpinUrl = efHelper.AddHttpForAli(xinpinUrl)
        efHelper.FetchProduct(xinpinUrl, "新品", "CA", planType, 16, DateTime.Now.AddDays(-30), DateTime.Now, siteId, issueId)
        'GetProducts(siteId, issueId, xinpinUrl, "新品", 12, Now, planType)
    End Sub

    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal subject As String)
        If (subject.Contains("】")) Then
            subject = subject.Substring(subject.IndexOf("】") + 1)
        End If
        efHelper.InsertIssueSubject(issueId, "诺力米特新品推荐:" & " 【" & subject & "】")
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
            efHelper.LogText(ex.ToString())
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
                        'If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now, planType)) Then
                        myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                        myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                        myProduct.PictureUrl = efHelper.AddHttpForAli(myProduct.PictureUrl)
                        myProduct.Description = productName
                        myProduct.LastUpdate = lastUpdate
                        myProduct.SiteID = siteId
                        listProducts.Add(myProduct)
                        'End If
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
        Dim pageUrl As String = "http://nollmet.tmall.com/p/rd855641.htm?spm=a1z10.3.w7519107-3258699832.16.vmsnav"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listAdId As New List(Of Integer)
            listAdId.Add(GetListAdids(pageUrl, siteId, categoryId, updateTime, planType))
            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)
        Catch ex As Exception
            efHelper.LogText("GetPromotionBanner Error:" & ex.Message.ToString)
        End Try
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As Integer

        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-16).ToShortDateString()
        Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='shop-custom abs']/div/div/table/tr/td")  '/div/div/a/
        Dim ad As New Ad()
        ad.Url = pageUrl
        ad.PictureUrl = efHelper.addDominForUrl(domin, banners(1).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim())
        If Not (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
            ad.SiteID = siteid
            ad.Lastupdate = Now
            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteid)
            adid = efHelper.InsertAd(ad, listAd, categroyId)
            Return adid
        End If

        pageUrl = "http://nollmet.tmall.com/category-313054830.htm"
        hd = efHelper.GetHtmlDocByUrlTmall(pageUrl)
        banners = hd.DocumentNode.SelectNodes("//div[@class='slide-content']/div[@class='ks-switchable-content']/div/a")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            ad.Url = efHelper.addDominForUrl(domin, banners(counter).GetAttributeValue("href", "").ToString.Trim())

            ad.PictureUrl = efHelper.AddHttpForAli(banners(counter).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim())

            If (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
                Continue For
            End If

            ad.SiteID = siteid
            ad.Lastupdate = Now
            If Not (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)
                adid = efHelper.InsertAd(ad, listAd, categroyId)
                Return adid
            End If
        Next
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


End Class
