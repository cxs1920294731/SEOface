Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Public Class ToYouth
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private listModelProductId As New List(Of String)
    'Private preSubject As String
    Private isNew As Boolean = vbTrue '获取早遇上新banner还是爆款王牌banner的标志，如果isNew= true，则获取早遇上新；如果isNew= false，则获取爆款王牌

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        If planType.Trim = "HA" Then '新品
            GetCategory(siteId)
            Dim bannerCategoryName As String = "xinpin"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetModelProducts(siteId, IssueID, "new", planType)
            Dim pageUrl As String = "http://chuyu.tmall.com/search.htm?spm=a1z10.3-b.w4011-2654337891.19.6ebijO&catId=396797358&search=y&orderType=newOn_desc"
            efHelper.FetchProduct(pageUrl, bannerCategoryName, "CA", planType, 6, DateTime.Now.AddDays(-20), DateTime.Now, siteId, IssueID)
            'GetNewProducts(siteId, IssueID, planType)
            GetNewBanner(siteId, IssueID, planType)
            Dim contactListArr() As String = {"13年6月止无yahoo中国区邮箱", "聚200以下1无yahoo中国区邮箱", "集市店1无yahoo中国区邮箱", "测试组"}
            efHelper.InsertContactList(IssueID, contactListArr, "draft")

            Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "来自初语—— " & querySubject.FirstOrDefault().Trim & "(AD)")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "xinpin"
        myCategory.Description = "xinpin"
        myCategory.Url = "http://chuyu.tmall.com/widgetAsync.htm?ids=5600967151%2C6723199119%2C6723205092&path=%2Fp%2Fcjxp640247.htm&callback=callbackGetMods5600967151&site_instance_id=152197176"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory1 As New Category
        myCategory1.Category1 = "modelR"
        myCategory1.Description = "modelR"
        myCategory1.Url = "http://chuyu.tmall.com/widgetAsync.htm?ids=6076563651&path=%2Fshop%2Fview_shop.htm&callback=callbackGetMods6076563651&site_instance_id=152197176"
        myCategory1.SiteID = siteId
        myCategory1.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory1, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "modelL"
        myCategory2.Description = "modelL"
        myCategory2.Url = "http://chuyu.tmall.com/widgetAsync.htm?ids=6076563651&path=%2Fshop%2Fview_shop.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "modelM"
        myCategory3.Description = "modelM"
        myCategory3.Url = "http://chuyu.tmall.com/widgetAsync.htm"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "modelB"
        myCategory4.Description = "modelB"
        myCategory4.Url = "http://chuyu.tmall.com/category.htm?search=y&orderType=hotsell_desc"
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory4, siteId)
    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://chuyu.tmall.com/"
        Dim iProIssueCount As Integer = 5 '每次只获取一个Banner信息

        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
        Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
        InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

        Dim listAdId2 As List(Of Integer) = GetListLittleAdids(pageUrl, siteId, categoryId, updateTime, planType)
        If (listAdId2.Count = 0) Then
            listAdId2 = GetListLittleAdids2(pageUrl, siteId, categoryId, updateTime, planType)
        End If
        If (listAdId2.Count > 2) Then
            InsertAdsIssue(siteId, Issueid, listAdId2, iProIssueCount)
        End If
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-10).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As New HtmlDocument '= efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Try
            hd = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Catch ex As Exception
            hd = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        End Try
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='J_TModule']/div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-self-defined']/div[@class='skin-box-bd clear-fix']/span/div/div/div[@class='ks-switchable-content']/div/a")
        For counter As Integer = 0 To banners.Count - 1
            Dim bannerUrl As String
            If Not (banners(counter).GetAttributeValue("href", "").ToString.ToLower.Contains("http://")) Then
                If (banners(counter).GetAttributeValue("href", "").ToString.Trim().Trim.StartsWith("/")) Then
                    bannerUrl = "http://chuyu.tmall.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                Else
                    bannerUrl = "http://chuyu.tmall.com/" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                End If
            Else
                bannerUrl = banners(counter).GetAttributeValue("href", "").ToString.Trim()
            End If

            If (efHelper.isAdSent(siteid, bannerUrl, startDate, endDate, planType)) Then '控制2周内获取的Ad不重复：
                Continue For
            End If

            Dim bannerImgUrl As HtmlNodeCollection = banners(counter).SelectNodes("img")
            If (bannerImgUrl.Count = 1) Then
                Dim i As Integer = 0
                Dim Ad As Ad = New Ad()
                Ad.Url = bannerUrl
                Try
                    
                    Ad.PictureUrl = bannerImgUrl(i).GetAttributeValue("data-ks-lazyload", "").ToString
                    Ad.PictureUrl = efHelper.AddHttpForAli(Ad.PictureUrl)
                Catch ex As Exception
                    'ignore
                End Try

                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)

                adid = InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)

            ElseIf (bannerImgUrl.Count = 2) Then
                For i As Integer = 0 To bannerImgUrl.Count - 1
                    Dim Ad As Ad = New Ad()
                    Ad.Url = bannerUrl
                    Try
                        Ad.PictureUrl = bannerImgUrl(i).GetAttributeValue("data-ks-lazyload", "").ToString
                        Ad.PictureUrl = efHelper.AddHttpForAli(Ad.PictureUrl)
                    Catch ex As Exception
                        'ignore
                    End Try

                    Ad.SiteID = siteid
                    Ad.Lastupdate = nowTime

                    Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                    listAd = efHelper.GetListAd(siteid)

                    adid = InsertAd(Ad, listAd, categroyId)
                    listAdIds.Add(adid)
                Next
            ElseIf (bannerImgUrl.Count = 3) Then
                Dim i As Integer = 1
                Dim Ad As Ad = New Ad()
                Ad.Url = bannerUrl
                Try
                    Ad.PictureUrl = bannerImgUrl(i).GetAttributeValue("data-ks-lazyload", "").ToString
                    Ad.PictureUrl = efHelper.AddHttpForAli(Ad.PictureUrl)
                Catch ex As Exception
                    'ignore
                End Try

                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)

                adid = InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
            ElseIf (bannerImgUrl.Count = 4) Then
                For i As Integer = 1 To bannerImgUrl.Count - 2
                    Dim Ad As Ad = New Ad()
                    Ad.Url = bannerUrl
                    Try
                        Ad.PictureUrl = bannerImgUrl(i).GetAttributeValue("data-ks-lazyload", "").ToString
                        Ad.PictureUrl = efHelper.AddHttpForAli(Ad.PictureUrl)
                    Catch ex As Exception
                        'ignore
                    End Try

                    Ad.SiteID = siteid
                    Ad.Lastupdate = nowTime

                    Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                    listAd = efHelper.GetListAd(siteid)

                    adid = InsertAd(Ad, listAd, categroyId)
                    listAdIds.Add(adid)
                Next
            End If
            Exit For '每次只获取一块banner，以防两个banner之间的部件填充时发生匹配混乱

        Next
        Return listAdIds
    End Function

    Public Function GetListLittleAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        pageUrl = "http://chuyu.tmall.com/widgetAsync.htm?ids=7041931527%2C5904810606%2C6550044835&path=%2Fshop%2Fview_shop.htm&callback=callbackGetMods7041931527&site_instance_id=152197176"
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim hdstring = GetHtmlString(pageUrl)
        Dim hd As New HtmlDocument   '= efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        hd.LoadHtml(hdstring)
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='tier10']/table/tbody/tr[2]/td")
        Try
            For counter As Integer = 0 To banners.Count - 1
                If (banners(counter).InnerHtml <> "") Then
                    Dim Ad As Ad = New Ad()

                    If Not (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.ToLower.Contains("http://")) Then
                        If (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim().Trim.StartsWith("/")) Then
                            Ad.Url = "http://chuyu.tmall.com" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        Else
                            Ad.Url = "http://chuyu.tmall.com/" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        End If
                    Else
                        Ad.Url = banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                    End If

                    If (Ad.Url = "") Then
                        Continue For
                    End If

                    Try
                        If Not (banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.ToLower().Contains("http://")) Then
                            If (banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.Trim.StartsWith("/")) Then
                                Ad.PictureUrl = "http://chuyu.tmall.com" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
                            Else
                                Ad.PictureUrl = "http://chuyu.tmall.com/" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
                            End If
                        Else
                            Ad.PictureUrl = banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
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

    Public Function GetListLittleAdids2(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        pageUrl = "http://chuyu.tmall.com/?spm=a1z10.1.w5001-3242313717.3.I34nwy&scene=taobao_shop#123"
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim hdstring = GetHtmlString(pageUrl)
        Dim hd As New HtmlDocument   '= efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        hd.LoadHtml(hdstring)
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='tier10']/table/tbody/tr[2]/td")
        Try
            For counter As Integer = 0 To 4
                If (banners(counter).InnerHtml <> "") Then
                    Dim Ad As Ad = New Ad()

                    If Not (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.ToLower.Contains("http://")) Then
                        If (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim().Trim.StartsWith("/")) Then
                            Ad.Url = "http://chuyu.tmall.com" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        Else
                            Ad.Url = "http://chuyu.tmall.com/" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                        End If
                    Else
                        Ad.Url = banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                    End If

                    If (Ad.Url = "") Then
                        Continue For
                    End If

                    Try
                        If Not (banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString.ToLower().Contains("http://")) Then
                            If (banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim.StartsWith("/")) Then
                                Ad.PictureUrl = "http://chuyu.tmall.com" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
                            Else
                                Ad.PictureUrl = "http://chuyu.tmall.com/" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
                            End If
                        Else
                            Ad.PictureUrl = banners(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").ToString
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

    Public Sub GetNewProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String)
        Dim pageUrl As String = "http://chuyu.tmall.com/category.htm?spm=a1z10.5-b.w4011-2654352007.128.RnlHMP&catId=396797358&search=y&orderType=newOn_desc"
        Dim hd As HtmlDocument
        Try
            hd = efHelper.GetAsynHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Catch ex As Exception
            hd = efHelper.GetAsynHtmlDocByUrlTmall(pageUrl)
        End Try
        Dim xinpin As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='attrValues']/ul/li")
        Dim xinpinStr As String = ""
        Dim index As Integer
        Dim xinpinTime As Date
        Dim timespan As TimeSpan
        Dim xinpinUrl As String
        For counter = 0 To xinpin.Count - 1
            'preSubject = xinpin(counter).InnerText.Trim
            xinpinStr = xinpin(counter).InnerText
            If (xinpinStr.Contains("新")) Then
                index = xinpinStr.IndexOf("新")
                'preSubject = xinpinStr.Substring(index + 1).Trim
                xinpinStr = Now.Year.ToString & "-" & xinpinStr.Remove(index).Trim.Replace(".", "-")

                xinpinTime = Format(Date.Parse(xinpinStr), "yyyy-MM-dd")
                timespan = Now - xinpinTime
                If (timespan < New TimeSpan(0, 0, 0, 0)) Then

                    timespan = Now.Date - xinpinTime.AddYears(-1)
                End If
                If (timespan > New TimeSpan(0, 0, 0, 0) AndAlso timespan < New TimeSpan(7, 0, 0, 0)) Then
                    xinpinUrl = xinpin(counter).SelectSingleNode("a").GetAttributeValue("href", "")
                    Exit For
                End If

            End If
        Next
        xinpinUrl = efHelper.AddHttpForAli(xinpinUrl)
        'Dim xinpinUrls As String = xinpin.GetAttributeValue("href", "").ToString.Trim
        GetProducts(siteId, issueId, xinpinUrl, "xinpin", 12, Now, "HO")
    End Sub
    

    Public Sub GetModelProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)

        Dim modelUrl As String = "http://chuyu.tmall.com/widgetAsync.htm?ids=6076563651&path=%2Fshop%2Fview_shop.htm&callback=callbackGetMods6076563651&site_instance_id=152197176"

        Try
            
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer = 10
            Dim updateTime As DateTime = Now

            Dim listProductId As List(Of Integer) = GetListModelProducts(modelUrl, siteId, updateTime, planType, iProIssueCount)

            InsertProductsIssue(siteId, Issueid, "CA", listProductId, iProIssueCount)

        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Public Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As New List(Of Product)
            If Not (String.IsNullOrEmpty(categoryUrl)) Then
                listProducts = efHelper.GetAsynProductList(categoryUrl, siteId) ' GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)
            End If

            If (listProducts.Count < 6) Then
                isNew = False
                categoryUrl = "http://chuyu.tmall.com/category.htm?search=y&orderType=hotsell_desc"
                listProducts = efHelper.GetAsynProductList(categoryUrl, siteId) 'GetListProducts(categoryUrl, siteId, Now, categoryName, planType, iProIssueCount)
            End If
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

    Public Sub GetNewBanner(ByVal siteId As Integer, ByVal issueId As Integer, ByVal plantype As String)
        Dim bannerProduct As New Product
        If (isNew) Then
            bannerProduct.Prodouct = "早遇上新"
            bannerProduct.Url = "http://chuyu.tmall.com/category-396797358.htm?orderType=newOn_desc&tsearch=y"
            bannerProduct.PictureUrl = "http://app.rspread.com/Spread5/spreaderfiles/29105/343585/toyouth03_html/images/toyouth03_06.jpg"
            bannerProduct.LastUpdate = Now
            bannerProduct.Description = bannerProduct.Prodouct
            bannerProduct.SiteID = siteId
            bannerProduct.PictureAlt = bannerProduct.Prodouct
        Else
            bannerProduct.Prodouct = "爆款王牌"
            bannerProduct.Url = "http://chuyu.tmall.com/category.htm?search=y&orderType=hotsell_desc"
            bannerProduct.PictureUrl = "http://app.rspread.com/Spread5/spreaderfiles/29105/343582/toyouth02_html/images/toyouth02_15.jpg"
            bannerProduct.LastUpdate = Now
            bannerProduct.Description = bannerProduct.Prodouct
            bannerProduct.SiteID = siteId
            bannerProduct.PictureAlt = bannerProduct.Prodouct
        End If
        Dim listProduct As New List(Of Product)
        listProduct = efHelper.GetProductList(siteId)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, "modelB")
        Dim returnId As Integer = efHelper.InsertProduct(bannerProduct, Now, categoryId, listProduct)
        listProductId.Add(returnId)
        InsertProductsIssue(siteId, issueId, "CA", listProductId, 1)
    End Sub

    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                  ByVal iProIssueCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim indexId As Integer
        Dim modelId As String
        Dim index As Integer
        Try
            Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")
            For counter = 0 To productDivs.Count - 1
                If (productDivs(counter).GetAttributeValue("class", "").ToString.ToLower() = "pagination") Then
                    Exit For
                End If

                Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("dl")
                For j = 0 To productDetails.Count - 1
                    Dim product As New Product()

                    product.Prodouct = productDetails(j).SelectSingleNode("dd[@class='detail']/a").InnerText
                    '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                    'If (product.Prodouct.Length > 70) Then
                    '    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                    'End If
                    If product.Prodouct.Contains("运费") Or product.Prodouct.Contains("差额") Or product.Prodouct.Contains("专拍") Or product.Prodouct.Contains("补拍") Or product.Prodouct.Contains("差价") Or product.Prodouct.Contains("補拍") Or product.Prodouct.Contains("鏈接") Or product.Prodouct.Contains("專拍") Or product.Prodouct.Contains("運費") Or product.Prodouct.Contains("链接") Then
                        Continue For
                    End If
                    product.Url = productDetails(j).SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If
                    If (product.Url.Contains("id")) Then
                        indexId = product.Url.IndexOf("id")
                        modelId = product.Url.Substring(indexId + 3, 11)
                        If (listModelProductId.Contains(modelId)) Then
                            Continue For
                        End If
                    End If
                    If (product.Url.Contains("&rn")) Then
                        index = product.Url.IndexOf("&rn")
                        product.Url = product.Url.Remove(index)
                    End If
                    product.Url = efHelper.AddHttpForAli(product.Url)
                    product.Discount = Double.Parse(productDetails(j).SelectSingleNode("dd[@class='detail']/div[@class='attribute']/div[@class='cprice-area']/span[@class='c-price']").InnerText.Trim())

                    product.Currency = "￥"

                    '特殊图片处理 pList2_mg
                    Dim productImage As HtmlNode = productDetails(j).SelectSingleNode("dt[@class='photo']")
                    product.PictureUrl = productImage.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "") '2013/4/7添加
                    product.PictureUrl = efHelper.AddHttpForAli(product.PictureUrl)

                    product.LastUpdate = nowTime
                    product.Description = product.Prodouct
                    product.SiteID = siteId
                    product.PictureAlt = product.Prodouct

                    If (Not isNew) Then
                        If (Not efHelper.IsProductSent(siteId, product.Url, Now.AddDays(-30), Now, planType)) Then '控制一个月内获取的新品不重复：
                            listProducts.Add(product)
                        End If
                    Else
                        listProducts.Add(product)
                    End If
                Next
                If (listProducts.Count = iProIssueCount * 2) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            LogText("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Function GetListModelProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal planType As String,
                                  ByVal iProIssueCount As Integer) As List(Of Integer)
        Dim listProductid As New List(Of Integer)
        Dim categoryName As String = "modelL"
        Dim listProduct As New List(Of Product)
        listProduct = efHelper.GetProductList(siteId)
        Try
            'Dim listAds As New List(Of Ad)
            Dim modelHtml As String = GetHtmlString(pageUrl) 'slideBox mt10 mb10
            'Dim hddstring As String = hd22  ToString.Replace("\", "")
            Dim index1 As Integer = modelHtml.IndexOf("<div")
            Dim index2 As Integer = modelHtml.LastIndexOf("div>")
            modelHtml = modelHtml.Substring(index1, index2 - index1 + 4)

            Dim rg As New Regex("<li\s*><a\s*[\s\S]*?/li>")
            Dim regurcollection As MatchCollection = rg.Matches(modelHtml)
            Dim productDetail As String
            Dim rehref As New Regex("href=""[\s\S]*?""")
            Dim resrc As New Regex("src=""[\s\S]*?""")
            Dim indexId As Integer = -1
            Dim modelId As String

            Dim categoryId As Integer
            For i As Integer = 0 To regurcollection.Count - 1
                Dim product As New Product
                productDetail = regurcollection.Item(i).Value.ToString()
                product.Url = rehref.Match(productDetail).Value.Replace("href=""", "").Replace("""", "")
                If (product.Url.Contains("id")) Then
                    indexId = product.Url.IndexOf("id")
                    modelId = product.Url.Substring(indexId + 3, 11)
                    listModelProductId.Add(modelId)
                End If
                product.PictureUrl = resrc.Match(productDetail).Value.Replace("src=""", "").Replace("""", "")
                product.LastUpdate = nowTime
                product.SiteID = siteId
                If (product.Url.Contains("chuyu.tmall.com/p")) Then
                    product.Description = "banner"
                End If
                If (i = 4) Then
                    categoryName = "modelM"
                ElseIf (i > 4) Then
                    categoryName = "modelR"
                End If
                categoryId = efHelper.GetCategoryId(siteId, categoryName)
                Dim returnId As Integer = efHelper.InsertProduct(product, Now, categoryId, listProduct)
                listProductid.Add(returnId)
            Next

        Catch ex As Exception
            efHelper.LogText(ex.StackTrace)
        End Try
        Return listProductid
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
