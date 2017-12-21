Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack

Public Class Dressvenus

    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper
    Private listProductUrl As New List(Of String)
    Private domin As String = "http://www.dressve.com"
    Private imgDomin As String = "http://s.dressve.com"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                    ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)

        'Website analysis
        InsertCategories(siteId)
        Dim bannerCategoryName = "Dresses"
        GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
        GetNewArrivals(siteId, IssueID, planType)

        'Add Subject
        AddIssueSubject(IssueID, siteId, planType)
        'Add 
        Dim contactListArr() As String = {"dressvenus_新品自动推荐"}
        efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

    End Sub

    ''' <summary>
    ''' 往数据库中categories插入数据
    ''' </summary>
    ''' <param name="siteID">siteid </param>
    ''' <remarks></remarks>
    Public Sub InsertCategories(ByVal siteID As Integer)
        Try
            Dim queryCategory = From c In efContext.Categories
                                Where c.SiteID = siteID
                                Select c
            Dim listCategoryName As New HashSet(Of String)
            For Each q In queryCategory
                listCategoryName.Add(q.Category1)
            Next

            '添加五大分类 
            Dim nowTime As DateTime = Now
            Dim cateOuterwears As New Category()
            Dim cateDresses As New Category()
            Dim catePromShoes As New Category()
            Dim cateBoots As New Category()
            Dim cateSandals As New Category()
            Dim cateJewelry As New Category()
            Dim cate1 As New Category()
            Dim cate2 As New Category()
            Dim cateShoesZone As New Category()
            If Not listCategoryName.Contains("Shoeszone") Then
                cateShoesZone.Category1 = "Shoeszone"
                cateShoesZone.SiteID = siteID
                cateShoesZone.Url = "http://www.dressve.com/Fashion/Shoes-Zone-c1-101701/"
                cateShoesZone.LastUpdate = nowTime
                cateShoesZone.Description = "Shoeszone"
                cateShoesZone.ParentID = -1
                cateShoesZone.Gender = "M"
                efContext.AddToCategories(cateShoesZone)
            End If

            If Not listCategoryName.Contains("WomenTops") Then
                cateOuterwears.Category1 = "WomenTops"
                cateOuterwears.SiteID = siteID
                cateOuterwears.Url = "http://www.dressve.com/Newest/Tops-101704"
                cateOuterwears.LastUpdate = nowTime
                cateOuterwears.Description = "WomenTops"
                cateOuterwears.ParentID = -1
                cateOuterwears.Gender = "M"
                efContext.AddToCategories(cateOuterwears)
            Else
                cateOuterwears = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "WomenTops" And m.SiteID = siteID)
                cateOuterwears.Url = "http://www.dressve.com/Best-Selling/Tops-101704/"
            End If

            If Not listCategoryName.Contains("Dresses") Then
                cateDresses.Category1 = "Dresses"
                cateDresses.SiteID = siteID
                cateDresses.Url = "http://www.dressve.com/Newest/Dresses-101703/"
                cateDresses.LastUpdate = nowTime
                cateDresses.Description = "Dresses"
                cateDresses.ParentID = -1
                cateDresses.Gender = "M"
                efContext.AddToCategories(cateDresses)
            Else
                cateDresses = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Dresses" And m.SiteID = siteID)
                cateDresses.Url = "http://www.dressve.com/Best-Selling/Dresses-101703/"
            End If

            If Not listCategoryName.Contains("PromShoes") Then
                catePromShoes.Category1 = "PromShoes"
                catePromShoes.SiteID = siteID
                catePromShoes.Url = "http://www.dressve.com/Newest/Prom-Shoes-7646/"
                catePromShoes.LastUpdate = nowTime
                catePromShoes.Description = "PromShoes"
                catePromShoes.ParentID = -1
                catePromShoes.Gender = "M"
                efContext.AddToCategories(catePromShoes)
            Else
                catePromShoes = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "PromShoes" And m.SiteID = siteID)
                catePromShoes.Url = "http://www.dressve.com/Best-Selling/Prom-Shoes-7646/"
            End If

            If Not listCategoryName.Contains("Boots") Then
                cateBoots.Category1 = "Boots"
                cateBoots.SiteID = siteID
                cateBoots.Url = "http://www.dressve.com/Newest/Boots-101065/"
                cateBoots.LastUpdate = nowTime
                cateBoots.Description = "Boots"
                cateBoots.ParentID = -1
                cateBoots.Gender = "M"
                efContext.AddToCategories(cateBoots)
            End If

            If Not listCategoryName.Contains("Sandals") Then
                cateSandals.Category1 = "Sandals"
                cateSandals.SiteID = siteID
                cateSandals.Url = "http://www.dressve.com/Newest/Sandals-101066"
                cateSandals.LastUpdate = nowTime
                cateSandals.Description = "Sandals"
                cateSandals.ParentID = -1
                cateSandals.Gender = "M"
                efContext.AddToCategories(cateSandals)
            Else
                cateSandals = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Sandals" And m.SiteID = siteID)
                cateSandals.Url = "http://www.dressve.com/Best-Selling/Sandals-101066/"
            End If

            If Not listCategoryName.Contains("Swimwear") Then
                cateJewelry.Category1 = "Swimwear"
                cateJewelry.SiteID = siteID
                cateJewelry.Url = "http://www.dressve.com/Newest/Swimwear-101043/"
                cateJewelry.LastUpdate = nowTime
                cateJewelry.Description = "Swimwear"
                cateJewelry.ParentID = -1
                cateJewelry.Gender = "M"
                efContext.AddToCategories(cateJewelry)
            Else
                cateJewelry = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Swimwear" And m.SiteID = siteID)
                cateJewelry.Url = "http://www.dressve.com/Best-Selling/Swimwear-101043/"
            End If

            If Not listCategoryName.Contains("Outerwears") Then
                cate1.Category1 = "Outerwears"
                cate1.SiteID = siteID
                cate1.Url = "http://www.dressve.com/Best-Selling/Outerwears-101706/"
                cate1.LastUpdate = nowTime
                cate1.Description = "Outerwears"
                cate1.ParentID = -1
                cate1.Gender = "M"
                efContext.AddToCategories(cate1)
            Else
                cate1 = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Outerwears" And m.SiteID = siteID)
                cate1.Url = "http://www.dressve.com/Best-Selling/Outerwears-101706/"
            End If

            If Not listCategoryName.Contains("Pants") Then
                cate2.Category1 = "Pants"
                cate2.SiteID = siteID
                cate2.Url = "http://www.dressve.com/Best-Selling/Pants-101705/"
                cate2.LastUpdate = nowTime
                cate2.Description = "Pants"
                cate2.ParentID = -1
                cate2.Gender = "M"
                efContext.AddToCategories(cate2)
            Else
                cate2 = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Pants" And m.SiteID = siteID)
                cate2.Url = "http://www.dressve.com/Best-Selling/Pants-101705/"
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try

    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        Dim pageUrl = "http://www.dressve.com/"
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

            'pageUrl = "http://www.dressve.com/TimeShowPage.html"
            'Dim listAdId2 As List(Of Integer) = GetListAdids2(pageUrl, siteId, categoryId, updateTime, planType)
            'InsertAdsIssue(siteId, Issueid, listAdId2, iProIssueCount)
        Catch ex As Exception
            'Throw ex
        End Try
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)

        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-16).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='imgPlay']/div[@class='imgs']/ul/li/a")
        Try
            For counter As Integer = 0 To banners.Count - 1
                Dim Ad As Ad = New Ad()
                Try
                    Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                Ad.Url = efHelpre.addDominForUrl(domin, banners(counter).GetAttributeValue("href", "").ToString.Trim())
                Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("taikoo_lazy_src", "")
                'Ad.SizeHeight = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString)
                'Ad.SizeWidth = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString)
                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelpre.GetListAd(siteid)
                If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                    adid = efHelpre.InsertAd(Ad, listAd, categroyId)
                    listAdIds.Add(adid)
                End If

            Next
        Catch ex As Exception
        End Try

        Return listAdIds
    End Function

    Public Function GetListAdids2(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Try
            Dim listAdIds As New List(Of Integer)
            Dim adid As Integer
            Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
            Dim banners As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@id='ad']")
            Dim Ad As Ad = New Ad()

            Dim tempString As String = banners.GetAttributeValue("onclick", "").ToString.Trim()
            Dim index As Integer = tempString.IndexOf("'")
            'tempString = tempString.Substring(index + 1).Trim("'")
            Ad.Url = tempString.Substring(index + 1).Trim("'")
            If (String.IsNullOrEmpty(Ad.Url)) Then
                Ad.Url = "http://www.dressve.com/"
            End If

            'tempString = banners(0).GetAttributeValue("style", "").ToString
            'Ad.PictureUrl = banners(0).GetAttributeValue("style", "").ToString
            Dim index1 As Integer = banners.GetAttributeValue("style", "").ToString.IndexOf("(")
            Dim index2 As Integer = banners.GetAttributeValue("style", "").ToString.IndexOf(")")
            If Not (banners.GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).ToLower.Contains("http")) Then
                If (banners.GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).Trim.StartsWith("/")) Then
                    Ad.PictureUrl = "http://s.dressvenus.com" & banners.GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                Else
                    Ad.PictureUrl = "http://s.dressvenus.com/" & banners.GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                End If
            Else
                Ad.PictureUrl = banners.GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
            End If
            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime
            Ad.Type = "B"

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelpre.GetListAd(siteid)
            adid = efHelpre.InsertAd(Ad, listAd, categroyId)
            listAdIds.Add(adid)

            Return listAdIds
        Catch ex As Exception
        End Try
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

    Private Sub GetNewArrivals(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal planType As String)
        Dim WomenTopsUrl As String = ""
        Dim DressesUrl As String = ""
        Dim PromShoesUrl As String = ""
        Dim cate2Url As String = ""
        Dim cate1Url As String = ""
        '从数据库中获取五个分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
        Dim queryURL = From c In efContext.Categories
                       Where c.SiteID = siteId
                       Select c
        For Each q In queryURL
            Select Case q.Category1.Trim()
                Case "Outerwears"
                    cate1Url = q.Url
                Case "Pants"
                    cate2Url = q.Url
                Case "Shoeszone"
                    PromShoesUrl = q.Url
                Case "Dresses"
                    DressesUrl = q.Url
                Case "WomenTops"
                    WomenTopsUrl = q.Url
            End Select
        Next

        Dim maxProductCount As Integer = 36
        Dim iTakeCount As Integer = 16
        Dim iProIssueCount As Integer = 3
        Dim updateTime As DateTime = Now
        Try
            GetProducts(siteId, IssueId, WomenTopsUrl, "WomenTops", maxProductCount, iTakeCount, iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, DressesUrl, "Dresses", maxProductCount, iTakeCount, iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, PromShoesUrl, "Shoeszone", maxProductCount, iTakeCount, iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, cate2Url, "Pants", maxProductCount, iTakeCount, iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, cate1Url, "Outerwears", 28, iTakeCount, iProIssueCount, updateTime, planType)
        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, ByVal iMaxCount As Integer, ByVal iTakeCount As Integer, _
                            ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, iMaxCount, siteId, updateTime, categoryName)

            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

            Dim endDate As String = Now.AddDays(1).ToShortDateString()
            Dim startDate As String = Now.AddDays(-33).ToShortDateString()
            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)

                If (Not efHelpre.IsProductSent(siteId, li.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
                    listProductId.Add(returnId)
                End If
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iTakeCount, iProIssueCount, categoryName)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 获取某个分类URL的产品信息，
    ''' </summary>
    ''' <param name="pageUrl">特定分类的URL</param>
    ''' <param name="productCount"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListProducts(ByVal pageUrl As String, ByVal productCount As Integer, _
                                    ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='gallery_con']/dl") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            If Not (categoryName = "Shoeszone") Then
                For counter As Integer = 0 To productCount - 1
                    Dim imageliNodes As HtmlNode = productDivs(counter).SelectSingleNode("dt/a") 'dd/ul/li 图片部分
                    Dim textliNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("dd/ul/li") '文字部分

                    Dim product As New Product()
                    Dim productNode As HtmlNode = textliNodes(0).SelectSingleNode("span[@class='tidely']/a")
                    Dim productNode2 As HtmlNode = textliNodes(1).SelectSingleNode("span[@class='h20']")
                    product.Prodouct = productNode.InnerText
                    product.Url = efHelpre.addDominForUrl(domin, productNode.GetAttributeValue("href", ""))
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If
                    'product.Discount =  '获取不到，所以不对其进行赋值
                    Dim pri As String() = productNode2.InnerText.Trim().ToString().Split(";")
                    product.Discount = Double.Parse(pri(1).ToString())
                    product.Currency = productNode2.InnerText.Trim().ToString().Substring(0, 3)

                    '特殊图片处理
                    product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("taikoo_lazy_src", "")
                    If (product.PictureUrl.Contains("loading.gif")) Then
                        product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("data-src", "")
                    End If

                    product.LastUpdate = nowTime
                    product.Description = productNode.InnerText
                    product.SiteID = siteId
                    product.PictureAlt = imageliNodes.SelectSingleNode("img").GetAttributeValue("alt", "")

                    listProducts.Add(product)
                Next
            Else 'Shoeszone
                productDivs = hd.DocumentNode.SelectNodes("//div[@id='right']/div[@class='garrery_01']/div[@class='gallery_con']/dl")
                For counter As Integer = 0 To productCount - 1
                    Dim imageliNodes As HtmlNode = productDivs(counter).SelectSingleNode("dt/a") 'dd/ul/li 图片部分
                    Dim textliNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("dd/ul/li") '文字部分

                    Dim product As New Product()
                    Dim productNode As HtmlNode = textliNodes(0).SelectSingleNode("a")
                    Dim productNode2 As HtmlNode = textliNodes(1).SelectSingleNode("span[@class='h20']") '/span[@class='multiPrice']
                    product.Prodouct = productNode.InnerText
                    product.Url = "http://www.dressve.com" + productNode.GetAttributeValue("href", "").TrimStart()
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If
                    'product.Discount =  '获取不到，所以不对其进行赋值
                    product.Discount = Double.Parse(productNode2.SelectSingleNode("span[@class='multiPrice']").InnerText.Trim)
                    'Try
                    '    Dim pri As String = productNode2.InnerText.Trim().ToString()
                    '    product.Discount = Double.Parse(pri.Substring(3, pri.Length - 3).Trim())
                    'Catch es As Exception
                    '    Dim pri As String() = productNode2.InnerText.Trim().ToString().Split(";")
                    '    product.Discount = Double.Parse(pri(1).ToString())
                    'End Try
                    product.Currency = "$"

                    '特殊图片处理
                    product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("taikoo_lazy_src", "")
                    If (imageliNodes.SelectSingleNode("img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                        product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("data-src", "")
                    End If

                    product.LastUpdate = nowTime
                    product.Description = productNode.InnerText
                    product.SiteID = siteId
                    product.PictureAlt = imageliNodes.SelectSingleNode("img").GetAttributeValue("alt", "")

                    listProducts.Add(product)
                Next
            End If

        Catch ex As Exception
            ' LogText(ex.ToString())
        End Try
        Return listProducts
    End Function

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer), _
                                    ByVal iTakeCount As Integer, ByVal iProIssueCount As Integer, ByVal categoryName As String)

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


    ''' <summary>
    ''' 获取某个店铺的一个CategoryId
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCategoryId(ByVal siteId As Integer, ByVal categoryName As String) As Integer
        Dim queryCategory = From c In efContext.Categories
                                  Where c.SiteID = siteId AndAlso c.Category1.Trim.ToUpper = categoryName.ToUpper
                                  Select c
        Dim categoryId As Integer = 0
        For Each q In queryCategory
            categoryId = q.CategoryID
        Next
        Return categoryId
    End Function

    Public Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal discount As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                          ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal categoryId As Integer,
                          ByVal list As List(Of Product)) As Integer
        url = url.Trim()
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = myProduct
        product.Url = url
        product.Discount = discount
        product.PictureUrl = pictureUrl
        product.LastUpdate = lastUpdate
        product.Description = description
        product.SiteID = siteId
        product.Currency = currency
        product.PictureAlt = pictureAlt

        If (JudgeProduct(product.Url, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Discount = product.Discount
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            'Dim returnId As Integer = updateProduct.ProdouctID
            'efContext.SaveChanges()

            '2013/4/17新增，以防止在填充模板时，产品丢失
            Dim updateCategory = updateProduct.Categories
            If Not (updateCategory.Contains(category)) Then
                category.Products.Add(updateProduct) '在查询分类时，添加一条关系在ProductCategory表中
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
    End Function

    ''' <summary>
    ''' 判断即将插入的数据URL是否在数据库中已经存在，如果存在，返回false
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function JudgeProduct(ByVal url As String, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url.Trim() = url.Trim()) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 将Products表中需要匹配列的信息添加到list中
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductList(ByVal siteId As Integer) As List(Of Product)
        '将Ads表中需要匹配的字段写入list中
        Dim query = From p In efContext.Products
                   Where p.SiteID = siteId
                   Select p
        Dim listProduct As New List(Of Product)
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Discount = q.Discount, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next
        Return listProduct
    End Function

    ''' <summary>
    ''' 获得邮件的subject
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String)

        Dim queryLastSubject = From iss In efContext.Issues
                             Where iss.IssueID < issueId AndAlso iss.SiteID = siteId AndAlso iss.PlanType = planType
                             Order By iss.IssueID Descending
                             Select iss
        Dim listLastSubject As New HashSet(Of String)
        Dim counter As Integer = 0
        For Each query In queryLastSubject
            If Not (String.IsNullOrEmpty(query.Subject)) AndAlso counter < 10 Then
                listLastSubject.Add(query.Subject)
                counter = counter + 1
            End If
            If (counter >= 10) Then
                Exit For
            End If
        Next
        Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct
        Dim issueSubject As String = ""
        For Each mySubject In querySubject
            Dim subject As String = ""
            subject = "Weekly New Arrival:" & mySubject
            If Not (listLastSubject.Contains(subject)) Then
                issueSubject = subject
                Exit For
            End If
        Next
        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = issueSubject
        efContext.SaveChanges()
    End Sub

    Private Sub InsertContactList(ByVal spreaLogin As String, ByVal appId As String, ByVal siteId As Integer, _
                                  ByVal planType As String, ByVal issueId As Integer, ByVal listName As String, _
                                  ByVal splitContactCount As Integer)
        Dim querySubscriber As New QuerySubscriber()

        querySubscriber.CountryList = New String() {}
        Dim criteria As String = querySubscriber.ToJsonString()
        Dim topN As Integer = Integer.MaxValue
        Dim saveAsList As String = ""

        If (splitContactCount >= 2) Then 'splitContactCount>=2
           
        Else 'splitContactCount=0或1
            Dim forceCreate As Boolean = True
            Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()

            '2013/4/11新增返回结果判断，如果没有创建联系人列表失败，则不创建campaign
            Dim returnResult As Integer = 0
            Try
                mySpread.Timeout = 600000
                returnResult = mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
            Catch ex As Exception
                mySpread.Timeout = 1200000
                mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
                LogText(ex.ToString())
            End Try
            If (returnResult > 0) Then
                Dim contactList As New ContactLists_Issue()
                contactList.IssueID = issueId
                contactList.ContactList = saveAsList
                efContext.AddToContactLists_Issue(contactList)
            End If

            Dim contact1 As New ContactLists_Issue()
            contact1.IssueID = issueId
            contact1.ContactList = "Opens"
            efContext.AddToContactLists_Issue(contact1)

            efContext.SaveChanges()
        End If
    End Sub

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
