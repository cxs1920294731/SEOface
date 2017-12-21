Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions

Public Class Tbdress
    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper
    Private listProductUrl As New List(Of String)
    Private domin As String = "http://www.tbdress.com/"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                    ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String)
        InsertCategories(siteId)
        Dim bannerCategoryName = "Outerwears"
        GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
        GetProducts(siteId, IssueID, planType)
        'Add Subject
        AddIssueSubject(IssueID, siteId, planType)
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
            Dim cateShoes As New Category()
            Dim cateTops As New Category()

            If Not listCategoryName.Contains("Outerwears") Then
                cateOuterwears.Category1 = "Outerwears"
                cateOuterwears.SiteID = siteID
                cateOuterwears.Url = "http://www.tbdress.com/Cheap-Outerwears-100596/"
                cateOuterwears.LastUpdate = nowTime
                cateOuterwears.Description = "Outerwears"
                cateOuterwears.ParentID = -1
                cateOuterwears.Gender = "M"
                efContext.AddToCategories(cateOuterwears)
            End If

            If Not listCategoryName.Contains("WomenTops") Then
                cateTops.Category1 = "WomenTops"
                cateTops.SiteID = siteID
                cateTops.Url = "http://www.tbdress.com/Cheap-Women-Tops-100508/"
                cateTops.LastUpdate = nowTime
                cateTops.Description = "WomenTops"
                cateTops.ParentID = -1
                cateTops.Gender = "M"
                efContext.AddToCategories(cateTops)
            End If

            If Not listCategoryName.Contains("Dresses") Then
                cateDresses.Category1 = "Dresses"
                cateDresses.SiteID = siteID
                cateDresses.Url = "http://www.tbdress.com/Cheap-Women-Dresses-100507/"
                cateDresses.LastUpdate = nowTime
                cateDresses.Description = "Dresses"
                cateDresses.ParentID = -1
                cateDresses.Gender = "M"
                efContext.AddToCategories(cateDresses)
            End If

            If Not listCategoryName.Contains("Shoes") Then
                cateShoes.Category1 = "Shoes"
                cateShoes.SiteID = siteID
                cateShoes.Url = "http://www.tbdress.com/Cheap-Shoes-101913/"
                cateShoes.LastUpdate = nowTime
                cateShoes.Description = "Shoes"
                cateShoes.ParentID = -1
                cateShoes.Gender = "M"
                efContext.AddToCategories(cateShoes)
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            EFHelper.LogText(ex.Message.ToString)
        End Try

    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        Dim pageUrl = domin
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
        Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
        InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)

        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-16).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='right']/div[@class='banner']")
        Dim bannerScript As String = banners(0).InnerText
        Try
            Dim rg As New Regex("ID:[\w\W]*?\}")
            Dim regurcollection As MatchCollection = rg.Matches(bannerScript)
            Dim adDetail As String
            Dim resrc As New Regex("ImgUrl:""[\w\W]*?""")
            Dim rehref As New Regex("http://[\w\W]*?""")

            For i As Integer = 0 To regurcollection.Count - 1
                Dim ad As New Ad()
                adDetail = regurcollection.Item(i).Value.ToString()
                Dim bannerCollection As MatchCollection = rehref.Matches(adDetail)
                ad.PictureUrl = efHelpre.addDominForUrl(domin, resrc.Match(adDetail).Value.Replace("ImgUrl:""", "").Replace("""", ""))
                If (String.IsNullOrEmpty(ad.PictureUrl)) Then
                    Exit For
                End If
                ad.Url = efHelpre.addDominForUrl(domin, bannerCollection.Item(1).Value.ToString.Replace("""", ""))

                ad.Lastupdate = nowTime
                ad.SiteID = siteid

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelpre.GetListAd(siteid)
                If (Not efHelpre.isAdSent(siteid, ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                    adid = efHelpre.InsertAd(ad, listAd, categroyId)
                    listAdIds.Add(adid)
                    'If (listAdIds.Count = 1) Then
                    '    Exit For
                    'End If
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

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal planType As String)
        Dim WomenTopsUrl As String = ""
        Dim DressesUrl As String = ""
        Dim ShoesUrl As String = ""
        Dim OuterwearsUrl As String = ""
        '从数据库中获取五个分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
        Dim queryURL = From c In efContext.Categories
                       Where c.SiteID = siteId
                       Select c
        For Each q In queryURL
            Select Case q.Category1.Trim()
                Case "Outerwears"
                    OuterwearsUrl = q.Url
                Case "WomenTops"
                    WomenTopsUrl = q.Url
                Case "Shoes"
                    ShoesUrl = q.Url
                Case "Dresses"
                    DressesUrl = q.Url
            End Select
        Next

        Dim iProIssueCount As Integer = 3
        Dim updateTime As DateTime = Now
        Try
            GetProducts(siteId, IssueId, WomenTopsUrl, "WomenTops", iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, DressesUrl, "Dresses", iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, ShoesUrl, "Shoes", iProIssueCount, updateTime, planType)

            GetProducts(siteId, IssueId, OuterwearsUrl, "Outerwears", iProIssueCount, updateTime, planType)
        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, _
                            ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, iProIssueCount, siteId, updateTime, categoryName)
            ' 当产品发送完成时实现翻页
            If (listProducts.Count < 2) Then
                listProducts = GetListProducts(categoryUrl & "2/", iProIssueCount, siteId, updateTime, categoryName)
            End If
            If (listProducts.Count < 2) Then
                listProducts = GetListProducts(categoryUrl & "3/", iProIssueCount, siteId, updateTime, categoryName)
            End If
            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)
                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            EFHelper.LogText(ex.ToString())
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

        Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
        Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//dl[@class='garrery']/dd") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")

        For counter As Integer = 0 To productDivs.Count - 1
            Dim imageliNodes As HtmlNode = productDivs(counter).SelectSingleNode("a") 'dd/ul/li 图片部分
            'Dim textliNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("dd/ul/li") '文字部分

            Dim product As New Product()

            product.Prodouct = productDivs(counter).SelectSingleNode("p").InnerText.Trim
            product.Url = efHelpre.addDominForUrl(domin, imageliNodes.GetAttributeValue("href", ""))
            If (listProductUrl.Contains(product.Url)) Then
                Continue For
            Else
                listProductUrl.Add(product.Url)
            End If

            product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("div").InnerText.Replace("USD $ ", ""))
            product.Currency = "USD"

            '特殊图片处理
            product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("src", "")
            If (product.PictureUrl.Contains("loading.gif")) Then
                product.PictureUrl = imageliNodes.SelectSingleNode("img").GetAttributeValue("data-src", "")
            End If

            product.LastUpdate = nowTime
            product.Description = product.Prodouct
            product.SiteID = siteId
            product.PictureAlt = product.Prodouct

            If (Not efHelpre.IsProductSent(siteId, product.Url, Now.AddDays(-160), Now)) Then '控制一个月内获取的新品不重复
                listProducts.Add(product)
                If (listProducts.Count = productCount * 2) Then
                    Exit For
                End If
            End If
        Next
        Return listProducts
    End Function

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer), _
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
        Dim querySubject = From p In efContext.Products
                Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                Select p.Prodouct
        efHelpre.InsertIssueSubject(issueId, "Hi [FIRSTNAME],Weekly Deals:" & querySubject.FirstOrDefault.ToString.Trim)
    End Sub

   

End Class
