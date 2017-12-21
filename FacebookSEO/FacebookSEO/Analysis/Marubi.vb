Imports HtmlAgilityPack
Imports System.Text.RegularExpressions
Imports Analysis.SpreadWebReference

''' <summary>
''' 该类用于支持丸美新品网站 数据抓取。
''' EmailAlerter系统会将抓取的数据自动填充到模板中。
''' 本类参考源自Siluguoguo类
''' </summary>
''' <remarks>
'''          create by: felix.liu
'''          create Date: 2014-04-15 
''' </remarks>
Public Class Marubi
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private hasAds As Boolean = vbFalse
    ''' <summary>
    ''' 每周只推送一个类别，按第2）点的类别顺序轮流推送，每个类别获取6个产品，同一个类别的邮件产品内容2个月内不能重复。
    ''' 采用每次从[AutomationPlan]中获取URL作为本次的获取类别，并通过Select Case 确认下一个类别的URL更新AutomationPlan
    ''' 主推类别5个：
    ''' 眼霜：http://marubi.tmall.com/category-763846386.htm 
    ''' 洗面奶：http://marubi.tmall.com/category-763846387.htm 
    ''' 爽肤水：http://marubi.tmall.com/category-763846388.htm
    ''' 乳液：http://marubi.tmall.com/category-763846389.htm 
    ''' 精华液：http://marubi.tmall.com/category-763846390.htm 
    ''' 面霜：http://marubi.tmall.com/category-763846391.htm
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="splitContactCount"></param>
    ''' <param name="spreadLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="url"></param>
    ''' <param name="categories"></param>
    ''' <remarks></remarks>
    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        GetCategory(siteId)
        'Dim bannerCategoryName As String = "marubi"
        'GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
        GetBanner(siteId, IssueID, planType)

        Dim ChineseCategoryName As String = "更多"
        Dim nextUrl As String = url
        Dim product As New Product()
        'If Not (hasAds) Then
        GetHOTProducts(url, "marubi", siteId, planType, IssueID)
        Select Case url.Trim
            Case "http://marubi.tmall.com/category-763846386.htm" '眼霜
                ChineseCategoryName = "眼霜"
                'CategoryName = "EyeCream"
                product.Url = "http://marubi.tmall.com/category-763846386.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846387.htm"
            Case "http://marubi.tmall.com/category-763846387.htm" '洗面奶
                ChineseCategoryName = "洗面奶"
                'CategoryName = "MilkyWash"
                product.Url = "http://marubi.tmall.com/category-763846387.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846388.htm"
            Case "http://marubi.tmall.com/category-763846388.htm" '爽肤水
                ChineseCategoryName = "爽肤水"
                'CategoryName = "Toner"
                product.Url = "http://marubi.tmall.com/category-763846388.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846389.htm"
            Case "http://marubi.tmall.com/category-763846389.htm" '乳液
                ChineseCategoryName = "乳液"
                'CategoryName = "Lotions"
                product.Url = "http://marubi.tmall.com/category-763846389.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846390.htm"
            Case "http://marubi.tmall.com/category-763846390.htm" '精华液
                ChineseCategoryName = "精华液"
                ' CategoryName = "Essence"
                product.Url = "http://marubi.tmall.com/category-763846390.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846391.htm"
            Case "http://marubi.tmall.com/category-763846391.htm" '面霜
                ChineseCategoryName = "面霜"
                'CategoryName = "FaceCream"
                product.Url = "http://marubi.tmall.com/category-763846391.htm"
                product.Prodouct = ChineseCategoryName
                product.SiteID = siteId
                product.LastUpdate = Now
                product.Description = ChineseCategoryName
                nextUrl = "http://marubi.tmall.com/category-763846386.htm"
        End Select
        'Else
        'product.Url = "http://marubi.tmall.com"
        'product.Prodouct = ChineseCategoryName
        'product.SiteID = siteId
        'product.LastUpdate = Now
        'product.Description = ChineseCategoryName
        'End If

        Dim listProduct As New List(Of Product)
        listProduct = efHelper.GetProductList(siteId)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, "title")
        Dim returnId As Integer = efHelper.InsertProduct(product, Now, categoryId, listProduct)
        listProductId.Add(returnId)
        InsertProductsIssue(siteId, IssueID, "CA", listProductId, 1)

        efHelper.InsertContactList(IssueID, "14年4月有效地址名字不含yahoo中国区邮箱", "draft")

        Dim querySubject = From p In efContext.Products
                     Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                     Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                     Select p.Prodouct
        Dim endSubject As String = ""
        For Each p In querySubject
            If (Not String.IsNullOrEmpty(p) AndAlso p.ToString.Length > 3) Then
                endSubject = p.ToString.Trim
                Exit For
            End If
        Next
        '8）标题暂定：[FIRSTNAME] 你好，为你精选 + 类别名称：+ 第一个产品名称(由于第一个产品为分组信息所以此处应该抓第二产品
        If (endSubject <> "") Then
            efHelper.InsertIssueSubject(IssueID, "[FIRSTNAME] 您好，本周为你精选 " & ChineseCategoryName.ToString & "：" & endSubject & "(AD)")
        Else
            efHelper.InsertIssueSubject(IssueID, "[FIRSTNAME] 您好，更多惊喜尽在丸美!" & "(AD)")
        End If

        '将url更新为下一期邮件要发送的category
        Dim queryCategory As AutomationPlan = efContext.AutomationPlans.Where(Function(c) c.PlanType = planType AndAlso c.SiteID = siteId).Single()
        queryCategory.URL = nextUrl
        efContext.SaveChanges()

    End Sub

    ''' <summary>
    ''' 静态添加分类Banner分类和Title分类基本不会改变
    ''' Title分类主要用于存放一个产品，该产品用于描述分组信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="CategoryName"></param>
    ''' <param name="url"></param>
    ''' <param name="IssueID"></param>
    ''' <remarks></remarks>
    Private Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "marubi"
        myCategory.Description = "marubi"
        myCategory.Url = "http://marubi.tmall.com/"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory1 As New Category
        myCategory1.Category1 = "title"
        myCategory1.Description = "title"
        myCategory1.Url = "http://marubi.tmall.com/search.htm"
        myCategory1.SiteID = siteId
        myCategory1.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory1, siteId)

    End Sub

    ''' <summary>
    ''' Banner: 获取首页（http://marubi.tmall.com ）上的Promotion Banner的image和Alt（没则不需要获取），
    ''' 以Banner的图片链接为重复判断，保证1个月内的Banner的image不重复，如果重复，则不获取, 则Deals邮件里不出Banner
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Sub GetBanner(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String)
        Dim helper As New EFHelper()
        Dim bannerUrl As String = "http://marubi.tmall.com/"  'Banner动态加载外部html

        Dim doc As HtmlDocument = helper.GetHtmlDocument2(bannerUrl, 120000)

        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, "marubi")
        Dim listAdIds As New List(Of Integer)
        Dim adsNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='ks-switchable-content']/div[@class='item']")
        If Not (adsNodes Is Nothing) Then
            For Each node In adsNodes
                Dim ad As New Ad
                Dim adLink As String
                Try
                    adLink = node.SelectSingleNode("div/map/area").GetAttributeValue("href", "").Trim()
                Catch e As Exception
                    Try
                        adLink = node.SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                    Catch ex As Exception
                        adLink = "http://marubi.tmall.com/"
                    End Try
                End Try

                ad.Url = adLink
                ad.Url = efHelper.AddHttpForAli(ad.Url)
                Try
                    ad.PictureUrl = node.SelectSingleNode("div/img").GetAttributeValue("data-ks-lazyload", "").Trim()
                Catch e As Exception
                    ad.PictureUrl = node.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").Trim()
                End Try
                ad.PictureUrl = efHelper.AddHttpForAli(ad.PictureUrl)
                If (efHelper.isAdSent(siteId, ad.PictureUrl, Now.Date.AddDays(-31), Now.Date, PlanType, True)) Then
                    Continue For
                End If
                ad.SiteID = siteId
                ad.Lastupdate = Date.Now
                
                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteId)
                Dim adid As Integer
                If (Not efHelper.isAdSent(siteId, ad.PictureUrl, Now.Date.AddDays(-31), Now, planType, True) AndAlso ad.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                    adid = InsertAd(ad, listAd, categoryId)
                    listAdIds.Add(adid)
                    hasAds = True
                    Exit For
                End If
            Next
        End If

        InsertAdsIssue(siteId, issueId, listAdIds, 1)
    End Sub





#Region "Product Area"
    Private Sub GetHOTProducts(ByVal url As String, ByVal categoryName As String, ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer)

        Try
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now

            'Dim maxProductCount As Integer = 34
            iProIssueCount = 4 '在模板中要填充产品的个数
            GetProducts(siteID, issueID, url, categoryName, iProIssueCount, updateTime, planType)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Function GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

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
    End Function

    Private Function GetListProducts(ByVal categoryUrl As String, ByVal siteId As Integer, ByVal lastUpdate As DateTime, _
                                          ByVal categoryname As String, ByVal planType As String, ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper()
        Try
            Dim categoryDoc As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")

            For Each productLine In productLines
                Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
                If Not (productDls Is Nothing) Then
                    For Each dl In productDls

                        Dim myProduct As New Product
                        Dim productName As String = dl.SelectSingleNode("dd[@class='detail']/a").InnerText
                        If productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
                            Continue For
                        End If
                        myProduct.Prodouct = productName
                        Dim tempUrl As String = dl.SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                        myProduct.Url = tempUrl.Split("&")(0)
                        If (listProductUrl.Contains(myProduct.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(myProduct.Url)
                        End If
                        myProduct.Url = efHelper.AddHttpForAli(myProduct.Url)
                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-60), DateTime.Now, planType)) Then
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
                If (listProducts.Count >= iProCount) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when categoryUrl:" & categoryUrl & " categoryName:" & categoryname & ex.StackTrace)
        End Try
        Return listProducts
    End Function

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
#End Region
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

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
