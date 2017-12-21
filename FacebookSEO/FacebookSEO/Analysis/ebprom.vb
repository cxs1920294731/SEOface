Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack

Public Class ebprom

    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper
    Private listProductUrl As New List(Of String)
    Private domin As String = "http://www.ebprom.com"
    Private proNoRepeatSpan As Integer = 60 '产品不重复时间长（天为单位）

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                    ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)

        'Website analysis
        InsertCategories(siteId)
        Dim bannerCategoryName = "Prom"
        GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
        If (planType.Trim = "HO") Then
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)
            AddIssueSubject(IssueID, siteId, planType)
        ElseIf (planType.Trim = "HA") Then
            GetHotOrNewProducts(siteId, IssueID, "new", planType)
            AddIssueSubject(IssueID, siteId, planType)
        End If

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

            If Not listCategoryName.Contains("Prom") Then
                cateOuterwears.Category1 = "Prom"
                cateOuterwears.SiteID = siteID
                cateOuterwears.Url = "http://www.ebprom.com/special-occasion-dresses-prom-dresses_c2?pagesize=48"
                cateOuterwears.LastUpdate = nowTime
                cateOuterwears.Description = "Prom"
                cateOuterwears.ParentID = -1
                cateOuterwears.Gender = "M"
                efContext.AddToCategories(cateOuterwears)
                'Else
                '    cateOuterwears = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Prom" And m.SiteID = siteID)
                '    cateOuterwears.Url = "http://www.dressve.com/Best-Selling/Tops-101704/"
            End If

            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try

    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        Dim pageUrl = "http://www.ebprom.com"
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

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
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='index_ad fl']/ul/li/a")
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
                Ad.PictureUrl = efHelpre.addDominForUrl(domin, banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString)
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

    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim promUrl As String = ""

        Try
            '从数据库中获取5个主推分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Prom"
                        If (HotORNew = "new") Then
                            promUrl = q.Url.Trim() & "&productsort=5"
                        Else
                            promUrl = q.Url.Trim()
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 4

            GetProducts(siteId, Issueid, promUrl, "Prom", iProIssueCount, updateTime, planType)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub


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

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = efHelpre.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)

            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            efHelper.LogText("siteID:" & siteId & "categoryName:" & categoryName & "ID:" & categoryId)

            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)

                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            efHelpre.InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 获取某个分类URL的产品信息，
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="planType"></param>
    ''' <param name="iProIssueCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                  ByVal iProIssueCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-proNoRepeatSpan).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@id='list_bg_img']/ul/li/div/ul")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                product.Prodouct = productDetails.SelectSingleNode("div[@class='_title']/a").InnerText
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 70) Then
                    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                End If
                product.Url = efHelpre.addDominForUrl(domin, productDetails.SelectSingleNode("div[@class='_title']/a").GetAttributeValue("href", "").Trim())
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If
                Try
                    product.Price = Double.Parse(productDetails.SelectSingleNode("div[@class='black line_120']/div/del").InnerText.ToLower.Replace("$", "").Replace("usd", "").Trim())
                Catch ex As Exception
                End Try

                product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='black line_120']/div/a").InnerText.ToLower.Replace("$", "").Replace("usd", "").Trim())
                product.Currency = "USD"

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("li[@class='relative _img']")
                product.PictureUrl = efHelpre.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("src", ""))
                If (product.PictureUrl.Contains("default.")) Then
                    product.PictureUrl = efHelpre.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("data-lazysrc", ""))
                End If

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct
                If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
                    listProducts.Add(product)
                End If
                If (listProducts.Count = 2 * iProIssueCount) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

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
        Dim addNum As Integer = 1
        Dim preSubject As String = ""
        If (planType.Contains("HO")) Then
            preSubject = "Hi [FirstName],ebProm Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "Hi [FirstName],ebProm New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")

        Dim query = From i In efContext.Issues
                     Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
        Dim subject As String = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
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
