Imports HtmlAgilityPack
Imports System.Security.Policy

Public Class LighTake

    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper
    Private Flage1 As Boolean = vbTrue
    Private Flage2 As Boolean = vbTrue
    Private Flage3 As Boolean = vbTrue
    Private listProductUrl As New List(Of String)
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        Dim contactlistCount As Integer
        If planType.Trim.Contains("HO") Then
            bannerCategoryName = "Puzzles & Magic Cube"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType) '获取主页Weekly Deals的部分，顺序获取2个产品，如果和上周的重复，则不获取;Section: DA
            GetHotOrNewProducts(siteId, IssueID, "hot", planType.Trim)
            ''Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "DA")
            ' ''Add 
            Dim contactListArr() As String = {"Alldata"}
            efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting
        ElseIf (planType.Trim.Contains("HA")) Then
            '分析html获取到banner条的分类
            '获取banner信息插入到ads表中
            bannerCategoryName = "Electronic Cigarettes"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "NE", planType) '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
            GetHotOrNewProducts(siteId, IssueID, "new", planType.Trim) '在获取分类url时，直接将最后的hot替换成new后直接使用
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "NE")
            ''Add 
            Dim contactListArr() As String = {"Alldata"}
            efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting
        ElseIf (planType.Trim.Contains("HP")) Then
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
                        categoryUrlL = q.Url & "?sort=4&pagesize=80"
                    Case cate(1).Trim
                        categoryUrlC = q.Url & "?sort=4&pagesize=80"
                End Select
            Next
            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "LightakeSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
        End If
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
            Dim nowTime As DateTime = Now
            '添加14大分类 
            Dim cateCell_Phones_Mobiles As New Category()
            Dim cateNotebook_Tablet_PC As New Category()
            Dim cateMobile_Power_Bank As New Category()
            Dim cateComputer_Peripherals As New Category()
            Dim cateConsumer_Electronics As New Category()

            If Not listCategoryName.Contains("Puzzles & Magic Cube") Then
                cateCell_Phones_Mobiles.Category1 = "Puzzles & Magic Cube"
                cateCell_Phones_Mobiles.SiteID = siteID
                cateCell_Phones_Mobiles.Url = "http://lightake.com/c/Puzzles-Magic-Cube_001"
                cateCell_Phones_Mobiles.LastUpdate = nowTime
                cateCell_Phones_Mobiles.Description = "Puzzles & Magic Cube"
                cateCell_Phones_Mobiles.ParentID = -1
                efContext.AddToCategories(cateCell_Phones_Mobiles)
            End If
            If Not listCategoryName.Contains("Electronic Cigarettes") Then
                cateNotebook_Tablet_PC.Category1 = "Electronic Cigarettes"
                cateNotebook_Tablet_PC.SiteID = siteID
                cateNotebook_Tablet_PC.Url = "http://lightake.com/c/Electronic-Cigarettes_011"
                cateNotebook_Tablet_PC.LastUpdate = nowTime
                cateNotebook_Tablet_PC.Description = "Electronic Cigarettes"
                cateNotebook_Tablet_PC.ParentID = -1
                efContext.AddToCategories(cateNotebook_Tablet_PC)
            End If
            If Not listCategoryName.Contains("Cell Phones & Tablets") Then
                cateMobile_Power_Bank.Category1 = "Cell Phones & Tablets"
                cateMobile_Power_Bank.SiteID = siteID
                cateMobile_Power_Bank.Url = "http://lightake.com/c/Cell-Phones-Tablets_008"
                cateMobile_Power_Bank.LastUpdate = nowTime
                cateMobile_Power_Bank.Description = "Cell Phones & Tablets"
                cateMobile_Power_Bank.ParentID = -1
                efContext.AddToCategories(cateMobile_Power_Bank)
            End If
            If Not listCategoryName.Contains("Consumer Electronics") Then
                cateComputer_Peripherals.Category1 = "Consumer Electronics"
                cateComputer_Peripherals.SiteID = siteID
                cateComputer_Peripherals.Url = "http://lightake.com/c/Consumer-Electronics_006"
                cateComputer_Peripherals.LastUpdate = nowTime
                cateComputer_Peripherals.Description = "Consumer Electronics"
                cateComputer_Peripherals.ParentID = -1
                efContext.AddToCategories(cateComputer_Peripherals)
            End If
            If Not listCategoryName.Contains("Health & Beauty") Then
                cateConsumer_Electronics.Category1 = "Health & Beauty"
                cateConsumer_Electronics.SiteID = siteID
                cateConsumer_Electronics.Url = "http://lightake.com/c/Health-Beauty_010"
                cateConsumer_Electronics.LastUpdate = nowTime
                cateConsumer_Electronics.Description = "Health & Beauty"
                cateConsumer_Electronics.ParentID = -1
                efContext.AddToCategories(cateConsumer_Electronics)
            End If
            If Not listCategoryName.Contains("Batteries & Chargers") Then
                cateConsumer_Electronics.Category1 = "Batteries & Chargers"
                cateConsumer_Electronics.SiteID = siteID
                cateConsumer_Electronics.Url = "http://lightake.com/c/Batteries-Chargers_006007"
                cateConsumer_Electronics.LastUpdate = nowTime
                cateConsumer_Electronics.Description = "Batteries & Chargers"
                cateConsumer_Electronics.ParentID = -1
                efContext.AddToCategories(cateConsumer_Electronics)
            End If
            If Not listCategoryName.Contains("Home & Garden") Then
                cateConsumer_Electronics.Category1 = "Home & Garden"
                cateConsumer_Electronics.SiteID = siteID
                cateConsumer_Electronics.Url = "http://lightake.com/c/Home-Garden_009"
                cateConsumer_Electronics.LastUpdate = nowTime
                cateConsumer_Electronics.Description = "Home & Garden"
                cateConsumer_Electronics.ParentID = -1
                efContext.AddToCategories(cateConsumer_Electronics)
            End If
            If Not listCategoryName.Contains("Toys,Hobbies & Gaming") Then
                cateConsumer_Electronics.Category1 = "Toys,Hobbies & Gaming"
                cateConsumer_Electronics.SiteID = siteID
                cateConsumer_Electronics.Url = "http://lightake.com/c/Toys-Hobbies-Gaming_013"
                cateConsumer_Electronics.LastUpdate = nowTime
                cateConsumer_Electronics.Description = "Toys,Hobbies & Gaming"
                cateConsumer_Electronics.ParentID = -1
                efContext.AddToCategories(cateConsumer_Electronics)
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try
    End Sub

    '获取主页Weekly Deals的部分的2个产品，如果和上周的重复，则不获取;Section: DA
    Public Sub GetWeeklyDealsOrNewArri(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetWeeklyDealOrNewArriProducts(siteId, Issueid, section, planType) '这个函数中已判断是否和上周重复
            Dim topCount As Integer = 2
            If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
                Dim categoryName As String
                Dim categoryId As Integer

                Dim listProduct As New List(Of Product)
                listProduct = GetProductList(siteId)
                Dim listProductId As New List(Of Integer)
                Dim counter As Integer = 0
                For Each li In listProducts
                    categoryName = getCategoryName(li.Url)
                    Try
                        categoryId = GetCategoryId(siteId, categoryName)
                    Catch ex As Exception
                        categoryId = GetCategoryId(siteId, "Puzzles & Magic Cube")
                    End Try
                    Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)
                    listProductId.Add(returnId)

                    '只插入排好序的前topCount个产品
                    counter += 1
                    If (counter = topCount) Then
                        Exit For
                    End If
                Next
                InsertProductsIssue(siteId, Issueid, section, listProductId, topCount)
            End If
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    '根据网站的12个分类获取热销部分，请按顺序获取，保证月内不重复，Section：CA
    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim puzzlesUrl As String = ""
        Dim electronicUrl As String = ""
        Dim phoneUrl As String = ""
        Dim healthUrl As String = ""
        Dim ConsumerUrl As String = ""
        Try
            '从数据库中获取5个主推分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Puzzles & Magic Cube"
                        If (HotORNew = "new") Then
                            puzzlesUrl = q.Url & "?sort=1"
                        Else
                            puzzlesUrl = q.Url & "?sort=5"
                        End If
                    Case "Electronic Cigarettes"
                        If (HotORNew = "new") Then
                            electronicUrl = q.Url & "?sort=1"
                        Else
                            electronicUrl = q.Url & "?sort=5"
                        End If
                    Case "Cell Phones & Tablets"
                        If (HotORNew = "new") Then
                            phoneUrl = q.Url & "?sort=1"
                        Else
                            phoneUrl = q.Url & "?sort=5"
                        End If
                    Case "Consumer Electronics"
                        If (HotORNew = "new") Then
                            ConsumerUrl = q.Url & "?sort=1"
                        Else
                            ConsumerUrl = q.Url & "?sort=5"
                        End If
                    Case "Health & Beauty"
                        If (HotORNew = "new") Then
                            healthUrl = q.Url & "?sort=1"
                        Else
                            healthUrl = q.Url & "?sort=5"
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, puzzlesUrl, "Puzzles & Magic Cube", iProIssueCount, updateTime, planType, 4, 8)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, electronicUrl, "Electronic Cigarettes", iProIssueCount, updateTime, planType, 10, 20)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, phoneUrl, "Cell Phones & Tablets", iProIssueCount, updateTime, planType, 50, 150)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, healthUrl, "Health & Beauty", iProIssueCount, updateTime, planType, 5, 9)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, ConsumerUrl, "Consumer Electronics", iProIssueCount, updateTime, planType, 6, 12)
        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String, ByVal lowPrice As Integer, ByVal highetPrice As Integer)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim listExcludeProductId As New List(Of Integer) '存放按照价格区间来匹配产品而没有匹配到的产品id，以备符合指定价格区间的产品数量不足时从此中再选择产品填充
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)

                If (isInPriceSpan(li, lowPrice, highetPrice)) Then
                    listProductId.Add(returnId)
                Else
                    listExcludeProductId.Add(returnId)
                End If
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            If (listProductId.Count < iProIssueCount) Then
                For Each li In listExcludeProductId
                    listProductId.Add(li)
                    If (listProductId.Count = iProIssueCount) Then
                        Exit For
                    End If
                Next
            End If

            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

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
            LogText(ex.ToString())
        End Try
    End Sub


    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                    ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-proNoRepeatSpan).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@id='dib_outBox']/span[@class='dib_box']")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)

                product.Url = productDetails.SelectSingleNode("a").GetAttributeValue("href", "").Trim().ToLower
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If
                product.Prodouct = productDetails.SelectSingleNode("span[@class='ovh_2line mt5']/a").InnerText.Trim
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 70) Then
                    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                End If

                'product.Price = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Discount = Double.Parse(productDetails.SelectSingleNode("span[@class='dib_txt']/span").InnerText.ToLower.Replace("us$", "").Trim())
                product.Currency = "US$"

                '特殊图片处理 pList2_mg
                product.PictureUrl = productDetails.SelectSingleNode("a/img").GetAttributeValue("src", "")
                If (product.PictureUrl.Contains("blank.gif")) Then
                    product.PictureUrl = productDetails.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                End If

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct

                If Not (isSameProductDiffSize(listProducts, product)) Then '避免同一封邮件抓取同款不同size或不同色的产品
                    If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
                        listProducts.Add(product)
                    End If
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Function InsertProduct(ByVal productInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal list As List(Of Product)) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = productInList.Prodouct
        product.Url = productInList.Url
        product.Price = productInList.Price
        product.PictureUrl = productInList.PictureUrl
        product.LastUpdate = now
        product.Description = productInList.Description
        product.SiteID = productInList.SiteID
        product.Currency = productInList.Currency
        product.PictureAlt = productInList.PictureAlt
        product.Discount = productInList.Discount
        product.ShipsImg = productInList.ShipsImg
        product.FreeShippingImg = productInList.FreeShippingImg

        If (JudgeProduct(product.Url, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = productInList.Url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = product.Description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = now
            updateProduct.Discount = product.Discount
            updateProduct.FreeShippingImg = product.FreeShippingImg
            updateProduct.ShipsImg = product.ShipsImg
            'Dim returnId As Integer = updateProduct.ProdouctID
            'efContext.SaveChanges()

            'test
            'If (productInList.Url = "http://www.ahappydeal.com/product-165223.html") Then
            '    Dim a As String = productInList.Url
            'End If

            '2014/1/21新增，防止一个产品有多个productCategory关系
            Dim updateCategory = updateProduct.Categories
            ''For Each pcate In updateCategory
            ''    product.Categories.Remove(pcate)
            ''Next
            If (updateCategory.Count = 0) Then
                category.Products.Add(updateProduct) '在查询分类时，添加一条关系在ProductCategory表中
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
    End Function


    Public Function GetWeeklyDealOrNewArriProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String) As List(Of Product)
        Dim pageUrl = "http://lightake.com/"
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-10).ToShortDateString() 'Deals控制一周内不重复
        Dim listProducts As New List(Of Product)
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            If (section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//ul[@class='tab_content tab_content_rel']/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
                '获取产品  mb10 newArr
                Dim counter As Integer
                'Dim dealtimeLeft As Integer
                If Not productDivs Is Nothing Then
                    For counter = 0 To productDivs.Count - 1
                        Dim product As New Product()
                        'Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("div[@class='proBox clearfix']/div")

                        product.Prodouct = productDivs(counter).SelectSingleNode("div[@class='ovh_2line']/a/div").GetAttributeValue("title", "").Trim
                        If (product.Prodouct.ToString.Length > 100) Then
                            product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                        End If
                        product.Url = productDivs(counter).SelectSingleNode("div[@class='ovh_2line']/a").GetAttributeValue("href", "").Trim().ToLower()
                        If (listProductUrl.Contains(product.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(product.Url)
                        End If

                        product.Price = Double.Parse(productDivs(counter).SelectSingleNode("div[@class='ovh_2line']/del").InnerText.ToLower.Replace("us$", "").Trim()) '
                        product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("div[@class='ovh_2line']/span").InnerText.ToLower.Replace("us$", "").Trim()) '
                        product.Currency = "US$"

                        product.PictureUrl = productDivs(counter).SelectSingleNode("span[@class='rel dib']/a/img").GetAttributeValue("src", "").Replace("180x180", "420x420") '2013/4/7添加
                        If (product.PictureUrl.Contains("blank.gif")) Then
                            product.PictureUrl = productDivs(counter).SelectSingleNode("span[@class='rel dib']/a/img").GetAttributeValue("data-ks-lazyload", "").Replace("180x180", "420x420")
                        End If

                        product.LastUpdate = nowTime
                        product.Description = product.Prodouct
                        product.SiteID = siteId
                        product.PictureAlt = product.Prodouct
                        If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制指定时间内获取的产品不重复：
                            listProducts.Add(product)
                        End If
                        'End If
                    Next
                End If
            ElseIf (section = "NE") Then
                productDivs = hd.DocumentNode.SelectNodes("//div[@id='pendant']/div[@class='fix pendant_box']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
                '获取产品  mb10 newArr
                Dim counter As Integer
                'Dim dealtimeLeft As Integer
                If Not productDivs Is Nothing Then
                    For counter = 0 To productDivs.Count - 1
                        Dim product As New Product()
                        'Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("div[@class='proBox clearfix']/div")

                        product.Prodouct = productDivs(counter).SelectSingleNode("div/div/div[@class='ovh_2line']/a").InnerText.Trim
                        If (product.Prodouct.ToString.Length > 100) Then
                            product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                        End If
                        product.Url = productDivs(counter).SelectSingleNode("a").GetAttributeValue("href", "").Trim().ToLower()
                        If (listProductUrl.Contains(product.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(product.Url)
                        End If

                        'product.Price = Double.Parse(productDivs(counter).SelectSingleNode("div/div/div/div[@class='co008 mt5 b']").InnerText.ToLower.Replace("us$", "").Trim()) '
                        product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("div/div/div[@class='co008 mt5 b']").InnerText.ToLower.Replace("us$", "").Trim()) '
                        product.Currency = "US$"

                        '特殊图片处理
                        product.PictureUrl = productDivs(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").Trim.Replace("180x180", "420x420")
                        If (product.PictureUrl.Contains("blank.gif")) Then
                            product.PictureUrl = productDivs(counter).SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "").Replace("180x180", "420x420")
                        End If

                        product.LastUpdate = nowTime
                        product.Description = product.Prodouct
                        product.SiteID = siteId
                        product.PictureAlt = product.Prodouct
                        If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制指定时间内获取的产品不重复：
                            listProducts.Add(product)
                        End If
                        'End If
                    Next
                End If
            End If
            Return listProducts
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:WeeklyDeals" & " error:" & ex.Message.ToString)
        End Try
    End Function

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        Dim pageUrl As String = "http://lightake.com/"
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)

            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 获取指定站点的指定CategoryName的CategoryId
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

    ''' <summary>
    ''' 获取指定URL页面的Banner条信息
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="siteid"></param>
    ''' <param name="nowTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10    js_slideBox   tempWrap
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='banner_box']/dl/dd/div/div/a")
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try

            Try
                Dim index1 As Integer = banners(counter).GetAttributeValue("style", "").ToString.IndexOf("(")
                Dim index2 As Integer = banners(counter).GetAttributeValue("style", "").ToString.IndexOf(")")
                If Not (banners(counter).GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).ToLower.Contains("http://")) Then
                    If (banners(counter).GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).Trim.StartsWith("/")) Then
                        Ad.PictureUrl = "http://e.lightake.com" & banners(counter).GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    Else
                        Ad.PictureUrl = "http://e.lightake.com/" & banners(counter).GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    End If
                Else
                    Ad.PictureUrl = banners(counter).GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                End If
            Catch ex As Exception
                'ignore
            End Try

            Ad.Url =  banners(counter).GetAttributeValue("href", "").ToString.Trim()

            'Ad.SizeHeight = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString)
            'Ad.SizeWidth = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString)
            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelpre.GetListAd(siteid)
            If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                adid = efHelpre.InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
                Exit For '获取到一个banner后则退出防止获取更多的banner
            End If
        Next
        Return listAdIds
    End Function

    ''' <summary>
    ''' 将获得的Ad信息保存至Ads表中。
    ''' </summary>
    ''' <param name="ad"></param>
    ''' <param name="listAd"></param>
    ''' <param name="categoryID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

        If (JudgeAd(newAd.Url, listAd)) Then
            newAd.Categories.Add(category)
            efContext.AddToAds(newAd)
            efContext.SaveChanges()
            Return newAd.AdID
        Else
            Dim updateAd = efContext.Ads.FirstOrDefault(Function(m) m.Url = ad.Url)
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
            efContext.SaveChanges()

            Return updateAd.AdID
        End If
    End Function

    ''' <summary>
    ''' 判断即将插入的Ad的URL是否在数据库中已经存在，如果存在，返回false
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function JudgeAd(ByVal url As String, ByVal list As List(Of Ad)) As Boolean
        For Each li In list
            If (li.Url.Trim() = url.Trim()) Then
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

    ''' <summary>
    ''' 获取指定站点的products表中的所有产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductList(ByVal siteId As Integer) As List(Of Product)
        Dim query = From p In efContext.Products
                   Where p.SiteID = siteId
                   Select p
        Dim listProduct As New List(Of Product)
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next
        Return listProduct
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

    Private Function getCategoryName(ByVal Url As String) As String
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(Url, 60000)
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@class='conStr']/div[@class='strIn']/div[@class='mt10']")
        Dim aHerfs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='conStr']/div[@class='strIn']/div[@class='mt10']/a")
        Dim categoryName As String = aHerfs(1).InnerText
        If (categoryName <> "Puzzles & Magic Cube" AndAlso categoryName <> "Electronic Cigarettes" AndAlso categoryName <> "Cell Phones & Tablets" AndAlso
            categoryName <> "Consumer Electronics" AndAlso categoryName <> "Health & Beauty") Then
            categoryName = "Puzzles & Magic Cube" '不属于5个主推分类中的任何一个时自动将其归为第一个分类：Puzzles & Magic Cube
        End If
        Return categoryName
    End Function

    ''' <summary>
    ''' 判断即将插入的Productd的URL是否在数据库中已经存在，如果存在，返回false
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
    ''' 获得邮件的subject
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal categoryName As String)
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then
            preSubject = "LighTake Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "LighTake New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HP")) Then
            preSubject = "LighTake Special " & categoryName & " For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")

        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        If (planType.Trim = "HO" OrElse planType.Trim = "HA" OrElse planType.Trim.Contains("HP")) Then
            addNum = 1
        Else
            addNum = 0
        End If
        If (planType.Contains("HO") OrElse planType.Contains("HA")) Then
            planType = planType.Substring(0, 2)
        End If
        Dim query = From i In efContext.Issues
                     Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
        Dim subject As String = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub

    Private Function isInPriceSpan(ByVal product As Product, ByVal lowPrice As Integer, ByVal highetPrice As Integer) As Boolean
        If (Flage1 AndAlso Double.Parse(product.Discount) < lowPrice) Then
            Flage1 = False
            Return True
        End If
        If (Flage2 AndAlso Double.Parse(product.Discount) >= lowPrice AndAlso Double.Parse(product.Discount) < highetPrice) Then
            Flage2 = False
            Return True
        End If
        If (Flage3 AndAlso Double.Parse(product.Discount) >= highetPrice) Then
            Flage3 = False
            Return True
        End If
        Return False
    End Function

    Private Function isSameProductDiffSize(ByRef listProducts As List(Of Product), ByVal product As Product) As Boolean
        For Each pro In listProducts
            If (pro.Description.ToString.Equals(product.Description.ToString)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
