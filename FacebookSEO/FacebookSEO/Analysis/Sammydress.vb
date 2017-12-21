Imports HtmlAgilityPack

Public Class Sammydress
    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper()
    Private pageUrl As String = "http://www.sammydress.com/"
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        If planType.Trim.Contains("HO") Then
            'Website analysis
            bannerCategoryName = "Women's Dresses" '将banner条的类名放到Women's Clothing一类
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType) '获取主页Weekly Deals的部分，获取折扣最高的3个产品，如果和上周的重复，则不获取;Section: DA
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "DA")

            If (planType.Trim = "HO") Then
                'Dim contactListArr() As String = {"S_SPREAD_001"} 'use seachContact to get the list who has opened but no clicked and not opened
                'efHelpre.InsertContactList(IssueID, contactListArr, "waiting") 'draft  'waiting

                Dim saveListName As String = "Sammydress_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"
                CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                saveListName = "Sammydress_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")
                CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
            ElseIf (planType.Trim.Contains("HO") AndAlso planType.Trim <> "HO") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "Sammydress_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
        ElseIf (planType.Trim.Contains("HA")) Then
            '分析html获取到banner条的分类
            '获取banner信息插入到ads表中
            'GetPromotionBannerForNewArri(siteId, IssueID, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "NE", planType) '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
            GetHotOrNewProducts(siteId, IssueID, "new", planType) '在获取分类url时，直接将最后的hot替换成new后直接使用
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "NE")

            If (planType.Trim = "HA") Then
                'Dim contactListArr() As String = {"S_SPREAD_001"} 'use seachContact to get the list who has opened but no clicked and not opened
                'efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

                Dim saveListName As String = "Sammydress_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"
                CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                saveListName = "Sammydress_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")
                CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
            ElseIf (planType.Trim.Contains("HA") AndAlso planType.Trim <> "HA") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "Sammydress_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
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
                        categoryUrlL = q.Url.Replace("?odr=hot", "?odr=likes")
                    Case cate(1).Trim
                        categoryUrlC = q.Url
                End Select
            Next

            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "SammydressSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
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

            '添加6大分类 
            Dim nowTime As DateTime = Now
            Dim cateWomen_s_Dresses As New Category()
            Dim cateMen_s_Clothing As New Category()
            Dim cateMen_s As New Category()
            Dim cateJewelry As New Category()
            Dim cateJewelryC As New Category()
            Dim cateWeddinEvents As New Category()
            Dim cateHandBags As New Category()
            Dim cateHandBagsC As New Category()
            Dim cateWomenShoes As New Category()
            Dim cateWomenShoesC As New Category()
            Dim cateSexy_Lingerie As New Category()
            Dim cateWatches As New Category()
            Dim cateBeauty_Accessories As New Category()
            Dim cateHome_Living As New Category()
            Dim cateWomen_s_Cloth As New Category()
            Dim cateKids As New Category()

            If Not listCategoryName.Contains("Women's Dresses") Then
                cateWomen_s_Dresses.Category1 = "Women's Dresses"
                cateWomen_s_Dresses.SiteID = siteID
                cateWomen_s_Dresses.Url = "http://www.sammydress.com/Wholesale-Dresses-c-2.html?odr=hot"
                cateWomen_s_Dresses.LastUpdate = nowTime
                cateWomen_s_Dresses.Description = "Women's Dresses"
                cateWomen_s_Dresses.ParentID = -1
                cateWomen_s_Dresses.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Dresses)
            Else
                cateWomen_s_Dresses = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Women's Dresses" And m.SiteID = siteID)
                cateWomen_s_Dresses.Url = "http://www.sammydress.com/Wholesale-Dresses-c-2.html?odr=hot"
            End If

            If Not listCategoryName.Contains("Women's Clothings Clearance") Then
                cateWomen_s_Cloth.Category1 = "Women's Clothings Clearance"
                cateWomen_s_Cloth.SiteID = siteID
                cateWomen_s_Cloth.Url = "http://www.sammydress.com/clearance-1.html?odr=hot"
                cateWomen_s_Cloth.LastUpdate = nowTime
                cateWomen_s_Cloth.Description = "Women's Clothings Clearance"
                cateWomen_s_Cloth.ParentID = -1
                cateWomen_s_Cloth.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Cloth)
            Else
                cateWomen_s_Cloth = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Women's Clothings Clearance" And m.SiteID = siteID)
                cateWomen_s_Cloth.Url = "http://www.sammydress.com/clearance-1.html?odr=hot"
            End If

            If Not listCategoryName.Contains("Men's Clothing") Then
                cateMen_s_Clothing.Category1 = "Men's Clothing"
                cateMen_s_Clothing.SiteID = siteID
                cateMen_s_Clothing.Url = "http://www.sammydress.com/Wholesale-Men-b-89.html?odr=hot"
                cateMen_s_Clothing.LastUpdate = nowTime
                cateMen_s_Clothing.Description = "Men's Clothing"
                cateMen_s_Clothing.ParentID = -1
                cateSexy_Lingerie.Gender = "W"
                efContext.AddToCategories(cateMen_s_Clothing)
            Else
                cateMen_s_Clothing = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Men's Clothing" And m.SiteID = siteID)
                cateMen_s_Clothing.Url = "http://www.sammydress.com/Wholesale-Men-b-89.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Men's Clothing Clearance") Then
                cateMen_s.Category1 = "Men's Clothing Clearance"
                cateMen_s.SiteID = siteID
                cateMen_s.Url = "http://www.sammydress.com/clearance-89.html?odr=hot"
                cateMen_s.LastUpdate = nowTime
                cateMen_s.Description = "Men's Clothing Clearance"
                cateMen_s.ParentID = -1
                cateMen_s.Gender = "W"
                efContext.AddToCategories(cateMen_s)
            Else
                cateMen_s = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Men's Clothing Clearance" And m.SiteID = siteID)
                cateMen_s.Url = "http://www.sammydress.com/clearance-89.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Women's Shoes") Then
                cateWomenShoes.Category1 = "Women's Shoes"
                cateWomenShoes.SiteID = siteID
                cateWomenShoes.Url = "http://www.sammydress.com/Wholesale-Women-s-Shoes-c-171.html?odr=hot"
                cateWomenShoes.LastUpdate = nowTime
                cateWomenShoes.Description = "Women's Shoes"
                cateWomenShoes.ParentID = -1
                cateWomenShoes.Gender = "W"
                efContext.AddToCategories(cateWomenShoes)
            Else
                cateWomenShoes = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Women's Shoes" And m.SiteID = siteID)
                cateWomenShoes.Url = "http://www.sammydress.com/Wholesale-Women-s-Shoes-c-171.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Shoes Clearance") Then
                cateWomenShoesC.Category1 = "Shoes Clearance"
                cateWomenShoesC.SiteID = siteID
                cateWomenShoesC.Url = "http://www.sammydress.com/clearance-170.html?odr=hot"
                cateWomenShoesC.LastUpdate = nowTime
                cateWomenShoesC.Description = "Shoes Clearance"
                cateWomenShoesC.ParentID = -1
                cateWomenShoesC.Gender = "W"
                efContext.AddToCategories(cateWomenShoesC)
            Else
                cateWomenShoesC = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Shoes Clearance" And m.SiteID = siteID)
                cateWomenShoesC.Url = "http://www.sammydress.com/clearance-170.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Jewelry") Then
                cateJewelry.Category1 = "Jewelry"
                cateJewelry.SiteID = siteID
                cateJewelry.Url = "http://www.sammydress.com/Wholesale-Jewelry-b-159.html?odr=hot"
                cateJewelry.LastUpdate = nowTime
                cateJewelry.Description = "Jewelry"
                cateJewelry.ParentID = -1
                cateJewelry.Gender = "W"
                efContext.AddToCategories(cateJewelry)
            Else
                cateJewelry = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Jewelry" And m.SiteID = siteID)
                cateJewelry.Url = "http://www.sammydress.com/Wholesale-Jewelry-b-159.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Jewelry Clearance") Then
                cateJewelryC.Category1 = "Jewelry Clearance"
                cateJewelryC.SiteID = siteID
                cateJewelryC.Url = "http://www.sammydress.com/clearance-159.html?odr=hot"
                cateJewelryC.LastUpdate = nowTime
                cateJewelryC.Description = "Jewelry Clearance"
                cateJewelryC.ParentID = -1
                cateJewelryC.Gender = "W"
                efContext.AddToCategories(cateJewelryC)
            Else
                cateJewelryC = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Jewelry Clearance" And m.SiteID = siteID)
                cateJewelryC.Url = "http://www.sammydress.com/clearance-159.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Bags") Then
                cateHandBags.Category1 = "Bags"
                cateHandBags.SiteID = siteID
                cateHandBags.Url = "http://www.sammydress.com/Wholesale-Bags-Accessories-b-44.html?odr=hot"
                cateHandBags.LastUpdate = nowTime
                cateHandBags.Description = "Bags"
                cateHandBags.ParentID = -1
                cateHandBags.Gender = "W"
                efContext.AddToCategories(cateHandBags)
            Else
                cateHandBags = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Bags" And m.SiteID = siteID)
                cateHandBags.Url = "http://www.sammydress.com/Wholesale-Bags-Accessories-b-44.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Bags Clearance") Then
                cateHandBagsC.Category1 = "Bags Clearance"
                cateHandBagsC.SiteID = siteID
                cateHandBagsC.Url = "http://www.sammydress.com/clearance-44.html?odr=hot"
                cateHandBagsC.LastUpdate = nowTime
                cateHandBagsC.Description = "Bags Clearance"
                cateHandBagsC.ParentID = -1
                cateHandBagsC.Gender = "W"
                efContext.AddToCategories(cateHandBagsC)
            Else
                cateHandBagsC = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Bags Clearance" And m.SiteID = siteID)
                cateHandBagsC.Url = "http://www.sammydress.com/clearance-44.html?odr=hot"
            End If
            If Not listCategoryName.Contains("Swimwear") Then
                cateWeddinEvents.Category1 = "Swimwear"
                cateWeddinEvents.SiteID = siteID
                cateWeddinEvents.Url = "http://www.sammydress.com/Wholesale-Swimwear-c-301.html"
                cateWeddinEvents.LastUpdate = nowTime
                cateWeddinEvents.Description = "Swimwear"
                cateWeddinEvents.ParentID = -1
                cateWeddinEvents.Gender = "W"
                efContext.AddToCategories(cateWeddinEvents)
            Else
                cateWeddinEvents = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Swimwear" And m.SiteID = siteID)
                cateWeddinEvents.Url = "http://www.sammydress.com/Wholesale-Swimwear-c-301.html"
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try

    End Sub

    '将从页面获取的banner分类或者不确定产品类型的分类插入到categories表中
    Public Sub InsertCategories(ByVal siteID As Integer, ByVal categoryName As String, ByVal url As String)
        Try
            Dim queryCategory = From c In efContext.Categories
                                Where c.SiteID = siteID
                                Select c
            Dim listCategoryName As New HashSet(Of String)
            For Each q In queryCategory
                listCategoryName.Add(q.Category1)
            Next

            Dim newCategory As New Category()
            If Not listCategoryName.Contains(categoryName) Then
                newCategory.Category1 = categoryName
                newCategory.SiteID = siteID
                newCategory.Url = url
                newCategory.LastUpdate = DateTime.Now
                newCategory.Description = categoryName
                newCategory.ParentID = -1
                efContext.AddToCategories(newCategory)
            End If

            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try

    End Sub

    '获取主页Promotion Banner的image和Alt，以Banner的图片链接为重复判断，保证一个月内的Banner的image不重复，如果重复，则不获取;Section：PO
    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

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


    '获取主页Weekly Deals的部分，获取折扣最高的2个产品，如果和上周的重复，则不获取;Section: DA
    '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
    Public Sub GetWeeklyDealsOrNewArri(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetWeeklyDealOrNewArriProducts(siteId, Issueid, section, planType) '这个函数中已判断是否和上周重复
            Dim topCount As Integer = 2
            If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
                '
                'getTopWeeklyDealProducts(listProducts, topCount)
                Dim categoryName As String
                Dim categoryId As Integer

                Dim listProduct As New List(Of Product)
                listProduct = GetProductList(siteId)
                Dim listProductId As New List(Of Integer)
                Dim counter As Integer = 0
                For Each li In listProducts
                    categoryName = "Women's Dresses"  'getCategoryName(li.Url)
                    categoryId = GetCategoryId(siteId, categoryName)
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
        Dim Women_s_ClothingUrl As String = ""
        Dim Men_s_ClothingUrl As String = ""
        Dim JewelryUrl As String = ""
        Dim WeddinEventsUrl As String = ""
        Dim BagsUrl As String = ""
        Dim SwimwearUrl As String = ""
        Dim WomenShoesUrl As String = ""
        Dim Sexy_LingerieUrl As String = ""
        Dim WatchesUrl As String = ""
        Dim Beauty_AccessoriesUrl As String = ""
        Dim Home_LivingUrl As String = ""
        Dim ToysUrl As String = ""
        Dim KidsUrl As String = ""
        Try
            '从数据库中获取12个分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()

                    Case "Swimwear"
                        If (HotORNew = "new") Then
                            SwimwearUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            SwimwearUrl = q.Url
                        End If
                    Case "Women's Shoes"
                        If (HotORNew = "new") Then
                            WomenShoesUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            WomenShoesUrl = q.Url
                        End If
                    Case "Bags"
                        If (HotORNew = "new") Then
                            BagsUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            BagsUrl = q.Url
                        End If
                    Case "Jewelry"
                        If (HotORNew = "new") Then
                            JewelryUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            JewelryUrl = q.Url
                        End If
                    Case "Men's Clothing"
                        If (HotORNew = "new") Then
                            Men_s_ClothingUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            Men_s_ClothingUrl = q.Url
                        End If
                    Case "Women's Dresses"
                        If (HotORNew = "new") Then
                            Women_s_ClothingUrl = q.Url.Replace("?odr=hot", "?odr=new")
                        Else
                            Women_s_ClothingUrl = q.Url
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            GetProducts(siteId, Issueid, Women_s_ClothingUrl, "Women's Dresses", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, JewelryUrl, "Jewelry", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, WomenShoesUrl, "Women's Shoes", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, BagsUrl, "Bags", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, SwimwearUrl, "Swimwear", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, Men_s_ClothingUrl, "Men's Clothing", iProIssueCount, updateTime, planType)

        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName)

            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

            Dim endDate As String = Now.AddDays(1).ToShortDateString()
            Dim startDate As String = Now.AddDays(-33).ToShortDateString()
            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, updateTime, categoryId, listProduct)

                If (Not efHelpre.IsProductSent(siteId, li.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
                    listProductId.Add(returnId)
                End If
            Next
            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 获取指定URL页面的产品信息
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection

            productDivs = hd.DocumentNode.SelectNodes("//div[@class='catePro_ListBox']/ul/li")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                product.Prodouct = productDetails.SelectSingleNode("p[@class='proName']").SelectSingleNode("a").InnerText
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 60) Then
                    product.Prodouct = product.Prodouct.Substring(0, 60) & "..."
                End If
                product.Url = productDetails.SelectSingleNode("p[@class='proName']").SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If
                'product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Discount = Double.Parse(productDetails.SelectSingleNode("p[@class='proPrice']").SelectNodes("span")(1).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Currency = productDetails.SelectSingleNode("p[@class='proPrice']").SelectNodes("span")(0).InnerText

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNodeCollection = productDetails.SelectNodes("div[@class='proImgBox']/a")
                product.PictureUrl = productImage(0).SelectSingleNode("img").GetAttributeValue("src", "") '2013/4/7添加
                If (productImage(0).SelectSingleNode("img").GetAttributeValue("src", "").Contains("lazyload")) Then
                    product.PictureUrl = productImage(0).SelectSingleNode("img").GetAttributeValue("data-original", "")
                End If

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct
                'product.SizeHeight = Integer.Parse(productImage.SelectSingleNode("a/img").GetAttributeValue("height", ""))
                Try
                    product.Sales = Integer.Parse(productDetails.SelectSingleNode("div[@class='proImgBox']/span/strong").InnerText)
                Catch
                    product.Sales = 0 '无折扣则为零
                End Try
                If (product.Sales <> 0) Then
                    product.Price = (product.Discount / (100 - product.Sales)) * 100
                End If

                'Dim shipsin24hrs As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/shipin24hrs.jpg"
                'Dim freeshipping As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/freeshipping.jpg"
                'If Not (productDetails.SelectSingleNode("p[@class='p_shipsInHrs']") Is Nothing) Then
                '    product.ShipsImg = shipsin24hrs
                'End If

                'If Not (productDetails.SelectSingleNode("p[@class='p_shipping pt5']") Is Nothing) Then
                '    product.FreeShippingImg = freeshipping
                'End If

                If Not (isSameProductDiffSize(listProducts, product)) Then '避免同一封邮件抓取同款不同size或不同色的产品
                    listProducts.Add(product)
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Function GetWeeklyDealOrNewArriProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String) As List(Of Product)
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String
        If (section = "DA") Then
            startDate = Now.AddDays(-8).ToShortDateString() 'Deals控制一周内不重复
        ElseIf (section = "NE") Then
            startDate = Now.AddDays(-15).ToShortDateString() 'NewArri控制两周内不重复  section[@id='miain']/
        End If
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            If (section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//section[@class='mb10']/div[@class='scrollBox']/textarea/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            ElseIf (section = "NE") Then
                productDivs = hd.DocumentNode.SelectNodes("//section[@class='mb15']/div[@class='scrollBox']/textarea/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            End If
            '获取产品  mb10 newArr
            Dim listProducts As New List(Of Product)
            Dim counter As Integer
            If Not productDivs Is Nothing Then
                For counter = 0 To productDivs.Count - 1
                    Dim product As New Product()
                    Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("p")

                    product.Prodouct = productDetails(1).SelectSingleNode("a").InnerText
                    If (product.Prodouct.ToString.Length > 100) Then
                        product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                    End If
                    product.Url = productDetails(1).SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If
                    'product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                    product.Discount = Double.Parse(productDetails(2).SelectNodes("span")(1).GetAttributeValue("orgp", "").Trim("$").Trim())
                    product.Currency = productDetails(2).SelectNodes("span")(0).InnerText

                    '特殊图片处理
                    product.PictureUrl = productDetails(0).SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                    If (productDetails(0).SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("lazyload.gif")) Then
                        product.PictureUrl = productDetails(0).SelectSingleNode("a/img").GetAttributeValue("data-original", "")
                    End If

                    product.LastUpdate = nowTime
                    product.Description = productDetails(0).SelectSingleNode("a").GetAttributeValue("title", "")
                    product.SiteID = siteId
                    product.PictureAlt = productDetails(0).SelectSingleNode("a/img").GetAttributeValue("alt", "")
                    'product.SizeHeight = Integer.Parse(productDetails(0).SelectSingleNode("a/img").GetAttributeValue("height", ""))
                    'Try
                    '    product.Sales = Integer.Parse(productDetails(0).SelectSingleNode("strong").InnerText)
                    'Catch
                    '    product.Sales = 0 '无折扣则为零
                    'End Try
                    'If (product.Sales <> 0) Then
                    '    product.Price = (product.Discount / (100 - product.Sales)) * 100
                    'End If

                    'Dim shipsin24hrs As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/shipin24hrs.jpg"
                    'Dim freeshipping As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/freeshipping.jpg"
                    'If (productDetails.Count = 4) Then
                    '    If (productDetails(3).GetAttributeValue("class", "").Contains("p_shipsInHrs")) Then
                    '        product.ShipsImg = shipsin24hrs
                    '    ElseIf (productDetails(3).GetAttributeValue("class", "").Contains("p_shipping pt5")) Then
                    '        product.FreeShippingImg = freeshipping
                    '    End If

                    'ElseIf (productDetails.Count >= 5) Then
                    '    If (productDetails(3).GetAttributeValue("class", "").Contains("p_shipsInHrs")) Then
                    '        product.ShipsImg = shipsin24hrs
                    '    ElseIf (productDetails(3).GetAttributeValue("class", "").Contains("p_shipping pt5")) Then
                    '        product.FreeShippingImg = freeshipping
                    '    End If

                    '    If (productDetails(4).GetAttributeValue("class", "").Contains("p_shipsInHrs")) Then
                    '        product.ShipsImg = shipsin24hrs
                    '    ElseIf (productDetails(4).GetAttributeValue("class", "").Contains("p_shipping pt5")) Then
                    '        product.FreeShippingImg = freeshipping
                    '    End If
                    'End If

                    If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的产品不重复：
                        listProducts.Add(product)
                    End If
                Next
            End If
            Return listProducts
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:WeeklyDealsOrNewArri" & " error:" & ex.Message.ToString)
        End Try
    End Function

    Private Function getTopWeeklyDealProducts(ByRef listProducts As List(Of Product), ByVal topCount As Integer)
        '交换排序，按照产品折扣Sales从大到小排序
        Dim maxProduct As New Product()
        Dim temp As New Product()
        Dim k As Integer
        Try
            '只排好折扣最高的前topCount个产品，其余的不关心
            For i As Integer = 0 To topCount - 1
                maxProduct = listProducts(i)
                k = i
                For j As Integer = i + 1 To listProducts.Count - 1
                    If (listProducts(j).Sales > maxProduct.Sales) Then
                        k = j
                        maxProduct = listProducts(k)
                    End If
                Next
                If Not (k = i) Then
                    temp = listProducts(i)
                    listProducts(i) = maxProduct
                    listProducts(k) = temp
                End If
            Next
        Catch ex As Exception
            Throw New Exception("error occured when sort products by sales ")
        End Try
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
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='banner fl']/ul/li/a")
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            Ad.Url = banners(counter).GetAttributeValue("href", "").ToString.Trim()
            If Not (Ad.Url.Trim.ToLower.Contains("http")) Then
                If (Ad.Url.Trim.StartsWith("/")) Then
                    Ad.Url = "http://www.sammydress.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                Else
                    Ad.Url = "http://www.sammydress.com/" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
                End If
            End If
            Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
            If Not (Ad.PictureUrl.Trim.ToLower.Contains("http")) Then
                If (Ad.PictureUrl.Trim.StartsWith("/")) Then
                    Ad.PictureUrl = "http://www.sammydress.com" & banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
                Else
                    Ad.PictureUrl = "http://www.sammydress.com/" & banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
                End If
            End If
            'Ad.SizeHeight = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString)
            'Ad.SizeWidth = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString)
            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelpre.GetListAd(siteid)
            If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                adid = InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
            End If

        Next

        'Try
        '    If (listAdIds.Count = 0) Then '转而去到S站的每周邮件页面（http://www.sammydress.com/html/newsletter-2014-01-06.html ）获取banner
        '        Dim mon As Integer = DateTime.Now.DayOfWeek
        '        pageUrl = "http://www.sammydress.com/html/newsletter-" & DateTime.Now.AddDays(1 - mon).ToString("yyyy-MM-dd") & ".html"
        '        hd = efHelpre.GetHtmlDocument(pageUrl, 60000)
        '        Dim banner As HtmlNode = hd.DocumentNode.SelectSingleNode("/html/body/table[2]/tr[6]")
        '        'Dim banners2 As HtmlNodeCollection = hd.DocumentNode.SelectNodes("/html/body/table/tr/td/table/tr")
        '        'Dim banner As HtmlNode = banners2(0).SelectSingleNode("td/table[1]/tr[6]")

        '        Dim Ad As Ad = New Ad()
        '        Try
        '            Ad.Ad1 = banner.SelectSingleNode("td/a[1]/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
        '            Ad.Description = banner.SelectSingleNode("td/a[1]/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
        '        Catch ex As Exception
        '            'ignore
        '        End Try
        '        Ad.Url = banner.SelectSingleNode("td/a[1]").GetAttributeValue("href", "").ToString.Trim()
        '        Ad.PictureUrl = banner.SelectSingleNode("td/a[1]/img").GetAttributeValue("src", "").ToString
        '        Ad.SiteID = siteid
        '        Ad.Lastupdate = nowTime

        '        Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
        '        listAd = efHelpre.GetListAd(siteid)
        '        If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate)) Then '控制一个月内获取的Ad不重复：
        '            adid = InsertAd(Ad, listAd, categroyId)
        '            listAdIds.Add(adid)
        '        End If
        '    End If
        'Catch ex As Exception
        '    'ignore
        'End Try

        Return listAdIds
    End Function

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

    Public Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                          ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String,
                          ByVal discount As Double, ByVal freeShips As String, ByVal shipINHrs As String, ByVal categoryId As Integer,
                          ByVal list As List(Of Product)) As Integer
        url = url.Trim()
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = myProduct
        product.Url = url
        product.Price = price
        product.PictureUrl = pictureUrl
        product.LastUpdate = lastUpdate
        product.Description = description
        product.SiteID = siteId
        product.Currency = currency
        product.PictureAlt = pictureAlt
        product.Discount = discount
        product.FreeShippingImg = freeShips
        product.ShipsImg = shipINHrs

        If (JudgeProduct(product.Url, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            updateProduct.Discount = product.Discount
            updateProduct.FreeShippingImg = product.FreeShippingImg
            updateProduct.ShipsImg = product.ShipsImg
            'Dim returnId As Integer = updateProduct.ProdouctID
            'efContext.SaveChanges()

            '2014/2/21新增，防止一个产品有多个productCategory关系
            Dim updateCategory = updateProduct.Categories

            'Dim queryCate = From p In efContext.Products
            '                Where p.ProdouctID = updateProduct.ProdouctID
            '                Select p
            'Dim cate = queryCate.Single.Categories
            Dim counter As Integer = updateCategory.Count
            For i As Integer = 0 To counter - 1
                updateCategory(0).Products.Remove(updateProduct)
            Next
            efContext.SaveChanges()

            If Not updateCategory.Contains(category) Then
                category.Products.Add(updateProduct)
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
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

    ''' <summary>
    ''' 获得邮件的subject
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal categoryName As String)
        Dim preSubject As String
        Dim subject As String
        Dim addNum As Integer
        If (planType.Contains("HO")) Then
            preSubject = "Sammydress Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "Sammydress New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HP")) Then
            preSubject = "Sammydress Special " & categoryName & " For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If

        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        If (planType.Trim = "HO" OrElse planType.Trim = "HA" OrElse planType.Trim.Contains("HP")) Then
            addNum = 1
        Else
            addNum = 0
        End If
        If (planType.Contains("HO") OrElse planType.Contains("HA")) Then
            planType = planType.Substring(0, 2)
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")
        Dim query = From i In efContext.Issues
                     Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
        subject = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")
        'Promition Banner有 Alt文字时：Hi [FIRSTNAME], WEEKLY DEALS FOR YOU: + [Alt文字, Weekly Deals第一件产品名称]
        'Promition Banner没有 Alt文字时：Hi [FIRSTNAME], WEEKLY DEALS FOR YOU: + [Weekly Deals第一件产品名称]
        'Dim subject As String
        'Dim querySubject = From p In efContext.Ads
        '                 Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
        '                 Where pi.IssueID = issueId AndAlso pi.SiteId = siteId
        '                 Select p.Ad1
        'If Not String.IsNullOrEmpty(querySubject.FirstOrDefault()) Then
        '    subject = preSubject & querySubject.FirstOrDefault()
        'Else
        '    querySubject = From p In efContext.Products
        '                 Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '                 Where pi.IssueID = issueId AndAlso pi.SectionID = section AndAlso pi.SiteId = siteId
        '                 Select p.Prodouct
        '    If Not String.IsNullOrEmpty(querySubject.FirstOrDefault()) Then
        '        subject = preSubject & querySubject.FirstOrDefault()
        '    Else
        '        querySubject = From p In efContext.Products
        '                 Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '                 Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
        '                 Select p.Prodouct
        '        If Not String.IsNullOrEmpty(querySubject.FirstOrDefault()) Then
        '            subject = preSubject & querySubject.FirstOrDefault()
        '        End If
        '    End If
        'End If
        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub


    Private Function getCategoryName(ByVal url As String) As String
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(url, 60000)
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@class='path']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
        Dim categoryName As String = productDivs.SelectNodes("a")(2).InnerText
        If (categoryName.ToLower.Contains("wedding")) Then
            categoryName = "Wedding & Events Clearance"
        ElseIf (categoryName.ToLower.Contains("wedding")) Then
        End If
        Select Case categoryName.Trim
            Case "Women's Clothing"
                categoryName = "Women's Dresses"
            Case "Shoes"
                categoryName = "Women's Shoes"
            Case "Bags"
                categoryName = "Women's Handbags"
        End Select
        Return categoryName
    End Function

    Public Function GetPromotionBannerForNewArri(ByVal siteID As Integer, ByVal IssueID As Integer, ByVal planType As String)
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Dim pageURL As String = "http://www.sammydress.com/new-products.html"
        Try
            Dim listAdId As List(Of Integer) = GetListAdidsForNewArri(pageURL, siteID, updateTime, planType)

            InsertAdsIssue(siteID, IssueID, listAdId, iProIssueCount)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetListAdidsForNewArri(ByVal pageURL As String, ByVal siteID As Integer, ByVal nowTime As DateTime, ByVal planType As String)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-64).ToShortDateString() ''控制两个个月内获取的Ad不重复：
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(pageURL, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//ul[@id='js_slideBox_Imgshow']/li/a")
        Dim categroyName As String = ""
        Dim categroyId As Integer
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            Ad.Url = "http://www.sammydress.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
            categroyName = getbannerCategoryNameForNewArri(Ad.Url)
            InsertCategories(siteID, categroyName, "")
            categroyId = GetCategoryId(siteID, categroyName)
            Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
            Ad.SizeHeight = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString)
            Ad.SizeWidth = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString)
            Ad.SiteID = siteID
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelpre.GetListAd(siteID)
            If (Not efHelpre.isAdSent(siteID, Ad.Url, startDate, endDate, planType)) Then '控制两个个月内获取的Ad不重复：
                adid = InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
            End If
        Next
        Return listAdIds
    End Function

    Private Function getbannerCategoryNameForNewArri(ByVal url As String) As String
        Dim hd As HtmlDocument = efHelpre.GetHtmlDocument(url, 60000)
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@id='mainWrap']/div[@class='path']")
        Dim aHerfs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='mainWrap']/div[@class='path']/a")
        If (productDivs Is Nothing) Then
            productDivs = hd.DocumentNode.SelectSingleNode("//div[@id='mainWrap']/div[@class='path goodsPath']")
            aHerfs = hd.DocumentNode.SelectNodes("//div[@id='mainWrap']/div[@class='path goodsPath']/a")
        End If
        Dim categoryName As String = ""

        Dim index = productDivs.InnerText.IndexOf(";")
        If (aHerfs.Count = 1) Then
            categoryName = productDivs.InnerText.Substring(index + 1, productDivs.InnerText.Length - index - 1).Trim()
        ElseIf (aHerfs.Count = 2) Then
            categoryName = aHerfs(1).InnerText
        End If
        If (categoryName.ToLower.Contains("sexy")) Then
            categoryName = "Sexy Lingerie"
        ElseIf (categoryName.ToLower.Contains("wedding")) Then
            categoryName = "Wedding & Events"
        ElseIf (categoryName.ToLower.Contains("shoes")) Then
            categoryName = "Women's Shoes"
        ElseIf (categoryName.ToLower.Contains("bag")) Then
            categoryName = "Women's Handbags"
        ElseIf (categoryName.ToLower.Contains("jewelry")) Then
            categoryName = "Jewelry"
        Else
            categoryName = "Women's Dresses"
        End If

        Return categoryName
    End Function

    Private Function isSameProductDiffSize(ByRef listProducts As List(Of Product), ByVal product As Product) As Boolean
        For Each pro In listProducts
            If (pro.Description.ToString.Equals(product.Description.ToString)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Sub CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal saveListName As String, ByVal daySpan As Integer, _
                                         ByVal strategy As ChooseStrategy, ByVal sendStatus As String, ByVal loginEmail As String, ByVal appId As String)
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(-daySpan).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = Date.Now.ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()
        Try
            mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
        'Dim arrContactList As String() = New String() {saveListName}
        Dim helper As New EFHelper
        helper.InsertContactList(issueId, saveListName, sendStatus)
    End Sub

    Public Shared Sub CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal categoryName As String, ByVal saveListName As String, _
                                         ByVal daySpan As Integer, ByVal strategy As ChooseStrategy, ByVal sendStatus As String, ByVal loginEmail As String, ByVal appId As String)
        Dim categoryId As String = EFHelper.GetCategoryId(siteId, categoryName.Trim())
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Favorite = categoryId
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(-daySpan).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = Date.Now.ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()
        Try
            mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
        Catch ex As Exception
            Throw New Exception(ex.ToString())

        End Try
        'Dim arrContactList As String() = New String() {saveListName}
        Dim helper As New EFHelper
        helper.InsertContactList(issueId, saveListName, sendStatus)
    End Sub

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
