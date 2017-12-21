Imports HtmlAgilityPack

Public Class DresslilyNew
    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper()
    Private pageUrl As String = "http://www.dresslily.com/"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
               ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        InsertCategories(siteId)
        Dim contactlistCount As Integer
        Dim bannerCategoryName As String
        If planType.Contains("HO") Then
            'Website analysis
            bannerCategoryName = "Women's Dresses2" '将banner条的类名放到Women's Clothing一类
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType) '获取主页Daily Deals 的部分，获取折扣最高的2个产品，如果和上周的重复，则不获取;Section: DA
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "DA")
            ' ''Add 
            'Dim contactListArr() As String = {"Opens", "OpensHO20140224", "D_SPREAD_001"}
            'efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

            If (planType.Trim = "HO") Then
                Dim contactListArr() As String = {"Opens20140402", "Opens"}
                efHelpre.InsertContactList(IssueID, contactListArr, "waiting") 'draft  'waiting

                'Dim saveListName As String = "dresslily_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                'Dim sendingStatus As String = "draft"

                'contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

                'saveListName = "dresslily_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")

                'contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.Open, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

            ElseIf (planType.Trim.Contains("HO") AndAlso planType.Trim <> "HO") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "dresslily_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                If (planType.Trim = "HO1") Then
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                    'HO3
                    saveListName = "dresslily_" & planType.Trim & "Women's Outerwear" & Now.ToString("yyyyMMdd")
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, "Women's Outerwear2", saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                    'HO4
                    saveListName = "dresslily_" & planType.Trim & "Women's Bottoms" & Now.ToString("yyyyMMdd")
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, "Women's Bottoms2", saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                ElseIf (planType.Trim <> "HO3" AndAlso planType.Trim <> "HO4") Then
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                End If
                
            End If
        ElseIf planType.Contains("HA") Then
            '分析html获取到banner条的分类
            '获取banner信息插入到ads表中
            'GetPromotionBannerForNewArri(siteId, IssueID)'没有Banner
            GetWeeklyDealsOrNewArri(siteId, IssueID, "NE", planType) '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
            GetHotOrNewProducts(siteId, IssueID, "new", planType) '在获取分类url时，直接将最后的hot替换成new后直接使用
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "NE")
            ' ''Add 
            'Dim contactListArr() As String = {"Opens", "OpensHO20140224", "D_SPREAD_001"}
            'efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

            If (planType.Trim = "HA") Then
                Dim contactListArr() As String = {"Opens20140402", "Opens"}
                efHelpre.InsertContactList(IssueID, contactListArr, "waiting") 'draft  'waiting

                'Dim saveListName As String = "dresslily_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                'Dim sendingStatus As String = "draft"

                'contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If


                'saveListName = "dresslily_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")
                'contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.Open, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

            ElseIf (planType.Trim.Contains("HA") AndAlso planType.Trim <> "HA") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "dresslily" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                If (planType.Trim = "HA1") Then
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                    'HA3
                    saveListName = "dresslily_" & planType.Trim & "Women's Outerwear" & Now.ToString("yyyyMMdd")
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, "Women's Outerwear2", saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                    'HA4
                    saveListName = "dresslily_" & planType.Trim & "Women's Bottoms" & Now.ToString("yyyyMMdd")
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, "Women's Bottoms2", saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                ElseIf (planType.Trim <> "HA3" AndAlso planType.Trim <> "HA4") Then
                    contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                    If (contactlistCount > 0) Then
                        efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                    End If
                End If
            End If
        ElseIf planType.Contains("HP") Then
            Dim cate As String() = categories.Split("^")

            Dim categoryUrlC As String
            Dim categoryUrlL As String
            Dim iProIssueCount As Integer = 8

            Select Case cate(0).Trim
                Case "Women's Dresses2"
                    categoryUrlL = "http://www.dresslily.com/dresses-2014-c-167.html"
                    categoryUrlC = "http://www.dresslily.com/tees-t-shirts-c-25.html"
                Case "Women's Tops2"
                    categoryUrlL = "http://www.dresslily.com/tees-t-shirts-c-25.html"
                    categoryUrlC = "http://www.dresslily.com/dresses-2014-c-167.html"
                Case "Women's Outerwear2"
                    categoryUrlL = "http://www.dresslily.com/blazers-c-33.html"
                    categoryUrlC = "http://www.dresslily.com/women-s-pumps-c-83.html"
                Case "Women's Bottoms2"
                    categoryUrlL = "http://www.dresslily.com/jumpsuits-c-42.html"
                    categoryUrlC = "http://www.dresslily.com/dresses-2014-c-167.html"
                Case "Bags2"
                    categoryUrlL = "http://www.dresslily.com/tote-bags-c-116.html"
                    categoryUrlC = "http://www.dresslily.com/earrings-c-58.html"
                Case "Shoes2"
                    categoryUrlL = "http://www.dresslily.com/women-s-pumps-c-83.html"
                    categoryUrlC = "http://www.dresslily.com/dresses-2014-c-167.html"
                Case "Jewelry2"
                    categoryUrlL = "http://www.dresslily.com/earrings-c-58.html"
                    categoryUrlC = "http://www.dresslily.com/tote-bags-c-116.html"
                Case "Lingerie2"
                    categoryUrlL = "http://www.dresslily.com/sexy-dresses-c-100.html"
                    categoryUrlC = "http://www.dresslily.com/peep-toe-c-89.html"
                Case "Watches2"
                    categoryUrlL = "http://www.dresslily.com/women-s-watches-c-124.html"
                    categoryUrlC = "http://www.dresslily.com/earrings-c-58.html"
            End Select

            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0).Trim.Replace("2", ""))
            ''Add 
            Dim sendingStatus = "waiting"
            Dim saveListName As String = "dresslilySpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
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

            Dim aCategory As Category
            '添加9大分类 
            Dim nowTime As DateTime = Now
            Dim cateWomen_s_Dresses As New Category()
            Dim cateWomen_s_Clothing As New Category()
            Dim cateWomen_s_Tops As New Category()
            Dim cateWomen_s_Outerwear As New Category()
            Dim cateWomen_s_Bottoms As New Category()
            Dim cateBags As New Category()
            Dim cateShoes As New Category()
            Dim cateJewelry As New Category()
            Dim cateLingerie As New Category()
            Dim cateWatches As New Category()
            'Dim cateKids As New Category()
            If Not listCategoryName.Contains("Women's Dresses2") Then
                cateWomen_s_Dresses.Category1 = "Women's Dresses2"
                cateWomen_s_Dresses.SiteID = siteID
                cateWomen_s_Dresses.Url = "http://www.dresslily.com/dresses-c-6.html"
                cateWomen_s_Dresses.LastUpdate = nowTime
                cateWomen_s_Dresses.Description = "Women's Dresses2"
                cateWomen_s_Dresses.ParentID = -1
                cateWomen_s_Dresses.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Dresses)
            End If
            If Not listCategoryName.Contains("Women's Tops2") Then
                cateWomen_s_Tops.Category1 = "Women's Tops2"
                cateWomen_s_Tops.SiteID = siteID
                cateWomen_s_Tops.Url = "http://www.dresslily.com/tops-c-24.html"
                cateWomen_s_Tops.LastUpdate = nowTime
                cateWomen_s_Tops.Description = "Women's Tops2"
                cateWomen_s_Tops.ParentID = -1
                cateWomen_s_Tops.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Tops)
            End If
            If Not listCategoryName.Contains("Women's Outerwear2") Then
                cateWomen_s_Outerwear.Category1 = "Women's Outerwear2"
                cateWomen_s_Outerwear.SiteID = siteID
                cateWomen_s_Outerwear.Url = "http://www.dresslily.com/outerwear-c-31.html"
                cateWomen_s_Outerwear.LastUpdate = nowTime
                cateWomen_s_Outerwear.Description = "Women's Outerwear2"
                cateWomen_s_Outerwear.ParentID = -1
                cateWomen_s_Outerwear.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Outerwear)
            End If
            If Not listCategoryName.Contains("Women's Bottoms2") Then
                cateWomen_s_Bottoms.Category1 = "Women's Bottoms2"
                cateWomen_s_Bottoms.SiteID = siteID
                cateWomen_s_Bottoms.Url = "http://www.dresslily.com/bottom-c-36.html"
                cateWomen_s_Bottoms.LastUpdate = nowTime
                cateWomen_s_Bottoms.Description = "Women's Bottoms2"
                cateWomen_s_Bottoms.ParentID = -1
                cateWomen_s_Bottoms.Gender = "W"
                efContext.AddToCategories(cateWomen_s_Bottoms)
            Else
            End If
            If Not listCategoryName.Contains("Bags2") Then
                cateBags.Category1 = "Bags2"
                cateBags.SiteID = siteID
                cateBags.Url = "http://www.dresslily.com/bags-b-3.html"
                cateBags.LastUpdate = nowTime
                cateBags.Description = "Bags2"
                cateBags.ParentID = -1
                cateBags.Gender = "W"
                efContext.AddToCategories(cateBags)
            End If
            If Not listCategoryName.Contains("Shoes2") Then
                cateShoes.Category1 = "Shoes2"
                cateShoes.SiteID = siteID
                cateShoes.Url = "http://www.dresslily.com/women-s-shoes-b-81.html"
                cateShoes.LastUpdate = nowTime
                cateShoes.Description = "Shoes2"
                cateShoes.ParentID = -1
                cateShoes.Gender = "W"
                efContext.AddToCategories(cateShoes)
            End If

            If Not listCategoryName.Contains("Jewelry2") Then
                cateJewelry.Category1 = "Jewelry2"
                cateJewelry.SiteID = siteID
                cateJewelry.Url = "http://www.dresslily.com/jewelry-b-56.html"
                cateJewelry.LastUpdate = nowTime
                cateJewelry.Description = "Jewelry2"
                cateJewelry.ParentID = -1
                cateJewelry.Gender = "W"
                efContext.AddToCategories(cateJewelry)
            End If
            If Not listCategoryName.Contains("Lingerie2") Then
                cateLingerie.Category1 = "Lingerie2"
                cateLingerie.SiteID = siteID
                cateLingerie.Url = "http://www.dresslily.com/sexy-lingerie-b-97.html"
                cateLingerie.LastUpdate = nowTime
                cateLingerie.Description = "Lingerie2"
                cateLingerie.ParentID = -1
                cateLingerie.Gender = "W"
                efContext.AddToCategories(cateLingerie)
            End If

            If Not listCategoryName.Contains("Watches2") Then
                cateWatches.Category1 = "Watches2"
                cateWatches.SiteID = siteID
                cateWatches.Url = "http://www.dresslily.com/watches-b-123.html"
                cateWatches.LastUpdate = nowTime
                cateWatches.Description = "Watches2"
                cateWatches.ParentID = -1
                cateWatches.Gender = "W"
                efContext.AddToCategories(cateWatches)
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


    '获取主页Dialy Deals的部分，获取折扣最高的2个产品，如果和上周的重复，则不获取;Section: DA
    '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
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
                    categoryName = "Women's Dresses2"
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


    '根据网站的9个分类获取热销部分，请按顺序获取，保证月内不重复，Section：CA
    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim Women_s_DressesUrl As String = ""
        Dim Women_s_TopsUrl As String = ""
        Dim Women_s_OuterwearUrl As String = ""
        Dim Women_s_BottomsUrl As String = ""
        Dim BagsUrl As String = ""
        Dim ShoesUrl As String = ""
        Dim JewelryUrl As String = ""
        Dim LingerieUrl As String = ""
        Dim WatchesUrl As String = ""

        Try
            '从数据库中获取9个分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Women's Dresses2"
                        If (HotORNew = "new") Then
                            Women_s_DressesUrl = q.Url & "?odr=new"
                        Else
                            Women_s_DressesUrl = q.Url & "?odr=hot"
                        End If
                    Case "Women's Tops2"
                        If (HotORNew = "new") Then
                            Women_s_TopsUrl = q.Url & "?odr=new"
                        Else
                            Women_s_TopsUrl = q.Url & "?odr=hot"
                        End If
                    Case "Women's Outerwear2"
                        If (HotORNew = "new") Then
                            Women_s_OuterwearUrl = q.Url & "?odr=new"
                        Else
                            Women_s_OuterwearUrl = q.Url & "?odr=hot"
                        End If
                    Case "Women's Bottoms2"
                        If (HotORNew = "new") Then
                            Women_s_BottomsUrl = q.Url & "?odr=new"
                        Else
                            Women_s_BottomsUrl = q.Url & "?odr=hot"
                        End If
                    Case "Bags2"
                        If (HotORNew = "new") Then
                            BagsUrl = q.Url & "?odr=new"
                        Else
                            BagsUrl = q.Url & "?odr=hot"
                        End If
                    Case "Shoes2"
                        If (HotORNew = "new") Then
                            ShoesUrl = q.Url & "?odr=new"
                        Else
                            ShoesUrl = q.Url & "?odr=hot"
                        End If
                    Case "Jewelry2"
                        If (HotORNew = "new") Then
                            JewelryUrl = q.Url & "?odr=new"
                        Else
                            JewelryUrl = q.Url & "?odr=hot"
                        End If
                    Case "Lingerie2"
                        If (HotORNew = "new") Then
                            LingerieUrl = q.Url & "?odr=new"
                        Else
                            LingerieUrl = q.Url & "?odr=hot"
                        End If
                    Case "Watches2"
                        If (HotORNew = "new") Then
                            WatchesUrl = q.Url & "?odr=new"
                        Else
                            WatchesUrl = q.Url & "?odr=hot"
                        End If
                End Select
            Next

            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 2
            GetProducts(siteId, Issueid, Women_s_DressesUrl, "Women's Dresses2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, Women_s_TopsUrl, "Women's Tops2", iProIssueCount, updateTime, planType)
            iProIssueCount = 1
            GetProducts(siteId, Issueid, Women_s_OuterwearUrl, "Women's Outerwear2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, Women_s_BottomsUrl, "Women's Bottoms2", iProIssueCount, updateTime, planType)

            iProIssueCount = 3
            GetProducts(siteId, Issueid, BagsUrl, "Bags2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, ShoesUrl, "Shoes2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, JewelryUrl, "Jewelry2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, LingerieUrl, "Lingerie2", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, WatchesUrl, "Watches2", iProIssueCount, updateTime, planType)

        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 添加数据到Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="section"></param>
    ''' <remarks></remarks>
    Private Sub InsertProductIssueAll(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal section As String, ByVal listProductId As List(Of Integer))
        For Each li In listProductId
            Dim pIssue As New Products_Issue()
            pIssue.ProductId = li
            pIssue.SiteId = siteId
            pIssue.IssueID = IssueId
            pIssue.SectionID = section
            'If (JudgeProductIssue(pIssue.ProductId, pIssue.SiteId, pIssue.IssueDate, listProductIssue)) Then
            efContext.AddToProducts_Issue(pIssue)
            'End If
        Next
        efContext.SaveChanges()
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            '此方法已控制一个月内获取的新品不重复：
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, iProIssueCount, planType)

            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)


            For Each li In listProducts
                Dim returnId As Integer = efHelpre.InsertProduct(li, Now, categoryId, listProduct)
                listProductId.Add(returnId)
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
    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                                    ByVal planType As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@class='catePro_ListBox']/ul/li")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                'Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectSingleNode("ul").SelectNodes("li")
                product.Prodouct = productDetails.SelectSingleNode("p[@class='proName']").SelectSingleNode("a").InnerText
                'product.Prodouct = productDetails(1).SelectSingleNode("h4").SelectSingleNode("a").InnerText
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 60) Then
                    product.Prodouct = product.Prodouct.Substring(0, 60) & "..."
                End If

                product.Url = productDetails.SelectSingleNode("p[@class='proName']").SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                If Not (product.Url.Contains("http:")) Then
                    If (product.Url.StartsWith("/")) Then
                        product.Url = "http://www.dresslily.com" & product.Url
                    Else
                        product.Url = "http://www.dresslily.com/" & product.Url
                    End If
                End If
                'product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                'product.Price = Double.Parse(productDetails.SelectSingleNode("p[@class='proPrice']").SelectSingleNode("span[@class='my_shop_price']").GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Discount = Double.Parse(productDetails.SelectSingleNode("p[@class='proPrice']").SelectSingleNode("span[@class='my_shop_price']").GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Currency = "$"

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("div[@class='proImgBox']")
                product.PictureUrl = productImage.SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                If (productImage.SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("lazy")) Then
                    product.PictureUrl = productImage.SelectSingleNode("a/img").GetAttributeValue("data-original", "")
                End If

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct

                'Try
                '    product.Sales = Integer.Parse(productImage.SelectSingleNode("strong").InnerText)
                'Catch
                '    product.Sales = 0 '无折扣则为零
                product.Sales = 0
                'End Try


                If Not (isSameProductDiffSize(listProducts, product)) Then '避免同一封邮件抓取同款不同size或不同色的产品
                    If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, PlanType)) Then '控制一个月内获取的新品不重复：
                        listProducts.Add(product)
                    End If
                End If
                If (listProducts.Count >= iProIssueCount * 2) Then
                    Exit For
                End If

            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Function GetWeeklyDealOrNewArriProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String) As List(Of Product)
        Dim helper As New EFHelper
        Dim pageUrl As String = "http://www.dresslily.com"
        Dim nowTime As DateTime = Now
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String
        startDate = Now.AddDays(-8).ToShortDateString() '控制一周内不重复
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            If (section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//section[@class='mb10']/div/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            ElseIf (section = "NE") Then
                productDivs = hd.DocumentNode.SelectNodes("//section[@class='mb15']/div/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            End If
            '获取产品  mb10 newArr
            Dim listProducts As New List(Of Product)
            Dim counter As Integer
            If Not productDivs Is Nothing Then
                For counter = 0 To productDivs.Count - 1
                    Dim product As New Product()
                    Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("p")
                    product.Prodouct = productDivs(counter).SelectSingleNode("p[@class='proName']/a").InnerText
                    If (product.Prodouct.ToString.Length > 100) Then
                        product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                    End If
                    product.Url = pageUrl + productDivs(counter).SelectSingleNode("p[@class='proName']/a").GetAttributeValue("href", "").Trim()
                    'product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                    product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("p[@class='proPrice']/span[@class='my_shop_price fb f14']").GetAttributeValue("orgp", "").Trim("$").Trim())

                    product.Currency = "$"
                    '特殊图片处理
                    product.PictureUrl = productDivs(counter).SelectSingleNode("p[@class='tc']/a/img").GetAttributeValue("src", "") '2013/4/7添加
                    If (productDivs(counter).SelectSingleNode("p[@class='tc']/a/img").GetAttributeValue("src", "").Contains("load")) Then
                        product.PictureUrl = productDivs(counter).SelectSingleNode("p[@class='tc']/a/img").GetAttributeValue("data-original", "")
                    End If

                    product.LastUpdate = nowTime
                    product.Description = product.Prodouct
                    product.SiteID = siteId
                    product.PictureAlt = product.Prodouct

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
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='js_banner']/ul/li/a")
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            Ad.Url = "http://www.dresslily.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
            Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
            'If (banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString <> "") Then
            '    Ad.SizeHeight = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("height", "").ToString)
            'End If
            'If (banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString <> "") Then
            '    Ad.SizeWidth = Integer.Parse(banners(counter).SelectSingleNode("img").GetAttributeValue("width", "").ToString)
            'End If

            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelpre.GetListAd(siteid)
            If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                adid = efHelpre.InsertAd(Ad, listAd, categroyId)
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

            '2014/1/21新增，防止一个产品有多个productCategory关系
            Dim updateCategory = updateProduct.Categories

            Dim queryCate = From p In efContext.Products
                            Where p.ProdouctID = updateProduct.ProdouctID
                            Select p
            Dim cate = queryCate.Single.Categories
            Dim counter As Integer = cate.Count
            For i As Integer = 0 To counter - 1
                cate(0).Products.Remove(updateProduct)
            Next
            efContext.SaveChanges()

            If Not cate.Contains(category) Then
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
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then
            preSubject = "Dresslily Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "Dresslily New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HP")) Then
            preSubject = "Dresslily Specials " & categoryName & " For "
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
        Dim subject As String = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub




    Public Function GetPromotionBannerForNewArri(ByVal siteID As Integer, ByVal IssueID As Integer, ByVal planType As String)
        Dim updateTime As DateTime = Now
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Dim pageURL As String = "http://www.dresslily.com/new-products.html"
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
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//section[@id='miain']/div[@class='newArrivalTop clearfix']/div[@class='na_slideBox fl']/ul/li/a")
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
        Dim categoryName As String = ""

        Dim index = productDivs.InnerText.IndexOf(";")
        If (aHerfs.Count = 1) Then
            categoryName = productDivs.InnerText.Substring(index + 1, productDivs.InnerText.Length - index - 1).Trim()
        ElseIf (aHerfs.Count = 2) Then
            categoryName = aHerfs(1).InnerText
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

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
