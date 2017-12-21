Imports HtmlAgilityPack
Imports System.Security.Policy

Public Class AhappyDeal

    Private efContext As New EmailAlerterEntities()
    Private efHelpre As New EFHelper
    Private Flage1 As Boolean = vbTrue
    Private Flage2 As Boolean = vbTrue
    Private Flage3 As Boolean = vbTrue
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        Dim contactlistCount As Integer
        If planType.Trim.Contains("HO") Then
            bannerCategoryName = "Cell Phones /Mobiles"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType) '获取主页Weekly Deals的部分，顺序获取2个产品，如果和上周的重复，则不获取;Section: DA
            GetHotOrNewProducts(siteId, IssueID, "hot", planType.Trim)
            ''Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "DA")
            ' ''Add 
            If (planType.Trim = "HO") Then
                'Dim contactListArr() As String = {"A_SPREAD_001"}
                'efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting
                Dim saveListName As String = "Ahappydeal_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If

                saveListName = "Ahappydeal_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")

                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If

            ElseIf (planType.Trim.Contains("HO") AndAlso planType.Trim <> "HO") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "Ahappydeal_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If
            End If
        ElseIf (planType.Trim.Contains("HA")) Then
            '分析html获取到banner条的分类
            '获取banner信息插入到ads表中
            'GetPromotionBannerForNewArri(siteId, IssueID)
            'GetWeeklyDealsOrNewArri(siteId, IssueID, "NE") '随机获取2个产品, 并且获取到该产品所属的分类，如果和上周的重复，则不获取，Section: NE
            GetHotOrNewProducts(siteId, IssueID, "new", planType.Trim) '在获取分类url时，直接将最后的hot替换成new后直接使用
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "NE")
            ''Add 
            If (planType.Trim = "HA") Then
                'Dim contactListArr() As String = {"A_SPREAD_001"}
                'efHelpre.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting
                Dim saveListName As String = "Ahappydeal_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If


                saveListName = "Ahappydeal_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")
                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If

            ElseIf (planType.Trim.Contains("HA") AndAlso planType.Trim <> "HA") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "Ahappydeal_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelpre.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelpre.InsertContactList(IssueID, saveListName, sendingStatus)
                End If
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
                        categoryUrlL = q.Url & "?sortby=reviews"
                    Case cate(1).Trim
                        categoryUrlC = q.Url & "?sortby=reviews"
                End Select
            Next
            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "AhappySpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
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
            '添加banner这一分类
            Dim cateBanner As New Category()
            If Not listCategoryName.Contains("Banner") Then
                cateBanner.Category1 = "Banner"
                cateBanner.SiteID = siteID
                cateBanner.Url = "http://www.dealsmachine.com/"
                cateBanner.LastUpdate = nowTime
                cateBanner.Description = "Banner"
                cateBanner.ParentID = -1
                efContext.AddToCategories(cateBanner)
            End If
            '添加14大分类 
            Dim cateCell_Phones_Mobiles As New Category()
            Dim cateNotebook_Tablet_PC As New Category()
            Dim cateMobile_Power_Bank As New Category()
            Dim cateHeadphones As New Category()
            Dim cateComputer_Peripherals As New Category()
            Dim cateiPhone_iPad_iPod As New Category()
            Dim cateConsumer_Electronics As New Category()
            Dim cateCar_Electronics As New Category()
            Dim cateWatches As New Category()
            Dim cateFlashlights_LED_lights As New Category()
            Dim cateOutdoors_Sports As New Category()
            Dim cateCell_Phone_Accessories As New Category()
            Dim cateHome_Office As New Category()
            Dim cateHobbies_Toys As New Category()
            If Not listCategoryName.Contains("Cell Phones /Mobiles") Then
                cateCell_Phones_Mobiles.Category1 = "Cell Phones /Mobiles"
                cateCell_Phones_Mobiles.SiteID = siteID
                cateCell_Phones_Mobiles.Url = "http://www.dealsmachine.com/cell-phones-mobiles-c-59.html"
                cateCell_Phones_Mobiles.LastUpdate = nowTime
                cateCell_Phones_Mobiles.Description = "Cell Phones /Mobiles"
                cateCell_Phones_Mobiles.ParentID = -1
                efContext.AddToCategories(cateCell_Phones_Mobiles)
            Else
                cateCell_Phones_Mobiles = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Cell Phones /Mobiles" And m.SiteID = siteID)
                cateCell_Phones_Mobiles.Url = "http://www.dealsmachine.com/cell-phones-mobiles-c-59.html"
            End If
            If Not listCategoryName.Contains("Notebook/Tablet PC") Then
                cateNotebook_Tablet_PC.Category1 = "Notebook/Tablet PC"
                cateNotebook_Tablet_PC.SiteID = siteID
                cateNotebook_Tablet_PC.Url = "http://www.dealsmachine.com/notebook-umpc-c-95.html"
                cateNotebook_Tablet_PC.LastUpdate = nowTime
                cateNotebook_Tablet_PC.Description = "Notebook/Tablet PC"
                cateNotebook_Tablet_PC.ParentID = -1
                efContext.AddToCategories(cateNotebook_Tablet_PC)
            Else
                cateNotebook_Tablet_PC = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Notebook/Tablet PC" And m.SiteID = siteID)
                cateNotebook_Tablet_PC.Url = "http://www.dealsmachine.com/notebook-umpc-c-95.html"
            End If
            If Not listCategoryName.Contains("Mobile Power Bank") Then
                cateMobile_Power_Bank.Category1 = "Mobile Power Bank"
                cateMobile_Power_Bank.SiteID = siteID
                cateMobile_Power_Bank.Url = "http://www.dealsmachine.com/mobile-power-bank-c-1832.html"
                cateMobile_Power_Bank.LastUpdate = nowTime
                cateMobile_Power_Bank.Description = "Mobile Power Bank"
                cateMobile_Power_Bank.ParentID = -1
                efContext.AddToCategories(cateMobile_Power_Bank)
            Else
                cateMobile_Power_Bank = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Mobile Power Bank" And m.SiteID = siteID)
                cateMobile_Power_Bank.Url = "http://www.dealsmachine.com/mobile-power-bank-c-1832.html"
            End If
            If Not listCategoryName.Contains("Computer Peripherals") Then
                cateComputer_Peripherals.Category1 = "Computer Peripherals"
                cateComputer_Peripherals.SiteID = siteID
                cateComputer_Peripherals.Url = "http://www.dealsmachine.com/computer-peripherals-c-86.html"
                cateComputer_Peripherals.LastUpdate = nowTime
                cateComputer_Peripherals.Description = "Computer Peripherals"
                cateComputer_Peripherals.ParentID = -1
                efContext.AddToCategories(cateComputer_Peripherals)
            Else
                cateComputer_Peripherals = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Computer Peripherals" And m.SiteID = siteID)
                cateComputer_Peripherals.Url = "http://www.dealsmachine.com/computer-peripherals-c-86.html"
            End If
            If Not listCategoryName.Contains("Consumer Electronics") Then
                cateConsumer_Electronics.Category1 = "Consumer Electronics"
                cateConsumer_Electronics.SiteID = siteID
                cateConsumer_Electronics.Url = "http://www.dealsmachine.com/consumer-electronics-c-1782.html"
                cateConsumer_Electronics.LastUpdate = nowTime
                cateConsumer_Electronics.Description = "Consumer Electronics"
                cateConsumer_Electronics.ParentID = -1
                efContext.AddToCategories(cateConsumer_Electronics)
            Else
                cateConsumer_Electronics = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Consumer Electronics" And m.SiteID = siteID)
                cateConsumer_Electronics.Url = "http://www.dealsmachine.com/consumer-electronics-c-1782.html"
            End If
            'If Not listCategoryName.Contains("Car Electronics") Then
            '    cateCar_Electronics.Category1 = "Car Electronics"
            '    cateCar_Electronics.SiteID = siteID
            '    cateCar_Electronics.Url = "http://www.dealsmachine.com/car-electronics-c-63.html"
            '    cateCar_Electronics.LastUpdate = nowTime
            '    cateCar_Electronics.Description = "Car Electronics"
            '    cateCar_Electronics.ParentID = -1
            '    efContext.AddToCategories(cateCar_Electronics)
            'End If
            'If Not listCategoryName.Contains("Watches") Then
            '    cateWatches.Category1 = "Watches"
            '    cateWatches.SiteID = siteID
            '    cateWatches.Url = "http://www.dealsmachine.com/watches-c-96.html"
            '    cateWatches.LastUpdate = nowTime
            '    cateWatches.Description = "Watches"
            '    cateWatches.ParentID = -1
            '    efContext.AddToCategories(cateWatches)
            'End If
            'If Not listCategoryName.Contains("Flashlights & LED lights") Then
            '    cateFlashlights_LED_lights.Category1 = "Flashlights & LED lights"
            '    cateFlashlights_LED_lights.SiteID = siteID
            '    cateFlashlights_LED_lights.Url = "http://www.dealsmachine.com/lasers-flashlights-c-91.html"
            '    cateFlashlights_LED_lights.LastUpdate = nowTime
            '    cateFlashlights_LED_lights.Description = "Flashlights & LED lights"
            '    cateFlashlights_LED_lights.ParentID = -1
            '    efContext.AddToCategories(cateFlashlights_LED_lights)
            'End If
            'If Not listCategoryName.Contains("Outdoors & Sports") Then
            '    cateOutdoors_Sports.Category1 = "Outdoors & Sports"
            '    cateOutdoors_Sports.SiteID = siteID
            '    cateOutdoors_Sports.Url = "http://www.dealsmachine.com/outdoors-sports-c-107.html"
            '    cateOutdoors_Sports.LastUpdate = nowTime
            '    cateOutdoors_Sports.Description = "Outdoors & Sports"
            '    cateOutdoors_Sports.ParentID = -1
            '    efContext.AddToCategories(cateOutdoors_Sports)
            'End If
            'If Not listCategoryName.Contains("Cell Phone Accessories") Then
            '    cateCell_Phone_Accessories.Category1 = "Cell Phone Accessories"
            '    cateCell_Phone_Accessories.SiteID = siteID
            '    cateCell_Phone_Accessories.Url = "http://www.dealsmachine.com/cell-phone-accessories-c-89.html"
            '    cateCell_Phone_Accessories.LastUpdate = nowTime
            '    cateCell_Phone_Accessories.Description = "Cell Phone Accessories"
            '    cateCell_Phone_Accessories.ParentID = -1
            '    efContext.AddToCategories(cateCell_Phone_Accessories)
            'End If
            'If Not listCategoryName.Contains("Home & Office") Then
            '    cateHome_Office.Category1 = "Home & Office"
            '    cateHome_Office.SiteID = siteID
            '    cateHome_Office.Url = "http://www.dealsmachine.com/home-office-c-105.html"
            '    cateHome_Office.LastUpdate = nowTime
            '    cateHome_Office.Description = "Home & Office"
            '    cateHome_Office.ParentID = -1
            '    efContext.AddToCategories(cateHome_Office)
            'End If
            'If Not listCategoryName.Contains("Hobbies & Toys") Then
            '    cateHobbies_Toys.Category1 = "Hobbies & Toys"
            '    cateHobbies_Toys.SiteID = siteID
            '    cateHobbies_Toys.Url = "http://www.dealsmachine.com/hobbies-toys-c-103.html"
            '    cateHobbies_Toys.LastUpdate = nowTime
            '    cateHobbies_Toys.Description = "Hobbies & Toys"
            '    cateHobbies_Toys.ParentID = -1
            '    efContext.AddToCategories(cateHobbies_Toys)
            'End If
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
                        categoryId = GetCategoryId(siteId, "Cell Phones /Mobiles")
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
        Dim Cell_Phones_MobilesUrl As String = ""
        Dim Notebook_Tablet_PCUrl As String = ""
        Dim Mobile_Power_BankUrl As String = ""
        Dim HeadphonesUrl As String = ""
        Dim Computer_PeripheralsUrl As String = ""
        Dim iPhone_iPad_iPodUrl As String = ""
        Dim Consumer_ElectronicsUrl As String = ""
        Dim Car_ElectronicsUrl As String = ""
        Dim WatchesUrl As String = ""
        Dim Flashlights_LED_lightsUrl As String = ""
        Dim Outdoors_SportsUrl As String = ""
        Dim Cell_Phone_AccessoriesUrl As String = ""
        Dim Home_OfficeUrl As String = ""
        Dim Hobbies_ToysUrl As String = ""
        Try
            '从数据库中获取5个主推分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    'Case "Hobbies & Toys"
                    '    If (HotORNew = "new") Then
                    '        Hobbies_ToysUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Hobbies_ToysUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Home & Office"
                    '    If (HotORNew = "new") Then
                    '        Home_OfficeUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Home_OfficeUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Cell Phone Accessories"
                    '    If (HotORNew = "new") Then
                    '        Cell_Phone_AccessoriesUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Cell_Phone_AccessoriesUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Outdoors & Sports"
                    '    If (HotORNew = "new") Then
                    '        Outdoors_SportsUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Outdoors_SportsUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Flashlights & LED lights"
                    '    If (HotORNew = "new") Then
                    '        Flashlights_LED_lightsUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Flashlights_LED_lightsUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Watches"
                    '    If (HotORNew = "new") Then
                    '        WatchesUrl = q.Url & "?sortby=new"
                    '    Else
                    '        WatchesUrl = q.Url & "?sortby=hot"
                    '    End If
                    'Case "Car Electronics"
                    '    If (HotORNew = "new") Then
                    '        Car_ElectronicsUrl = q.Url & "?sortby=new"
                    '    Else
                    '        Car_ElectronicsUrl = q.Url & "?sortby=hot"
                    '    End If
                    Case "Consumer Electronics"
                        If (HotORNew = "new") Then
                            Consumer_ElectronicsUrl = q.Url & "?sortby=new"
                        Else
                            Consumer_ElectronicsUrl = q.Url & "?sortby=hot"
                        End If
                        'Case "iPhone & iPad & iPod"
                        '    If (HotORNew = "new") Then
                        '        iPhone_iPad_iPodUrl = q.Url & "?sortby=new"
                        '    Else
                        '        iPhone_iPad_iPodUrl = q.Url & "?sortby=hot"
                        '    End If
                    Case "Computer Peripherals"
                        If (HotORNew = "new") Then
                            Computer_PeripheralsUrl = q.Url & "?sortby=new"
                        Else
                            Computer_PeripheralsUrl = q.Url & "?sortby=hot"
                        End If
                        'Case "Headphones"
                        '    If (HotORNew = "new") Then
                        '        HeadphonesUrl = q.Url & "?sortby=new"
                        '    Else
                        '        HeadphonesUrl = q.Url & "?sortby=hot"
                        '    End If
                    Case "Mobile Power Bank"
                        If (HotORNew = "new") Then
                            Mobile_Power_BankUrl = q.Url & "?sortby=new"
                        Else
                            Mobile_Power_BankUrl = q.Url & "?sortby=hot"
                        End If
                    Case "Notebook/Tablet PC"
                        If (HotORNew = "new") Then
                            Notebook_Tablet_PCUrl = q.Url & "?sortby=new"
                        Else
                            Notebook_Tablet_PCUrl = q.Url & "?sortby=hot"
                        End If
                    Case "Cell Phones /Mobiles"
                        If (HotORNew = "new") Then
                            Cell_Phones_MobilesUrl = q.Url & "?sortby=new"
                        Else
                            Cell_Phones_MobilesUrl = q.Url & "?sortby=hot"
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Cell_Phones_MobilesUrl, "Cell Phones /Mobiles", iProIssueCount, updateTime, planType, 50, 100)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Notebook_Tablet_PCUrl, "Notebook/Tablet PC", iProIssueCount, updateTime, planType, 60, 80)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Mobile_Power_BankUrl, "Mobile Power Bank", iProIssueCount, updateTime, planType, 5, 10)
            'GetProducts(siteId, Issueid, HeadphonesUrl, "Headphones", iProIssueCount, updateTime)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Computer_PeripheralsUrl, "Computer Peripherals", iProIssueCount, updateTime, planType, 6, 12)
            'GetProducts(siteId, Issueid, iPhone_iPad_iPodUrl, "iPhone & iPad & iPod", iProIssueCount, updateTime)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Consumer_ElectronicsUrl, "Consumer Electronics", iProIssueCount, updateTime, planType, 5, 10)
            'GetProducts(siteId, Issueid, Car_ElectronicsUrl, "Car Electronics", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, WatchesUrl, "Watches", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, Flashlights_LED_lightsUrl, "Flashlights & LED lights", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, Outdoors_SportsUrl, "Outdoors & Sports", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, Cell_Phone_AccessoriesUrl, "Cell Phone Accessories", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, Home_OfficeUrl, "Home & Office", iProIssueCount, updateTime)
            'GetProducts(siteId, Issueid, Hobbies_ToysUrl, "Hobbies & Toys", iProIssueCount, updateTime)
        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String, ByVal lowPrice As Integer, ByVal highetPrice As Integer)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType)

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
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType)

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


    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@id='mainbar']/div[@class='section m_box']/ul/li")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                product.Prodouct = productDetails.SelectSingleNode("p[@class='p_name pt10']").SelectSingleNode("a").InnerText
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 70) Then
                    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                End If
                product.Url = productDetails.SelectSingleNode("p[@class='p_name pt10']").SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If

                'product.Price = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Discount = Double.Parse(productDetails.SelectSingleNode("p[@class='p_price  fb red']").SelectSingleNode("span[@class='my_shop_price fl pl5']").GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Currency = productDetails.SelectSingleNode("p[@class='p_price  fb red']").SelectNodes("span")(0).InnerText

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("p[@class='p_mg pr']")
                product.PictureUrl = productImage.SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                If (productImage.SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("grey.gif")) Then
                    product.PictureUrl = productImage.SelectSingleNode("a/img").GetAttributeValue("data-original", "")
                End If

                product.LastUpdate = nowTime
                product.Description = productImage.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                product.SiteID = siteId
                product.PictureAlt = productImage.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                product.SizeHeight = Integer.Parse(productImage.SelectSingleNode("a/img").GetAttributeValue("height", ""))
                product.SizeWidth = Integer.Parse(productImage.SelectSingleNode("a/img").GetAttributeValue("width", ""))

                Dim shipsin24hrs As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/shipin24hrs.jpg"
                Dim freeshipping As String = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/freeshipping.jpg"
                If Not (productDetails.SelectSingleNode("p[@class='p_Ship']") Is Nothing) Then
                    product.ShipsImg = shipsin24hrs
                End If

                If Not (productDetails.SelectSingleNode("p[@class='p_free en_p_free']") Is Nothing) Then
                    product.FreeShippingImg = freeshipping
                End If

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
            'If (productInList.Url = "http://www.dealsmachine.com/product-165223.html") Then
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
        Dim pageUrl = "http://www.dealsmachine.com/DailyDeals.html"
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-8).ToShortDateString() 'Deals控制一周内不重复
        If (section = "DA") Then
            startDate = Now.AddDays(-8).ToShortDateString() 'Deals控制一周内不重复
            'ElseIf (section = "NE") Then
            '    startDate = Now.AddDays(-15).ToShortDateString() 'NewArri控制两周内不重复
        End If
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@id='group_mainbar']/div[@class='group_box']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            If (section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//div[@id='group_mainbar']/div[@class='group_box']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
                'ElseIf (section = "NE") Then
                '    productDivs = hd.DocumentNode.SelectNodes("//section[@id='miain']/section[@class='mb10 newArr']/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            End If
            '获取产品  mb10 newArr
            Dim listProducts As New List(Of Product)
            Dim counter As Integer
            'Dim dealtimeLeft As Integer
            If Not productDivs Is Nothing Then
                For counter = 0 To productDivs.Count - 1
                    Dim product As New Product()
                    Dim productDetails As HtmlNodeCollection = productDivs(counter).SelectNodes("div[@class='proBox clearfix']/div")

                    product.Prodouct = productDivs(counter).SelectSingleNode("h3/a").InnerText
                    If (product.Prodouct.ToString.Length > 100) Then
                        product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                    End If
                    product.Url = productDivs(counter).SelectSingleNode("h3/a").GetAttributeValue("href", "").Trim()
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If

                    product.Price = Double.Parse(productDetails(0).SelectSingleNode("p[@class='deal-discount']/strong").InnerText.Trim("$").Trim()) 'deal-discount
                    product.Discount = Double.Parse(productDetails(0).SelectSingleNode("p[@class='deal-price']/strong").InnerText.Trim("$").Trim()) 'deal-discount
                    product.Currency = "$"

                    '特殊图片处理
                    product.PictureUrl = productDetails(1).SelectSingleNode("p[@class='deal_img pr']/a/img").GetAttributeValue("src", "") '2013/4/7添加
                    If (product.PictureUrl.Contains("lazyload.gif")) Then
                        product.PictureUrl = productDetails(1).SelectSingleNode("p[@class='deal_img pr']/a/img").GetAttributeValue("data-original", "")
                    End If

                    product.LastUpdate = nowTime
                    product.Description = productDivs(counter).SelectSingleNode("h3/a").InnerText
                    product.SiteID = siteId
                    product.PictureAlt = productDetails(1).SelectSingleNode("p[@class='deal_img pr']/a/img").GetAttributeValue("alt", "")
                    product.SizeHeight = Integer.Parse(productDetails(1).SelectSingleNode("p[@class='deal_img pr']/a/img").GetAttributeValue("height", ""))
                    product.SizeWidth = Integer.Parse(productDetails(1).SelectSingleNode("p[@class='deal_img pr']/a/img").GetAttributeValue("width", ""))
                    'Try
                    '    product.Sales = Integer.Parse(productDetails(0).SelectSingleNode("strong").InnerText)
                    'Catch
                    '    product.Sales = 0 '无折扣则为零
                    'End Try
                    'product.Discount = (product.Price / (100 - product.Sales)) * 100

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
                    '        product.FreeShippingImg = freeshippingFFF
                    '    End If

                    '    If (productDetails(4).GetAttributeValue("class", "").Contains("p_shipsInHrs")) Then
                    '        product.ShipsImg = shipsin24hrs
                    '    ElseIf (productDetails(4).GetAttributeValue("class", "").Contains("p_shipping pt5")) Then
                    '        product.FreeShippingImg = freeshipping
                    '    End If
                    'End If
                    'Dim dealtimeLeft As HtmlNodeCollection = productDetails(0).SelectNodes("dl[@class='deal_time mt10']/dd")
                    'Dim timeLeft As String = dealtimeLeft(1).InnerText
                    'If (timeLeft.ToLower().Trim() = "finished") Then '控制获取的产品不过期
                    If (Not efHelpre.IsProductSent(siteId, product.Url, startDate, endDate, PlanType)) Then '控制指定时间内获取的产品不重复：
                        listProducts.Add(product)
                    End If
                    'End If
                Next
            End If
            Return listProducts
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:WeeklyDeals" & " error:" & ex.Message.ToString)
        End Try
    End Function

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        Dim pageUrl As String = "http://www.dealsmachine.com/"
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
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='js_slideBox']/div[@class='slideBox_Imgshow']/ul/li/a")
        '删除这一句是因为获取到的bannes.count = 5，并不是重复了3遍后的15，所以不需要做这样一个控制
        'Dim getBannersCount As Integer = 5 '在这里定义好获取多少个banner条，这样设定的原因是ahappydeal分析html得来的banner条重复了3遍
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            Ad.Url = "http://www.dealsmachine.com" & banners(counter).GetAttributeValue("href", "").ToString.Trim()
            Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString
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

        Try
            If (listAdIds.Count = 0) Then '转而去到A站的每周邮件页面（http://www.dealsmachine.com/mail/2014/A0120.html ）获取banner
                Dim mon As Integer = DateTime.Now.DayOfWeek
                pageUrl = "http://www.dealsmachine.com/mail/" & DateTime.Now.ToString("yyyy") & "/A" & DateTime.Now.AddDays(1 - mon).ToString("MMdd") & ".html"
                hd = efHelpre.GetHtmlDocument(pageUrl, 60000)
                Dim banners1 As HtmlNodeCollection = hd.DocumentNode.SelectNodes("/html/body/table/tr/td")
                Dim banners2 As HtmlNodeCollection = hd.DocumentNode.SelectNodes("/html/body/table/tr/td/table/tr")
                Dim banner As HtmlNode = banners2(0).SelectSingleNode("td/table[1]/tr[6]")

                Dim Ad As Ad = New Ad()
                Try
                    Ad.Ad1 = banner.SelectSingleNode("td/a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    Ad.Description = banner.SelectSingleNode("td/a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                Ad.Url = banner.SelectSingleNode("td/a").GetAttributeValue("href", "").ToString.Trim()
                Ad.PictureUrl = banner.SelectSingleNode("td/a/img").GetAttributeValue("src", "").ToString
                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelpre.GetListAd(siteid)
                If (Not efHelpre.isAdSent(siteid, Ad.Url, startDate, endDate, planType)) Then '控制一个月内获取的Ad不重复：
                    adid = efHelpre.InsertAd(Ad, listAd, categroyId)
                    listAdIds.Add(adid)
                End If
            End If
        Catch ex As Exception
            'ignore
        End Try

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
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@class='nav_path mt10']/div[@class='fl curPath']")
        Dim aHerfs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='nav_path mt10']/div[@class='fl curPath']/a")
        Dim categoryName As String = ""

        Dim index = productDivs.InnerText.IndexOf(";")
        If (aHerfs.Count = 1) Then
            categoryName = productDivs.InnerText.Substring(index + 1, productDivs.InnerText.Length - index - 1).Trim()
        ElseIf (aHerfs.Count > 1) Then
            categoryName = aHerfs(1).InnerText
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
            preSubject = "Dealsmachine Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "Dealsmachine New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HP")) Then
            preSubject = "Dealsmachine Special " & categoryName & " For " 'Sammydress New Arrivals For Jan.2014.Vol.01
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
