Imports HtmlAgilityPack
Imports System.Math

Public Class Eachbuyer

    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private Flage1 As Boolean = vbTrue
    Private Flage2 As Boolean = vbTrue
    Private Flage3 As Boolean = vbTrue
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）
    Private domin As String = "http://www.eachbuyer.com"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        Dim contactlistCount As Integer
        If planType.Trim.Contains("HO") Then
            bannerCategoryName = "Consumer Electronics"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            ''获取的首页的deals
            GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType)
            ''主推4个分类,每个分类获取4个产品
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)
            ''Add Subject
            AddIssueSubject(IssueID, siteId, "HO", "DA", spreadLogin) 'AddIssueSubject(IssueID, siteId, planType, "DA")
            'Dim contactListArr() As String = {"Reasonablers"}
            Dim contactListArr() As String = {"All Open 20140318"}
            If (spreadLogin.ToLower.Trim = "lihuojian@hofan.cn") Then
                contactListArr = {"Opens201404181210"}
            End If
            efHelper.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

        ElseIf (planType.Trim.Contains("HA")) Then
            bannerCategoryName = "Consumer Electronics"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetWeeklyDealsOrNewArri(siteId, IssueID, "NE", planType)
            '主推4个分类,每个分类获取4个产品
            GetHotOrNewProducts(siteId, IssueID, "new", planType)
            ''Add Subject
            AddIssueSubject(IssueID, siteId, "HA", "NE", spreadLogin) 'AddIssueSubject(IssueID, siteId, planType, "NE")
            'Dim contactListArr() As String = {"Reasonablers"}
            Dim contactListArr() As String = {"All Open 20140318"}
            If (spreadLogin.ToLower.Trim = "lihuojian@hofan.cn") Then
                contactListArr = {"Opens201404181210"}
            End If
            efHelper.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting
        ElseIf (planType.Trim.Contains("HP")) Then
            Dim cate As String() = categories.Split("^")

            Dim categoryUrlC As String
            Dim categoryUrlL As String
            Dim iProIssueCount As Integer = 3
            Dim cateName As String = ""

            If (cate(0).Trim = "Consumer Electronics") Then
                categoryUrlL = "http://www.eachbuyer.com/business-industrial-c15016/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/camera-photo-c15014/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/sound-vision-c15017/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)

                iProIssueCount = 1
                cateName = "Computer & Networking"
                categoryUrlL = "http://www.eachbuyer.com/computer-hardwares-c15172/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/speakers-earphones-c15169/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/usb-accessories-c15167/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
            ElseIf (cate(0).Trim = "Lighting") Then
                categoryUrlL = "http://www.eachbuyer.com/led-light-bulbs-c15354/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/flashlights-c15304/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/Decoration-Lights-c15642/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)

                iProIssueCount = 1
                cateName = "Watches"
                categoryUrlL = "http://www.eachbuyer.com/men-s-watches-c15286/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/women-s-watches-c15287/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/jewelry-c15277/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
            ElseIf (cate(0).Trim = "Car Accessories") Then
                categoryUrlL = "http://www.eachbuyer.com/led-lights-c15082/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/motorcycle-parts-c15382/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/Car-Parts-&-Accessories-c15601/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)

                iProIssueCount = 1
                cateName = "Health & Beauty"
                categoryUrlL = "http://www.eachbuyer.com/Intimate-Gadgets-c15559/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/wigs-c15326/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
                categoryUrlL = "http://www.eachbuyer.com/Health-Care-c15390/3.html"
                GetProducts(siteId, IssueID, categoryUrlL, cateName, iProIssueCount, Now, planType)
            Else
                iProIssueCount = 6
                Dim queryURL = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
                For Each q In queryURL
                    Select Case q.Category1.Trim()
                        Case cate(0).Trim
                            categoryUrlL = q.Url & "2.html"
                        Case cate(1).Trim
                            categoryUrlC = q.Url & "2.html"
                    End Select
                Next
                GetProducts(siteId, IssueID, categoryUrlL, cate(0).Trim, iProIssueCount, Now, planType)
                GetProducts(siteId, IssueID, categoryUrlC, cate(1).Trim, iProIssueCount, Now, planType)
            End If
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0), spreadLogin)
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "Eachbuyer" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 365, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            efHelper.LogText("siteID:" & siteId & "cate(0):" & cate(0).Trim & "savelistName:" & saveListName & "contactlistCount:" & contactlistCount)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
            If (planType.Trim = "HP4" OrElse planType.Trim = "HP5" OrElse planType.Trim = "HP6" OrElse planType.Trim = "HP7" OrElse planType.Trim = "HP8") Then
                Dim saveListName1 As String = "Eachbuyer" & cate(1).Trim & Now.ToString("yyyyMMdd")
                contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(1).Trim, saveListName1, 365, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                efHelper.LogText("siteID:" & siteId & "cate(1):" & cate(1).Trim & "savelistName:" & saveListName & "contactlistCount:" & contactlistCount)
                If (contactlistCount > 0) Then
                    efHelper.InsertContactList(IssueID, saveListName1, sendingStatus)
                End If
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
            '添加13大分类 
            Dim cateCell_Phones_Mobiles As New Category()
            Dim cateNotebook_Tablet_PC As New Category()
            Dim cateMobile_Power_Bank As New Category()
            Dim cateHeadphones As New Category()
            Dim cateComputer_Peripherals As New Category()
            Dim cateiPhone_iPad_iPod As New Category()
            Dim cateJewelryWatch As New Category()
            Dim cateCar_Electronics As New Category()
            Dim cateWatches As New Category()
            Dim cateFlashlights_LED_lights As New Category()
            Dim cateOutdoors_Sports As New Category()
            Dim cateCell_Phone_Accessories As New Category()
            Dim cateHome_Office As New Category()
            Dim cateHobbies_Toys As New Category()
            If Not listCategoryName.Contains("Consumer Electronics") Then
                cateCell_Phones_Mobiles.Category1 = "Consumer Electronics"
                cateCell_Phones_Mobiles.SiteID = siteID
                cateCell_Phones_Mobiles.Url = "http://www.eachbuyer.com/business-industrial-c15016/"
                cateCell_Phones_Mobiles.LastUpdate = nowTime
                cateCell_Phones_Mobiles.Description = "Consumer Electronics"
                cateCell_Phones_Mobiles.ParentID = -1
                efContext.AddToCategories(cateCell_Phones_Mobiles)
            Else
                cateCell_Phones_Mobiles = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Consumer Electronics" And m.SiteID = siteID)
                cateCell_Phones_Mobiles.Url = "http://www.eachbuyer.com/business-industrial-c15016/"
            End If
            If Not listCategoryName.Contains("Lighting") Then
                cateNotebook_Tablet_PC.Category1 = "Lighting"
                cateNotebook_Tablet_PC.SiteID = siteID
                cateNotebook_Tablet_PC.Url = "http://www.eachbuyer.com/led-light-bulbs-c15354/"
                cateNotebook_Tablet_PC.LastUpdate = nowTime
                cateNotebook_Tablet_PC.Description = "Lighting"
                cateNotebook_Tablet_PC.ParentID = -1
                efContext.AddToCategories(cateNotebook_Tablet_PC)
            Else
                cateNotebook_Tablet_PC = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Lighting" And m.SiteID = siteID)
                cateNotebook_Tablet_PC.Url = "http://www.eachbuyer.com/led-light-bulbs-c15354/"
            End If
            If Not listCategoryName.Contains("Car Accessories") Then
                cateMobile_Power_Bank.Category1 = "Car Accessories"
                cateMobile_Power_Bank.SiteID = siteID
                cateMobile_Power_Bank.Url = "http://www.eachbuyer.com/led-lights-c15082/"
                cateMobile_Power_Bank.LastUpdate = nowTime
                cateMobile_Power_Bank.Description = "Car Accessories"
                cateMobile_Power_Bank.ParentID = -1
                efContext.AddToCategories(cateMobile_Power_Bank)
            Else
                cateMobile_Power_Bank = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Car Accessories" And m.SiteID = siteID)
                cateMobile_Power_Bank.Url = "http://www.eachbuyer.com/led-lights-c15082/"
            End If
            If Not listCategoryName.Contains("Clothing & Accessories") Then
                cateHeadphones.Category1 = "Clothing & Accessories"
                cateHeadphones.SiteID = siteID
                cateHeadphones.Url = "http://www.eachbuyer.com/accessories-c15245/"
                cateHeadphones.LastUpdate = nowTime
                cateHeadphones.Description = "Clothing & Accessories"
                cateHeadphones.ParentID = -1
                efContext.AddToCategories(cateHeadphones)
            Else
                cateHeadphones = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Clothing & Accessories" And m.SiteID = siteID)
                cateHeadphones.Url = "http://www.eachbuyer.com/accessories-c15245/"
            End If
            If Not listCategoryName.Contains("Cell Phone & Tablets") Then
                cateComputer_Peripherals.Category1 = "Cell Phone & Tablets"
                cateComputer_Peripherals.SiteID = siteID
                cateComputer_Peripherals.Url = "http://www.eachbuyer.com/cell-phone-c15096/"
                cateComputer_Peripherals.LastUpdate = nowTime
                cateComputer_Peripherals.Description = "Cell Phone & Tablets"
                cateComputer_Peripherals.ParentID = -1
                efContext.AddToCategories(cateComputer_Peripherals)
            Else
                cateComputer_Peripherals = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Cell Phone & Tablets" And m.SiteID = siteID)
                cateComputer_Peripherals.Url = "http://www.eachbuyer.com/cell-phone-c15096/"
            End If
            If Not listCategoryName.Contains("Health & Beauty") Then
                cateiPhone_iPad_iPod.Category1 = "Health & Beauty"
                cateiPhone_iPad_iPod.SiteID = siteID
                cateiPhone_iPad_iPod.Url = "http://www.eachbuyer.com/Health-Care-c15390/"
                cateiPhone_iPad_iPod.LastUpdate = nowTime
                cateiPhone_iPad_iPod.Description = "Health & Beauty"
                cateiPhone_iPad_iPod.ParentID = -1
                efContext.AddToCategories(cateiPhone_iPad_iPod)
            Else
                cateiPhone_iPad_iPod = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Health & Beauty" And m.SiteID = siteID)
                cateiPhone_iPad_iPod.Url = "http://www.eachbuyer.com/Health-Care-c15390/"
            End If
            If Not listCategoryName.Contains("Watches") Then
                cateJewelryWatch.Category1 = "Watches"
                cateJewelryWatch.SiteID = siteID
                cateJewelryWatch.Url = "http://www.eachbuyer.com/men-s-watches-c15286/"
                cateJewelryWatch.LastUpdate = nowTime
                cateJewelryWatch.Description = "Watches"
                cateJewelryWatch.ParentID = -1
                efContext.AddToCategories(cateJewelryWatch)
            Else
                cateJewelryWatch = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Watches" And m.SiteID = siteID)
                cateJewelryWatch.Url = "http://www.eachbuyer.com/men-s-watches-c15286/"
            End If
            If Not listCategoryName.Contains("Cell Phone & Tablet ACC") Then
                cateCar_Electronics.Category1 = "Cell Phone & Tablet ACC"
                cateCar_Electronics.SiteID = siteID
                cateCar_Electronics.Url = "http://www.eachbuyer.com/iphone-accessories-c15143/"
                cateCar_Electronics.LastUpdate = nowTime
                cateCar_Electronics.Description = "Cell Phone & Tablet ACC"
                cateCar_Electronics.ParentID = -1
                efContext.AddToCategories(cateCar_Electronics)
            Else
                cateCar_Electronics = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Cell Phone & Tablet ACC" And m.SiteID = siteID)
                cateCar_Electronics.Url = "http://www.eachbuyer.com/iphone-accessories-c15143/"
            End If
            If Not listCategoryName.Contains("Home & Garden") Then
                cateWatches.Category1 = "Home & Garden"
                cateWatches.SiteID = siteID
                cateWatches.Url = "http://www.eachbuyer.com/home-d-cor-c15293/"
                cateWatches.LastUpdate = nowTime
                cateWatches.Description = "Home & Garden"
                cateWatches.ParentID = -1
                efContext.AddToCategories(cateWatches)
            Else
                cateWatches = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Home & Garden" And m.SiteID = siteID)
                cateWatches.Url = "http://www.eachbuyer.com/home-d-cor-c15293/"
            End If
            If Not listCategoryName.Contains("Security System") Then
                cateFlashlights_LED_lights.Category1 = "Security System"
                cateFlashlights_LED_lights.SiteID = siteID
                cateFlashlights_LED_lights.Url = "http://www.eachbuyer.com/access-control-systems-c15037/"
                cateFlashlights_LED_lights.LastUpdate = nowTime
                cateFlashlights_LED_lights.Description = "Security System"
                cateFlashlights_LED_lights.ParentID = -1
                efContext.AddToCategories(cateFlashlights_LED_lights)
            Else
                cateFlashlights_LED_lights = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Security System" And m.SiteID = siteID)
                cateFlashlights_LED_lights.Url = "http://www.eachbuyer.com/access-control-systems-c15037/"
            End If
            If Not listCategoryName.Contains("Computer & Networking") Then
                cateOutdoors_Sports.Category1 = "Computer & Networking"
                cateOutdoors_Sports.SiteID = siteID
                cateOutdoors_Sports.Url = "http://www.eachbuyer.com/computer-networking-c15004/"
                cateOutdoors_Sports.LastUpdate = nowTime
                cateOutdoors_Sports.Description = "Computer & Networking"
                cateOutdoors_Sports.ParentID = -1
                efContext.AddToCategories(cateOutdoors_Sports)
            Else
                cateOutdoors_Sports = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Computer & Networking" And m.SiteID = siteID)
                cateOutdoors_Sports.Url = "http://www.eachbuyer.com/computer-networking-accessories-c15004/"
            End If
            If Not listCategoryName.Contains("Sports & Outdoors") Then
                cateCell_Phone_Accessories.Category1 = "Sports & Outdoors"
                cateCell_Phone_Accessories.SiteID = siteID
                cateCell_Phone_Accessories.Url = "http://www.eachbuyer.com/camping-hiking-c15207/"
                cateCell_Phone_Accessories.LastUpdate = nowTime
                cateCell_Phone_Accessories.Description = "Sports & Outdoors"
                cateCell_Phone_Accessories.ParentID = -1
                efContext.AddToCategories(cateCell_Phone_Accessories)
            Else
                cateCell_Phone_Accessories = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Sports & Outdoors" And m.SiteID = siteID)
                cateCell_Phone_Accessories.Url = "http://www.eachbuyer.com/camping-hiking-c15207/"
            End If
            If Not listCategoryName.Contains("Toy & Hobby") Then
                cateHome_Office.Category1 = "Toy & Hobby"
                cateHome_Office.SiteID = siteID
                cateHome_Office.Url = "http://www.eachbuyer.com/rc-toys-hobbies-c15110/"
                cateHome_Office.LastUpdate = nowTime
                cateHome_Office.Description = "Toy & Hobby"
                cateHome_Office.ParentID = -1
                efContext.AddToCategories(cateHome_Office)
            Else
                cateHome_Office = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "Toy & Hobby" And m.SiteID = siteID)
                cateHome_Office.Url = "http://www.eachbuyer.com/rc-toys-hobbies-c15110/"
            End If

            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
            Throw ex
        End Try
    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://www.eachbuyer.com/"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)

            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)
        Catch ex As Exception
            LogText("GetPromotionBanner error :" & ex.ToString)
        End Try
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        Try
            'Dim listAds As New List(Of Ad)
            Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
            Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='focus']/div/ul/li")  '/div/div/a/
            For counter As Integer = 0 To banners.Count - 1
                Dim Ad As Ad = New Ad()
                Try
                    Ad.Ad1 = banners(counter).SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    Ad.Description = banners(counter).SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                If Not (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.ToLower.Contains("http://")) Then
                    If (banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim().Trim.StartsWith("/")) Then
                        Ad.Url = "http://www.eachbuyer.com" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                    Else
                        Ad.Url = "http://www.eachbuyer.com/" & banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                    End If
                Else
                    Ad.Url = banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim()
                End If

                Try
                    If Not (banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.ToLower().Contains("http://")) Then
                        If (banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.Trim.StartsWith("/")) Then
                            Ad.PictureUrl = "http://www.eachbuyer.com" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
                        Else
                            Ad.PictureUrl = "http://www.eachbuyer.com/" & banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
                        End If
                    Else
                        Ad.PictureUrl = banners(counter).SelectSingleNode("a/img").GetAttributeValue("src", "").ToString
                    End If
                Catch ex As Exception
                    'ignore
                End Try

                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)
                If (Not efHelper.isAdSent(siteid, Ad.Url, startDate, endDate, planType) AndAlso Ad.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                    adid = efHelper.InsertAd(Ad, listAd, categroyId)
                    listAdIds.Add(adid)
                End If
            Next

        Catch ex As Exception
            LogText("Get Banner list error:" & ex.ToString)
        End Try

        Return listAdIds
    End Function

    Public Sub GetWeeklyDealsOrNewArri(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String)
        Try
            Dim topCount As Integer = 2
            Dim categoryName As String = "Car Accessories"
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listProducts As List(Of Product) = GetWeeklyDealOrNewArriProducts(siteId, Issueid, section, planType)
            If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
                Dim listProduct As New List(Of Product)
                listProduct = efHelper.GetProductList(siteId)
                Dim listProductId As New List(Of Integer)
                Dim counter As Integer = 0
                For Each li In listProducts
                    Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)
                    listProductId.Add(returnId)

                    '只插入排好序的前topCount个产品
                    counter += 1
                    If (counter = topCount) Then
                        Exit For
                    End If
                Next
                efHelper.InsertProductsIssue(siteId, Issueid, section, listProductId, topCount)
            End If

            'categoryName = "Consumer Electronics"
            'categoryId = efHelper.GetCategoryId(siteId, categoryName)
            'listProducts = GetListProducts("http://www.eachbuyer.com/Measurement-&-Analysis-Instruments-c15635/", siteId, Now, categoryName, planType, topCount) '这个函数中已判断产品是否重复
            'If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
            '    Dim listProduct As New List(Of Product)
            '    listProduct = efHelper.GetProductList(siteId)
            '    Dim listProductId As New List(Of Integer)
            '    Dim counter As Integer = 0
            '    For Each li In listProducts
            '        Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)
            '        listProductId.Add(returnId)

            '        '只插入排好序的前topCount个产品
            '        counter += 1
            '        If (counter = topCount) Then
            '            Exit For
            '        End If
            '    Next
            '    efHelper.InsertProductsIssue(siteId, Issueid, section, listProductId, topCount)
            'End If
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    Public Function GetWeeklyDealOrNewArriProducts(ByVal siteID As Integer, ByVal Issueid As Integer, ByVal Section As String, ByVal planType As String) As List(Of Product)
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim pageUrl As String = "http://www.eachbuyer.com/"
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String
        startDate = Now.AddDays(-30).ToShortDateString() 'Deals控制一月内不重复
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            If (Section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//div[@class='lists clear']")(0).SelectNodes("div[@class='list-right']/ul/li")
            ElseIf (Section = "NE") Then
                productDivs = hd.DocumentNode.SelectNodes("//div[@class='lists clear']")(1).SelectNodes("div[@class='list-right']/ul/li")
            End If
            '获取产品  mb10 newArr
            Dim listProducts As New List(Of Product)
            Dim counter As Integer
            If Not productDivs Is Nothing Then
                For counter = 0 To productDivs.Count - 1
                    Dim product As New Product()

                    product.Prodouct = productDivs(counter).SelectSingleNode("div/div[2]/a").InnerText
                    'If (product.Prodouct.ToString.Length > 100) Then
                    '    product.Prodouct = product.Prodouct.Substring(0, 100) & "..."
                    'End If
                    product.Url = efHelper.addDominForUrl(domin, productDivs(counter).SelectSingleNode("div/div[2]/a").GetAttributeValue("href", "").Trim())
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If

                    Try
                        product.Price = Double.Parse(productDivs(counter).SelectSingleNode("div/div[@class='p-price']/span[@class='price-sm-o']").InnerText.Replace("$", "").Trim())
                    Catch ex As Exception
                        'ingore
                    End Try
                    product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("div/div[@class='p-price']/span[@class='price-sm-n']").InnerText.Replace("$", "").Trim())
                    product.Currency = "$"

                    '特殊图片处理
                    product.PictureUrl = productDivs(counter).SelectSingleNode("div/div[@class='p-pic']/a/img").GetAttributeValue("src", "") '2013/4/7添加
                    If (product.PictureUrl.Contains("default.")) Then
                        product.PictureUrl = efHelper.addDominForUrl(domin, productDivs(counter).SelectSingleNode("div/div[@class='p-pic']/a/img").GetAttributeValue("data-lazysrc", ""))
                    End If
                    product.LastUpdate = nowTime
                    product.Description = product.Prodouct
                    product.SiteID = siteID
                    product.FreeShippingImg = "http://www.eachbuyer.com/images/common/fs_us.jpg"

                    If (Not efHelper.IsProductSent(siteID, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的产品不重复：
                        listProducts.Add(product)
                    End If
                    If (listProducts.Count = 6) Then
                        Exit For
                    End If
                Next
            End If
            Return listProducts
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:HotStyle" & " error:" & ex.Message.ToString)
        End Try
    End Function





    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim Consumer_ElectronicsUrl As String = ""
        Dim lightsUrl As String = ""
        Dim CarUrl As String = ""
        Dim clothingUrl As String = ""
        Dim Cell_Phones_MobilesUrl As String = ""
        Dim healthUrl As String = ""
        Dim WatchesUrl As String = ""
        Dim Cell_Phone_AccessoriesUrl As String = ""
        Dim HomeUrl As String = ""
        Dim securityUrl As String = ""
        Dim Computer_NetworkUrl As String = ""
        Dim Outdoors_SportsUrl As String = ""
        Dim Hobbies_ToysUrl As String = ""

        Try
            '从数据库中获取5个主推分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取new arrive的产品，在获取分类url时，将最后的hot替换成new后直接使用
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Consumer Electronics"
                        If (HotORNew = "new") Then
                            Consumer_ElectronicsUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Consumer_ElectronicsUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Lighting"
                        If (HotORNew = "new") Then
                            lightsUrl = q.Url.Trim() & "?sort=2"
                        Else
                            lightsUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Car Accessories"
                        If (HotORNew = "new") Then
                            CarUrl = q.Url.Trim() & "?sort=2"
                        Else
                            CarUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Clothing & Accessories"
                        If (HotORNew = "new") Then
                            clothingUrl = q.Url.Trim() & "?sort=2"
                        Else
                            clothingUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Cell Phone & Tablets"
                        If (HotORNew = "new") Then
                            Cell_Phones_MobilesUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Cell_Phones_MobilesUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Health & Beauty"
                        If (HotORNew = "new") Then
                            healthUrl = q.Url.Trim() & "?sort=2"
                        Else
                            healthUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Watches"
                        If (HotORNew = "new") Then
                            WatchesUrl = q.Url.Trim() & "?sort=2"
                        Else
                            WatchesUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Cell Phone & Tablet ACC"
                        If (HotORNew = "new") Then
                            Cell_Phone_AccessoriesUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Cell_Phone_AccessoriesUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Home & Garden"
                        If (HotORNew = "new") Then
                            HomeUrl = q.Url.Trim() & "?sort=2"
                        Else
                            HomeUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Security System"
                        If (HotORNew = "new") Then
                            securityUrl = q.Url.Trim() & "?sort=2"
                        Else
                            securityUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Computer & Networking"
                        If (HotORNew = "new") Then
                            Computer_NetworkUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Computer_NetworkUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Sports & Outdoors"
                        If (HotORNew = "new") Then
                            Outdoors_SportsUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Outdoors_SportsUrl = q.Url.Trim() '& "2.html"
                        End If
                    Case "Toy & Hobby"
                        If (HotORNew = "new") Then
                            Hobbies_ToysUrl = q.Url.Trim() & "?sort=2"
                        Else
                            Hobbies_ToysUrl = q.Url.Trim() '& "2.html"
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Consumer_ElectronicsUrl, "Consumer Electronics", iProIssueCount, updateTime, planType, 5, 7)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, lightsUrl, "Lighting", iProIssueCount, updateTime, planType, 5, 10)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, CarUrl, "Car Accessories", iProIssueCount, updateTime, planType, 5, 10)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, WatchesUrl, "Watches", iProIssueCount, updateTime, planType, 5, 7)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, Outdoors_SportsUrl, "Sports & Outdoors", iProIssueCount, updateTime, planType, 3, 6)

            iProIssueCount = 1
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, healthUrl, "Health & Beauty", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, clothingUrl, "Clothing & Accessories", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, HomeUrl, "Home & Garden", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, securityUrl, "Security System", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, Computer_NetworkUrl, "Computer & Networking", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True
            GetProducts(siteId, Issueid, Cell_Phone_AccessoriesUrl, "Cell Phone & Tablet ACC", iProIssueCount, updateTime, planType, 1, 1)
            Flage1 = True : Flage2 = True : Flage3 = True 
            GetProducts(siteId, Issueid, Hobbies_ToysUrl, "Toy & Hobby", iProIssueCount, updateTime, planType, 1, 1)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String, ByVal lowPrice As Integer, ByVal highetPrice As Integer)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim listExcludeProductId As New List(Of Integer) '存放按照价格区间来匹配产品而没有匹配到的产品id，以备符合指定价格区间的产品数量不足时从此中再选择产品填充
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)


            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

                If (isInPriceSpan(li, lowPrice, highetPrice)) And returnId > 0 Then
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
                    If li > 0 Then
                        listProductId.Add(li)
                    End If
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

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)

            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            efHelper.LogText("siteID:" & siteId & "categoryName:" & categoryName & "ID:" & categoryId)

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

    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                  ByVal iProIssueCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-proNoRepeatSpan).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@class='primary-list']/ul/li")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                Try
                    product.Prodouct = productDetails.SelectSingleNode("div/div[@class='p-name']/a").InnerText
                Catch ex As Exception
                    product.Prodouct = productDetails.SelectSingleNode("div/div[2]/a").InnerText
                End Try
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                If (product.Prodouct.Length > 70) Then
                    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                End If
                Try
                    product.Url = efHelper.addDominForUrl(domin, productDetails.SelectSingleNode("div/div[@class='p-name']/a").GetAttributeValue("href", "").Trim())
                Catch ex As Exception
                    product.Url = efHelper.addDominForUrl(domin, productDetails.SelectSingleNode("div/div[2]/a").GetAttributeValue("href", "").Trim())
                End Try
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If

                Try
                    product.Price = Double.Parse(productDetails.SelectSingleNode("div/div[@class='p-price']/span[@class='p-price-o']").InnerText.Replace("$", "").Trim())
                Catch ex As Exception
                    
                End Try
                Try
                    product.Discount = Double.Parse(productDetails.SelectSingleNode("div/div[@class='p-price']/span[@class='p-price-n']").InnerText.Replace("$", "").Trim())
                Catch

                End Try
                product.Currency = "$"

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("div/div[@class='p-pic']")
                If (productImage Is Nothing) Then
                    productImage = productDetails.SelectSingleNode("div/div[1]")
                End If
                product.PictureUrl = efHelper.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("src", ""))
                If (product.PictureUrl.Contains("default.")) Then
                    product.PictureUrl = efHelper.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("data-lazysrc", ""))
                End If

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct
                If (Not efHelper.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
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

    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal categoryName As String, ByVal spreadLogin As String)
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then
            preSubject = "EachBuyer Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "EachBuyer New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HP")) Then
            preSubject = "EachBuyer Special For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")

        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        If (planType.Trim = "HO" OrElse planType.Trim = "HA" OrElse planType.Trim.Contains("HP")) Then
            If (spreadLogin.ToLower.Trim = "lihuojian@hofan.cn") Then
                If (planType.Trim.Contains("HP")) Then
                    addNum = 1
                Else
                    addNum = 8
                End If
            Else
                addNum = 1
            End If
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
        Try
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
            myIssue.Subject = subject
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("AddIssueSubject Error issueId:" & issueId & " siteId：" & siteId & " plantype:" & planType)
            Throw ex
        End Try

    End Sub

    Public Sub InsertAdsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal listAdId As List(Of Integer), ByVal iProIssueCount As Integer)
        Try
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
        Catch ex As Exception
            LogText("InsertAdsIssue error " & "siteId:" & siteId & " issueId:" & issueId)
            Throw ex
        End Try


    End Sub

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
                                    ByVal iProIssueCount As Integer)

        Try
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
        Catch ex As Exception
            LogText("eacherbuyer InsertProductsIssue error")
        End Try

    End Sub

    Private Function sortTopWeeklyDealProducts(ByRef listProducts As List(Of Product), ByVal topCount As Integer)
        '交换排序，按照产品价格product.Discount从大到小排序
        Dim maxProduct As New Product()
        Dim temp As New Product()
        Dim k As Integer
        Try
            '只排好价格product.Discount最高的前topCount个产品，其余的不关心
            For i As Integer = 0 To topCount - 1
                maxProduct = listProducts(i)
                k = i
                For j As Integer = i + 1 To listProducts.Count - 1
                    If (listProducts(j).Discount > maxProduct.Discount) Then
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
            Throw New Exception("error occured when sort products by discount ")
        End Try
    End Function

    Private Function getCategoryName(ByVal Url As String) As String
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(Url, 60000)
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@class='breadcrumbs']/ul/li[@class='home']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
        Dim dfdfdf As String = productDivs.InnerText.Replace("&amp;", "#").Replace("&gt;", "@")
        Dim dfa As String() = dfdfdf.Split("@")
        Dim categoryName As String = dfa(1).Replace("#", "&").Trim()
        Return categoryName
    End Function

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

    Private Function isSoldOut(ByVal productUrl As String) As Boolean
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(productUrl, 60000)
        Try
            Dim productDivs As String = hd.DocumentNode.SelectSingleNode("//div[@class='d_noproduct']").InnerText.Trim  'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            Return True
        Catch
            Return False
        End Try
    End Function

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
