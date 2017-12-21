Imports HtmlAgilityPack

Public Class Oasap
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim contactlistCount As Integer
        Dim bannerCategoryName As String
        If planType.Trim.Contains("HO") Then
            '首页（http://www.oasap.com/）大Banner，每期轮流抓取，保证一个月内根据Banner不重复
            bannerCategoryName = "Dresses"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            'hot style ，获取2个产品，月内不重复，如果没有获取到hotstyle则邮件不出hotstyle。Section="DA"
            GetHotStyle(siteId, IssueID, "DA", planType)
            '主推4个分类,每个分类获取4个产品
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "DA", spreadLogin, "HO") 'AddIssueSubject(IssueID, siteId, planType, "DA")

            If (planType.Trim = "HO") Then
                Dim contactListArr() As String
                If (spreadLogin.Trim.ToLower = "ji.yanan@shiyibofeng.com") Then
                    contactListArr = {"zhangdong all open 20140319", "jiyanan all open 20140319", "OASAP_Subscriber"}
                Else
                    contactListArr = {"jiyanan not open 20140319", "zhangdong not open 20140319", "Oasap_Unregistered_11", "Oasap_Unregistered_12", "Oasap_Unregistered_13", "Oasap_Unregistered_14"}
                End If
                efHelper.InsertContactList(IssueID, contactListArr, "draft") 'draft  'waiting

                'Dim saveListName As String = "oasap_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                'Dim sendingStatus As String = "draft"

                'contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

                'saveListName = "oasap_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")

                'contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.Open, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

            ElseIf (planType.Trim.Contains("HO") AndAlso planType.Trim <> "HO") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "oasap_" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                End If
            End If
        ElseIf (planType.Trim.Contains("HA")) Then
            bannerCategoryName = "Dresses"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            'New Arrivals页面（http://www.oasap.com/46-new-arrivals）获取2个产品
            GetNewArrivals(siteId, IssueID, "NE", planType)
            '主推4个分类,每个分类获取4个产品
            GetHotOrNewProducts(siteId, IssueID, "new", planType)
            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, "NE", spreadLogin, "HA") 'AddIssueSubject(IssueID, siteId, planType, "NE")
            'Dim contactListArr() As String = {"Reasonablers"}
            If (planType.Trim = "HA") Then
                Dim contactListArr() As String
                If (spreadLogin.Trim.ToLower = "ji.yanan@shiyibofeng.com") Then
                    contactListArr = {"zhangdong all open 20140319", "jiyanan all open 20140319", "OASAP_Subscriber"}
                Else  'If (spreadLogin.Trim.ToLower = "zhang.dong@shiyibofeng.com") Then
                    contactListArr = {"jiyanan not open 20140319", "zhangdong not open 20140319", "Oasap_Unregistered_11", "Oasap_Unregistered_12", "Oasap_Unregistered_13", "Oasap_Unregistered_14"}
                End If
                efHelper.InsertContactList(IssueID, contactListArr, "draft")

                'Dim saveListName As String = "oasap_NotOpen" & planType.Trim & Now.ToString("yyyyMMdd")
                'Dim sendingStatus As String = "draft"

                'contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.NotOpen, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If


                'saveListName = "oasap_OpenNotClick" & planType.Trim & Now.ToString("yyyyMMdd")
                'contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 30, ChooseStrategy.Open, sendingStatus, spreadLogin, appId)
                'If (contactlistCount > 0) Then
                '    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                'End If

            ElseIf (planType.Trim.Contains("HA") AndAlso planType.Trim <> "HA") Then
                Dim cate As String() = categories.Split("^")
                Dim saveListName As String = "oasap" & planType.Trim & cate(0).Trim & Now.ToString("yyyyMMdd")
                Dim sendingStatus As String = "draft"

                contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
                If (contactlistCount > 0) Then
                    efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
                End If
            End If
        ElseIf (planType.Trim.Contains("HP")) Then
            Dim cate As String() = categories.Split("^")

            Dim categoryUrlC As String
            Dim categoryUrlL As String
            Dim iProIssueCount As Integer = 8

            Dim queryURL = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case cate(0).Trim
                        categoryUrlL = q.Url & "++6"
                    Case cate(1).Trim
                        categoryUrlC = q.Url & "++6"
                End Select
            Next

            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            AddIssueSubject(IssueID, siteId, planType.Trim, cate(0), spreadLogin, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "oasapSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 30, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
        End If
    End Sub

    Public Sub AddIssueSubject(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal section As String, ByVal spreadLogin As String, ByVal categoryName As String)
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then 'Weekly Deals From OASAP
            preSubject = "Weekly Deals From OASAP"
        ElseIf (planType.Contains("HA")) Then
            preSubject = "Weekly New Arrivals From OASAP"
        ElseIf (planType.Contains("HP")) Then
            preSubject = "Weekly Special " & categoryName & " From OASAP"
        End If
        'Promition Banner有 Alt文字时：Hi [FIRSTNAME], WEEKLY DEALS FOR YOU: + [Alt文字, Weekly Deals第一件产品名称]
        'Promition Banner没有 Alt文字时：Hi [FIRSTNAME], WEEKLY DEALS FOR YOU: + [Weekly Deals第一件产品名称]

        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        If (planType.Trim = "HO" OrElse planType.Trim = "HA" OrElse planType.Trim.Contains("HP")) Then
            If (spreadLogin.Trim.ToLower = "ji.yanan@shiyibofeng.com") Then   'ji.yanan@shiyibofeng.com
                addNum = 1
            Else
                If (planType.Trim.Contains("HP")) Then
                    addNum = 1
                Else
                    addNum = 9
                End If
            End If
        Else
            addNum = 0
        End If

        If (planType.Contains("HO") OrElse planType.Contains("HA")) Then
            planType = planType.Substring(0, 2)
        End If
        Dim query = From i In efContext.Issues
                    Where i.IssueDate > "2014-03-12" AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                    Select i
        Dim subject As String = preSubject & " Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueID)
        'subject = subject.Replace("&nbsp;", "")
        'subject = subject.TrimEnd(".")
        'subject = subject & "Just For Test"
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub

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
            Dim cateBags As New Category()
            If Not listCategoryName.Contains("Bags") Then
                cateBags.Category1 = "Bags"
                cateBags.SiteID = siteID
                cateBags.Url = "http://www.oasap.com/20-bags-purses"
                cateBags.LastUpdate = nowTime
                cateBags.Description = "Bags"
                cateBags.ParentID = -1
                efContext.AddToCategories(cateBags)
            End If
            '添加4大分类 
            Dim cateDresses As New Category()
            Dim cateClothing As New Category()
            Dim cateShoes As New Category()
            Dim cateAccessories As New Category()
            If Not listCategoryName.Contains("Dresses") Then
                cateDresses.Category1 = "Dresses"
                cateDresses.SiteID = siteID
                cateDresses.Url = "http://www.oasap.com/7-dresses"
                cateDresses.LastUpdate = nowTime
                cateDresses.Description = "Dresses"
                cateDresses.ParentID = -1
                cateDresses.Gender = "W"
                efContext.AddToCategories(cateDresses)
            End If
            If Not listCategoryName.Contains("Apparel") Then
                cateClothing.Category1 = "Apparel"
                cateClothing.SiteID = siteID
                cateClothing.Url = " http://www.oasap.com/2-apparel"
                cateClothing.LastUpdate = nowTime
                cateClothing.Description = "Apparel"
                cateClothing.ParentID = -1
                cateClothing.Gender = "W"
                efContext.AddToCategories(cateClothing)
            End If
            If Not listCategoryName.Contains("Shoes") Then
                cateShoes.Category1 = "Shoes"
                cateShoes.SiteID = siteID
                cateShoes.Url = " http://www.oasap.com/26-shoes"
                cateShoes.LastUpdate = nowTime
                cateShoes.Description = "Shoes"
                cateShoes.ParentID = -1
                cateShoes.Gender = "W"
                efContext.AddToCategories(cateShoes)
            End If
            If Not listCategoryName.Contains("Accessories") Then
                cateAccessories.Category1 = "Accessories"
                cateAccessories.SiteID = siteID
                cateAccessories.Url = "http://www.oasap.com/13-accessories"
                cateAccessories.LastUpdate = nowTime
                cateAccessories.Description = "Accessories"
                cateAccessories.ParentID = -1
                cateAccessories.Gender = "W"
                efContext.AddToCategories(cateAccessories)
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try
    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://www.oasap.com"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)

            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub GetHotStyle(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal section As String, ByVal planType As String)
        Try
            Dim topCount As Integer = 2
            Dim listProducts As List(Of Product) = GetHotStyleProducts(siteId, Issueid, section, 4, planType) '这个函数中已判断是否和上周重复
            If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
                Dim categoryName As String
                Dim categoryId As Integer

                Dim listProduct As New List(Of Product)
                listProduct = efHelper.GetProductList(siteId)
                Dim listProductId As New List(Of Integer)
                Dim counter As Integer = 0
                For Each li In listProducts
                    Try
                        categoryName = getCategoryName(li.Url)
                        categoryId = efHelper.GetCategoryId(siteId, categoryName)
                    Catch ex As Exception
                        Continue For
                    End Try
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
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    '根据网站的4个分类获取热销部分，请按顺序获取，保证月内不重复，Section：CA
    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim DressesUrl As String = ""
        Dim ApparelUrl As String = ""
        Dim ShoesUrl As String = ""
        Dim AccessoriesUrl As String = ""
        Try
            '从数据库中获取4个分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可
            '如果是获取hot的产品，在获取分类url时，++5
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Dresses"
                        If (HotORNew = "hot") Then
                            DressesUrl = q.Url & "++5"
                        Else
                            DressesUrl = q.Url
                        End If
                    Case "Apparel"
                        If (HotORNew = "hot") Then
                            ApparelUrl = q.Url & "++5"
                        Else
                            ApparelUrl = q.Url
                        End If
                    Case "Shoes"
                        If (HotORNew = "hot") Then
                            ShoesUrl = q.Url & "++5"
                        Else
                            ShoesUrl = q.Url
                        End If
                    Case "Accessories"
                        If (HotORNew = "hot") Then
                            AccessoriesUrl = q.Url & "++5"
                        Else
                            AccessoriesUrl = q.Url
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 4
            GetProducts(siteId, Issueid, DressesUrl, "Dresses", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, ApparelUrl, "Apparel", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, ShoesUrl, "Shoes", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, AccessoriesUrl, "Accessories", iProIssueCount, updateTime, planType)
        Catch ex As Exception
            'LogText(ex.ToString())
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='topWrapper']/div/div/ul/li")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            If Not (banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("href", "").ToString.ToLower.Contains("www.oasap.com")) Then
                If (banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("href", "").ToString.Trim().Trim.StartsWith("/")) Then
                    Ad.Url = "http://www.oasap.com" & banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("href", "").ToString.Trim()
                Else
                    Ad.Url = "http://www.oasap.com/" & banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("href", "").ToString.Trim()
                End If
            Else
                Ad.Url = banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("href", "").ToString.Trim()
            End If

            Try
                Dim index1 As Integer = banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.IndexOf("(")
                Dim index2 As Integer = banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.IndexOf(")")
                If Not (banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).ToLower.Contains("www.oasap.com")) Then
                    If (banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).Trim.StartsWith("/")) Then
                        Ad.PictureUrl = "http://www.oasap.com" & banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    Else
                        Ad.PictureUrl = "http://www.oasap.com/" & banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    End If
                Else
                    Ad.PictureUrl = banners(counter).SelectSingleNode("div").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                End If
            Catch ex As Exception
                'ignore
            End Try

            Try
                Dim index1 As Integer = banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.IndexOf("(")
                Dim index2 As Integer = banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.IndexOf(")")
                If Not (banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).ToLower.Contains("www.oasap.com")) Then
                    If (banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1).Trim.StartsWith("/")) Then
                        Ad.PictureUrl = "http://www.oasap.com" & banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    Else
                        Ad.PictureUrl = "http://www.oasap.com/" & banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                    End If
                Else
                    Ad.PictureUrl = banners(counter).SelectSingleNode("div/div/a").GetAttributeValue("style", "").ToString.Substring(index1 + 1, index2 - index1 - 1)
                End If
            Catch ex As Exception
                'ignore
            End Try

            Try
                If Not (banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("src", "").ToString.ToLower().Contains("www.oasap.com")) Then
                    If (banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("src", "").ToString.Trim.StartsWith("/")) Then
                        Ad.PictureUrl = "http://www.oasap.com" & banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("src", "").ToString
                    Else
                        Ad.PictureUrl = "http://www.oasap.com/" & banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("src", "").ToString
                    End If
                Else
                    Ad.PictureUrl = banners(counter).SelectSingleNode("div/div/a/img").GetAttributeValue("src", "").ToString
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

    Public Function GetHotStyleProducts(ByVal siteID As Integer, ByVal Issueid As Integer, ByVal Section As String, ByVal getProductCount As Integer, ByVal planType As String) As List(Of Product)
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim pageUrl As String = "http://www.oasap.com/"
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String
        startDate = Now.AddDays(-30).ToShortDateString() 'Deals控制一月内不重复
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            If (Section = "DA") Then
                productDivs = hd.DocumentNode.SelectNodes("//ul[@id='slider2']/li/div") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
                'ElseIf (Section = "NE") Then
                '    productDivs = hd.DocumentNode.SelectNodes("//section[@id='miain']/section[@class='mb10 newArr']/ul/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            End If
            '获取产品  mb10 newArr
            Dim listProducts As New List(Of Product)
            Dim counter As Integer
            If Not productDivs Is Nothing Then
                For counter = 0 To productDivs.Count - 1
                    Dim product As New Product()

                    product.Prodouct = productDivs(counter).SelectSingleNode("a/span").InnerText
                    If (product.Prodouct.ToString.Length > 40) Then
                        product.Prodouct = product.Prodouct.Substring(0, 40) & "..."
                    End If
                    product.Url = productDivs(counter).SelectSingleNode("a").GetAttributeValue("href", "").Trim()
                    If (listProductUrl.Contains(product.Url)) Then
                        Continue For
                    Else
                        listProductUrl.Add(product.Url)
                    End If
                    'product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                    product.Discount = Double.Parse(productDivs(counter).SelectSingleNode("span").InnerText.Trim("$").Trim())
                    product.Currency = "$"

                    '特殊图片处理
                    product.PictureUrl = productDivs(counter).SelectSingleNode("a/div/div/img").GetAttributeValue("src", "") '2013/4/7添加
                    product.LastUpdate = nowTime
                    'product.picturealt = getDescription(product.Url)
                    product.SiteID = siteID
                    product.PictureAlt = "http://pinterest.com/pin/create/button/?url=" & product.Url & "&media=" & product.PictureUrl
                    If (Not efHelper.IsProductSent(siteID, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的产品不重复：
                        listProducts.Add(product)
                    End If
                    If (listProducts.Count = getProductCount) Then
                        Exit For
                    End If
                Next
            End If
            Return listProducts
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:HotStyle" & " error:" & ex.Message.ToString)
        End Try
    End Function

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product)
            'If (planType.Contains("HP")) Then
            '    listProducts = GetSPListProducts(categoryUrl, siteId, updateTime, categoryName)
            'Else
            listProducts = GetListProducts(categoryUrl, siteId, updateTime, categoryName)
            'End If

            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName) ' GetCategoryId(siteId, categoryName)

            Dim endDate As String = Now.AddDays(1).ToShortDateString()
            Dim startDate As String
            If (categoryName = "Accessories") Then
                startDate = Now.AddDays(-15).ToShortDateString()
            Else
                startDate = Now.AddDays(-33).ToShortDateString()
            End If
            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

                If (returnId > 0 AndAlso Not efHelper.IsProductSent(siteId, li.Url, startDate, endDate, planType)) Then '控制一个月内获取的新品不重复：
                    listProductId.Add(returnId)
                End If
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            efHelper.InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper

        Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
        Dim productDivs As HtmlNodeCollection
        productDivs = hd.DocumentNode.SelectNodes("//div[@id='ColumnContainer']/div")
        For counter = 0 To productDivs.Count - 1
            Dim product As New Product()
            Dim productDetails As HtmlNode = productDivs(counter)
            product.Prodouct = productDetails.SelectSingleNode("a").InnerText
            '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
            If (product.Prodouct.Length > 40) Then
                product.Prodouct = product.Prodouct.Substring(0, 40) & "..."
            End If
            product.Url = productDetails.SelectSingleNode("a").GetAttributeValue("href", "").Trim()

            If (listProductUrl.Contains(product.Url)) Then
                Continue For
            Else
                listProductUrl.Add(product.Url)
            End If

            'If (productDetails.SelectSingleNode("div[@class='productprice']").InnerHtml.Contains("price_o")) Then
            Try
                product.Price = Double.Parse(productDetails.SelectSingleNode("div[@class='p_price']/s").InnerText.Trim("$").Trim())
            Catch ex As Exception
                'ignore
            End Try

            product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='p_price']/span[@class='p_price_now price_n']").InnerText.Trim("$").Trim())
            product.Currency = "$"

            '特殊图片处理 pList2_mg  categoryImageDiv
            'Dim productImage As HtmlNode = productDetails.SelectSingleNode("p[@class='pList2_mg']")
            product.PictureUrl = productDetails.SelectSingleNode("div[@class='p_list_img']/a/img").GetAttributeValue("src", "") '2013/4/7添加
            If (product.PictureUrl.Contains("lazyload")) Then
                product.PictureUrl = productDetails.SelectSingleNode("div[@class='p_list_img']/a/img").GetAttributeValue("data-original", "")
            End If

            product.LastUpdate = nowTime
            product.SiteID = siteId

            product.PictureAlt = "http://pinterest.com/pin/create/button/?url=" & product.Url & "&media=" & product.PictureUrl

            listProducts.Add(product)
        Next

        Return listProducts
    End Function

    'Public Function GetSPListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String) As List(Of Product)
    '    Dim listProducts As New List(Of Product)
    '    Dim helper As New EFHelper
    '    Try
    '        Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
    '        Dim productDivs As HtmlNodeCollection
    '        productDivs = hd.DocumentNode.SelectNodes("//ul[@id='items']/li")
    '        For counter = 0 To productDivs.Count - 1
    '            Dim product As New Product()
    '            Dim productDetails As HtmlNode = productDivs(counter)
    '            product.Prodouct = productDetails.SelectSingleNode("a").InnerText
    '            '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
    '            If (product.Prodouct.Length > 60) Then
    '                product.Prodouct = product.Prodouct.Substring(0, 60) & "..."
    '            End If
    '            product.Url = productDetails.SelectSingleNode("a").GetAttributeValue("href", "").Trim()

    '            'If (listProductUrl.Contains(product.Url)) Then
    '            '    Continue For
    '            'Else
    '            '    listProductUrl.Add(product.Url)
    '            'End If

    '            'If (productDetails.SelectSingleNode("div[@class='productprice']").InnerHtml.Contains("price_o")) Then
    '            Try
    '                product.Price = Double.Parse(productDetails.SelectSingleNode("div[@class='productprice']/span[@class='price_o']").InnerText.Trim("$").Trim())
    '            Catch ex As Exception
    '                'ignore
    '            End Try

    '            product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='productprice']/span[@class='saleprice']/span[@class='price_n']").InnerText.Trim("$").Trim())
    '            product.Currency = "$"

    '            '特殊图片处理 pList2_mg  categoryImageDiv
    '            'Dim productImage As HtmlNode = productDetails.SelectSingleNode("p[@class='pList2_mg']")
    '            product.PictureUrl = productDetails.SelectSingleNode("div[@class='categoryImageDiv']/a/img").GetAttributeValue("src", "") '2013/4/7添加
    '            If (product.PictureUrl.Contains("lazyload")) Then
    '                product.PictureUrl = productDetails.SelectSingleNode("div[@class='categoryImageDiv']/a/img").GetAttributeValue("data-original", "")
    '            End If

    '            product.LastUpdate = nowTime
    '            product.SiteID = siteId

    '            product.PictureAlt = "http://pinterest.com/pin/create/button/?url=" & product.Url & "&media=" & product.PictureUrl

    '            listProducts.Add(product)
    '        Next
    '    Catch ex As Exception
    '        Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
    '    End Try
    '    Return listProducts
    'End Function

    Private Sub GetNewArrivals(ByVal siteId As Integer, ByVal issueId As Integer, ByVal section As String, ByVal planType As String)
        Try
            Dim topCount As Integer = 2
            Dim listProducts As List(Of Product) = GetNewArriProducts(siteId, issueId, section, 4, planType) '这个函数中已判断是否和上周重复
            If (listProducts.Count >= topCount) Then '如果个数不足，则不处理
                Dim categoryName As String
                Dim categoryId As Integer

                Dim listProduct As New List(Of Product)
                listProduct = efHelper.GetProductList(siteId)
                Dim listProductId As New List(Of Integer)
                Dim counter As Integer = 0
                For Each li In listProducts
                    Try
                        categoryName = getCategoryName(li.Url)
                        categoryId = efHelper.GetCategoryId(siteId, categoryName)
                    Catch ex As Exception
                        Continue For
                    End Try
                    Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)
                    listProductId.Add(returnId)

                    '只插入排好序的前topCount个产品
                    counter += 1
                    If (counter = topCount) Then
                        Exit For
                    End If
                Next
                efHelper.InsertProductsIssue(siteId, issueId, section, listProductId, topCount)
            End If
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    Public Function GetNewArriProducts(ByVal siteID As Integer, ByVal Issueid As Integer, ByVal Section As String, ByVal getProductCount As Integer, ByVal planType As String) As List(Of Product)
        Dim helper As New EFHelper
        Dim nowTime As DateTime = Now
        Dim pageUrl As String = "http://www.oasap.com/46-new-arrivals"
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String
        startDate = Now.AddDays(-30).ToShortDateString() 'Deals控制一月内不重复
        Dim listProducts As New List(Of Product)

        Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
        Dim productDivs As HtmlNodeCollection
        productDivs = hd.DocumentNode.SelectNodes("//div[@id='ColumnContainer']/div")
        For counter = 0 To productDivs.Count - 1
            Dim product As New Product()
            Dim productDetails As HtmlNode = productDivs(counter)
            product.Prodouct = productDetails.SelectSingleNode("a").InnerText
            '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
            If (product.Prodouct.Length > 60) Then
                product.Prodouct = product.Prodouct.Substring(0, 60) & "..."
            End If
            product.Url = productDetails.SelectSingleNode("a").GetAttributeValue("href", "").Trim()

            If (listProductUrl.Contains(product.Url)) Then
                Continue For
            Else
                listProductUrl.Add(product.Url)
            End If

            'If (productDetails.SelectSingleNode("div[@class='productprice']").InnerHtml.Contains("price_o")) Then
            Try
                product.Price = Double.Parse(productDetails.SelectSingleNode("div[@class='p_price']/s").InnerText.Trim("$").Trim())
            Catch ex As Exception
                'ignore
            End Try

            product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='p_price']/span[@class='p_price_now price_n']").InnerText.Trim("$").Trim())
            product.Currency = "$"

            '特殊图片处理 pList2_mg  categoryImageDiv
            'Dim productImage As HtmlNode = productDetails.SelectSingleNode("p[@class='pList2_mg']")
            product.PictureUrl = productDetails.SelectSingleNode("div[@class='p_list_img']/a/img").GetAttributeValue("src", "") '2013/4/7添加
            If (product.PictureUrl.Contains("lazyload")) Then
                product.PictureUrl = productDetails.SelectSingleNode("div[@class='p_list_img']/a/img").GetAttributeValue("data-original", "")
            End If

            product.LastUpdate = nowTime
            'product.Description = getDescription(product.Url)
            product.SiteID = siteID

            product.PictureAlt = "http://pinterest.com/pin/create/button/?url=" & product.Url & "&media=" & product.PictureUrl


            If (Not efHelper.IsProductSent(siteID, product.Url, startDate, endDate, planType)) Then '控制一个月内获取的产品不重复：
                listProducts.Add(product)
            End If
            If (listProducts.Count = getProductCount) Then
                Exit For
            End If
        Next
        Return listProducts
    End Function

    Private Function getCategoryName(ByVal url As String) As String

        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(url, 60000)
        Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//ul[@class='breadcrumbs']/li") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
        Dim categoryName As String = productDivs(1).SelectSingleNode("a").InnerText
        If (categoryName.Trim.ToLower.Contains("clothing")) Then
            categoryName = "Apparel"
        ElseIf (categoryName.Trim.ToLower.Contains("shoe")) Then
            categoryName = "Shoes"
        ElseIf (categoryName.Trim.ToLower.Contains("accessories")) Then
            categoryName = "Accessories"
        Else
            categoryName = "Dresses"
        End If
        If (productDivs(2).SelectSingleNode("a").InnerText.Trim.ToLower().Contains("dresses")) Then
            categoryName = "Dresses"
        End If
        Return categoryName
    End Function

    'Private Function getDescription(ByVal url As String) As String  '<span itemprop="description">
    '    Dim description As String = ""
    '    Try
    '        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(url, 60000)
    '        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//span[@itemprop='description']/p/font/b") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
    '        description = productDivs.InnerText
    '        If (description.Length > 170) Then
    '            description = description.Substring(0, 170) & "..."
    '        End If
    '        Return description
    '    Catch ex As Exception
    '        'ignore
    '    End Try
    'End Function

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
