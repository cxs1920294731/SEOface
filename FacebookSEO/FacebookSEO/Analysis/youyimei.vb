
Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Public Class youyimei
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        Dim contactlistCount As Integer = 0
        GetCategory(siteId)
        If planType.Trim = "HO" Then '新品
            Dim bannerCategoryName As String = "新品"
            GetBanner(siteId, IssueID, bannerCategoryName, planType, "http://youyimei.tmall.com/", 3)

            Dim pageUrl As String = "http://youyimei.tmall.com/search.htm?spm=a1z10.3-b.w4011-2636531587.87.Yg1UqT&scene=taobao_shop&search=y&orderType=newOn_desc&tsearch=y"
            efHelper.FetchProduct(pageUrl, bannerCategoryName, "CA", planType, 6, DateTime.Now.AddDays(-20), DateTime.Now, siteId, IssueID)
            'GetNewProducts(siteId, IssueID, planType)
            'GetCategoryProducts(siteId, IssueID, "new", planType)

            Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "亲爱的[FIRSTNAME]，本周优衣美为您带来：" & querySubject.FirstOrDefault.ToString & "(AD)")
        ElseIf (planType.Trim = "HP") Then  '针对开启没点击的邮件，去Hot Seller页面抓取6个产品)，则将产品归为openNoClick1,openNoClick2两个分类，则思路同其余关联销售的专题。
            Dim categoryUrl As String = "http://youyimei.tmall.com/search.htm?search=y&orderType=hotsell_desc&tsearch=y"
            Dim iProIssueCount = 3
            GetProducts(siteId, IssueID, categoryUrl, "热卖专辑", iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrl, "掌柜推荐", iProIssueCount, Now, planType)

            Dim sendingStatus = "draft"
            Dim saveListName As String = "youyimeiSpecialNoclick" & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 90, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If

            'Dim querySubject = From p In efContext.Products
            '         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
            '         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
            '         Select p.Prodouct
            'Dim endSubject As String = ""
            'For Each p In querySubject
            '    If (Not String.IsNullOrEmpty(p) AndAlso p.ToString.Length > 3) Then
            '        endSubject = p.ToString.Trim
            '        Exit For
            '    End If
            'Next
            Dim query = From i In efContext.Issues
                     Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
            efHelper.InsertIssueSubject(IssueID, "优衣美本周为您特别推荐：" & "热卖专辑/掌柜推荐" & "（第" & (query.Count + 1).ToString & "期）" & "(AD)")
        Else '个性化邮件
            '按销量排序url参数?orderType=hotsell_desc
            Dim cate As String() = categories.Split("^")

            Dim categoryUrlC As String
            Dim categoryUrlL As String
            Dim iProIssueCount As Integer = 3

            Dim queryURL = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case cate(0).Trim
                        categoryUrlL = q.Url.Trim & "?orderType=hotsell_desc"
                    Case cate(1).Trim
                        categoryUrlC = q.Url.Trim & "?orderType=hotsell_desc"
                End Select
            Next
            If (cate(0).Trim = "新品") Then
                categoryUrlL = "http://youyimei.tmall.com/search.htm?orderType=newOn_desc&tsearch=y"
            End If
            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            'AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "youyimeiSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 90, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If
            Dim saveListName1 As String = "youyimeiSpecial" & cate(1).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(1).Trim, saveListName1, 90, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName1, sendingStatus)
            End If

            Dim nowYear As String = Date.Now.ToString("yyyy")
            Dim query = From i In efContext.Issues
                     Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
            efHelper.InsertIssueSubject(IssueID, "优衣美本周为您特别推荐：" & cate(0).Trim & "/" & cate(1).Trim & "（第" & (query.Count + 1).ToString & "期）" & "(AD)")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper

        Dim myCategory6 As New Category
        myCategory6.Category1 = "热卖专辑"
        myCategory6.Description = "热卖专辑"
        myCategory6.Url = "http://youyimei.tmall.com/search.htm?search=y&orderType=hotsell_desc&tsearch=y"
        myCategory6.SiteID = siteId
        myCategory6.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory6, siteId)
        Dim myCategory7 As New Category
        myCategory7.Category1 = "掌柜推荐"
        myCategory7.Description = "掌柜推荐"
        myCategory7.Url = "http://youyimei.tmall.com/search.htm?spm=a1z10.3.w4011-2636531587.1.U1HkAL"
        myCategory7.SiteID = siteId
        myCategory7.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory7, siteId)

        Dim myCategory As New Category
        myCategory.Category1 = "新品"
        myCategory.Description = "新品"
        myCategory.Url = "http://youyimei.tmall.com/"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory1 As New Category
        myCategory1.Category1 = "上衣"
        myCategory1.Description = "上衣"
        myCategory1.Url = "http://youyimei.tmall.com/category-798328409.htm"
        myCategory1.SiteID = siteId
        myCategory1.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory1, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "休闲裤"
        myCategory2.Description = "休闲裤"
        myCategory2.Url = "http://youyimei.tmall.com/category-798327969.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "打底裤"
        myCategory3.Description = "打底裤"
        myCategory3.Url = "http://youyimei.tmall.com/category-798327972.htm"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "牛仔裤"
        myCategory4.Description = "牛仔裤"
        myCategory4.Url = "http://youyimei.tmall.com/category-798327970.htm"
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory4, siteId)

        Dim myCategory5 As New Category
        myCategory5.Category1 = "西裤/正裤"
        myCategory5.Description = "西裤/正裤"
        myCategory5.Url = "http://youyimei.tmall.com/category-798327971.htm"
        myCategory5.SiteID = siteId
        myCategory5.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory5, siteId)
    End Sub

    Private Sub GetNewProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String)
        Dim pageUrl As String = "http://youyimei.tmall.com/category-798327969.htm"
        Dim hd As HtmlDocument
        Try
            hd = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10
        Catch ex As Exception
            hd = efHelper.GetHtmlDocByUrlTmall(pageUrl)
        End Try
        Dim xinpinUrl As HtmlNode = hd.DocumentNode.SelectNodes("//li[@class='cat snd-cat']")(1).SelectSingleNode("h4/a")
        Dim xinpinUrls As String = xinpinUrl.GetAttributeValue("href", "").ToString.Trim
        GetProducts(siteId, issueId, xinpinUrls, "新品", 3, Now, "HO")
    End Sub

    Public Sub GetCategoryProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim shangyiUrl As String = ""
        Dim dadikuUrl As String = ""
        Dim xiuxiankuUrl As String = ""
        Dim xikuzhengkuUrl As String = ""
        Dim niuzaikuUrl As String = ""

        Try
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "上衣"
                        shangyiUrl = q.Url.Trim()
                    Case "休闲裤"
                        xiuxiankuUrl = q.Url.Trim()
                    Case "打底裤"
                        dadikuUrl = q.Url.Trim()
                    Case "西裤/正裤"
                        xikuzhengkuUrl = q.Url.Trim()
                    Case "牛仔裤"
                        niuzaikuUrl = q.Url.Trim()
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            GetProducts(siteId, Issueid, shangyiUrl, "上衣", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, xiuxiankuUrl, "休闲裤", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, dadikuUrl, "打底裤", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, xikuzhengkuUrl, "西裤/正裤", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, niuzaikuUrl, "牛仔裤", iProIssueCount, updateTime, planType)

        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Public Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product)
            listProducts = efHelper.GetAsynProductList(categoryUrl, siteId)
            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)

            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount * 2) Then
                    Exit For
                End If
            Next

            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String,
                     ByVal pageUrl As String, ByVal bannerImgIndex As Integer)

        Dim ad As New Ad()
        Dim imgRegex As String = "background:url\(([\w\W]*?)\)"
        Dim txtDocHtml As String = EFHelper.GetHtmlStringByUrlAli(pageUrl)
        Dim mCollection As MatchCollection = Regex.Matches(txtDocHtml, imgRegex, RegexOptions.Multiline)
        Dim matched As Match = mCollection(bannerImgIndex - 1)
        If (matched.Success) Then
            Dim bannerImg As String = mCollection(bannerImgIndex - 1).Groups(1).Value.Trim
            Dim srcRegex As String = bannerImg & "[\w\W]*?<a[\w\W]*?href=""([\w\W]*?)"""
            ad.PictureUrl = bannerImg
            Dim srcMatched As Match = Regex.Match(txtDocHtml, srcRegex, RegexOptions.Multiline)
            If (srcMatched.Success) Then
                ad.Url = srcMatched.Groups(1).Value.Trim
            End If
        End If

        If Not String.IsNullOrEmpty(ad.PictureUrl) Then '对banner不重复时长未做限定
            If Not (String.IsNullOrEmpty(ad.Url)) Then
                ad.Url = efHelper.AddHttpForAli(ad.Url)
            End If

            ad.PictureUrl = efHelper.AddHttpForAli(ad.PictureUrl)
            ad.SiteID = siteId
            ad.Lastupdate = Now

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteId)
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim adid As Integer = efHelper.InsertAd(ad, listAd, categoryId)
            efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)
        End If
    End Sub




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

    

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
