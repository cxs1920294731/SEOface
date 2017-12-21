Imports HtmlAgilityPack

Public Class BostantenTmall

    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        Dim contactlistCount
        GetCategory(siteId)
        If planType.Trim = "HO" Then '热卖
            Dim bannerCategoryName As String = "公文包"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            GetHOTProducts(siteId, planType, IssueID)
            'If (productCounter < 9) Then
            '    Dim categoryUrl As String = "http://octlegendast.tmall.com/p/rd415809.htm"
            '    Dim requestUrl As String = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
            '    GetTopSellProducts(categoryUrl, requestUrl, siteId, planType, IssueID, lastUpdate, 9 - productCounter)
            'End If
            efHelper.InsertContactList(IssueID, "Opens20140401", "draft")
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
            efHelper.InsertIssueSubject(IssueID, "波斯丹顿本周精选：" & endSubject & "(AD)")
        ElseIf (planType.Trim = "HP") Then  '针对开启没点击的邮件，去Hot Seller页面抓取6个产品)，则将产品归为openNoClick1,openNoClick2两个分类，则思路同其余关联销售的专题。
            Dim categoryUrl As String = "http://bosidandun.tmall.com/search.htm?search=y&orderType=hotsell_desc&tsearch=y"
            Dim iProIssueCount = 3
            GetProducts(siteId, IssueID, categoryUrl, "特别呈现", iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrl, "精彩为您", iProIssueCount, Now, planType)

            'Add Subject
            'AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "BostantenSpecialNoclick" & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, saveListName, 90, ChooseStrategy.OpenExcludeCategory, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If

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
            efHelper.InsertIssueSubject(IssueID, "波斯丹顿本周特别呈现：" & endSubject & "(AD)")
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
            GetProducts(siteId, IssueID, categoryUrlL, cate(0), iProIssueCount, Now, planType)
            GetProducts(siteId, IssueID, categoryUrlC, cate(1), iProIssueCount, Now, planType)

            'Add Subject
            'AddIssueSubject(IssueID, siteId, planType.Trim, cate(0))
            ''Add 
            Dim sendingStatus = "draft"
            Dim saveListName As String = "BostantenSpecial" & cate(0).Trim & Now.ToString("yyyyMMdd")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate(0).Trim, saveListName, 90, ChooseStrategy.Favorite, sendingStatus, spreadLogin, appId)
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName, sendingStatus)
            End If

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
            efHelper.InsertIssueSubject(IssueID, "波斯丹顿本周特别呈现：" & endSubject & "(AD)")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "公文包"
        myCategory.Description = "gongwenbao"
        myCategory.Url = "http://bosidandun.tmall.com/category-246022099.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "手提包"
        myCategory2.Description = "shoutibao"
        myCategory2.Url = "http://bosidandun.tmall.com/category-233557762.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "单肩包"
        myCategory3.Description = "danjianbao"
        myCategory3.Url = "http://bosidandun.tmall.com/category-233557761.htm"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "钱包"
        myCategory4.Description = "qianbao"
        myCategory4.Url = "http://bosidandun.tmall.com/category-278506777.htm "
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory4, siteId)

        Dim myCategory5 As New Category
        myCategory5.Category1 = "腰包/胸包"
        myCategory5.Description = "yaobaoxiongbao"
        myCategory5.Url = "http://bosidandun.tmall.com/category-406419920.htm"
        myCategory5.SiteID = siteId
        myCategory5.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory5, siteId)

        Dim myCategory6 As New Category
        myCategory6.Category1 = "皮带"
        myCategory6.Description = "pidai"
        myCategory6.Url = "http://bosidandun.tmall.com/category-351865825.htm"
        myCategory6.SiteID = siteId
        myCategory6.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory6, siteId)

        Dim myCategory7 As New Category
        myCategory7.Category1 = "特别呈现"
        myCategory7.Description = "特别呈现"
        myCategory7.Url = "http://bosidandun.tmall.com/search.htm?search=y&orderType=hotsell_desc&tsearch=y"
        myCategory7.SiteID = siteId
        myCategory7.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory7, siteId)

        Dim myCategory8 As New Category
        myCategory8.Category1 = "精彩为您"
        myCategory8.Description = "精彩为您"
        myCategory8.Url = "http://bosidandun.tmall.com/search.htm?search=y&orderType=hotsell_desc"
        myCategory8.SiteID = siteId
        myCategory8.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory8, siteId)
    End Sub

    Private Sub GetHOTProducts(ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer)
        Dim gongwenbaoUrl As String = ""
        Dim shoutibaoUrl As String = ""
        Dim danjianbaoUrl As String = ""
        Dim qianbaoUrl As String = ""
        Dim yaobao_xiongbaoUrl As String = ""
        Dim pidaiUrl As String = ""
        Try
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteID
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "公文包"
                        gongwenbaoUrl = q.Url
                    Case "手提包"
                        shoutibaoUrl = q.Url
                    Case "单肩包"
                        danjianbaoUrl = q.Url
                    Case "钱包"
                        qianbaoUrl = q.Url
                    Case "腰包/胸包"
                        yaobao_xiongbaoUrl = q.Url
                    Case "皮带"
                        pidaiUrl = q.Url
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            iProIssueCount = 2 '在模板中要填充产品的个数
            GetProducts(siteID, issueID, gongwenbaoUrl, "公文包", iProIssueCount, updateTime, planType)
            GetProducts(siteID, issueID, shoutibaoUrl, "手提包", iProIssueCount, updateTime, planType)
            GetProducts(siteID, issueID, danjianbaoUrl, "单肩包", iProIssueCount, updateTime, planType)
            iProIssueCount = 1
            GetProducts(siteID, issueID, qianbaoUrl, "钱包", iProIssueCount, updateTime, planType)
            GetProducts(siteID, issueID, yaobao_xiongbaoUrl, "腰包/胸包", iProIssueCount, updateTime, planType)
            GetProducts(siteID, issueID, pidaiUrl, "皮带", iProIssueCount, updateTime, planType)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://bosidandun.tmall.com/"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Try
            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
            InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-30).ToShortDateString()
        Dim hd As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(pageUrl) 'slideBox mt10 mb10

        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='main']/div/a")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            Dim ad As New Ad()
            ad.Url = banners(counter).GetAttributeValue("href", "").ToString.Trim()
            ad.Url = efHelper.addDominForUrl("http://bosidandun.tmall.com", ad.Url)
            Try
                ad.PictureUrl = banners(counter).GetAttributeValue("style", "").ToString.Trim()
                Dim index1 As Integer = ad.PictureUrl.IndexOf("(")
                Dim index2 As Integer = ad.PictureUrl.IndexOf(")")
                ad.PictureUrl = ad.PictureUrl.Substring(index1 + 1, index2 - index1 - 1)
            Catch ex As Exception
                ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "").ToString.Trim()
            End Try

            ad.PictureUrl = efHelper.addDominForUrl("http://bosidandun.tmall.com", ad.PictureUrl)

            If (efHelper.isAdSent(siteid, ad.PictureUrl, startDate, endDate, planType, True)) Then '控制2周内获取的Ad不重复：
                Continue For
            End If

            ad.SiteID = siteid
            ad.Lastupdate = Now
            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteid)

            adid = efHelper.InsertAd(ad, listAd, categroyId)
            listAdIds.Add(adid)
        Next
        Return listAdIds
    End Function

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

                If (listProductId.Count = iProIssueCount * 2) Then
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
            Dim index As Integer
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
                        myProduct.Url = dl.SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                        If (myProduct.Url.Contains("&")) Then
                            index = myProduct.Url.IndexOf("&")
                            myProduct.Url = myProduct.Url.Remove(index)
                        End If
                        If (listProductUrl.Contains(myProduct.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(myProduct.Url)
                        End If

                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now, planType)) Then
                            myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                            myProduct.Description = productName
                            myProduct.LastUpdate = lastUpdate
                            myProduct.SiteID = siteId
                            listProducts.Add(myProduct)
                        End If
                    Next
                End If
                If (listProducts.Count >= iProCount * 2) Then
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
