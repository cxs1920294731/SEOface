Imports HtmlAgilityPack

Public Class michaelkors
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private domin As String = "http://www.michaelkorsplazas.com"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String, ByVal dlltype As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        planType = planType.Trim
        If planType.Trim.Contains("HO") Then
            bannerCategoryName = "Grayson"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType, url)
            ' ''hot style ，获取2个产品，月内不重复，如果没有获取到hotstyle则邮件不出hotstyle。Section="DA"
            'GetWeeklyDealsOrNewArri(siteId, IssueID, "DA", planType)
            ' ''主推4个分类,每个分类获取4个产品
            GetHotOrNewProducts(siteId, IssueID, "hot", planType, url)
            ' ''Add Subject
            AddIssueSubject(IssueID, siteId, "HO", "DA") 'AddIssueSubject(IssueID, siteId, planType, "DA")
            '操作联系人名单步骤：
            '第1步：手动在废表SplitContactLists手动插入数据：ShopName为 dlltype 的ContactList
            Dim splitContactLists As New List(Of String)
            splitContactLists = (From s In efContext.SplitContactLists _
                                 Where (s.ShopName = dlltype.Trim AndAlso s.SiteID = siteId AndAlso s.PlanType = planType AndAlso s.Flag = "False")
                                 Select s.ContactListName).ToList()
            '第2步：去读废表SplitContactLists里ShopName为 michaelkorsplazas 中Flag为 False的ContactList
            Dim currentContactName As String = splitContactLists(0).ToString.Trim
            efHelper.InsertContactList(IssueID, currentContactName, "draft") 'draft  'waiting
            '第3部：拿到第一个ContactList插入到表ContactLists_Issue中，并且在表SplitContactLists中将此记录中Flag更新为True

            Dim update = (From s In efContext.SplitContactLists
                          Where (s.ShopName = dlltype.Trim AndAlso s.SiteID = siteId AndAlso s.PlanType = planType _
                          AndAlso s.ContactListName = currentContactName)).FirstOrDefault()
            update.Flag = True
            efContext.SaveChanges()
            If (splitContactLists.Count = 1) Then
                splitContactLists = (From s In efContext.SplitContactLists _
                     Where (s.ShopName = dlltype.Trim AndAlso s.SiteID = siteId AndAlso s.PlanType = planType AndAlso s.Flag = "True")
                     Select s.ContactListName).ToList()

                For Each sc In splitContactLists
                    Dim update1 = (From s In efContext.SplitContactLists
                                 Where (s.ShopName = dlltype.Trim AndAlso s.SiteID = siteId AndAlso s.PlanType = planType _
                                 AndAlso s.ContactListName = sc)).FirstOrDefault()
                    update1.Flag = False
                Next
            End If
            efContext.SaveChanges()
            '第4部：当废表SplitContactLists里ShopName为 michaelkorsplazas 中所有Contactlist的Flag都为Ture时，更新所有相关记录Flag为False,重复第2步
        End If
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
            '添加13大分类 
            Dim cateSatchelsgrayson As New Category()
            Dim cateToteshamilton As New Category()
            Dim cateTotesjetset As New Category()
            Dim cateWallets As New Category()
            Dim cateShoes As New Category()
            Dim cateWatches As New Category()
            Dim cateJewelry As New Category()

            If Not listCategoryName.Contains("Grayson") Then
                cateSatchelsgrayson.Category1 = "Grayson"
                cateSatchelsgrayson.SiteID = siteID
                cateSatchelsgrayson.Url = "http://www.michaelkorsplazas.com/michael-kors-satchels-grayson-c-79_80.html?page=1"
                cateSatchelsgrayson.LastUpdate = nowTime
                cateSatchelsgrayson.Description = "Grayson"
                cateSatchelsgrayson.ParentID = -1
                efContext.AddToCategories(cateSatchelsgrayson)
            End If
            If Not listCategoryName.Contains("Hamilton") Then
                cateToteshamilton.Category1 = "Hamilton"
                cateToteshamilton.SiteID = siteID
                cateToteshamilton.Url = "http://www.michaelkorsplazas.com/michael-kors-totes-hamilton-c-88_96.html?page=1"
                cateToteshamilton.LastUpdate = nowTime
                cateToteshamilton.Description = "Hamilton"
                cateToteshamilton.ParentID = -1
                efContext.AddToCategories(cateToteshamilton)
            End If
            If Not listCategoryName.Contains("Jet Set") Then
                cateTotesjetset.Category1 = "Jet Set"
                cateTotesjetset.SiteID = siteID
                cateTotesjetset.Url = "http://www.michaelkorsplazas.com/michael-kors-totes-jet-set-c-88_97.html?page=1"
                cateTotesjetset.LastUpdate = nowTime
                cateTotesjetset.Description = "Jet Set"
                cateTotesjetset.ParentID = -1
                efContext.AddToCategories(cateTotesjetset)
            End If
            If Not listCategoryName.Contains("Wallets Jet Set") Then
                cateWallets.Category1 = "Wallets Jet Set"
                cateWallets.SiteID = siteID
                cateWallets.Url = "http://www.michaelkorsplazas.com/michael-kors-wallets-jet-set-c-101_104.html?page=1"
                cateWallets.LastUpdate = nowTime
                cateWallets.Description = "Wallets Jet Set"
                cateWallets.ParentID = -1
                efContext.AddToCategories(cateWallets)
            End If
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.Message.ToString)
        End Try
    End Sub

    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String, ByVal Url As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = Url
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
        Dim startDate As String = Now.AddDays(-33).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='brand']/div/div/a")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()

            Try
                Ad.Url = efHelper.addDominForUrl(domin, banners(counter).GetAttributeValue("href", "").ToString)
            Catch ex As Exception
                'ignore
            End Try

            Try
                Ad.PictureUrl = efHelper.addDominForUrl(domin, banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString)
            Catch ex As Exception
                'ignore
            End Try

            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteid)
            If (Not efHelper.isAdSent(siteid, Ad.PictureUrl, startDate, endDate, planType, True) AndAlso Ad.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                adid = InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
            End If
        Next

        Return listAdIds
    End Function

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

    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String, ByVal url As String)
        Dim GraysonUrl As String = ""
        Dim HamiltonUrl As String = ""
        Dim Jet_SetUrl As String = ""
        Dim Wallets_Jet_SetUrl As String = ""
        Try
            '从数据库中获取5个主推分类的url，不直接赋值的原因是，考虑如果某个分类url有变动，只需要修改数据库即可

            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 4
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "Grayson"
                        GraysonUrl = q.Url
                    Case "Hamilton"
                        HamiltonUrl = q.Url
                    Case "Jet Set"
                        Jet_SetUrl = q.Url
                    Case "Wallets Jet Set"
                        Wallets_Jet_SetUrl = q.Url
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            GetProducts(siteId, Issueid, GraysonUrl, "Grayson", iProIssueCount, updateTime, planType, url)
            GetProducts(siteId, Issueid, HamiltonUrl, "Hamilton", iProIssueCount, updateTime, planType, url)
            GetProducts(siteId, Issueid, Jet_SetUrl, "Jet Set", iProIssueCount, updateTime, planType, url)
            GetProducts(siteId, Issueid, Wallets_Jet_SetUrl, "Wallets Jet Set", iProIssueCount, updateTime, planType, url)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String, ByVal url As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, url)
            If (listProducts.Count < iProIssueCount) Then
                Dim index As Integer = categoryUrl.LastIndexOf("=")
                If (Integer.Parse(categoryUrl.Substring(index + 1, 1)) + 1 > 6) Then
                    categoryUrl = categoryUrl.Substring(0, index + 1) & "1"
                Else
                    categoryUrl = categoryUrl.Substring(0, index + 1) & Integer.Parse(categoryUrl.Substring(index + 1, 1)) + 1
                End If

                Dim updateCategory = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = categoryName)
                updateCategory.Url = categoryUrl
                efContext.SaveChanges()
                listProducts = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, url)
            End If
            Dim listProduct As List(Of Product) = efHelper.GetProductList(siteId)
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
    End Sub

    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                    ByVal url As String) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-140).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@class='deptCrossSellBox']/div")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)
                product.Prodouct = productDetails.SelectSingleNode("div[2]").SelectSingleNode("a").InnerText
                '限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                product.Url = efHelper.addDominForUrl(domin, productDetails.SelectSingleNode("div[2]").SelectSingleNode("a").GetAttributeValue("href", "").Trim())

                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If

                product.Price = Double.Parse(productDetails.SelectSingleNode("span[1]").InnerText.Trim("$").Trim())
                product.Discount = Double.Parse(productDetails.SelectSingleNode("span[2]").InnerText.Trim("$").Replace("(USD)", "").Trim())
                product.Currency = "$"

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("div[1]")

                product.PictureUrl = efHelper.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("src", "").ToString)

                If (product.PictureUrl.Contains("grey.gif")) Then
                    product.PictureUrl = efHelper.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("data-original", ""))
                End If
                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId

                If (Not efHelper.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制获取的产品不重复：
                    listProducts.Add(product)
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal categoryName As String)
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then
            preSubject = "MichaelKors Deals For " 'Sammydress Deals For Jan.2014.Vol.01
            'ElseIf (planType.Contains("HA")) Then
            '    preSubject = "UTsource Electronic Components For " 'Sammydress New Arrivals For Jan.2014.Vol.01
            'ElseIf (planType.Contains("HP")) Then
            '    preSubject = "UTsource  Special " & categoryName & " For " 'Sammydress New Arrivals For Jan.2014.Vol.01
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
        Dim subject As String = preSubject & "Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
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
