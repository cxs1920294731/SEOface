Imports HtmlAgilityPack
Public Class DressLilyWomen
    Private listCategory As New List(Of Category)
    Private listProductUrl As New List(Of String)
    Private mySubject As String = ""

    ''' <summary>
    ''' 主程序入口
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="splitContactCount"></param>
    ''' <param name="spreadLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="url"></param>
    ''' <remarks></remarks>
    Public Sub Start(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        Dim nowTime As DateTime = Now
        Dim helper As New EFHelper()
        Try
            listProductUrl = helper.GetListProduct(siteId)
            GetCategory(siteId, issueId, nowTime)

            Try
                '获取Women's Clothing的产品，并使用1个产品作为Promotion
                Dim menClothingUrl As String = "http://www.dresslily.com/women-s-clothing-b-1.html?odr=new"
                GetProductAndProductIssue(menClothingUrl, 8, "Women's Clothing", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            Try
                '获取Sexy Lingerie的产品
                Dim menShoesUrl As String = "http://www.dresslily.com/sexy-lingerie-b-97.html?odr=new"
                GetProductAndProductIssue(menShoesUrl, 8, "Sexy Lingerie", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            Try
                '获取Women's Shoes的产品
                Dim menBagsUrl As String = "http://www.dresslily.com/women-s-shoes-c-81.html?odr=new"
                GetProductAndProductIssue(menBagsUrl, 8, "Women's Shoes", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            Try
                '获取Women's Dresses的产品
                Dim menWatchUrl As String = "http://www.dresslily.com/dresses-c-6.html?odr=new"
                GetProductAndProductIssue(menWatchUrl, 8, "Women's Dresses", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            Try
                '获取Bags的产品
                Dim womenClothingUrl As String = "http://www.dresslily.com/bags-b-3.html?odr=new"
                GetProductAndProductIssue(womenClothingUrl, 8, "Bags", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            Try
                '获取Men's Watches的产品
                Dim sexyUrl As String = "http://www.dresslily.com/men-s-watches-c-125.html?odr=new"
                GetProductAndProductIssue(sexyUrl, 8, "Men's Watches", siteId, issueId, nowTime)
            Catch ex As Exception
                '2013/07/19 added,throw the exception to outer layer
                Throw New Exception(ex.ToString())
                'Ignore
            End Try

            ' ''添加Issues表中的Subject内容
            AddIssueSubject(issueId) '2013/4/8修改

            ''获取WeeklyDeals产品,Removed
            'GetWeeklyDeals(siteId, issueId, nowTime)

            '获取联系人列表
            CreateContacts(siteId, planType, spreadLogin, appId, splitContactCount, issueId)

            helper.InsertInbox(issueId) '2013/4/15新增
        Catch ex As Exception
            '2013/07/19 added,throw the exception to outer layer
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 添加分类到Categories表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Public Sub GetCategory(ByVal siteId As Integer, ByVal issueId As Integer, ByVal nowTime As DateTime)
        Try
            Dim helper As New EFHelper()
            listCategory = helper.GetListCategory(siteId)
            Dim urlList As New List(Of String)
            For Each li In listCategory
                urlList.Add(li.Url)
            Next
            Dim insertCategorys As New List(Of Category)

            Dim cate1 As New Category()
            cate1.ParentID = -1
            cate1.Category1 = "Women's Clothing"
            cate1.SiteID = siteId
            cate1.LastUpdate = nowTime
            cate1.Description = "Women's Clothing"
            cate1.Url = "http://www.dresslily.com/women-s-clothing-b-1.html"
            cate1.Gender = "W"
            If Not (urlList.Contains(cate1.Url)) Then
                insertCategorys.Add(cate1)
                listCategory.Add(cate1)
            End If

            Dim cate2 As New Category()
            cate2.ParentID = -1
            cate2.Category1 = "Sexy Lingerie"
            cate2.SiteID = siteId
            cate2.LastUpdate = nowTime
            cate2.Description = "Sexy Lingerie"
            cate2.Url = "http://www.dresslily.com/sexy-lingerie-b-97.html"
            cate2.Gender = "W"
            If Not (urlList.Contains(cate2.Url)) Then
                insertCategorys.Add(cate2)
                listCategory.Add(cate2)
            End If

            Dim cate3 As New Category()
            cate3.ParentID = -1
            cate3.Category1 = "Women's Shoes"
            cate3.SiteID = siteId
            cate3.LastUpdate = nowTime
            cate3.Description = "Women's Shoes"
            cate3.Url = "http://www.dresslily.com/women-s-shoes-c-81.html"
            cate3.Gender = "W"
            If Not (urlList.Contains(cate3.Url)) Then
                insertCategorys.Add(cate3)
                listCategory.Add(cate3)
            End If

            Dim cate4 As New Category()
            cate4.ParentID = -1
            cate4.Category1 = "Women's Dresses"
            cate4.SiteID = siteId
            cate4.LastUpdate = nowTime
            cate4.Description = "Women's Dresses"
            cate4.Url = "http://www.dresslily.com/Dresses-c-6.html"
            cate4.Gender = "W"
            If Not (urlList.Contains(cate4.Url)) Then
                insertCategorys.Add(cate4)
                listCategory.Add(cate4)
            End If

            Dim cate5 As New Category()
            cate5.ParentID = -1
            cate5.Category1 = "Men's Clothing"
            cate5.SiteID = siteId
            cate5.LastUpdate = nowTime
            cate5.Description = "Men's Clothing"
            cate5.Url = "http://www.dresslily.com/men-s-clothing-b-2.html"
            cate5.Gender = "M"
            If Not (urlList.Contains(cate5.Url)) Then
                insertCategorys.Add(cate5)
                listCategory.Add(cate5)
            End If

            Dim cate6 As New Category()
            cate6.ParentID = -1
            cate6.Category1 = "Men's Shoes"
            cate6.SiteID = siteId
            cate6.LastUpdate = nowTime
            cate6.Description = "Men's Shoes"
            cate6.Url = "http://www.dresslily.com/men-s-shoes-c-82.html"
            cate6.Gender = "M"
            If Not (urlList.Contains(cate6.Url)) Then
                insertCategorys.Add(cate6)
                listCategory.Add(cate6)
            End If

            Dim cate7 As New Category()
            cate7.ParentID = -1
            cate7.Category1 = "Women's Bags"
            cate7.SiteID = siteId
            cate7.LastUpdate = nowTime
            cate7.Description = "Women's Bags"
            cate7.Url = "http://www.dresslily.com/bags-b-3.html"
            cate7.Gender = "W"
            If Not (urlList.Contains(cate7.Url)) Then
                insertCategorys.Add(cate7)
                listCategory.Add(cate7)
            End If

            helper.InsertListCategory(insertCategorys)
        Catch ex As Exception
            '2013/07/19 added,throw the exception to outer layer
            Throw New Exception(ex.ToString())
            'Ignore
        End Try
    End Sub


    ''' <summary>
    ''' 将产品插入到Products和Products_Issue表中，并写入数据到ProductCategory表中
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="maxCount"></param>
    ''' <remarks></remarks>
    Public Sub GetProductAndProductIssue(ByVal pageUrl As String, ByVal maxCount As Integer, _
                                         ByVal categoryName As String, ByVal siteId As Integer, _
                                         ByVal issueId As Integer, ByVal nowTime As DateTime)
        Dim helper As New EFHelper()
        Dim insertProducts As New List(Of Product)
        Dim updateProducts As New List(Of Product)
        Dim listProductId As New List(Of Integer)

        '特殊图片Url
        'Dim specialUrl As String = "css.dresslily.com/imagecache/DRSLILY/temp/skin2/ximages/loading.gif"
        'Dim listBigPicProduct As New List(Of Product) '2013/4/5

        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            For counter As Integer = 0 To maxCount - 1
                Dim liNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("ul/li")
                Dim liNode1 As HtmlNode = liNodes(0)
                Dim liNode2 As HtmlNode = liNodes(1)
                Dim liNode3 As HtmlNode = liNodes(2) '原价
                ' Dim liNode4 As HtmlNode = liNodes(3) '折扣价
                Dim product As New Product()
                product.Prodouct = liNode2.SelectSingleNode("h4/a").InnerText
                product.Url = ParseLink(liNode2.SelectSingleNode("h4/a").GetAttributeValue("href", "")).Trim()
                '2013/07/19 added,price and discount price layer changed
                'product.Discount = Double.Parse(liNode3.SelectNodes("span")(2).InnerText.Trim())
                'product.Price = Double.Parse(liNode3.SelectNodes("span")(0).SelectNodes("span")(1).InnerText.Trim())
                product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Price = Double.Parse(liNode3.SelectNodes("span")(1).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Currency = "USD" 'Dora 2013-11-25 不好获取，且经常改版故直接赋值

                '特殊图片处理
                product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                If (liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                    product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("data-src", "")
                End If

                product.LastUpdate = nowTime
                product.Description = liNode2.SelectSingleNode("h4/a").InnerText
                product.SiteID = siteId
                product.PictureAlt = liNode1.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                product.SizeHeight = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("width", ""))
                product.SizeWidth = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("height", ""))

                If (liNodes.Count = 4) Then
                    Dim myPicture1 As String = ""
                    Try
                        myPicture1 = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try

                    If (myPicture1.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture1
                    ElseIf (myPicture1.Contains("freepic.gif") OrElse myPicture1.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture1
                    End If

                ElseIf (liNodes.Count >= 5) Then
                    Dim myPicture1 As String = ""
                    Dim myPicture2 As String = ""
                    Try
                        myPicture1 = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try
                    Try
                        myPicture2 = liNodes(4).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try
                    If (myPicture1.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture1
                    ElseIf (myPicture1.Contains("freepic.gif") OrElse myPicture1.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture1
                    End If
                    If (myPicture2.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture2
                    ElseIf (myPicture2.Contains("freepic.gif") OrElse myPicture2.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture2
                    End If
                End If

                If Not (listProductUrl.Contains(product.Url)) Then
                    insertProducts.Add(product)
                    listProductUrl.Add(product.Url)
                Else
                    updateProducts.Add(product)
                End If
            Next
            If Not (insertProducts.Count = 0) Then
                listProductId.AddRange(helper.InsertListProduct(insertProducts, categoryName, siteId))
            End If
            If Not (updateProducts.Count = 0) Then
                listProductId.AddRange(helper.UpdateListProduct(updateProducts, categoryName, siteId))
            End If

            '2013/4/8修改
            'If (categoryName = "Women's Clothing") Then
            '    helper.InsertPOProductIssue(listProductId, siteId, issueId)
            '    'listBigPicProduct = helper.InsertPOProductIssue(listProductId, siteId, issueId) '2013/4/5
            'Else
            'helper.InsertProductIssue(listProductId, siteId, issueId, "CA")
            'End If
            helper.InsertProductIssue(listProductId, siteId, issueId, "CA")

            '2013/4/5不使用Banner条，不使用大图
            '更新Men's Clothing分类中的SectionId="CA"的产品的图片，将小图更新为大图
            'Dim efContext As New EmailAlerterEntities()
            'For Each liP In listBigPicProduct
            '    Dim doc As HtmlDocument = helper.GetHtmlDocument(liP.Url, 60000)
            '    Dim imgNode As HtmlNode = doc.DocumentNode.SelectNodes("//div[@id='myImagesSlideBox']/div")(1).SelectSingleNode("span/img")
            '    Dim imgUrl As String = imgNode.GetAttributeValue("src", "")
            '    Dim queryProduct = efContext.Products.Single(Function(p) p.ProdouctID = liP.ProdouctID)
            '    queryProduct.PictureUrl = imgUrl
            '    queryProduct.SizeWidth = Integer.Parse(imgNode.GetAttributeValue("width", ""))
            '    queryProduct.SizeHeight = Integer.Parse(imgNode.GetAttributeValue("height", ""))
            '    efContext.SaveChanges()
            'Next

        Catch ex As Exception
            '2013/07/19 added,throw the exception to outer layer
            Throw New Exception(ex.ToString())
            'Ignore
        End Try
    End Sub

    ''' <summary>
    ''' 添加Subject到Issues表中
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal issueId As Integer)
        Dim helper As New EFHelper()
        Dim subject As String
        subject = "Hi [FIRSTNAME],Special For Women ★" & mySubject & "Weekly Deals"
        helper.InsertIssueSubject(subject, issueId)
    End Sub

    ''' <summary>
    ''' 获取Weekly Deals的产品信息并添加到Products中（Women's Clothing，SectionID="DA"）
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub GetWeeklyDeals(ByVal siteId As Integer, ByVal issueId As Integer, ByVal nowTime As DateTime)
        Dim insertProducts As New List(Of Product)
        Dim updateProducts As New List(Of Product)
        Dim listProductId As New List(Of Integer)
        Dim helper As New EFHelper()
        Dim weeklyUrl As String = "http://www.dresslily.com/women-s-clothing-b-1.html"
        Dim doc As HtmlDocument = helper.GetHtmlDocument(weeklyUrl, 60000)
        Dim liNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='inlistw']/ul/li")
        For Each li In liNodes
            Dim weeklyProduct As New Product()
            Dim imgNode As HtmlNode = li.SelectNodes("div")(0).SelectSingleNode("a/img")
            Dim linkNode As HtmlNode = li.SelectNodes("div")(0).SelectSingleNode("a")
            Dim moneyNode As HtmlNode = li.SelectNodes("div")(2)
            weeklyProduct.Prodouct = li.SelectNodes("div")(1).SelectSingleNode("a").InnerText
            weeklyProduct.Url = ParseLink(linkNode.GetAttributeValue("href", ""))
            weeklyProduct.Price = Double.Parse(moneyNode.SelectNodes("span")(1).InnerText.Trim())

            '2013/4/8添加，防止有特殊图片
            weeklyProduct.PictureUrl = imgNode.GetAttributeValue("src", "")
            If (imgNode.GetAttributeValue("src", "").Contains("loading.gif")) Then
                weeklyProduct.PictureUrl = imgNode.GetAttributeValue("data-src", "")
            End If

            weeklyProduct.LastUpdate = nowTime
            weeklyProduct.Description = li.SelectNodes("div")(0).SelectSingleNode("span").InnerText & "%OFF"
            weeklyProduct.SiteID = siteId
            weeklyProduct.Currency = moneyNode.SelectNodes("span")(0).InnerText.Trim()
            weeklyProduct.PictureAlt = imgNode.GetAttributeValue("alt", "")
            weeklyProduct.SizeWidth = Integer.Parse(imgNode.GetAttributeValue("width", "").Replace("px", ""))
            weeklyProduct.SizeHeight = Integer.Parse(imgNode.GetAttributeValue("height", "").Replace("px", ""))
            If Not (listProductUrl.Contains(weeklyProduct.Url)) Then
                insertProducts.Add(weeklyProduct)
            Else
                updateProducts.Add(weeklyProduct)
            End If
        Next
        If Not (insertProducts.Count = 0) Then
            listProductId.AddRange(helper.InsertListProduct(insertProducts, "Women's Clothing", siteId))
        End If
        If Not (updateProducts.Count = 0) Then
            listProductId.AddRange(helper.UpdateListProduct(updateProducts, "Women's Clothing", siteId))
        End If

        'helper.InsertProductIssue(listProductId, siteId, issueId, "DA") '2013/4/8修改
        helper.InsertPOProductIssue(listProductId, siteId, issueId, "DA", "Hi [FIRSTNAME],Special For Women ★ ")
    End Sub

    ''' <summary>
    ''' 转化链接，使之成为完整的链接
    ''' </summary>
    ''' <param name="strUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ParseLink(ByVal strUrl As String) As String
        Dim homeUrl As String = "http://www.dresslily.com"
        If (Not strUrl.Contains("http")) Then
            Return homeUrl + strUrl
        Else
            Return strUrl
        End If
    End Function

    ''' <summary>
    ''' 添加数据到SplitContactLists表和ContactLists_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="spreaLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="splitContactCount"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub CreateContacts(ByVal siteId As Integer, ByVal planType As String, ByVal spreaLogin As String, _
                              ByVal appId As String, ByVal splitContactCount As Integer, _
                              ByVal issueId As Integer)
        Dim efContext As New EmailAlerterEntities()
        Dim listCategoryWomen As List(Of Category) = (From c In efContext.Categories
                                                    Where (c.SiteID = siteId AndAlso c.Gender = "W")
                                                    Select c).ToList()
        Dim listContact As List(Of String) = (From c In efContext.ContactLists_Issue Select c.ContactList).ToList()
        For Each li In listCategoryWomen
            Dim querySubscriber As New QuerySubscriber()
            '2013/4/5
            querySubscriber.Strategy = ChooseStrategy.Favorite
            querySubscriber.Favorite = li.CategoryID.ToString()

            'Add By Gary Start 2013/06/09
            '调整为根据上一周的点击记录发送分类邮件
            querySubscriber.StartDate = Date.Now.AddDays(-Date.Now.DayOfWeek - 7).ToString("yyyy-MM-dd")
            querySubscriber.EndDate = Date.Now.AddDays(-Date.Now.DayOfWeek).ToString("yyyy-MM-dd")
            'Add By Gary End

            querySubscriber.CountryList = New String() {}
            Dim criteria As String = querySubscriber.ToJsonString()
            Dim topN As Integer = Integer.MaxValue

            Dim cateName As String = li.Category1.Trim()
            If (cateName.Contains("'") OrElse cateName.Contains(" ")) Then
                cateName = cateName.Replace("'", "-")
                cateName = cateName.Replace(" ", "-")
            End If
            '实际环境：Dim saveAsList As String = "ActiveWomen" & Now.ToString("yyyyMMdd") & cateName
            Dim saveAsList As String = "ActiveWomen" & Now.ToString("yyyyMMdd") & cateName

            Dim forceCreate As Boolean = True
            Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()

            'Dim contactList As New ContactLists_Issue()
            'contactList.IssueID = issueId
            'contactList.ContactList = saveAsList
            Dim returnResult As Integer = 0
            'If Not (listContact.Contains(saveAsList)) Then
            'efContext.AddToContactLists_Issue(contactList)
            Try
                mySpread.Timeout = 600000
                returnResult = mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
            Catch ex As Exception
                Try
                    mySpread.Timeout = 600000
                    returnResult = mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
                Catch ex1 As Exception
                    LogText(ex1.ToString())
                End Try
            End Try
            If Not (returnResult = 0) Then
                Dim contactList As New ContactLists_Issue()
                contactList.IssueID = issueId
                contactList.ContactList = saveAsList
                contactList.SendingStatus = "draft" '2013/12/31 added，Dresslily Men创建草稿状态
                'If Not (listContact.Contains(saveAsList)) Then
                efContext.AddToContactLists_Issue(contactList)
            End If
            'End If
        Next
        Dim contactList1 As New ContactLists_Issue()
        contactList1.IssueID = issueId
        contactList1.ContactList = "Women"
        efContext.AddToContactLists_Issue(contactList1)
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 写日志
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <remarks></remarks>
    Sub Log(ByVal Ex As Exception)
        Try
            LogText(Now & ", " & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
