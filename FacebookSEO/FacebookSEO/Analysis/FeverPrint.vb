Imports HtmlAgilityPack

Public Class FeverPrint

    Private Shared pageUrl As String = "http://fever-print.com/"
    Private Shared helper As New EFHelper()
    Private Shared doc As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
    Private Shared listCategory As New List(Of Category)
    Private Shared listCategoryUrl As New HashSet(Of String)
    Private Shared listProductUrl As New List(Of String)
    Private Shared iTakeCounter As Integer = 4 '模板中产品的个数
    Private Shared strListCategory As New HashSet(Of String) '2013/4/27添加
    Private Shared strListCategoryUrl As New HashSet(Of String) '2013/4/27添加

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        Dim nowTime As DateTime = Now
        '2013/4/27添加，增加分类名判断
        Dim listCategory As List(Of Category) = helper.GetListCategory(siteId)
        For Each li In listCategory
            strListCategory.Add(li.Category1)
            strListCategoryUrl.Add(li.Url)
        Next
        listProductUrl = helper.GetListProduct(siteId)

        GetCategory(siteId, IssueID, planType, nowTime)

        'GetAds(siteId, IssueID, planType, nowTime) '2013/07/05,新模板不使用Ads

        GetBestArrival(siteId, IssueID, planType, nowTime) '2013/07/05，新模板不适用Best Arrival

        'GetPromotionProduct(siteId, IssueID, nowTime) '2013/07/05 added，使用Promotion产品,2013/07/07网站上无此产品

        GetNewArrival(siteId, IssueID, planType, nowTime)

        GetCardProduct(siteId, IssueID, planType, nowTime)

        GetHotArrival(siteId, IssueID, planType, nowTime)

        GetSubject(siteId, IssueID, nowTime)

        'helper.InsertContactList(IssueID, "Reasonablers", "")
        'helper.InsertContactList(IssueID, "Reasonable Test", "")

        Dim contactListArr() As String = {"New", "DJC", "議員", "FeverPrint客戶", "e-mail list", "email", "JCI", "Facebook for Jeffrey", "SecondarySchool from Jeffrey", _
                                          "小學", "中學", "Hotel", "LEGCO", "扶輪社", "扶輪社(new)", "Reasonablers", "Reasonable Test"}
        helper.InsertContactList(IssueID, contactListArr, "draft")
        'helper.InsertInbox(IssueID)


    End Sub

    ''' <summary>
    ''' 获取所有的大分类
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetCategory(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)
        listCategory = helper.GetListCategory(siteId)
        Dim insertCategorys As New List(Of Category)
        Dim updateCategorys As New List(Of Category)
        For Each li In listCategory
            listCategoryUrl.Add(li.Url)
        Next
        Dim dtNodes As HtmlNodeCollection = doc.GetElementbyId("category_tree").SelectNodes("//div[@class='content']/dl/dt")
        For Each node In dtNodes
            Dim cateUrl As String = ParseLink(node.SelectSingleNode("a").GetAttributeValue("href", ""))
            Dim category As New Category()
            category.ParentID = -1
            category.Category1 = node.SelectSingleNode("a").InnerText.Trim()
            category.SiteID = siteId
            category.LastUpdate = nowTime
            category.Description = node.SelectSingleNode("a").InnerText
            category.Url = cateUrl
            If Not (listCategoryUrl.Contains(cateUrl)) Then
                insertCategorys.Add(category)
                strListCategory.Add(category.Category1) '2013/4/27添加
                strListCategoryUrl.Add(category.Url) '2013/4/27添加
            Else
                updateCategorys.Add(category)
            End If
        Next
        If Not (insertCategorys.Count = 0) Then
            helper.InsertListCategory(insertCategorys)
        End If
        If Not (updateCategorys.Count = 0) Then
            helper.UpdateListCategory(updateCategorys)
        End If
    End Sub

    ''' <summary>
    ''' 获取页面Banner条的信息，并填充到Ads和Ads_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetAds(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)
        Dim myHelper As New EFHelper()
        Dim listAds As List(Of Ad) = myHelper.GetListAd(siteId)
        Dim insertAds As New List(Of Ad)
        Dim updateAds As New List(Of Ad)
        Dim listAdsUrl As New HashSet(Of String)
        Dim listAdsId As New List(Of Integer)
        For Each li In listAds
            listAdsUrl.Add(li.Url)
        Next
        Dim jsUrl As String = "http://fever-print.com/data/flashdata/redfocus/data.js"
        Dim helper As New Analysis.EFHelper()
        Dim hd As HtmlDocument = helper.GetHtmlDocument(jsUrl, 60000)
        Dim jsStr As String = hd.DocumentNode.InnerHtml
        Dim matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(jsStr, "\u0022\w.*\u0022")
        Dim bannerCounter As Integer = matches.Count / 3
        For i As Integer = 0 To matches.Count - 1 Step 3
            Dim myBanner As New Ad()
            myBanner.Ad1 = matches.Item(i + 1).Value
            If (matches.Item(i + 1).Value.Contains(Chr(34))) Then
                myBanner.Ad1 = matches.Item(i + 1).Value.Replace(Chr(34), "").Trim()
            End If
            myBanner.PictureUrl = ParseLink(matches.Item(i).Value)
            If (matches.Item(i).Value.Contains(Chr(34))) Then
                myBanner.PictureUrl = ParseLink(matches.Item(i).Value.Replace(Chr(34), "").Trim())
            End If
            myBanner.Description = matches.Item(i + 1).Value
            If (matches.Item(i + 1).Value.Contains(Chr(34))) Then
                myBanner.Description = matches.Item(i + 1).Value.Replace(Chr(34), "").Trim()
            End If
            myBanner.Url = ParseLink(matches.Item(i + 2).Value.Trim())
            If (matches.Item(i + 2).Value.Contains(Chr(34))) Then
                myBanner.Url = ParseLink(matches.Item(i + 2).Value.Replace(Chr(34), "").Trim())
            End If
            myBanner.SiteID = siteId
            myBanner.Lastupdate = nowTime
            myBanner.Type = "P"
            If Not (listAdsUrl.Contains(myBanner.Url)) Then
                insertAds.Add(myBanner)
            Else
                updateAds.Add(myBanner)
            End If
        Next i
        If (insertAds.Count > 0) Then
            listAdsId.AddRange(myHelper.InsertAds(insertAds, siteId))
        End If
        If (updateAds.Count > 0) Then
            listAdsId.AddRange(myHelper.UpdateAds(updateAds, siteId))
        End If
        myHelper.InsertAdsIssue(listAdsId, siteId, issueId, 2, bannerCounter)
    End Sub

    ''' <summary>
    ''' 获取页面Best Arrival产品信息，并填充Products和Products_Issue表
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetBestArrival(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)

        Dim nodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='demo']/div[@id='demo1']/ul/li")
        Dim listProductId As New List(Of Integer)
        Dim itotalProductCount As Integer = nodes.Count '2013/4/26新增，每次轮番出产品
        For Each n In nodes
            Dim myProduct As New Product()
            Dim hrefNode As HtmlNode = n.SelectSingleNode("a[1]")
            Dim picNode As HtmlNode = n.SelectSingleNode("a[1]/img")
            Dim moneyNode As HtmlNode = n.SelectSingleNode("div[@class='txt']/div[@class='content']")
            myProduct.Prodouct = moneyNode.SelectSingleNode("a").InnerText
            myProduct.Url = ParseLink(hrefNode.GetAttributeValue("href", "").Trim())
            myProduct.Price = Convert.ToDecimal(moneyNode.LastChild.InnerHtml.Trim().Replace("HKD", "").Trim())
            myProduct.PictureUrl = ParseLink(picNode.GetAttributeValue("src", ""))
            myProduct.LastUpdate = nowTime
            myProduct.Description = picNode.GetAttributeValue("alt", "")
            myProduct.SiteID = siteId
            myProduct.Currency = "HKD"
            myProduct.PictureAlt = picNode.GetAttributeValue("alt", "")
            Dim hd As HtmlDocument = helper.GetHtmlDocument(myProduct.Url, 60000)
            Dim strCategory As String = hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).InnerText.Trim()
            Dim strCategoryUrl As String = ParseLink(hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).GetAttributeValue("href", ""))

            CheckCategory(strCategory, strCategoryUrl, siteId, nowTime)
            
            If (listProductUrl.Contains(myProduct.Url)) Then
                listProductId.Add(helper.UpdateSingleProduct(myProduct, strCategory, siteId, nowTime))
            Else
                listProductId.Add(helper.InsertSingleProduct(myProduct, strCategory, siteId))
                listProductUrl.Add(myProduct.Url)
            End If

            'If (listProductUrl.Contains(myProduct.Url)) Then
            '    listProductId.Add(helper.UpdateSingleProduct(myProduct, strCategory, siteId, nowTime))
            'Else
            '    listProductId.Add(helper.InsertSingleProduct(myProduct, strCategory, siteId))
            '    listProductUrl.Add(myProduct.Url)
            'End If
        Next
        helper.InsertProductIssue(listProductId, siteId, issueId, "BR", 2, itotalProductCount)
    End Sub

    ''' <summary>
    ''' 获取页面New Arrival产品信息，并填充Products和Products_Issue表
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetNewArrival(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='show_new_area']/div")
        Dim listProductId As New List(Of Integer)
        Dim iTotalProductCount As Integer = divNodes.Count '2013/4/26新增，每次轮番出产品
        For i As Integer = 0 To iTotalProductCount - 1
            Dim hrefNode As HtmlNode = divNodes(i).SelectSingleNode("p/a")
            Dim picNode As HtmlNode = divNodes(i).SelectSingleNode("a/img")
            Dim moneyNode As HtmlNode = divNodes(i).SelectSingleNode("font")
            Dim myProduct As New Product()
            myProduct.Prodouct = hrefNode.InnerText.Trim()
            myProduct.Url = ParseLink(hrefNode.GetAttributeValue("href", "").Trim())
            myProduct.Price = Decimal.Parse(moneyNode.InnerText.Replace("HKD", "").Trim())
            myProduct.PictureUrl = ParseLink(picNode.GetAttributeValue("src", ""))
            myProduct.LastUpdate = nowTime
            myProduct.Description = hrefNode.InnerText.Trim()
            myProduct.SiteID = siteId
            myProduct.Currency = "HKD"
            myProduct.PictureAlt = picNode.GetAttributeValue("alt", "")
            Dim hd As HtmlDocument = helper.GetHtmlDocument(myProduct.Url, 60000)
            Dim strCategory As String = hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).InnerText.Trim()
            Dim strCategoryUrl As String = ParseLink(hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).GetAttributeValue("href", ""))

            CheckCategory(strCategory, strCategoryUrl, siteId, nowTime) '2013/4/27新增

            If (listProductUrl.Contains(myProduct.Url)) Then
                listProductId.Add(helper.UpdateSingleProduct(myProduct, strCategory, siteId, nowTime))
            Else
                listProductId.Add(helper.InsertSingleProduct(myProduct, strCategory, siteId))
                listProductUrl.Add(myProduct.Url)
            End If
        Next
        helper.InsertProductIssue(listProductId, siteId, issueId, "NE", iTakeCounter, iTotalProductCount)
    End Sub

    ''' <summary>
    ''' 获取页面"卡片模版"产品信息，并填充Products和Products_Issue表
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetCardProduct(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='areaRight']/div")(5).SelectNodes("div[@class='content']/div[@class='contentR']/div[@class='General_content clearfix']/div")
        Dim listProductId As New List(Of Integer)
        Dim updateProducts As New List(Of Product)
        Dim insertProducts As New List(Of Product)
        Dim iTotalProductCount As Integer = divNodes.Count '2013/4/26新增，每次轮番出产品
        For Each div In divNodes
            Dim hrefNode As HtmlNode = div.SelectSingleNode("p/a")
            Dim picNode As HtmlNode = div.SelectSingleNode("a/img")
            Dim moneyNode As HtmlNode = div.SelectSingleNode("font")
            Dim myProduct As New Product()
            myProduct.Prodouct = hrefNode.InnerText.Trim()
            myProduct.Url = ParseLink(hrefNode.GetAttributeValue("href", "").Trim())
            myProduct.Price = Decimal.Parse(moneyNode.InnerText.Replace("HKD", "").Trim())
            myProduct.PictureUrl = ParseLink(picNode.GetAttributeValue("src", ""))
            myProduct.LastUpdate = nowTime
            myProduct.Description = hrefNode.InnerText.Trim()
            myProduct.SiteID = siteId
            myProduct.Currency = "HKD"
            myProduct.PictureAlt = picNode.GetAttributeValue("alt", "")
            If (listProductUrl.Contains(myProduct.Url)) Then
                updateProducts.Add(myProduct)
            Else
                insertProducts.Add(myProduct)
                listProductUrl.Add(myProduct.Url)
            End If
        Next

        '2013/4/27新增
        Dim strCategory As String = "平面設計"
        Dim strCategoryUrl As String = "http://www.fever-print.com/category.php?id=12"
        CheckCategory(strCategory, strCategoryUrl, siteId, nowTime)

        If (updateProducts.Count > 0) Then
            listProductId.AddRange(helper.UpdateListProduct(updateProducts, "平面設計", siteId))
        End If
        If (insertProducts.Count > 0) Then
            listProductId.AddRange(helper.InsertListProduct(insertProducts, "平面設計", siteId))
        End If
        helper.InsertProductIssue(listProductId, siteId, issueId, "CA", iTakeCounter, iTotalProductCount)
    End Sub

    ''' <summary>
    ''' 获取页面Hot Arrival信息，并填充到Products和Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetHotArrival(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                           ByVal nowTime As DateTime)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='show_hot_area']/div")
        Dim listProductId As New List(Of Integer)
        Dim iTotalProductCount As Integer = divNodes.Count
        For i As Integer = 0 To iTotalProductCount - 1
            Dim hrefNode As HtmlNode = divNodes(i).SelectSingleNode("p/a")
            Dim picNode As HtmlNode = divNodes(i).SelectSingleNode("a/img")
            Dim moneyNode As HtmlNode = divNodes(i).SelectSingleNode("font")
            Dim myProduct As New Product()
            myProduct.Prodouct = hrefNode.InnerText.Trim()
            myProduct.Url = ParseLink(hrefNode.GetAttributeValue("href", "").Trim())
            myProduct.Price = Decimal.Parse(moneyNode.InnerText.Replace("HKD", "").Trim())
            myProduct.PictureUrl = ParseLink(picNode.GetAttributeValue("src", ""))
            myProduct.LastUpdate = nowTime
            myProduct.Description = hrefNode.InnerText.Trim()
            myProduct.SiteID = siteId
            myProduct.Currency = "HKD"
            myProduct.PictureAlt = picNode.GetAttributeValue("alt", "")
            Dim hd As HtmlDocument = helper.GetHtmlDocument(myProduct.Url, 60000)
            Dim strCategory As String = hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).InnerText.Trim()

            '2013/4/27新增
            Dim strCategoryUrl As String = ParseLink(hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).GetAttributeValue("href", ""))
            CheckCategory(strCategory, strCategoryUrl, siteId, nowTime)

            If (listProductUrl.Contains(myProduct.Url)) Then
                listProductId.Add(helper.UpdateSingleProduct(myProduct, strCategory, siteId, nowTime))
            Else
                listProductId.Add(helper.InsertSingleProduct(myProduct, strCategory, siteId))
                listProductUrl.Add(myProduct.Url)
            End If
        Next
        helper.InsertProductIssue(listProductId, siteId, issueId, "BE", iTakeCounter, iTotalProductCount)
    End Sub

    ''' <summary>
    ''' 根据Ads和Ads_Issue表，Type="P"获取Ad并拼接写入到Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetSubject(ByVal siteId As Integer, ByVal issueId As Integer, ByVal nowTime As DateTime)
        'Dim listAd As List(Of Ad) = helper.GetListAd(siteId, issueId)
        'Dim subject As String = "Hi [FIRSTNAME],"
        'Dim counter As Integer = 1
        'For Each li In listAd
        '    If (counter < listAd.Count) Then
        '        subject = subject & li.Ad1 & "★"
        '    Else
        '        subject = subject & li.Ad1 & "★" & "Weekly Deals"
        '    End If
        '    counter = counter + 1
        'Next
        'helper.InsertIssueSubject(issueId, subject)
        Dim listProductName As List(Of String) = helper.SearchPromotionProduct(issueId, "BR")
        Dim subject As String = "Hi [FIRSTNAME],非凡印刷本週促銷商品★"
        For Each name In listProductName
            subject = subject & name & "★"
        Next
        subject = subject.Substring(0, subject.LastIndexOf("★"))
        helper.InsertIssueSubject(issueId, subject)
    End Sub

    ''' <summary>
    ''' 转换链接，使链接成为完整链接
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ParseLink(ByVal str As String) As String
        If (Not str.Contains("http")) Then
            Return pageUrl + str
        Else
            Return str
        End If
    End Function

    ''' <summary>
    ''' 判断该分类是否存在Categories表中，如果不存在，则添加一个分类
    ''' </summary>
    ''' <param name="strCategory"></param>
    ''' <param name="strCategoryUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub CheckCategory(ByVal strCategory As String, ByVal strCategoryUrl As String, _
                                          ByVal siteId As Integer, ByVal nowTime As DateTime)
        Dim counter As Integer = 0
        For Each strs In strListCategory
            If (strs = strCategory) Then
                Exit For
            End If
            counter = counter + 1
        Next
        If (counter = strListCategory.Count) Then
            Dim cate As New Category()
            cate.Category1 = strCategory
            cate.ParentID = -1
            cate.SiteID = siteId
            cate.LastUpdate = nowTime
            cate.Description = cate.Category1
            cate.Url = strCategoryUrl
            If (strListCategoryUrl.Contains(strCategoryUrl)) Then
                helper.UpdateCategory(cate)
            Else
                helper.InsertCategory(cate)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 获取“促銷商品”信息，并添加到Products和Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetPromotionProduct(ByVal siteId As Integer, ByVal issueId As Integer, ByVal nowTime As DateTime)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='areaRight']/div")(1).SelectNodes("div[@class='content']/div[@class='contentR']/div[@class='General_content clearfix']/div")
        Dim listProductId As New List(Of Integer)
        Dim iTotalProductCount As Integer = divNodes.Count
        For i As Integer = 0 To iTotalProductCount - 1
            Dim hrefNode As HtmlNode = divNodes(i).SelectSingleNode("p/a")
            Dim picNode As HtmlNode = divNodes(i).SelectSingleNode("a/img")
            Dim moneyNode As HtmlNode = divNodes(i).SelectSingleNode("font")
            Dim myProduct As New Product()
            myProduct.Prodouct = hrefNode.InnerText.Trim()
            myProduct.Url = ParseLink(hrefNode.GetAttributeValue("href", "").Trim())
            myProduct.Price = Decimal.Parse(moneyNode.InnerText.Replace("HKD", "").Trim())
            myProduct.PictureUrl = ParseLink(picNode.GetAttributeValue("src", ""))
            myProduct.LastUpdate = nowTime
            myProduct.Description = hrefNode.InnerText.Trim()
            myProduct.SiteID = siteId
            myProduct.Currency = "HKD"
            myProduct.PictureAlt = picNode.GetAttributeValue("alt", "")
            Dim hd As HtmlDocument = helper.GetHtmlDocument(myProduct.Url, 60000)
            Dim strCategory As String = hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).InnerText.Trim()

            '2013/4/27新增
            Dim strCategoryUrl As String = ParseLink(hd.DocumentNode.SelectNodes("//div[@class='block ur_here']/a")(1).GetAttributeValue("href", ""))
            CheckCategory(strCategory, strCategoryUrl, siteId, nowTime)

            If (listProductUrl.Contains(myProduct.Url)) Then
                listProductId.Add(helper.UpdateSingleProduct(myProduct, strCategory, siteId, nowTime))
            Else
                listProductId.Add(helper.InsertSingleProduct(myProduct, strCategory, siteId))
                listProductUrl.Add(myProduct.Url)
            End If
        Next
        helper.InsertProductIssue(listProductId, siteId, issueId, "PO")
    End Sub
End Class
