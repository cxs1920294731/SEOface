Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq

Public Class EverBuyingDeals
    Private Shared efContext As New EmailAlerterEntities()
    Private Shared listProductUrl As New List(Of String)
    Private Shared listProductName As New List(Of String)
    Private Shared listProductSku As New List(Of String)

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal spreadLogin As String, _
                     ByVal appId As String, ByVal updateTime As DateTime)
        Dim url As String = "http://www.everbuying.net/"
        Dim greatDealUrl = "http://www.everbuying.net/m-index-a-view_deals.html"
        Dim helper As New EFHelper
        Dim doc As HtmlDocument = helper.GetHtmlDocument(url, 120000)
        GetBanner(siteId, updateTime, IssueID, doc)
        'GetGreatDeal(siteId, IssueID, updateTime, doc)
        GetGreatDealsFromJson(siteId, IssueID, greatDealUrl)
        GetCatesProducts(siteId, IssueID, updateTime, planType)
        InsertSubject(IssueID, siteId, planType, "Deals")

        '2014/04/02 added
        'helper.InsertContactList(IssueID, "opens 20140214", "draft")
        'helper.InsertContactList(IssueID, "E_SPREAD_001", "draft")
        helper.InsertContactList(IssueID, "Opens20140402-1", "waiting")
    End Sub

    ''' <summary>
    ''' 获取Banner条信息，并更新Ads_Issue表的记录
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="updateTime"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetBanner(ByVal siteId As Integer, ByVal updateTime As DateTime, ByVal issueId As Integer, ByVal doc As HtmlDocument)
        Dim helper As New EFHelper
        'Dim url As String = "http://www.everbuying.net/"
        'Dim doc As HtmlDocument = helper.GetHtmlDocument(url, 60000)
        Dim listBannerUrl As New List(Of String)
        listBannerUrl = EFHelper.GetTopNAdsUrl(siteId, 4)
        Dim bannerNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='n_bannerList']/ul/li")
        For Each li In bannerNodes
            Dim bannerUrl As String = li.SelectSingleNode("a").GetAttributeValue("href", "").Trim()
            bannerUrl = AddDomain("http://www.everbuying.net", bannerUrl)
            If (listBannerUrl.Contains(bannerUrl)) Then
                Continue For
            Else
                Dim imgUrl As String = li.SelectSingleNode("a/img").GetAttributeValue("src", "")
                Dim ad As New Ad
                ad.Url = bannerUrl
                ad.PictureUrl = imgUrl
                ad.SiteID = siteId
                ad.Lastupdate = updateTime
                Dim adId As Integer = helper.InsertOrUpdateAd(ad)
                helper.InsertSingleAdsIssue(adId, siteId, issueId)
            End If
        Next
    End Sub

    ''' <summary>
    ''' 获取页面块“Great Deals”作为邮件模板的“Weekly Deals部分”，SectionID="WE"
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="updateTime"></param>
    ''' <param name="doc"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetGreatDeal(ByVal siteId As Integer, ByVal issueId As Integer, ByVal updateTime As DateTime, ByVal doc As HtmlDocument)
        'Dim listLastProducts As List(Of Product)=
        Dim listProduct As New List(Of Product)
        listProduct = GetProductList(siteId)
        Dim categoryId As Integer = EFHelper.GetCategoryId(siteId, "Cell Phones")
        Dim helper As New EFHelper
        Dim listProductId As New List(Of Integer)
        Dim productNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='n_indexProList mb15'][1]/ul/li")

        Dim sku As String
        For Each node In productNodes
            Dim product As New Product
            product.Prodouct = node.SelectSingleNode("div/p[1]").InnerText.Trim()
            product.Description = node.SelectSingleNode("div/p[1]").InnerText
            product.Url = AddDomain("http://www.everbuying.net", node.SelectSingleNode("div/p[1]/a").GetAttributeValue("href", "").Trim())
            If (listProductUrl.Contains(product.Url)) Then
                Continue For
            Else
                listProductUrl.Add(product.Url)
            End If
            sku = getSKU(product.Url)
            If (listProductSku.Contains(sku)) Then
                Continue For
            Else
                listProductSku.Add(sku)
            End If
            Dim productPrice As String = node.SelectSingleNode("div/p[2]").InnerText.Trim()
            Dim arrPrice As String() = System.Text.RegularExpressions.Regex.Split(productPrice, "USD", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            For Each arr In arrPrice
                If Not (String.IsNullOrEmpty(arr)) Then
                    'Dim price As String = arr
                    If (product.Price Is Nothing) Then
                        product.Price = Double.Parse(arr.Trim())
                    Else
                        product.Discount = Double.Parse(arr.Trim())
                    End If
                End If
            Next
            Dim pictureUrl As String = node.SelectSingleNode("div/a/img").GetAttributeValue("data-src", "")
            If (String.IsNullOrEmpty(pictureUrl)) Then
                pictureUrl = node.SelectSingleNode("div/a/img").GetAttributeValue("src", "")
            End If
            product.PictureUrl = pictureUrl  'node.SelectSingleNode("div/a/img").GetAttributeValue("data-src", "")
            product.LastUpdate = updateTime
            product.SiteID = siteId
            product.Currency = "$"
            product.SizeHeight = 150
            product.SizeWidth = 150

            Dim productId As Integer = helper.InsertProduct(product, DateTime.Now, categoryId, listProduct)
            listProductId.Add(productId)
        Next
        helper.InsertProductIssue(listProductId, siteId, issueId, "WE", 2, productNodes.Count)
    End Sub


    Public Shared Function GetGreatDealsFromJson(ByVal siteID As Integer, ByVal Issueid As Integer, ByVal requestUrl As String) As List(Of Product)
        Dim listProduct As New List(Of Product)
        listProduct = GetProductList(siteID)
        Dim categoryId As Integer = EFHelper.GetCategoryId(siteID, "Cell Phones")
        Dim helper As New EFHelper
        Dim listProductId As New List(Of Integer)
        ' Dim productNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='n_indexProList mb15'][1]/ul/li")
        Dim postsStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")
        Dim postsJson As JObject = JObject.Parse(postsStr)
        Dim postsJArr As JArray = postsJson("recommend_items")
        Dim pdlist As List(Of Product) = New List(Of Product)
        Dim prductCount As Integer = 0
        For i As Integer = 0 To postsJArr.Count - 1
            Dim product As New Product
            Dim productUrl As String = ""
            Dim item As JObject = postsJArr(i)
            item.TryGetValue("goods_title", product.Prodouct)
            item.TryGetValue("short_name", product.Description)
            item.TryGetValue("url_title", productUrl)
            product.Url = "http://www.everbuying.net" & productUrl
            If (listProductUrl.Contains(product.Url)) Then
                Continue For
            Else
                listProductUrl.Add(product.Url)
            End If
            Dim priceStr As String = ""
            item.TryGetValue("market_price", priceStr)
            product.Price = Double.Parse(priceStr)
            item.TryGetValue("shop_price", priceStr)
            product.Discount = Double.Parse(priceStr)
            item.TryGetValue("goods_img", product.PictureUrl)
            product.LastUpdate = Now
            product.SiteID = siteID
            product.Currency = "$"
            product.SizeHeight = 150
            product.SizeWidth = 150
            Dim productId As Integer = helper.InsertProduct(product, DateTime.Now, categoryId, listProduct)
            listProductId.Add(productId)
            prductCount += 1
        Next
        helper.InsertProductIssue(listProductId, siteID, Issueid, "WE", 2, prductCount)
    End Function

    ''' <summary>
    ''' 获取主推分类的产品，产品页面按"hot"进行排序，如：http://www.everbuying.net/Wholesale-Cell-Phones-b-22.html?odr=hot
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="updateTime"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetCatesProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Dim helper As New EFHelper
        Dim cate1 As String = "http://www.everbuying.net/Wholesale-Cell-Phones-b-22.html?display=g&page_size=48&odr=hot"
        Dim cate2 As String = "http://www.everbuying.net/Wholesale-iPhone-iPad-iPod-b-493.html?display=g&page_size=48&odr=hot"
        Dim cate3 As String = "http://www.everbuying.net/Wholesale-Cell-Phone-Accessories-b-667.html?display=g&odr=hot&page_size=48"
        Dim cate4 As String = "http://www.everbuying.net/Wholesale-Notebook-UMPC-MID-b-815.html?page_size=48&odr=hot&display=g"
        Dim cate5 As String = "http://www.everbuying.net/smlclass1674.html?page_size=48&odr=hot&display=g"
        Dim cate6 As String = "http://www.everbuying.net/Wholesale-Computers-Networking-b-926.html?page_size=48&odr=hot&display=g"
        Dim cate7 As String = "http://www.everbuying.net/Wholesale-LED-Lights-b-669.html?page_size=48&odr=hot&display=g"
        Dim cate8 As String = "http://www.everbuying.net/Wholesale-Consumer-Electronics-b-853.html?page_size=48&odr=hot&display=g"
        Dim arrCategorys As String() = New String() {cate1, cate2, cate3, cate4, cate5, cate6, cate7, cate8}
        Dim arrCategoryNames As String() = New String() {"Cell Phones", "Apple Accessories", "Cell Phone Accessories", "Tablet PC", "Mobile Power Bank", "Computers & Networking",
                                                         "LED Lights", "Consumer Electronics"}
        Dim sku As String
        For i = 0 To 7
            Dim categoryId As Integer = EFHelper.GetCategoryId(siteId, arrCategoryNames(i))
            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)

            Dim cate As String = arrCategorys(i)
            Dim cateDoc As HtmlDocument = helper.GetHtmlDocument(cate, 60000)
            Dim productNodes As HtmlNodeCollection = cateDoc.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            Dim counter As Integer = 0
            For Each productNode In productNodes
                Try
                    If (counter >= 3) Then
                        Exit For
                    Else
                        Dim product As New Product
                        Dim productName As String = productNode.SelectSingleNode("ul/li[2]").InnerText.Trim()
                        product.Prodouct = productName
                        product.Url = AddDomain("http://www.everbuying.net", productNode.SelectSingleNode("ul/li[2]/a").GetAttributeValue("href", "").Trim())
                        If (listProductUrl.Contains(product.Url) OrElse listProductName.Contains(product.Prodouct)) Then '该产品在本期的邮件中了
                            Continue For
                        Else
                            listProductUrl.Add(product.Url)
                            listProductName.Add(product.Prodouct)
                        End If
                        sku = getSKU(product.Url)
                        If (listProductSku.Contains(sku)) Then
                            Continue For
                        Else
                            listProductSku.Add(sku)
                        End If
                        If Not (helper.IsProductSent(siteId, product.Url, Now.AddDays(-23).ToString(), Now.ToString(), planType)) Then  '该产品在前几期邮件中没有发送
                            Dim productPrice As String = productNode.SelectSingleNode("ul/li[3]").InnerText.Trim()
                            Dim arrPrice As String() = System.Text.RegularExpressions.Regex.Split(productPrice, "USD", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            If (arrPrice.Count >= 3) Then
                                For Each arr In arrPrice
                                    If Not (String.IsNullOrEmpty(arr)) Then
                                        'Dim price As String = arr
                                        If (product.Price Is Nothing) Then
                                            product.Price = Double.Parse(arr.Trim())
                                        Else
                                            product.Discount = Double.Parse(arr.Trim())
                                        End If
                                    End If
                                Next
                            Else
                                productPrice = productPrice.Replace("USD", "").Trim()  'productPrice.Trim("USD").Trim()
                                product.Discount = Double.Parse(productPrice)
                            End If
                            Dim pictureUrl As String = productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("data-src", "")
                            If (String.IsNullOrEmpty(pictureUrl)) Then
                                pictureUrl = productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                            End If
                            product.PictureUrl = pictureUrl  'productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                            product.LastUpdate = updateTime
                            product.Description = productName
                            product.SiteID = siteId
                            product.Currency = "$"
                            product.PictureAlt = productName
                            product.SizeHeight = 150
                            product.SizeWidth = 150
                            Dim img1 As String = ""
                            Dim img2 As String = ""
                            Try
                                img1 = productNode.SelectSingleNode("ul/li[4]/img").GetAttributeValue("src", "")
                            Catch ex As Exception
                                'Ignore 
                            End Try
                            Try
                                img2 = productNode.SelectSingleNode("ul/li[5]/img").GetAttributeValue("src", "")
                            Catch ex As Exception
                                'Ignore
                            End Try
                            If (String.IsNullOrEmpty(img1)) Then
                                product.FreeShippingImg = img1
                            ElseIf (img1.Contains("freepic.gif")) Then
                                product.FreeShippingImg = img1
                            ElseIf (img1.Contains("icon_ships24.gif")) Then
                                product.ShipsImg = img1
                            End If
                            If (String.IsNullOrEmpty(img2)) Then
                                product.ShipsImg = img2
                            ElseIf (img2.Contains("freepic.gif")) Then
                                product.FreeShippingImg = img2
                            ElseIf (img2.Contains("icon_ships24.gif")) Then
                                product.ShipsImg = img2
                            End If

                            Dim productId As Integer = helper.InsertProduct(product, DateTime.Now, categoryId, listProduct)
                            helper.InsertSinglePIssue(productId, siteId, issueId, "CA")

                            counter = counter + 1
                        End If
                    End If
                Catch ex As Exception
                    'Ignore
                End Try

                'listProductId.Add(productId)
            Next
            '如果获取的产品数量仍不够3个，则转而去第二页获取产品
            If (counter < 3) Then
                Dim htmlIndex As Integer = cate.IndexOf(".html")
                cate = cate.Insert(htmlIndex, "-page-2")
                cateDoc = helper.GetHtmlDocument(cate, 60000)
                productNodes = cateDoc.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
                For Each productNode In productNodes
                    Try
                        If (counter >= 3) Then
                            Exit For
                        Else
                            Dim product As New Product
                            Dim productName As String = productNode.SelectSingleNode("ul/li[2]").InnerText.Trim()
                            product.Prodouct = productName
                            product.Url = AddDomain("http://www.everbuying.net", productNode.SelectSingleNode("ul/li[2]/a").GetAttributeValue("href", "").Trim())
                            If (listProductUrl.Contains(product.Url) OrElse listProductName.Contains(product.Prodouct)) Then '该产品在本期的邮件中了
                                Continue For
                            Else
                                listProductUrl.Add(product.Url)
                                listProductName.Add(product.Prodouct)
                            End If
                            sku = getSKU(product.Url)
                            If (listProductSku.Contains(sku)) Then
                                Continue For
                            Else
                                listProductSku.Add(sku)
                            End If
                            If Not (helper.IsProductSent(siteId, product.Url, Now.AddDays(-23).ToString(), Now.ToString(), planType)) Then  '该产品在前几期邮件中没有发送
                                Dim productPrice As String = productNode.SelectSingleNode("ul/li[3]").InnerText.Trim()
                                Dim arrPrice As String() = System.Text.RegularExpressions.Regex.Split(productPrice, "USD", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                                If (arrPrice.Count >= 3) Then
                                    For Each arr In arrPrice
                                        If Not (String.IsNullOrEmpty(arr)) Then
                                            'Dim price As String = arr
                                            If (product.Price Is Nothing) Then
                                                product.Price = Double.Parse(arr.Trim())
                                            Else
                                                product.Discount = Double.Parse(arr.Trim())
                                            End If
                                        End If
                                    Next
                                Else
                                    productPrice = productPrice.Replace("USD", "").Trim()  'productPrice.Trim("USD").Trim()
                                    product.Discount = Double.Parse(productPrice)
                                End If
                                Dim pictureUrl As String = productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("data-src", "")
                                If (String.IsNullOrEmpty(pictureUrl)) Then
                                    pictureUrl = productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                                End If
                                product.PictureUrl = pictureUrl  'productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                                product.LastUpdate = updateTime
                                product.Description = productName
                                product.SiteID = siteId
                                product.Currency = "$"
                                product.PictureAlt = productName
                                product.SizeHeight = 150
                                product.SizeWidth = 150
                                Dim img1 As String = ""
                                Dim img2 As String = ""
                                Try
                                    img1 = productNode.SelectSingleNode("ul/li[4]/img").GetAttributeValue("src", "")
                                Catch ex As Exception
                                    'Ignore 
                                End Try
                                Try
                                    img2 = productNode.SelectSingleNode("ul/li[5]/img").GetAttributeValue("src", "")
                                Catch ex As Exception
                                    'Ignore
                                End Try
                                If (String.IsNullOrEmpty(img1)) Then
                                    product.FreeShippingImg = img1
                                ElseIf (img1.Contains("freepic.gif")) Then
                                    product.FreeShippingImg = img1
                                ElseIf (img1.Contains("icon_ships24.gif")) Then
                                    product.ShipsImg = img1
                                End If
                                If (String.IsNullOrEmpty(img2)) Then
                                    product.ShipsImg = img2
                                ElseIf (img2.Contains("freepic.gif")) Then
                                    product.FreeShippingImg = img2
                                ElseIf (img2.Contains("icon_ships24.gif")) Then
                                    product.ShipsImg = img2
                                End If

                                Dim productId As Integer = helper.InsertProduct(product, DateTime.Now, categoryId, listProduct)
                                helper.InsertSinglePIssue(productId, siteId, issueId, "CA")

                                counter = counter + 1
                            End If
                        End If
                    Catch ex As Exception
                        'Ignore
                    End Try

                    'listProductId.Add(productId)
                Next
            End If
            'helper.InsertProductIssue(listProductId, siteId, issueId, "CA", 3, productNodes.Count)
        Next
    End Sub

    ''' <summary>
    ''' 获取指定站点的products表中的所有产品，Dora，2014.02.11
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetProductList(ByVal siteId As Integer) As List(Of Product)
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
    ''' 获取邮件标题
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="periodicalType">期刊名，如：Deals，New Arrivals</param>
    ''' <remarks></remarks>
    Public Shared Sub InsertSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal periodicalType As String)
        Dim helper As New EFHelper
        Dim iMonth As Integer = DateTime.Now.Month
        Dim strMonth As String = ""
        Select Case iMonth
            Case 1
                strMonth = "Jan"
            Case 2
                strMonth = "Feb"
            Case 3
                strMonth = "Mar"
            Case 4
                strMonth = "Apr"
            Case 5
                strMonth = "May"
            Case 6
                strMonth = "Jun"
            Case 7
                strMonth = "Jul"
            Case 8
                strMonth = "Aug"
            Case 9
                strMonth = "Sep"
            Case 10
                strMonth = "Oct"
            Case 11
                strMonth = "Nov"
            Case 12
                strMonth = "Dec"
        End Select
        Dim strCounter As String = ""
        Dim iCounter As Integer = helper.GetSentCount(issueId, siteId, planType)
        If (iCounter < 10) Then
            strCounter = "0" & iCounter.ToString()
        Else
            strCounter = iCounter
        End If
        Dim subject As String = "Everbuying " & periodicalType & " For " & strMonth & "." & Now.Year & ".Vol." & strCounter
        helper.InsertIssueSubject(issueId, subject)

    End Sub

    ''' <summary>
    ''' 把不完整的链接加上域名url
    ''' </summary>
    ''' <param name="domainUrl">域名url，最后一个字符不含"/"</param>
    ''' <param name="orginalUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddDomain(ByVal domainUrl As String, ByVal orginalUrl As String)
        If (orginalUrl.Contains(domainUrl)) Then
            Return orginalUrl
        Else
            If Not (orginalUrl.IndexOf("/") = 0) Then
                orginalUrl = String.Concat("/", orginalUrl)
            End If
            orginalUrl = String.Concat(domainUrl, orginalUrl)
            Return orginalUrl
        End If
    End Function

    ''' <summary>
    ''' 把不完整的链接加上域名url
    ''' </summary>
    ''' <param name="domainUrl">域名url，最后一个字符不含"/"</param>
    ''' <param name="orginalUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getSKU(ByVal productURL As String) As String
        Dim helper As New EFHelper
        Dim productDetail As HtmlDocument = helper.GetHtmlDocument(productURL, 60000)
        Dim skuNode As HtmlNode = productDetail.DocumentNode.SelectSingleNode("//div[@id='h1_goodsTitle']/em")
        Dim sku As String = skuNode.InnerText.Trim.Substring(0, 7)
        Return sku
    End Function
End Class
