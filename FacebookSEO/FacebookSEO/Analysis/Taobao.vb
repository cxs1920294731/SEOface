Imports HtmlAgilityPack
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports Enumerable = System.Linq.Enumerable

Public Class Taobao
    Private Shared helper As New EFHelper()
    Private Shared efContext As New EmailAlerterEntities()

    ''' <summary>
    ''' 程序开始入口
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="Url"></param>
    ''' <remarks></remarks>
    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        InsertCategory2(siteId, IssueID, url)
        InsertProductData2(siteId, IssueID, url)
        GetSubject(IssueID)
        Dim contactLists As List(Of SplitContactList) = helper.GetSplitContactLists(siteId)
        If contactLists.Count > 0 Then
            For Each contactList As SplitContactList In contactLists
                helper.InsertContactList(IssueID, contactList.ContactListName, contactList.SendingStatus)
            Next
        End If
    End Sub

    Public Shared Sub Start2(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        InsertCategory2(siteId, IssueID, url)
        InsertProductData2(siteId, IssueID, url)
    End Sub

    ''' <summary>
    ''' 从Promotion产品中获得Subject信息
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetSubject(ByVal issueId As Integer)
        Dim subject As String = "[FIRSTNAME] 你好," + helper.GetPromotionProduct(issueId)
        helper.InsertIssueSubject(issueId, subject)
    End Sub

    ''' <summary>
    ''' For Test only
    ''' </summary>
    ''' <param name="ShopUrl"></param>
    ''' <remarks></remarks>
    Public Shared Sub TestProductData(ShopUrl As String)
        Dim hotSellUrl As String = ShopUrl + "/search.htm?orderType=_hotsell"
        Dim doc As HtmlDocument = GetHtmlDocByUrl(hotSellUrl)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//ul[@class='shop-list']/li/div")
        Dim nowTime As DateTime = Now
        If divNodes IsNot Nothing Then
            Dim r As Integer = GetRnd(0, 7)
            For i As Integer = 0 To 8
                If i = r Then
                    Dim imgNode = divNodes(r).SelectSingleNode("div[@class='pic']")
                    Dim rateNode = divNodes(r).SelectSingleNode("div[@class='rating']")
                    Dim currency As String = "CNY"
                    Dim price As Decimal = Decimal.Parse(divNodes(r).SelectSingleNode("div[@class='price']/strong").InnerText.Trim())
                    Dim product As String = divNodes(r).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                    Dim url As String = divNodes(r).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                    Dim lastUpdate As DateTime = nowTime
                    Dim description As String = divNodes(r).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                    Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                    Dim discount As Decimal = Decimal.Parse(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim())

                    doc = GetHtmlDocByUrl(url)
                    Dim innerHtml As String = doc.DocumentNode.InnerHtml


                    Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                    Dim rateCount As String = GetJsonFromTaobao(rateCountApi)

                    Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                    Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                    'Dim test As String = GetJsonDataByParameter("""quantity"":" + GetValue(itemInfo, "quantity:", "}") + "}", "quantity", "quanity")
                    Dim sales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                    Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")

                    Dim rating As Decimal = Convert.ToDecimal(GetTotalRate(rateCount.Split(",")(1)))
                    Dim comment As Integer = Convert.ToInt32(GetTotalComment(rateCount))
                End If
            Next
        Else
            Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='item4line1']/dl")
            For i As Integer = 0 To itemNodes.Count - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("//dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("src", "")
                Dim picAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")
                Dim detailNode As HtmlNode = items.SelectSingleNode("//dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText
                Dim descrption As String = detailNode.SelectSingleNode("a").InnerText
                Dim discount As Decimal = Decimal.Parse(items.SelectSingleNode("//span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = items.SelectSingleNode("//span[@class='sprice-area']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If
                Dim sales As Long = Long.Parse(items.SelectSingleNode("//span[@class='sale-num']").InnerText)
                Dim lastUpdate As DateTime = DateTime.Now
                Dim currency As String = "CNY"

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)


                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                'Dim test As String = GetJsonDataByParameter("""quantity"":" + GetValue(itemInfo, "quantity:", "}") + "}", "quantity", "quanity")
                'Dim monthlysales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                'promoInfoApi 促销价信息
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                If Not String.IsNullOrEmpty(promPrice) Then
                    discount = Decimal.Parse(promPrice)
                End If

                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")

                Dim rating As Decimal = Convert.ToDecimal(GetTotalRate(rateCount.Split(",")(1)))
                Dim comment As Integer = Convert.ToInt32(GetTotalComment(rateCount))

            Next

        End If

    End Sub

    Public Shared Sub TestProductData2(ShopUrl As String)
        Dim count As Integer = 8
        Dim nowTime As DateTime = DateTime.Now
        Dim hotSellUrl As String = ShopUrl
        Dim doc As HtmlDocument = GetHtmlDocByUrl(hotSellUrl)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid big']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To IIf(count > divNodes.Count, divNodes.Count, count) - 1

                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If


                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim sales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                'Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                'listProductId.Add(returnId)
            Next
            'InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim priceNode As HtmlNode = divNodes(i).SelectSingleNode("div[@class='price']/strong")
                Dim price As Decimal
                Dim discount As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(RemoveStringCode(priceNode.InnerText.Trim()))
                    discount = price
                End If


                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim sales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                'Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                'listProductId.Add(returnId)
            Next
            'InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then
            For i As Integer = 0 To count - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If
                Dim sales As Long = Long.Parse(attrNode.SelectSingleNode("div/span[@class='sale-num']").InnerText)
                Dim lastUpdate As DateTime = nowTime
                Dim currency As String = "CNY"

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 & 评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                'Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                'Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                'Dim monthlysales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Or pictureUrl.Contains(".gif") Then
                    pictureUrl = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")
                    If String.IsNullOrEmpty(pictureUrl) Then
                        pictureUrl = picUrl
                    End If
                End If

                'Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                'listProductId.Add(returnId)
            Next
            'InsertProductIssue(siteId, IssueId, cateType, listProductId)

        End If

    End Sub

    ''' <summary>
    ''' 插入N件产品到指定的类别下
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="count"></param>
    ''' <param name="category"></param>
    ''' <param name="cateType"></param>
    ''' <param name="doc"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertProductByUrl(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal count As Integer,
                                         ByVal category As String, ByVal cateType As String, ByVal doc As HtmlDocument)
        Dim queryCate = (From m In efContext.Categories
                  Where m.Category1 = category AndAlso m.SiteID = siteId
                  Select m).First()
        Dim categoryId As Integer = queryCate.CategoryID
        Dim listProductId As New List(Of Integer)
        Dim nowTime As DateTime = DateTime.Now
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid big']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To IIf(count > divNodes.Count, divNodes.Count, count) - 1

                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If


                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim sales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim priceNode As HtmlNode = divNodes(i).SelectSingleNode("div[@class='price']/strong")
                Dim price As Decimal
                Dim discount As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(RemoveStringCode(priceNode.InnerText.Trim()))
                    discount = price
                End If


                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim sales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then
            For i As Integer = 0 To count - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If
                Dim sales As Long = Long.Parse(attrNode.SelectSingleNode("div/span[@class='sale-num']").InnerText)
                Dim lastUpdate As DateTime = nowTime
                Dim currency As String = "CNY"

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 & 评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                'Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                'Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                'Dim monthlysales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Or pictureUrl.Contains(".gif") Then
                    pictureUrl = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")
                    If String.IsNullOrEmpty(pictureUrl) Then
                        pictureUrl = picUrl
                    End If
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

    End Sub

    Public Shared Sub InsertProductByUrlForTest(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal count As Integer,
                                         ByVal category As String, ByVal cateType As String, ByVal doc As HtmlDocument)
        Dim queryCate = (From m In efContext.Categories
                      Where m.Category1 = category AndAlso m.SiteID = siteId
                      Select m).First()
        Dim categoryId As Integer = queryCate.CategoryID
        Dim listProductId As New List(Of Integer)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid big']/ul[@class='shop-list']/li/div")

        If divNodes IsNot Nothing Then
            Dim nowTime As DateTime = DateTime.Now
            For i As Integer = 0 To count
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim test As String = imgNode.InnerHtml
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim sales As Long = 100
                Dim pictureUrl As String = picUrl
                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            Dim nowTime As DateTime = DateTime.Now
            For i As Integer = 0 To count

                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim sales As Long = 100
                Dim pictureUrl As String = picUrl
                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then
            For i As Integer = 0 To count - 1
                Dim items As HtmlNode = itemNodes(i)

                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If
                Dim salesNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='sale-num']")
                Dim sales As Long
                If salesNode IsNot Nothing Then
                    sales = Long.Parse(salesNode.InnerText)
                Else
                    sales = 100
                End If

                Dim lastUpdate As DateTime = DateTime.Now
                Dim currency As String = "CNY"
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim pictureUrl As String = picUrl

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)
        End If
    End Sub


    ''' <summary>
    ''' 随机获取促销产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="count"></param>
    ''' <param name="category"></param>
    ''' <param name="cateType"></param>
    ''' <param name="doc"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertPromotionItem(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal count As Integer,
                                         ByVal category As String, ByVal cateType As String, ByVal doc As HtmlDocument)
        Dim queryCate = (From m In efContext.Categories
                              Where m.Category1 = category AndAlso m.SiteID = siteId
                              Select m).First()
        Dim categoryId As Integer = queryCate.CategoryID
        Dim listProductId As New List(Of Integer)
        Dim nowTime As DateTime = DateTime.Now
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid big']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count


                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim quanity As String = GetValue(itemInfo, "quanity:", ",")
                Dim sales As Long
                If Not String.IsNullOrEmpty(quanity) Then
                    sales = Long.Parse(quanity)
                End If

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            'InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim priceNode As HtmlNode = divNodes(i).SelectSingleNode("div[@class='price']/strong")
                Dim price As Decimal
                Dim discount As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(RemoveStringCode(priceNode.InnerText.Trim()))
                    discount = price
                End If

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 &评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim quanity As String = GetValue(itemInfo, "quanity:", ",")
                Dim sales As Long
                If Not String.IsNullOrEmpty(quanity) Then
                    sales = Long.Parse(quanity)
                End If

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Then
                    pictureUrl = picUrl
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then

            For i As Integer = 0 To IIf(count > itemNodes.Count, itemNodes.Count, count) - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If

                Dim lastUpdate As DateTime = nowTime
                Dim currency As String = "CNY"

                doc = GetHtmlDocByUrl(url)
                Dim innerHtml As String = doc.DocumentNode.InnerHtml

                '打分 & 评论
                Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                'promoInfoApi 促销价信息Api
                Dim promoPriceApi As String = GetValue(innerHtml, """umpStockUrl"":""", """")
                If Not String.IsNullOrEmpty(promoPriceApi) Then
                    Dim promPrice As String = GetValue(GetPromotionPriceJsonFromTaobao(url, promoPriceApi), "price:""", """")
                    If Not String.IsNullOrEmpty(promPrice) Then
                        discount = Decimal.Parse(promPrice)
                    End If
                End If

                'itemInfoApi 获得月销数量Api
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                Dim quanity As String = GetValue(itemInfo, "quanity:", ",")
                Dim sales As Long
                If Not String.IsNullOrEmpty(quanity) Then
                    sales = Long.Parse(quanity)
                End If

                '大图Url
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                If String.IsNullOrEmpty(pictureUrl) Or pictureUrl.Contains(".gif") Then
                    pictureUrl = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")
                    If String.IsNullOrEmpty(pictureUrl) Then
                        pictureUrl = picUrl
                    End If
                End If

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If
    End Sub

    Public Shared Sub InsertPromotionItemForTest(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal count As Integer,
                                     ByVal category As String, ByVal cateType As String, ByVal doc As HtmlDocument)
        Dim queryCate = (From m In efContext.Categories
                              Where m.Category1 = category AndAlso m.SiteID = siteId
                              Select m).First()
        Dim categoryId As Integer = queryCate.CategoryID
        Dim listProductId As New List(Of Integer)
        Dim nowTime As DateTime = DateTime.Now
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid big']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count - 1
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim sales As Long = 100
                Dim pictureUrl As String = picUrl
                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            'update picture
            Dim promproduct = efContext.Products.Single(Function(m) m.ProdouctID = productId)
            doc = GetHtmlDocByUrl(promproduct.Url)
            '大图Url
            Dim pictureUrl1 As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
            promproduct.PictureUrl = pictureUrl1
            efContext.SaveChanges()

            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
        If divNodes IsNot Nothing Then
            For i As Integer = 0 To count - 1
                Dim price As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim description As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim currency As String = "CNY"
                Dim product As String = divNodes(i).SelectSingleNode("div[@class='desc']").InnerText.Trim()
                Dim url As String = divNodes(i).SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "").Trim()
                Dim lastUpdate As DateTime = nowTime
                Dim imgNode = divNodes(i).SelectSingleNode("div[@class='pic']")
                Dim pictureAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim picUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-ks-lazyload", "")
                Dim discount As Decimal = Decimal.Parse(RemoveStringCode(divNodes(i).SelectSingleNode("div[@class='price']/strong").InnerText.Trim()))
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim sales As Long = 100
                Dim pictureUrl As String = picUrl
                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            'update picture
            Dim promproduct = efContext.Products.Single(Function(m) m.ProdouctID = productId)
            doc = GetHtmlDocByUrl(promproduct.Url)
            '大图Url
            Dim pictureUrl1 As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
            promproduct.PictureUrl = pictureUrl1
            efContext.SaveChanges()

            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then

            For i As Integer = 0 To count - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

                '过滤掉运费差价专拍
                If description.Contains("差额") Or description.Contains("专拍") Or description.Contains("补拍") Or description.Contains("差价") Or description.Contains("補拍") Or description.Contains("鏈接") Or description.Contains("專拍") Or description.Contains("運費") Or description.Contains("链接") Then
                    Continue For
                End If

                Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                Dim price As Decimal
                If priceNode IsNot Nothing Then
                    price = Decimal.Parse(priceNode.InnerText)
                Else
                    price = discount
                End If
                Dim salesNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='sale-num']")
                Dim sales As Long
                If salesNode IsNot Nothing Then
                    sales = Long.Parse(salesNode.InnerText)
                Else
                    sales = 100
                End If
                'Dim sales As Long = Long.Parse(attrNode.SelectSingleNode("div/span[@class='sale-num']").InnerText)
                Dim lastUpdate As DateTime = DateTime.Now
                Dim currency As String = "CNY"
                Dim rating As Decimal = 4.8D
                Dim comment As Integer = 1234
                Dim pictureUrl As String = picUrl

                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, categoryId, sales, discount, rating, comment)
                listProductId.Add(returnId)
            Next
            InsertProductIssue(siteId, IssueId, cateType, listProductId)

            '插入promotion记录
            Dim r As Integer = GetRnd(0, listProductId.Count - 2)
            Dim productId As Integer = listProductId(r)
            'update picture
            Dim promproduct = efContext.Products.Single(Function(m) m.ProdouctID = productId)
            doc = GetHtmlDocByUrl(promproduct.Url)
            '大图Url
            Dim pictureUrl1 As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("src", "")
            If String.IsNullOrEmpty(pictureUrl1) Then
                pictureUrl1 = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
            End If
            promproduct.PictureUrl = pictureUrl1

            efContext.SaveChanges()

            Dim plist As New List(Of Integer)
            plist.Add(productId)
            InsertProductIssue(siteId, IssueId, "PO", plist)

            '添加Subject
            Dim promotion = (From p In efContext.Products
                             Where p.ProdouctID = productId And p.SiteID = siteId
                             Select p).First()
            Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = IssueId)
            myIssue.Subject = "最新上市;" + promotion.Prodouct
            efContext.SaveChanges()
        End If

    End Sub

    ''' <summary>
    ''' 插入所有产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="ShopUrl"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertProductData(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)
        Dim hotSellUrl As String = ShopUrl + "/search.htm?orderType=_hotsell"
        Dim newOnUrl As String = ShopUrl + "/search.htm?orderType=newon"
        Dim hotKeepUrl As String = ShopUrl + "/search.htm?orderType=_hotkeep"
        Dim renqiUrl As String = ShopUrl + "/search.htm?orderType=_coefp"
        Dim dealsUrl As String = ShopUrl + "/search.htm?orderType=price"

        'Dim hots As String = "http://rjfashion.taobao.com/search.htm?scid=717123914"
        'Dim value As String = "http://rjfashion.taobao.com/search.htm?scid=459642221"

        InsertProductByUrl(siteId, IssueId, 8, "newon", "NE", GetHtmlDocByUrl(newOnUrl))
        InsertProductByUrl(siteId, IssueId, 8, "hotkeep", "HK", GetHtmlDocByUrl(hotKeepUrl))
        InsertProductByUrl(siteId, IssueId, 8, "popular", "MP", GetHtmlDocByUrl(renqiUrl))
        InsertProductByUrl(siteId, IssueId, 9, "price", "DE", GetHtmlDocByUrl(dealsUrl))
        InsertPromotionItem(siteId, IssueId, 9, "hotsell", "BE", GetHtmlDocByUrl(hotSellUrl))
        'InsertProductByUrl(siteId, IssueId, 9, "热卖小西装", "CA", GetHtmlDocByUrl(hots))
        'InsertProductByUrl(siteId, IssueId, 10, "超值两件套", "CA", GetHtmlDocByUrl(value))

    End Sub

    Public Shared Sub InsertProductData2(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)
        Dim queryTemplateInfo = From ti In efContext.TemplateInfoes
                                Where ti.SiteId = siteId
                                Select ti
        Dim tlist As List(Of TemplateInfo) = queryTemplateInfo.ToList()
        If (tlist.Count > 0) Then
            Dim catelist As List(Of String) = Split(tlist(0).SortedCates, ",").ToList()
            For Each s As String In catelist
                If Not String.IsNullOrEmpty(s) Then
                    Dim cateinfo As List(Of String) = Split(s, ":").ToList()
                    Dim url As String = ShopUrl + "/search.htm?scid=" + cateinfo(0)
                    InsertProductByUrl(siteId, IssueId, 8, cateinfo(1), "CA", GetHtmlDocByUrl(url))
                End If
            Next
        End If

        'Best Seller &　Promotion
        Dim hotSellUrl As String = ShopUrl + "/search.htm?orderType=_hotsell"
        InsertPromotionItem(siteId, IssueId, 9, "hotsell", "BE", GetHtmlDocByUrl(hotSellUrl))

    End Sub

    Public Shared Sub InsertProductDataForTest(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)
        Dim hotSellUrl As String = ShopUrl + "/search.htm?orderType=_hotsell"
        Dim newOnUrl As String = ShopUrl + "/search.htm?orderType=newon"
        Dim hotKeepUrl As String = ShopUrl + "/search.htm?orderType=_hotkeep"
        Dim renqiUrl As String = ShopUrl + "/search.htm?orderType=_coefp"
        Dim dealsUrl As String = ShopUrl + "/search.htm?orderType=price"

        'Dim hots As String = "http://rjfashion.taobao.com/search.htm?scid=524990480"
        'Dim value As String = "http://rjfashion.taobao.com/search.htm?scid=459642221"


        InsertProductByUrlForTest(siteId, IssueId, 8, "newon", "NE", GetHtmlDocByUrl(newOnUrl))
        InsertProductByUrlForTest(siteId, IssueId, 8, "hotkeep", "HK", GetHtmlDocByUrl(hotKeepUrl))
        InsertProductByUrlForTest(siteId, IssueId, 8, "popular", "MP", GetHtmlDocByUrl(renqiUrl))
        InsertProductByUrlForTest(siteId, IssueId, 9, "price", "DE", GetHtmlDocByUrl(dealsUrl))
        InsertPromotionItemForTest(siteId, IssueId, 9, "hotsell", "BE", GetHtmlDocByUrl(hotSellUrl))
        'InsertProductByUrl(siteId, IssueId, 9, "热卖小西装", "CA", GetHtmlDocByUrl(hots))
        'InsertProductByUrl(siteId, IssueId, 10, "超值两件套", "CA", GetHtmlDocByUrl(value))

    End Sub

    ''' <summary>
    ''' 添加数据到Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="section"></param>
    ''' <remarks></remarks>
    Private Shared Sub InsertProductIssue(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal section As String, ByVal listProductId As List(Of Integer))
        'Dim listProductIssue As New List(Of Products_Issue)
        'Dim queryIssue = efContext.Products_Issue
        'For Each q In queryIssue
        '    listProductIssue.Add(New Products_Issue With {.ProductId = q.ProductId, .SiteId = q.SiteId, .IssueDate = q.IssueDate, .SectionID = q.SectionID})
        'Next
        'Dim queryProduct = efContext.Products
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

    ''' <summary>
    ''' 判断将要插入Categories表中URL是否重复，若数据重复，则返回false
    ''' </summary>
    ''' <param name="category"></param>
    ''' <param name="siteId"></param>
    ''' <param name="Url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeCategory(ByVal category As String, ByVal siteId As Integer, ByVal Url As String) As Boolean
        '从数据库读取分类填充到List里面，判断记录是否重复
        Dim queryCategory = From c In efContext.Categories
                          Where c.SiteID = siteId
                          Select c
        Dim list As New List(Of Category)
        For Each q In queryCategory
            list.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next
        For Each li In list
            If (li.Url = Url) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 通过Url返回指定的HtmlDocument对象
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlDocByUrl(pageUrl As String) As HtmlDocument
        Dim request As WebRequest = WebRequest.Create(pageUrl)
        request.Timeout = 300000
        request.Headers.Add("Accept-Language", "zh-CN")
        'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
        Dim response As WebResponse = request.GetResponse()
        'WebRequest.GetResponse方法，返回对 Internet 请求的响应
        Dim resStream As Stream = response.GetResponseStream()
        'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
        Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
        Dim document As New HtmlDocument()
        document.Load(resStream, pageEncoding)
        Return document
    End Function

    ''' <summary>
    ''' 插入所有分类信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="ShopUrl"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertCategory(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)
        '从数据库读取分类填充到List里面，判断记录是否重复
        Dim queryCategory = From c In efContext.Categories
                          Where c.SiteID = siteId
                          Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next

        Dim cate As New Category()
        cate.Category1 = "hotsell"
        cate.SiteID = siteId
        cate.LastUpdate = DateTime.Now
        cate.Description = "hotsell"
        cate.Url = ShopUrl + "/search.htm?orderType=_hotsell"
        cate.ParentID = -1
        If JudgeCategory(cate.Category1, cate.SiteID, cate.Url) Then
            efContext.AddToCategories(cate)
        End If

        Dim cateNewon As New Category()
        cateNewon.Category1 = "newon"
        cateNewon.SiteID = siteId
        cateNewon.LastUpdate = DateTime.Now
        cateNewon.Description = "newon"
        cateNewon.Url = ShopUrl + "/search.htm?orderType=newon"
        cateNewon.ParentID = -1
        If JudgeCategory(cateNewon.Category1, cateNewon.SiteID, cateNewon.Url) Then
            efContext.AddToCategories(cateNewon)
        End If

        Dim cateHotKeep As New Category()
        cateHotKeep.Category1 = "hotkeep"
        cateHotKeep.SiteID = siteId
        cateHotKeep.LastUpdate = DateTime.Now
        cateHotKeep.Description = "hotkeep"
        cateHotKeep.Url = ShopUrl + "/search.htm?orderType=_hotkeep"
        cateHotKeep.ParentID = -1
        If JudgeCategory(cateHotKeep.Category1, cateHotKeep.SiteID, cateHotKeep.Url) Then
            efContext.AddToCategories(cateHotKeep)
        End If

        Dim cateRenqi As New Category()
        cateRenqi.Category1 = "popular"
        cateRenqi.SiteID = siteId
        cateRenqi.LastUpdate = DateTime.Now
        cateRenqi.Description = "popular"
        cateRenqi.Url = ShopUrl + "/search.htm?orderType=_coefp"
        cateRenqi.ParentID = -1
        If JudgeCategory(cateRenqi.Category1, cateRenqi.SiteID, cateRenqi.Url) Then
            efContext.AddToCategories(cateRenqi)
        End If

        Dim cateDeals As New Category()
        cateDeals.Category1 = "price"
        cateDeals.SiteID = siteId
        cateDeals.LastUpdate = DateTime.Now
        cateDeals.Description = "price"
        cateDeals.Url = ShopUrl + "/search.htm?orderType=price"
        cateDeals.ParentID = -1
        If JudgeCategory(cateDeals.Category1, cateDeals.SiteID, cateDeals.Url) Then
            efContext.AddToCategories(cateDeals)
        End If


        'Dim cateHots As New Category()
        'cateHots.Category1 = "热卖小西装"
        'cateHots.SiteID = siteId
        'cateHots.LastUpdate = DateTime.Now
        'cateHots.Description = "热卖小西装"
        'cateHots.Url = "http://rjfashion.taobao.com/search.htm?scid=717123914"
        'cateHots.ParentID = -1
        'If JudgeCategory(cateHots.Category1, cateHots.SiteID, cateHots.Url) Then
        '    efContext.AddToCategories(cateHots)
        'End If

        'Dim cateDouble As New Category()
        'cateDouble.Category1 = "超值两件套"
        'cateDouble.SiteID = siteId
        'cateDouble.LastUpdate = DateTime.Now
        'cateDouble.Description = "超值两件套"
        'cateDouble.Url = "http://rjfashion.taobao.com/search.htm?scid=459642221"
        'cateDouble.ParentID = -1
        'If JudgeCategory(cateDouble.Category1, cateDouble.SiteID, cateDouble.Url) Then
        '    efContext.AddToCategories(cateDouble)
        'End If

        efContext.SaveChanges()
    End Sub

    Public Shared Sub InsertCategory2(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)
        '从数据库读取分类填充到List里面，判断记录是否重复
        Dim queryCategory = From c In efContext.Categories
                          Where c.SiteID = siteId
                          Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next

        Dim queryTemplateInfo = From ti In efContext.TemplateInfoes
                                Where ti.SiteId = siteId
                                Select ti
        Dim tlist As List(Of TemplateInfo) = queryTemplateInfo.ToList()
        If (tlist.Count > 0) Then
            Dim catelist As List(Of String) = Split(tlist(0).SortedCates, ",").ToList()
            For Each s As String In catelist
                If Not String.IsNullOrEmpty(s) Then
                    Dim cateinfo As List(Of String) = Split(s, ":").ToList()
                    Dim cateHots As New Category()
                    cateHots.Category1 = cateinfo(1)
                    cateHots.SiteID = siteId
                    cateHots.LastUpdate = DateTime.Now
                    cateHots.Description = cateinfo(1)
                    cateHots.Url = ShopUrl + "/search.htm?scid=" + cateinfo(0)
                    cateHots.ParentID = -1
                    If JudgeCategory(cateHots.Category1, cateHots.SiteID, cateHots.Url) Then
                        efContext.AddToCategories(cateHots)
                    End If
                End If
            Next
        End If

        Dim cate As New Category()
        cate.Category1 = "hotsell"
        cate.SiteID = siteId
        cate.LastUpdate = DateTime.Now
        cate.Description = "hotsell"
        cate.Url = ShopUrl + "/search.htm?orderType=_hotsell"
        cate.ParentID = -1
        If JudgeCategory(cate.Category1, cate.SiteID, cate.Url) Then
            efContext.AddToCategories(cate)
        End If

        efContext.SaveChanges()
    End Sub

    Public Shared Sub ClearTestData(ByVal siteId As Integer)
        Dim efContext As New EmailAlerterEntities()
        Dim cates As List(Of Category) = (From c In efContext.Categories
                    Where c.SiteID = siteId
                    Select c).ToList()
        Dim proucts As List(Of Product) = (From p In efContext.Products
                      Where p.SiteID = siteId
                      Select p).ToList()
        For Each category As Category In cates
            For Each product As Product In proucts
                category.Products.Remove(product)
            Next
        Next
        'efContext.SaveChanges()


        Dim issues As List(Of Issue) = (From i In efContext.Issues
                     Where i.SiteID = siteId
                     Select i).ToList()
        Dim pro_issues As List(Of Products_Issue) = (From pi In efContext.Products_Issue
                         Where pi.SiteId = siteId
                         Select pi).ToList()


        For Each issue As Issue In issues
            For Each productsIssue As Products_Issue In pro_issues
                issue.Products_Issue.Remove(productsIssue)
            Next
        Next
        'efContext.SaveChanges()

        For Each category As Category In cates
            efContext.Categories.DeleteObject(category)
        Next

        For Each issue As Issue In issues
            efContext.Issues.DeleteObject(issue)
        Next

        For Each product As Product In proucts
            efContext.Products.DeleteObject(product)
        Next
        efContext.SaveChanges()

    End Sub

    Public Shared Sub InsertCategoryForTest(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal ShopUrl As String)

        Dim cate As New Category()
        cate.Category1 = "hotsell"
        cate.SiteID = siteId
        cate.LastUpdate = DateTime.Now
        cate.Description = "hotsell"
        cate.Url = ShopUrl + "/search.htm?orderType=_hotsell"
        cate.ParentID = -1
        efContext.AddToCategories(cate)

        Dim cateNewon As New Category()
        cateNewon.Category1 = "newon"
        cateNewon.SiteID = siteId
        cateNewon.LastUpdate = DateTime.Now
        cateNewon.Description = "newon"
        cateNewon.Url = ShopUrl + "/search.htm?orderType=newon"
        cateNewon.ParentID = -1
        efContext.AddToCategories(cateNewon)


        Dim cateHotKeep As New Category()
        cateHotKeep.Category1 = "hotkeep"
        cateHotKeep.SiteID = siteId
        cateHotKeep.LastUpdate = DateTime.Now
        cateHotKeep.Description = "hotkeep"
        cateHotKeep.Url = ShopUrl + "/search.htm?orderType=_hotkeep"
        cateHotKeep.ParentID = -1
        efContext.AddToCategories(cateHotKeep)

        Dim cateRenqi As New Category()
        cateRenqi.Category1 = "popular"
        cateRenqi.SiteID = siteId
        cateRenqi.LastUpdate = DateTime.Now
        cateRenqi.Description = "popular"
        cateRenqi.Url = ShopUrl + "/search.htm?orderType=_coefp"
        cateRenqi.ParentID = -1
        efContext.AddToCategories(cateRenqi)

        Dim cateDeals As New Category()
        cateDeals.Category1 = "price"
        cateDeals.SiteID = siteId
        cateDeals.LastUpdate = DateTime.Now
        cateDeals.Description = "price"
        cateDeals.Url = ShopUrl + "/search.htm?orderType=price"
        cateDeals.ParentID = -1
        efContext.AddToCategories(cateDeals)
        efContext.SaveChanges()
    End Sub


    ''' <summary>
    ''' 更新或插入数据到Products表中,并返回添加产品的Id
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="url"></param>
    ''' <param name="price"></param>
    ''' <param name="pictureUrl"></param>
    ''' <param name="lastUpdate"></param>
    ''' <param name="description"></param>
    ''' <param name="siteId"></param>
    ''' <param name="currency"></param>
    ''' <param name="pictureAlt"></param>
    ''' <param name="sizeWidth"></param>
    ''' <param name="sizeHeight"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="list"></param>
    ''' <returns>产品的Id</returns>
    ''' <remarks></remarks>
    Public Shared Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                              ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal categoryId As Integer, ByVal sales As Integer, ByVal discount As Decimal, ByVal score As Decimal, totalComment As Integer) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.Single()
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
        product.Sales = sales
        product.TbScore = score
        product.TbComment = totalComment
        product.Discount = discount

        Dim list As List(Of Product) = GetProductList(siteId)
        If (JudgeProduct(product.Prodouct, product.Url, product.Price, product.PictureUrl, product.SiteID, product.Currency, list)) Then
            efContext.AddToProducts(product)
            product.Categories.Add(category)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = (From m In efContext.Products Where m.Url = url And m.SiteID = siteId).Single()
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            updateProduct.Sales = sales
            updateProduct.TbScore = score
            updateProduct.TbComment = totalComment
            updateProduct.Discount = discount
            Dim queryCate = From p In efContext.Products
                                Where p.ProdouctID = updateProduct.ProdouctID
                                Select p
            Dim cate = queryCate.Single.Categories
            If Not cate.Contains(category) Then
                category.Products.Add(updateProduct)
            End If
            efContext.SaveChanges()
            Return updateProduct.ProdouctID
        End If
    End Function


    ''' <summary>
    ''' 判断即将插入的数据URL是否在数据库中已经存在，如果存在，返回false
    ''' </summary>
    ''' <param name="product"></param>
    ''' <param name="url"></param>
    ''' <param name="price"></param>
    ''' <param name="pictureUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="currency"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeProduct(ByVal product As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal siteId As Integer, ByVal currency As String, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url = url) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 将Products表中需要匹配列的信息添加到list中
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetProductList(ByVal siteId As Integer) As List(Of Product)
        '将Ads表中需要匹配的字段写入list中
        Dim query = From p In efContext.Products
                   Where p.SiteID = siteId
                   Select p
        Dim listProduct As New List(Of Product)
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next
        Return listProduct
    End Function

#Region "Tempalte"
    ''' <summary>
    ''' 添加模板
    ''' </summary>
    ''' <param name="t"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddTempalte(t As Template)
        efContext.AddToTemplates(t)
        efContext.SaveChanges()
        Return True
    End Function

    Public Shared Function UpdateTemplate(tid As Integer, content As String)
        Dim template = efContext.Templates.Single(Function(t) t.Tid = tid)
        template.Contents = content
        efContext.SaveChanges()
    End Function
#End Region

#Region "Func"

    ''' <summary>
    ''' 获取随机数
    ''' </summary>
    ''' <param name="min"></param>
    ''' <param name="max"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRnd(min As Long, max As Long) As Integer
        Randomize() '没有这个 产生的数会一样
        GetRnd = Rnd() * (max - min + 1) + min
    End Function

    ''' <summary>
    ''' 从字符串中提取浮点数
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveStringCode(str As String)
        Dim reg As New Regex("(-?\d+)(\.\d+)?")
        Dim m As Match = reg.Match(str)
        Return m.Groups(0).Value()
    End Function

    Public Shared Function GetTotalComment(str As String)
        Dim rg As New Regex("(?<=("":))[.\s\S]*?(?=(,""))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Public Shared Function GetTotalRate(str As String)
        Dim rg As New Regex("(?<=("":))[.\s\S]*?(?=})", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    ''' <summary>
    ''' Read rateCounter Json
    ''' </summary>
    ''' <param name="jText"></param>
    ''' <param name="name"></param>
    ''' <param name="param"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJsonDataByParameter(jText As String, name As String, param As String)
        Dim jsonText As String = "{""" + jText + "}"
        Dim jo As JObject = DirectCast(JsonConvert.DeserializeObject(jsonText), JObject)
        Dim r As String = jo(name)(param).ToString()
        Return r
    End Function

    Public Shared Function GetJsonFromTaobao(pageUrl As String) As String
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

    Public Shared Function GetValue(str As String, s As String, e As String) As String
        Dim rg As New Regex("(?<=(" & s & "))[.\s\S]*?(?=(" & e & "))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Public Shared Function GetPromotionPriceJsonFromTaobao(pageUrl As String, apiUrl As String) As String
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(apiUrl)
        hwr.Referer = pageUrl
        hwr.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:19.0) Gecko/20100101 Firefox/19.0"
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

#End Region
End Class
