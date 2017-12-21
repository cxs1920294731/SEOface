Imports HtmlAgilityPack
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports Enumerable = System.Linq.Enumerable

Public Class RjNews
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
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal preSubject As String)
        Dim newonUrl As String = helper.GetXinpinUrlandSubject("http://rjfashion.taobao.com/search.htm?spm=a1z10.5-c.w5001-6317386778.6.c0m0yX&orderType=_hotsell&scene=taobao_shop")(0)
        'Dim newonUrl As String = "http://rjfashion.taobao.com/category-782160892.htm?spm=a1z10.5.w4010-162191306.8.gLi8o4&search=y&parentCatId=177697977&parentCatName=%D0%C2%BF%EE-New+Arrival&catName=%26gt%3B8%D4%C226%C8%D5%D0%C2%C6%B7#bd"
        'InsertDefaultCategory(siteId, IssueID, newonUrl)
        helper.FetchProduct(newonUrl, "newon", "NE", planType, 20, DateTime.Now, DateTime.Now, siteId, IssueID)
        'InsertProductByUrl(siteId, IssueID, 20, "newon", "NE", GetPageContentWithAgent(newonUrl))
        Dim hotsaleURL = "http://rjfashion.taobao.com/search.htm?spm=a1z10.3.w4002-162183302.26.mzYGwJ&mid=w-162183302-0&scene=taobao_shop&orderType=hotsell_desc&pageNo=1"
        helper.FetchProduct(hotsaleURL, "newon", "DA", planType, 4, DateTime.Now.AddDays(-30), DateTime.Now, siteId, IssueID)
        'GetProducts(siteId, IssueID, hotsaleURL, "newon", 4, DateTime.Now, planType)
        GetSubject(IssueID, preSubject)
    End Sub

    Private Shared Function GetNewonUrl(ByVal url As String)
        Try
            Dim document As HtmlDocument = GetHtmlDocByUrl(url)
            Dim rootNode As HtmlNode = document.DocumentNode
            Return rootNode.SelectSingleNode("//div[@class='nav_c']/p/a[1]").GetAttributeValue("href", "")
        Catch ex As Exception
            Throw New Exception("Error occurs when getting the url of new arrivals")
        End Try
    End Function

    Private Shared Sub InsertDefaultCategory(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal newonUrl As String)
        '从数据库读取分类填充到List里面，判断记录是否重复
        Dim queryCategory = From c In efContext.Categories
                          Where c.SiteID = siteId
                          Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next

        Dim cate As New Category()
        cate.Category1 = "newon"
        cate.SiteID = siteId
        cate.LastUpdate = DateTime.Now
        cate.Description = "newon"
        cate.Url = newonUrl
        cate.ParentID = -1
        If JudgeCategory(cate.Category1, cate.SiteID, cate.Url) Then
            efContext.AddToCategories(cate)
        End If
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 从Promotion产品中获得Subject信息
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetSubject(ByVal issueId As Integer, ByVal preSubject As String)
        'Dim subject As String = "[FIRSTNAME] 你好,罗密欧与朱丽叶高端女装今日新品上线"
        Dim subject As String = "[FIRSTNAME] 你好" & "【今日爆款/新品，限时9折】"
        '2013/09/16 added,原因：请在现有的Subject的后面加上第一个产品的名称，设置为早上7点钟开始自动发送
        Dim productName As String = EFHelper.GetFirstProductSubject(issueId, "NE")
        'If Not (productName.Contains("【")) Then
        '    productName = "【" & productName & "】"
        'End If
        subject = subject & productName
        subject = preSubject.Replace("[FIRST_PRODUCT]", productName)

        '------------------------------------------------------------------------------------------------
        helper.InsertIssueSubject(issueId, subject)
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
    Private Shared Sub InsertPromotionItem(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal count As Integer,
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
                'doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-main-pic']/a/img").GetAttributeValue("data-src", "")
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
                'doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-main-pic']/a/img").GetAttributeValue("data-src", "")
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

            For i As Integer = 0 To count - 1
                Dim items As HtmlNode = itemNodes(i)
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                Dim url As String = imgNode.GetAttributeValue("href", "")
                Dim picUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("src", "")
                Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")

                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                Dim product As String = detailNode.SelectSingleNode("a").InnerText.Trim()
                Dim description As String = detailNode.SelectSingleNode("a").InnerText.Trim()

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
                Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                'Dim monthlysales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                '大图Url
                'doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-s310']/a/img").GetAttributeValue("data-src", "")
                Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-main-pic']/a/img").GetAttributeValue("data-src", "")
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
    Private Shared Function GetHtmlDocByUrl(ByVal pageUrl As String) As HtmlDocument
        Try
            Dim request As WebRequest = WebRequest.Create(pageUrl)
            request.Timeout = 30000
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
        Catch ex As Exception
            Throw New Exception("Request time out or bad url")
        End Try

    End Function

    Public Shared Function GetPageContentWithAgent(ByVal url As String) As HtmlDocument
        Dim htmlDoc As HtmlDocument = Nothing
        Try

            htmlDoc = New HtmlDocument()
            Dim httpWebRequest As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            httpWebRequest.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            httpWebRequest.Method = "GET"

            httpWebRequest.Headers.Add("Cookie", "cna=ZOKeDJ8EcEcCAQ6by/xOCWXH;")
            httpWebRequest.Headers.Add("Cookie", "uc1=cookie14=UoW29D54EhZafg%3D%3D;")
            httpWebRequest.Headers.Add("Cookie", "t=afbb4106e530015091a1dfea302587b4;")
            httpWebRequest.Headers.Add("Cookie", "_tb_token_=e36118a65e796;")
            httpWebRequest.Headers.Add("Cookie", "cookie2=1ca32845ef6267e5f5e4bee6a7d62e79;")
            httpWebRequest.Headers.Add("Cookie", "pnm_cku822=001fCJmZk4PGRVHHxtNZnIvYn02YSd8NGIZIw%3D%3D%7CfyJ6ZyByN2ggYXIlYHAkaRg%3D%7CfiB4D15%2BZH9geTp%2FJyN8OjRrKg8PEQFIWF9bakM%3D%7CeSRiYjNhJ384emQ2d2E8fmoheDp9On9oPH1vN3BvKX81ZCxpfDkQ%7CeCVoaEATTRVdGxFHBQ9BHUtSRVwIRUQMXl1SRUsNCktkUA%3D%3D%7CeyR8C0gHRQBBAB5KEgVQFwJdD0kUVAwTRAAKVxcPTANJFlZ%2FBA%3D%3D%7CeiJmeiV2KHMvangudmM6eXk%2BAA%3D%3D;")
            httpWebRequest.Headers.Add("Cookie", "cq=ccp%3D1;")
            httpWebRequest.Headers.Add("Cookie", "isg=472A72370CC2704ED90EEF41B43E86AC;")
            httpWebRequest.Headers.Add("Cookie", "v=0;")
            httpWebRequest.Headers.Add("Cookie", "mt=ci%3D-1_0;")

            httpWebRequest.AllowAutoRedirect = False
            httpWebRequest.Timeout = 30000
            httpWebRequest.Headers.Add("Accept-Language", "zh-CN")

            '2013/12/28,淘宝需要使用Cookie才能获取到Search页面的产品信息
            httpWebRequest.Headers.Add("Cookie", "cna=xQ5NC/YMS0MCAQ6byx3sGY56")

            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
            Dim response As HttpWebResponse = TryCast(httpWebRequest.GetResponse(), HttpWebResponse)
            htmlDoc.Load(response.GetResponseStream(), pageEncoding)
            response.Close()
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try

        Return htmlDoc
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


        Dim cateHots As New Category()
        cateHots.Category1 = "热卖小西装"
        cateHots.SiteID = siteId
        cateHots.LastUpdate = DateTime.Now
        cateHots.Description = "热卖小西装"
        cateHots.Url = "http://rjfashion.taobao.com/search.htm?scid=717123914"
        cateHots.ParentID = -1
        If JudgeCategory(cateHots.Category1, cateHots.SiteID, cateHots.Url) Then
            efContext.AddToCategories(cateHots)
        End If

        Dim cateDouble As New Category()
        cateDouble.Category1 = "超值两件套"
        cateDouble.SiteID = siteId
        cateDouble.LastUpdate = DateTime.Now
        cateDouble.Description = "超值两件套"
        cateDouble.Url = "http://rjfashion.taobao.com/search.htm?scid=459642221"
        cateDouble.ParentID = -1
        If JudgeCategory(cateDouble.Category1, cateDouble.SiteID, cateDouble.Url) Then
            efContext.AddToCategories(cateDouble)
        End If

        efContext.SaveChanges()
    End Sub

    Public Shared Sub ClearTestData(ByVal siteId As Integer)
        Dim cates = From c In efContext.Categories
                    Where c.SiteID = siteId
                    Select c
        Dim proucts = From p In efContext.Products
                      Where p.SiteID = siteId
                      Select p
        For Each category As Category In cates
            For Each product As Product In proucts
                category.Products.Remove(product)
            Next
        Next
        efContext.SaveChanges()


        Dim issues = From i In efContext.Issues
                     Where i.SiteID = siteId
                     Select i
        Dim pro_issues = From pi In efContext.Products_Issue
                         Where pi.SiteId = siteId
                         Select pi

        For Each issue As Issue In issues
            For Each productsIssue As Products_Issue In pro_issues
                issue.Products_Issue.Remove(productsIssue)
            Next
        Next
        efContext.SaveChanges()

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
                              ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal categoryId As Integer, ByVal sales As Integer, ByVal discount As Decimal, ByVal score As Decimal, ByVal totalComment As Integer) As Integer
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

    Public Shared Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product)
            listProducts = helper.GetAsynProductList(categoryUrl, siteId) ' GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount) '读第一页
            Dim count As Integer = listProducts.Count
            If (count < iProIssueCount) Then
                Dim pageNo As Integer = Integer.Parse(categoryUrl.Substring(categoryUrl.Length - 1))
                categoryUrl = categoryUrl.Replace(pageNo.ToString, (pageNo + 1).ToString) '读第二页
                listProducts = helper.GetAsynProductList(categoryUrl, siteId)
                count = listProducts.Count
                If (count < iProIssueCount) Then
                    pageNo = Integer.Parse(categoryUrl.Substring(categoryUrl.Length - 1))
                    categoryUrl = categoryUrl.Replace(pageNo.ToString, (pageNo + 1).ToString) '读第三页
                    listProducts = helper.GetAsynProductList(categoryUrl, siteId)
                End If
            End If
            Dim listProduct As New List(Of Product)
            listProduct = GetProductList(siteId)
            Dim listProductId As New List(Of Integer)
            Dim queryCate = (From m In efContext.Categories
                             Where m.Category1 = categoryName AndAlso m.SiteID = siteId
                             Select m).First()
            Dim categoryId As Integer = queryCate.CategoryID

            For Each li In listProducts
                Dim returnId As Integer = InsertProduct(li, Now, categoryId, listProduct)

                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount * 2) Then
                    Exit For
                End If
            Next

            InsertProductsIssue(siteId, IssueId, "DA", listProductId, iProIssueCount)
        Catch ex As Exception
            EFHelper.LogText(ex.ToString())
        End Try
    End Sub

    


    ''' <summary>
    ''' 将一个product数据插入到Product表中，如果Product表中已存在此条记录，则update，同时不修改此product的CategoryID  Dora 2014-02-11
    ''' </summary>
    ''' <param name="aProductInList"></param>
    ''' <param name="now"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertProduct(ByVal aProductInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal list As List(Of Product)) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = aProductInList.Prodouct
        product.Url = aProductInList.Url
        product.Price = aProductInList.Price
        product.PictureUrl = aProductInList.PictureUrl
        product.LastUpdate = now
        product.Description = aProductInList.Description
        product.SiteID = aProductInList.SiteID
        product.Currency = aProductInList.Currency
        product.PictureAlt = aProductInList.PictureAlt
        product.Discount = aProductInList.Discount
        product.ShipsImg = aProductInList.ShipsImg
        product.FreeShippingImg = aProductInList.FreeShippingImg

        If (JudgeProduct(product.Url, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = aProductInList.Url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.Description = product.Description
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = now
            updateProduct.Discount = product.Discount
            updateProduct.FreeShippingImg = product.FreeShippingImg
            updateProduct.ShipsImg = product.ShipsImg

            '2014/2/21新增，防止一个产品有多个productCategory关系
            Dim updateCategory = updateProduct.Categories

            'Dim queryCate = From p In efContext.Products
            '                Where p.ProdouctID = updateProduct.ProdouctID
            '                Select p
            'Dim cate = queryCate.Single.Categories
            Dim counter As Integer = updateCategory.Count
            For i As Integer = 0 To counter - 1
                updateCategory(0).Products.Remove(updateProduct)
            Next
            efContext.SaveChanges()

            If Not updateCategory.Contains(category) Then
                category.Products.Add(updateProduct)
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
    End Function

    ''' <summary>
    ''' 判断即将插入的Productd的URL是否在数据库中已经存在，如果存在，返回false Dora 2014-02-11
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeProduct(ByVal url As String, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url.Trim() = url.Trim()) Then
                Return False
            End If
        Next
        Return True
    End Function


    Private Shared Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
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

#Region "Template"
    ''' <summary>
    ''' 添加模板
    ''' </summary>
    ''' <param name="t"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddTempalte(ByVal t As Template)
        Dim temp = Enumerable.[Single](efContext.Templates, Function(m) m.Tid = t.Tid)
        If temp Is Nothing Then
            efContext.AddToTemplates(t)
        Else
            temp.Contents = t.Contents
        End If
        efContext.SaveChanges()
        Return True
    End Function


    Public Shared Function UpdateTemplate(ByVal tid As Integer, ByVal content As String)
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
    Public Shared Function GetRnd(ByVal min As Long, ByVal max As Long) As Integer
        Randomize() '没有这个 产生的数会一样
        GetRnd = Rnd() * (max - min + 1) + min
    End Function

    ''' <summary>
    ''' 从字符串中提取浮点数
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveStringCode(ByVal str As String)
        Dim reg As New Regex("(-?\d+)(\.\d+)?")
        Dim m As Match = reg.Match(str)
        Return m.Groups(0).Value()
    End Function

    Public Shared Function GetTotalComment(ByVal str As String)
        Dim rg As New Regex("(?<=("":))[.\s\S]*?(?=(,""))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Public Shared Function GetTotalRate(ByVal str As String)
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
    Public Shared Function GetJsonDataByParameter(ByVal jText As String, ByVal name As String, ByVal param As String)
        Dim jsonText As String = "{""" + jText + "}"
        Dim jo As JObject = DirectCast(JsonConvert.DeserializeObject(jsonText), JObject)
        Dim r As String = jo(name)(param).ToString()
        Return r
    End Function

    Public Shared Function GetJsonFromTaobao(ByVal pageUrl As String) As String
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

    Public Shared Function GetValue(ByVal str As String, ByVal s As String, ByVal e As String) As String
        Dim rg As New Regex("(?<=(" & s & "))[.\s\S]*?(?=(" & e & "))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Public Shared Function GetPromotionPriceJsonFromTaobao(ByVal pageUrl As String, ByVal apiUrl As String) As String
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
