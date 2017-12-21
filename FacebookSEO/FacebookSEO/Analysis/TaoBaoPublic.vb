Imports TaobaoAuto
Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json

Public Class TaoBaoPublic
    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal spreadLogin As String, _
                         ByVal appId As String, ByVal url As String, ByVal mylist As String, ByVal nowTime As DateTime)
        Dim helper As New EFHelper
        Dim listCategory As New List(Of Category)
        Dim listExistCategory As List(Of String) = EFHelper.GetListCateUrl(siteId)
        Dim cate1 As New Category
        cate1.Category1 = "新品"
        cate1.Url = url & "/search.htm?" & "orderType=" & TaobaoSortedType.newOn_asc.ToString()
        cate1.SiteID = siteId
        cate1.LastUpdate = nowTime
        If Not (listExistCategory.Contains(cate1.Url)) Then
            listCategory.Add(cate1)
        End If
        Dim cate2 As New Category
        cate2.Category1 = "热销"
        cate2.Url = url & "/search.htm?" & "orderType=" & TaobaoSortedType.hotsell_desc.ToString()
        cate2.SiteID = siteId
        cate2.LastUpdate = nowTime
        If Not (listExistCategory.Contains(cate2.Url)) Then
            listCategory.Add(cate2)
        End If
        If (listCategory.Count > 0) Then
            helper.InsertListCategory(listCategory)
        Else
            listCategory.Add(cate1)
            listCategory.Add(cate2)
        End If
        For Each cate As Category In listCategory
            GetProductsByUrl(cate.Url, 9, siteId, planType, IssueID)
        Next
        If Not (String.IsNullOrEmpty(mylist)) Then
            If (mylist.Contains(",") OrElse mylist.Contains("，")) Then
                Dim arrList As String()
                Try
                    arrList = mylist.Split(",")
                Catch ex As Exception
                    arrList = mylist.Split("，")
                End Try
                helper.InsertContactList(IssueID, arrList, "waiting")
            End If
        End If


        'Dim cates As New List(Of TaobaoAuto.Category)
        'Dim cate1 As New TaobaoAuto.Category
        'cate1.CateUrl = url.Trim() & "/search.htm?"
        'cate1.SortedType = TaobaoSortedType.newOn_asc.ToString()
        'cate1.CateName = "新品"
        'cates.Add(cate1)
        'Dim cate2 As New TaobaoAuto.Category
        'cate2.CateName = "热销"
        'cate2.CateUrl = url.Trim() & "/search.htm?"
        'cate2.SortedType = TaobaoSortedType.hotsell_desc.ToString()
        'cates.Add(cate2)
    End Sub

    ''' <summary>
    ''' 插入N件产品到指定的类别下
    ''' </summary>
    Private Shared Sub GetProductsByUrl(ByVal curl As String, ByVal count As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal issueId As Integer)
        Try
            'TODO 根据出入的参数生成url，获得产品信息
            Dim doc As HtmlDocument = GetHtmlDocByUrl(curl)  'RjNews.GetPageContentWithAgent(curl)
            Dim productlist As New List(Of Product)
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

                    doc = GetHtmlDocByUrl(url, curl)
                    Dim innerHtml As String = doc.DocumentNode.InnerHtml

                    '打分 &评论
                    'Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                    'Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                    'Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                    'Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                    'Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                    'Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

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

                    Dim prod As New Product
                    prod.Prodouct = product
                    prod.Url = url
                    prod.SiteID = siteId
                    prod.Description = description
                    prod.PictureUrl = pictureUrl
                    prod.PictureUrl = pictureAlt
                    prod.Price = price
                    prod.Discount = discount
                    'prod.Rating = rating
                    'prod.TotalComment = comment
                    prod.Sales = sales
                    prod.Currency = currency

                    productlist.Add(prod)
                Next
            End If

            divNodes = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/ul[@class='shop-list']/li/div")
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

                    Dim prod As New Product
                    prod.Prodouct = product
                    prod.Url = url
                    prod.SiteID = siteId
                    prod.Description = description
                    prod.PictureUrl = pictureUrl
                    prod.PictureAlt = pictureAlt
                    prod.Price = price
                    prod.Discount = discount
                    'prod. = rating
                    'prod.TotalComment = comment
                    prod.Sales = sales
                    prod.Currency = currency
                    productlist.Add(prod)
                Next
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
                    'Dim sales As Long = Long.Parse(attrNode.SelectSingleNode("div/span[@class='sale-num']").InnerText)
                    Dim lastUpdate As DateTime = nowTime
                    Dim currency As String = "CNY"

                    doc = GetHtmlDocByUrl(url, curl)
                    Dim innerHtml As String = doc.DocumentNode.InnerHtml

                    '打分 & 评论
                    'Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                    'Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                    'Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                    'Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                    'Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                    'Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))

                    'Dim rateCountApi As String = GetValue(innerHtml, "rateCounterApi:""", """") + "&callback=DT.mods.SKU.CountCenter.setReviewCount"
                    'Dim rateCount As String = GetJsonFromTaobao(rateCountApi)
                    'Dim rateParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(0)
                    'Dim commentParamString As String = GetValue(rateCountApi, "keys=", "&callback").Split(",")(1)
                    'Dim rating As Decimal = Decimal.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", rateParamString))
                    'Dim comment As Integer = Integer.Parse(GetJsonDataByParameter(rateCount.Replace("(", """:").Replace(");", ""), "DT.mods.SKU.CountCenter.setReviewCount", commentParamString))


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

                    'itemInfoApi 获得月销数量Api
                    'Dim itemInfoApi As String = GetValue(innerHtml, """apiItemInfo"":""", """")
                    'Dim itemInfo As String = GetJsonFromTaobao(itemInfoApi)
                    'Dim monthlysales As Long = Long.Parse(GetValue(itemInfo, "quanity:", ","))

                    '大图Url
                    Dim pictureUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-main-pic']/a/img").GetAttributeValue("data-src", "")
                    If String.IsNullOrEmpty(pictureUrl) Or pictureUrl.Contains(".gif") Then
                        pictureUrl = doc.DocumentNode.SelectSingleNode("//div[@class='tb-booth tb-pic tb-main-pic']/a/img").GetAttributeValue("src", "")
                        If String.IsNullOrEmpty(pictureUrl) Then
                            pictureUrl = picUrl
                        End If
                    End If

                    Dim prod As New Product
                    prod.Prodouct = product
                    prod.Url = url
                    prod.SiteID = siteId
                    prod.Description = description
                    prod.PictureUrl = pictureUrl
                    prod.PictureAlt = pictureAlt
                    prod.Price = price
                    prod.Discount = discount
                    'prod.Rating = rating
                    'prod.TotalComment = comment
                    prod.Sales = sales
                    prod.Currency = currency
                    productlist.Add(prod)
                Next
            End If
            'Return productlist
            Dim helper As New EFHelper
            Dim subject As String = "[FIRSTNAME] 你好"
            Dim counter As Integer = 0
            For Each prod As Product In productlist
                If (prod.Price <= 10) Then
                    Continue For
                End If
                If Not (helper.IsProductSent2(siteId, prod.Url, nowTime.ToString, nowTime.AddDays(-7), planType)) Then '上一期还未发送
                    If (counter < 1) Then
                        If (prod.Prodouct.Contains("【")) Then
                            subject = subject & prod.Prodouct
                        Else
                            subject = subject & "【" & prod.Prodouct & "】"
                        End If
                        counter = counter + 1
                    End If
                    Dim productId As Integer = helper.InsertOrUpdateProduct(prod, curl, siteId, nowTime)
                    Try
                        helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                    Catch ex As Exception
                        'Ignore
                    End Try

                End If
            Next
            helper.InsertIssueSubject(issueId, subject)
        Catch ex As Exception
            'Return Nothing
            Throw New Exception(ex.ToString())
        End Try
    End Sub
#Region "Func"
    ''' <summary>
    ''' 通过Url返回指定的HtmlDocument对象
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetHtmlDocByUrl(ByVal pageUrl As String) As HtmlDocument
        Try
            If IsUrl(pageUrl) Then
                Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
                request.Timeout = 300000
                request.Headers.Add("Accept-Language", "zh-CN")
                request.Referer = pageUrl
                request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
                request.Method = "GET"
                request.AllowAutoRedirect = False
                'request.Headers.Add("Cookie", "v=0; uc3=nk2=GcA1v%2ByEHrqEsPjN&id2=UoLZWGE5rRrT&vt3=F8dATH%2BfWolxFq60EV0%3D&lg2=UIHiLt3xD8xYTw%3D%3D; existShop=MTQwMTg1NDM3Mw%3D%3D; lgc=zhenzhen2510; tracknick=zhenzhen2510; sg=041; skt=a1983a6470c0d6a9; publishItemObj=Ng%3D%3D; _cc_=U%2BGCWk%2F7og%3D%3D; tg=0; _l_g_=Ug%3D%3D; cna=pocVDM+cJQQCAbcL/F9K2+oR; pnm_cku822=141fCJmZk4PGRVHHxtEb3EtbnA3YSd%2FN2EaIA%3D%3D%7CfyJ6ZyBzMGEkZHEvb38hZhc%3D%7CfiB4D15%2BZH9geTp%2FJyN8OTdoKQwMEgJLW1xYaUA%3D%7CeSRiYjNhJ34%2Fc2AzdGszcW8vdjRzM3NtPHhuNXFoI3A0bCprdCBgZFU%3D%7CeCVoaEASTBxUFx1LCAJMXRAIUEFJSltVGl4AX00EC1gIPBE%3D%7CeyR8C0gHRQBBBxdGHgxSEQlWBEAZWQEeSQ0HWhgERwhCHV10Dw%3D%3D%7CeiJmeiV2KHMvangudmM6eXk%2BAA%3D%3D; cookie2=6476a3c2bc1ec91ec66145b14bd2ccb9; cookie1=W88eHkUrJr%2FBkifCvPNe1qpqnj9mHFcyg3O3UXXtKOc%3D; unb=135483794; mt=ci=0_1&cyk=0_0; t=8603225f41301b09103a2a66c4f60480; _nk_=zhenzhen2510; cookie17=UoLZWGE5rRrT; uc1=lltime=1401849973&cookie14=UoW3vseDyNQnXw%3D%3D&existShop=true&cookie16=UIHiLt3xCS3yM2h4eKHS9lpEOw%3D%3D&cookie21=V32FPkk%2Fhodroid0QSjisQ%3D%3D&tag=2&cookie15=U%2BGCWk%2F75gdr5Q%3D%3D")
                'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
                Dim response As WebResponse = request.GetResponse()
                'WebRequest.GetResponse方法，返回对 Internet 请求的响应
                Dim resStream As Stream = response.GetResponseStream()
                'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
                Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
                Dim document As New HtmlDocument()
                document.Load(resStream, pageEncoding)
                Return document
            Else
                Throw New Exception("Invalid url")
            End If

        Catch ex As Exception
            Try
                Dim random As New Random()
                Dim timeSpam As Integer = random.Next(120000, 300000)
                System.Threading.Thread.Sleep(timeSpam)
                Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
                request.Timeout = 300000
                request.Headers.Add("Accept-Language", "zh-CN")
                request.Referer = pageUrl
                request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
                request.Method = "GET"
                request.AllowAutoRedirect = False
                'request.Headers.Add("Cookie", "v=0; uc3=nk2=GcA1v%2ByEHrqEsPjN&id2=UoLZWGE5rRrT&vt3=F8dATH%2BfWolxFq60EV0%3D&lg2=UIHiLt3xD8xYTw%3D%3D; existShop=MTQwMTg1NDM3Mw%3D%3D; lgc=zhenzhen2510; tracknick=zhenzhen2510; sg=041; skt=a1983a6470c0d6a9; publishItemObj=Ng%3D%3D; _cc_=U%2BGCWk%2F7og%3D%3D; tg=0; _l_g_=Ug%3D%3D; cna=pocVDM+cJQQCAbcL/F9K2+oR; pnm_cku822=141fCJmZk4PGRVHHxtEb3EtbnA3YSd%2FN2EaIA%3D%3D%7CfyJ6ZyBzMGEkZHEvb38hZhc%3D%7CfiB4D15%2BZH9geTp%2FJyN8OTdoKQwMEgJLW1xYaUA%3D%7CeSRiYjNhJ34%2Fc2AzdGszcW8vdjRzM3NtPHhuNXFoI3A0bCprdCBgZFU%3D%7CeCVoaEASTBxUFx1LCAJMXRAIUEFJSltVGl4AX00EC1gIPBE%3D%7CeyR8C0gHRQBBBxdGHgxSEQlWBEAZWQEeSQ0HWhgERwhCHV10Dw%3D%3D%7CeiJmeiV2KHMvangudmM6eXk%2BAA%3D%3D; cookie2=6476a3c2bc1ec91ec66145b14bd2ccb9; cookie1=W88eHkUrJr%2FBkifCvPNe1qpqnj9mHFcyg3O3UXXtKOc%3D; unb=135483794; mt=ci=0_1&cyk=0_0; t=8603225f41301b09103a2a66c4f60480; _nk_=zhenzhen2510; cookie17=UoLZWGE5rRrT; uc1=lltime=1401849973&cookie14=UoW3vseDyNQnXw%3D%3D&existShop=true&cookie16=UIHiLt3xCS3yM2h4eKHS9lpEOw%3D%3D&cookie21=V32FPkk%2Fhodroid0QSjisQ%3D%3D&tag=2&cookie15=U%2BGCWk%2F75gdr5Q%3D%3D")
                'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
                Dim response As WebResponse = request.GetResponse()
                'WebRequest.GetResponse方法，返回对 Internet 请求的响应
                Dim resStream As Stream = response.GetResponseStream()
                'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
                Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
                Dim document As New HtmlDocument()
                document.Load(resStream, pageEncoding)
                Return document
            Catch ex1 As Exception
                Throw New Exception("Request time out " + ex1.ToString())
            End Try
        End Try

    End Function

    Private Shared Function GetHtmlDocByUrl(ByVal pageUrl As String, ByVal referer As String) As HtmlDocument
        Try
            If IsUrl(pageUrl) Then
                Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
                request.Referer = referer
                request.Timeout = 300000
                request.Headers.Add("Accept-Language", "zh-CN")
                request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
                request.Method = "GET"
                request.AllowAutoRedirect = False
                'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
                Dim response As WebResponse = request.GetResponse()
                'WebRequest.GetResponse方法，返回对 Internet 请求的响应
                Dim resStream As Stream = response.GetResponseStream()
                'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
                Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
                Dim document As New HtmlDocument()
                document.Load(resStream, pageEncoding)
                Return document
            Else
                Throw New Exception("Invalid url")
            End If
        Catch ex As Exception
            Try
                Dim random As New Random()
                Dim timeSpam As Integer = random.Next(120000, 300000)
                System.Threading.Thread.Sleep(timeSpam)
                Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
                request.Referer = referer
                request.Timeout = 300000
                request.Headers.Add("Accept-Language", "zh-CN")
                request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
                request.Method = "GET"
                request.AllowAutoRedirect = False
                'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
                Dim response As WebResponse = request.GetResponse()
                'WebRequest.GetResponse方法，返回对 Internet 请求的响应
                Dim resStream As Stream = response.GetResponseStream()
                'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
                Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
                Dim document As New HtmlDocument()
                document.Load(resStream, pageEncoding)
                Return document
            Catch ex1 As Exception
                Throw New Exception("Request time out " + ex1.ToString())
            End Try
        End Try

    End Function


    Private Shared Function IsUrl(ByVal str_url As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(str_url, "http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?")
    End Function

    ''' <summary>
    ''' 获取随机数
    ''' </summary>
    ''' <param name="min"></param>
    ''' <param name="max"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetRnd(ByVal min As Long, ByVal max As Long) As Integer
        Randomize() '没有这个 产生的数会一样
        GetRnd = Rnd() * (max - min + 1) + min
    End Function

    ''' <summary>
    ''' 从字符串中提取浮点数
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function RemoveStringCode(ByVal str As String)
        Dim reg As New Regex("(-?\d+)(\.\d+)?")
        Dim m As Match = reg.Match(str)
        Return m.Groups(0).Value()
    End Function

    Private Shared Function GetTotalComment(ByVal str As String)
        Dim rg As New Regex("(?<=("":))[.\s\S]*?(?=(,""))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Private Shared Function GetTotalRate(ByVal str As String)
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
    Private Shared Function GetJsonDataByParameter(ByVal jText As String, ByVal name As String, ByVal param As String)
        Dim jsonText As String = "{""" + jText + "}"
        Dim jo As JObject = DirectCast(JsonConvert.DeserializeObject(jsonText), JObject)
        Dim r As String = jo(name)(param).ToString()
        Return r
    End Function

    Private Shared Function GetJsonFromTaobao(ByVal pageUrl As String) As String
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

    Private Shared Function GetValue(ByVal str As String, ByVal s As String, ByVal e As String) As String
        Dim rg As New Regex("(?<=(" & s & "))[.\s\S]*?(?=(" & e & "))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function

    Private Shared Function GetPromotionPriceJsonFromTaobao(ByVal pageUrl As String, ByVal apiUrl As String) As String
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
