Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Tmall
    Private Shared helper As New EFHelper()
    Private Shared efContext As New EmailAlerterEntities()
    Private Shared listProductUrl As New List(Of String)

    ''' <summary>
    ''' 程序入口
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="Url"></param>
    ''' <remarks></remarks>
    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal splitContactCount As Integer, ByVal spreadLogin As String, _
                            ByVal appId As String, ByVal url As String, ByVal categories As String, ByVal siteName As String, ByVal dllType As String,
                            ByVal iProCount As Integer, bannerUrl As String, ByVal bannerRegex As String, ByVal bannerIndex As Integer, ByVal preSubject As String)
        'InsertCategory(siteId, url, dllType)
        Dim category() As String = categories.Split("^")

        If (bannerIndex >= 0 AndAlso Not String.IsNullOrEmpty(bannerUrl)) Then
            helper.GetBanner(siteId, IssueID, category(0), planType, bannerUrl, bannerRegex, bannerIndex)
            '成功则向Ads_issue中插入一条记录
        End If


        InsertProductData(category, siteId, IssueID, url, planType, dllType, iProCount)
        GetSubject(IssueID, siteId, siteName, planType, preSubject, category(0))
        listProductUrl.Clear() '因为这个为shared属性，需求清空
        If (planType.Contains("HP")) Then
            Dim cate As String = category(0).Trim
            Dim contactlistCount As Integer
            Dim saveListName1 As String = siteName & cate & Now.ToString("yyyyMMdd")
            contactlistCount = helper.CreateContactList(siteId, IssueID, planType.Trim, cate, saveListName1, 15, ChooseStrategy.Favorite, "draft", spreadLogin, appId)

            'debug# set  contactlistCount = 1
            'contactlistCount = 1
            If (contactlistCount > 0) Then
                helper.InsertContactList(IssueID, saveListName1, "draft")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 从Promotion产品中获得Subject信息
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal siteName As String, ByVal planType As String,
                                  ByVal preSubject As String, ByVal cateName As String)
        Dim subject As String = ""
        Dim querySubject As String = (From p In efContext.Products
                             Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                             Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                             Select p.Prodouct).FirstOrDefault()
        Dim query = From i In efContext.Issues
                        Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                        Select i
        If (String.IsNullOrEmpty(preSubject)) Then
            If (planType.Contains("HO") Or planType.Contains("HA")) Then

                If Not (String.IsNullOrEmpty(querySubject)) Then
                    subject = "[FIRSTNAME] 你好," & siteName & "为你带来： " & querySubject & "(AD)"
                Else
                    subject = "[FIRSTNAME] 你好,更多惊喜尽在" & siteName
                End If
            ElseIf (planType.Contains("HP")) Then

                subject = "[FIRSTNAME] 您的" & siteName & "专属资讯（第" & (query.Count + 1).ToString.PadLeft(2, "0") & "）期"
            End If
        Else
            If Not (String.IsNullOrEmpty(querySubject)) Then
                subject = preSubject.Replace("[FIRST_PRODUCT]", querySubject.Trim)
            Else
                subject = preSubject.Replace("[FIRST_PRODUCT]", "")
            End If
            subject = subject.Replace("[VOL_NUMBER]", (query.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
        End If

        helper.InsertIssueSubject(issueId, subject)
    End Sub

    ''' <summary>
    ''' 插入所有产品信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="ShopUrl"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertProductData(ByRef category As String(), ByVal siteId As Integer, ByVal issueId As Integer, ByVal ShopUrl As String, ByVal planType As String,
                                        ByVal dllType As String, ByVal iProductCount As Integer)

        '单个分类每期邮件中获取产品数
        If (iProductCount > 0) Then
            For Each cateName In category
                Dim cate As Category = EFHelper.GetCategory(siteId, cateName)
                helper.FetchProduct(cate.Url, cate.Category1, "CA", planType, iProductCount, DateTime.Now.AddDays(-30), DateTime.Now, siteId, issueId)
            Next
        End If

    End Sub



    Public Shared Function GetCategoryDict(ByRef category As String(), ByVal siteId As Integer, ByVal IssueId As Integer, ByVal planType As String,
                                    ByVal dllType As String)

        Dim CategoryDict As Dictionary(Of String, Category) = New Dictionary(Of String, Category)

        '获取分类

        For Each cateName In category
            Dim cate As Category = CategoryDAL.GetCategory(siteId, cateName)
            If Not CategoryDict.Keys.Contains(cateName) Then
                CategoryDict.Add(cateName, cate)
            End If
        Next

        Return CategoryDict

    End Function



    Public Shared Function GetProdcutDict(ByVal CategoryDict As Dictionary(Of String, Category))

        Dim ProductDict As Dictionary(Of String, List(Of Product)) = New Dictionary(Of String, List(Of Product))
        Dim ProductList As List(Of Product) = New List(Of Product)


        For Each cateName In CategoryDict.Keys
            Dim category As Category = CategoryDict(cateName)
            'to do
            'ProductList = EFHelper.GetAsynProductList(category.Url, category.SiteID)
            'If Not ProductDict.Keys.Contains(cateName) Then
            '    ProductDict.Add(cateName, ProductList)
            'End If
        Next

        Return ProductDict

    End Function






    Private Shared Function GetListProductsTmall(ByVal cateUrl As String, ByVal siteId As Integer, ByVal cateName As String, ByVal iProductCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim index As Integer
        Dim nowTime As DateTime = DateTime.Now
        Dim doc As HtmlDocument = EFHelper.GetHtmlDocByUrlTmall(cateUrl)
        Dim divNodes As HtmlNodeCollection

        divNodes = doc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")
        For Each productLine In divNodes
            If (productLine.GetAttributeValue("class", "").ToString.ToLower() = "pagination") Then
                Exit For
            End If
            Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
            If Not (productDls Is Nothing) Then
                For Each dl In productDls
                    Dim myProduct As New Product
                    Dim productName As String = dl.SelectSingleNode("dd[@class='detail']/a").InnerText
                    If productName.Contains("运费") Or productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
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
                    myProduct.Url = EFHelper.AddHttpForAli(myProduct.Url)
                    If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-33), DateTime.Now)) Then
                        myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                        myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                        myProduct.PictureUrl = EFHelper.AddHttpForAli(myProduct.PictureUrl)
                        myProduct.Description = productName
                        myProduct.LastUpdate = nowTime
                        myProduct.SiteID = siteId
                        listProducts.Add(myProduct)
                    End If
                Next
            End If
            If (listProducts.Count >= iProductCount + 1) Then
                Exit For
            End If
        Next
        Return listProducts
    End Function

    Private Shared Function GetListProductsTaobao(ByVal cateUrl As String, ByVal siteId As Integer, ByVal cateName As String, ByVal iProductCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim index As Integer
        Dim nowTime As DateTime = DateTime.Now
        Dim doc As HtmlDocument = EFHelper.GetHtmlDocByUrlTmall(cateUrl)

        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div/dl")
        If itemNodes IsNot Nothing Then
            For i As Integer = 0 To itemNodes.Count - 1
                Dim myProduct As New Product
                Dim items As HtmlNode = itemNodes(i)
                Dim detailNode As HtmlNode = items.SelectSingleNode("dd[@class='detail']")
                myProduct.Prodouct = detailNode.SelectSingleNode("a").InnerText.Trim()
                myProduct.Description = myProduct.Prodouct
                '过滤掉运费差价专拍
                If myProduct.Description.Contains("差额") Or myProduct.Description.Contains("专拍") Or myProduct.Description.Contains("补拍") Or myProduct.Description.Contains("差价") _
                    Or myProduct.Description.Contains("補拍") Or myProduct.Description.Contains("鏈接") Or myProduct.Description.Contains("專拍") Or _
                    myProduct.Description.Contains("運費") Or myProduct.Description.Contains("链接") Then
                    Continue For
                End If
                Dim imgNode As HtmlNode = items.SelectSingleNode("dt[@class='photo']/a")
                myProduct.Url = imgNode.GetAttributeValue("href", "")
                If (myProduct.Url.Contains("&")) Then
                    index = myProduct.Url.IndexOf("&")
                    myProduct.Url = myProduct.Url.Remove(index)
                End If
                If (listProductUrl.Contains(myProduct.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(myProduct.Url)
                End If
                myProduct.Url = EFHelper.AddHttpForAli(myProduct.Url)
                If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-33), DateTime.Now)) Then
                    myProduct.PictureUrl = imgNode.SelectSingleNode("img").GetAttributeValue("data-ks-lazyload", "")
                    If (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                        myProduct.PictureUrl = imgNode.SelectSingleNode("img").GetAttributeValue("src", "")
                    End If
                    myProduct.PictureUrl = EFHelper.AddHttpForAli(myProduct.PictureUrl)
                    myProduct.PictureAlt = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")
                    Dim attrNode As HtmlNode = detailNode.SelectSingleNode("div[@class='attribute']")
                    Dim discount As Decimal = Decimal.Parse(attrNode.SelectSingleNode("div/span[@class='c-price']").InnerText)
                    Dim priceNode As HtmlNode = attrNode.SelectSingleNode("div/span[@class='s-price']")
                    Dim price As Decimal
                    If priceNode IsNot Nothing Then
                        price = Decimal.Parse(priceNode.InnerText)
                    Else
                        price = discount
                    End If
                    myProduct.Price = price
                    myProduct.Discount = discount
                    myProduct.Currency = "CNY"
                    myProduct.LastUpdate = nowTime
                    myProduct.SiteID = siteId
                    listProducts.Add(myProduct)
                End If
                If (listProducts.Count >= iProductCount + 1) Then
                    Exit For
                End If
            Next
        End If
        Return listProducts
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

    ' ''' <summary>
    ' ''' 通过Url返回HtmlDocument对象
    ' ''' </summary>
    ' ''' <param name="pageUrl"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Shared Function GetHtmlDocByUrl(pageUrl As String) As HtmlDocument
    '    Dim request As WebRequest = WebRequest.Create(pageUrl)
    '    request.Timeout = 300000
    '    request.Headers.Add("Accept-Language", "zh-CN")

    '    request.Method = "GET"
    '    request.Headers.Add("Cookie", "cna=ZOKeDJ8EcEcCAQ6by/xOCWXH;")
    '    request.Headers.Add("Cookie", "__tmall_fp_ab=__804b;")
    '    request.Headers.Add("Cookie", "t=afbb4106e530015091a1dfea302587b4;")
    '    request.Headers.Add("Cookie", "_tb_token_=QHXfB3OmZ82L;")
    '    request.Headers.Add("Cookie", "cookie2=a863eb3585391a2fd690d1d93cc7b774;")
    '    request.Headers.Add("Cookie", "pnm_cku822=;")
    '    request.Headers.Add("Cookie", "cq=ccp%3D1;")
    '    request.Headers.Add("Cookie", "isg=821C375C044161990D73479E11B56BBF;")
    '    'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
    '    Dim response As WebResponse = request.GetResponse()
    '    'WebRequest.GetResponse方法，返回对 Internet 请求的响应
    '    Dim resStream As Stream = response.GetResponseStream()
    '    'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
    '    Dim pageEncoding As Encoding = Encoding.GetEncoding("GB18030")

    '    Dim document As New HtmlDocument()
    '    document.Load(resStream, pageEncoding)
    '    Return document
    'End Function

    ''' <summary>
    ''' 插入所有分类
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="ShopUrl"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertCategory(ByVal siteId As Integer, ByVal ShopUrl As String, ByVal dllType As String)
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
        cate.Description = "热销"
        cate.Url = ShopUrl + "/search.htm?orderType=hotsell_desc"
        cate.ParentID = -1
        If JudgeCategory(cate.Category1, cate.SiteID, cate.Url, listCategory) Then
            efContext.AddToCategories(cate)
        End If

        Dim cateNewon As New Category()
        cateNewon.Category1 = "newon"
        cateNewon.SiteID = siteId
        cateNewon.LastUpdate = DateTime.Now
        cateNewon.Description = "新品"
        cateNewon.Url = ShopUrl + "/search.htm?orderType=newOn_desc"
        cateNewon.ParentID = -1
        If JudgeCategory(cateNewon.Category1, cateNewon.SiteID, cateNewon.Url, listCategory) Then
            efContext.AddToCategories(cateNewon)
        End If

        Dim cateHotKeep As New Category()
        cateHotKeep.Category1 = "hotkeep"
        cateHotKeep.SiteID = siteId
        cateHotKeep.LastUpdate = DateTime.Now
        cateHotKeep.Description = "收藏"
        cateHotKeep.Url = ShopUrl + "/search.htm?orderType=hotkeep_desc"
        cateHotKeep.ParentID = -1
        If JudgeCategory(cateHotKeep.Category1, cateHotKeep.SiteID, cateHotKeep.Url, listCategory) Then
            efContext.AddToCategories(cateHotKeep)
        End If

        Dim cateRenqi As New Category()
        cateRenqi.Category1 = "popular"
        cateRenqi.SiteID = siteId
        cateRenqi.LastUpdate = DateTime.Now
        cateRenqi.Description = "口碑"
        cateRenqi.Url = ShopUrl + "/search.htm?orderType=koubei"
        If (dllType.ToLower.Trim = "taobao") Then 'taobao店中为coefp_desc
            cateRenqi.Url = ShopUrl + "/search.htm?orderType=coefp_desc"
        End If
        cateRenqi.ParentID = -1
        If JudgeCategory(cateRenqi.Category1, cateRenqi.SiteID, cateRenqi.Url, listCategory) Then
            efContext.AddToCategories(cateRenqi)
        End If

        Dim cateDeals As New Category()
        cateDeals.Category1 = "price"
        cateDeals.SiteID = siteId
        cateDeals.LastUpdate = DateTime.Now
        cateDeals.Description = "价格"
        cateDeals.Url = ShopUrl + "/search.htm?orderType=price"
        cateDeals.ParentID = -1
        If JudgeCategory(cateDeals.Category1, cateDeals.SiteID, cateDeals.Url, listCategory) Then
            efContext.AddToCategories(cateDeals)
        End If
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
    Private Shared Function JudgeCategory(ByVal category As String, ByVal siteId As Integer, ByVal Url As String, ByVal list As List(Of Category)) As Boolean
        For Each li In list
            If (li.Url = Url) Then
                Return False
            End If
        Next
        Return True
    End Function







    ''' <summary>
    ''' 添加模板
    ''' </summary>
    ''' <param name="t"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddTempalte(t As Template)
        Dim temp = Enumerable.[Single](efContext.Templates, Function(m) m.Tid = t.Tid)
        If temp Is Nothing Then
            efContext.AddToTemplates(t)
        Else
            temp.Contents = t.Contents
        End If
        efContext.SaveChanges()
        Return True
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
        Dim jsonText As String = "{" + jText + "}"
        Dim jo As JObject = DirectCast(JsonConvert.DeserializeObject(jsonText), JObject)
        Dim r As String = jo(name)(param).ToString()
        Return r
    End Function

    ''' <summary>
    ''' Return Iteminof(Json Format) by Url
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="apiUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetItemInfoJsonFromTmall(pageUrl As String, apiUrl As String) As String
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(apiUrl)
        hwr.Referer = pageUrl
        hwr.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:19.0) Gecko/20100101 Firefox/19.0"
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

    ''' <summary>
    ''' Return RateInfo(Json Format) By Url
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRateJsonFromTmall(pageUrl)
        Dim hwr As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        Dim response As HttpWebResponse = hwr.GetResponse()
        Dim resStream As StreamReader = New StreamReader(response.GetResponseStream(), Encoding.Default)
        Dim result As String = resStream.ReadToEnd()
        resStream.Close()
        Return result
    End Function

    ''' <summary>
    ''' find specific string by Regex
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetValue(str As String, s As String, e As String) As String
        Dim rg As New Regex("(?<=(" & s & "))[.\s\S]*?(?=(" & e & "))", RegexOptions.Multiline Or RegexOptions.Singleline)
        Return rg.Match(str).Value
    End Function


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


    Public Shared Function RemoveStringCode(str As String)
        Dim reg As New Regex("(-?\d+)(\.\d+)?")
        Dim m As Match = reg.Match(str)
        Return m.Groups(0).Value()
    End Function

    Public Shared Function ShowTemplate(tid As Integer)
        Dim queryTemplate = From t In efContext.Templates Where t.Tid = tid Select t.Contents
        Return queryTemplate.Single().ToString()
    End Function


End Class
