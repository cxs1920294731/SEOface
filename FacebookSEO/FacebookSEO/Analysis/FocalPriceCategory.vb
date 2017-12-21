Imports HtmlAgilityPack
Imports System.Text.RegularExpressions

Public Class FocalPriceCategory

    Private Shared listNowProductUrl As New List(Of String)

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                         ByVal splitContactCount As Integer, ByVal spreadLogin As String, _
                         ByVal appId As String, ByVal url As String, ByVal nowTime As DateTime, ByVal categorys As String)
        Dim listNowProductUrl As New List(Of String)
        Dim helper As New EFHelper
        Dim lastUpdate As DateTime = nowTime
        Dim pageUrl As String = "http://es.focalprice.com/"
        'Dim doc As HtmlDocument = helper.GetHtmlDocument2(pageUrl, 120000)
        'FocalPrice.GetBanner(siteId, doc, IssueID, nowTime)
        GetSortedCateProducts(siteId, IssueID, categorys, lastUpdate, planType)
        CreateContactList(siteId, IssueID, categorys, nowTime, planType, spreadLogin, appId)
        Dim subject As String = EFHelper.GetFirstProductSubject(IssueID)
        helper.InsertIssueSubject(IssueID, subject)
        'GetCategoryProduct(siteId, IssueID, lastUpdate)
    End Sub

    ''' <summary>
    ''' 获取用户点击了的分类和该分类关联分类的产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueID"></param>
    ''' <param name="categorys"></param>
    ''' <param name="lastUpdate"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetSortedCateProducts(ByVal siteId As Integer, ByVal IssueID As Integer, ByVal categorys As String, _
                                                 ByVal lastUpdate As DateTime, ByVal planType As String)
        Dim helper As New EFHelper
        'If (categorys.Contains("^")) Then
        '    Dim splitCategory As String() = categorys.Split("^")
        'End If
        Dim arrCategory() As String = New String() {"Acceosrios para Apple", "Celulares", "Tablet PCs", "Electrónica de Consumo", "Relojes y Joyas", "Juguetes, Hobbies & Games"}  '获取固定分类的产品，只是分类产品的排列顺序不同
        Dim listCategory As List(Of Category) = EFHelper.GetListCategorys(siteId, arrCategory)
        Dim cateCounter As Integer = 0
        For Each cate In listCategory
            Dim doc1 As HtmlDocument = helper.GetHtmlDocument2(cate.Url, 120000)
            'Dim cateProdNodes As HtmlNodeCollection = doc1.DocumentNode.SelectNodes("//div[@class='hotitem bt_border mt10']/div[2]/div[1]/div")
            Dim cateProdNodes As HtmlNodeCollection = doc1.DocumentNode.SelectNodes("//div[@class='col_m mt10']/div[3]/ul/li")
            'If (cateCounter = 0) Then '大分类
            '    cateProdNodes = doc1.DocumentNode.SelectNodes("//div[@class='hotitem bt_border mt10']/div[2]/div[1]/div")
            'Else '小分类
            '    cateProdNodes = doc1.DocumentNode.SelectNodes("//div[@id='list_content']/div")
            'End If
            Dim counter As Integer = 0
            For Each cateProdNode In cateProdNodes
                Dim product As New Product
                Dim productName As String = cateProdNode.SelectSingleNode("p[@class='listProduct_name']").InnerText.Trim()
                Dim productUrl As String = FocalPrice.ParseLink(cateProdNode.SelectSingleNode("p[@class='listProduct_img']/a").GetAttributeValue("href", "").Trim())
                If (counter >= 3) Then '插入3个产品
                    Exit For
                Else
                    If Not (listNowProductUrl.Contains(productUrl)) Then '该产品在其他模块中没有出现
                        If Not (helper.IsProductSent2(siteId, productUrl, Now.AddDays(-14).ToString, Now.ToString(), planType)) Then '产品在前7天中没有发送过
                            If (productName.Length >= 63) Then
                                productName = productName.Substring(0, 63).Trim()
                                'productName = productName.Substring(0, productName.LastIndexOf(" ")) & "..."
                            End If
                            product.Prodouct = productName
                            product.Url = productUrl
                            'If (cateCounter = 0) Then '大分类
                            '    product.Discount = Double.Parse(Regex.Split(cateProdNode.SelectSingleNode("ul/li[3]").InnerText.Trim(), "US\$")(1).Trim())
                            'Else '小分类
                            '    product.Discount = Double.Parse(cateProdNode.SelectSingleNode("ul/li[3]/span[1]").InnerText.Replace("US$", "").Trim())
                            'End If
                            product.Discount = Double.Parse(Regex.Split(cateProdNode.SelectSingleNode("p[@class='listProduct_price']/span[1]").InnerText.Trim(), "US\$")(1).Trim())

                            product.PictureUrl = cateProdNode.SelectSingleNode("p[@class='listProduct_img']/a/img").GetAttributeValue("src", "")
                            product.LastUpdate = lastUpdate
                            product.SiteID = siteId
                            Dim productId As Integer = helper.InsertOrUpdateProduct(product, cate.Url.Trim(), siteId, lastUpdate)
                            helper.InsertSinglePIssue(productId, siteId, IssueID, "CA")
                            listNowProductUrl.Add(product.Url)
                            counter = counter + 1
                        End If
                    End If
                End If
            Next
            cateCounter = cateCounter + 1
        Next
    End Sub

    Private Shared Sub CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal categorys As String, _
                                         ByVal lastUpdate As DateTime, ByVal planType As String, ByVal spreadloginEmail As String, _
                                         ByVal appId As String)
        Dim splitCategory As String()  '= categorys.Split("^")
        Dim saveListName As String  '= "FocalPrice-" & splitCategory(0) & "-" & DateTime.Now.ToString("yyyyMMdd")
        Dim categoryId As String  '= EFHelper.GetCategoryId(siteId, splitCategory(0).Trim())
        If (categorys.Contains("^")) Then
            splitCategory = categorys.Split("^")
            saveListName = "FocalPrice-" & splitCategory(0) & "-" & DateTime.Now.ToString("yyyyMMdd")
            categoryId = EFHelper.GetCategoryId(siteId, splitCategory(0).Trim())
        Else
            saveListName = "FocalPrice-" & categorys & "-" & DateTime.Now.ToString("yyyyMMdd")
            categoryId = EFHelper.GetCategoryId(siteId, categorys)
        End If
        EFHelper.LogText("categoryID:" & categoryId)
        'Dim QuerySubscriber As New QuerySubscriber
        'QuerySubscriber.Favorite = categoryId
        'QuerySubscriber.Strategy = ChooseStrategy.Favorite
        'QuerySubscriber.CountryList = New String() {}
        'QuerySubscriber.StartDate = Date.Now.AddDays(-7).ToString("yyyy-MM-dd")

        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Favorite = categoryId
        QuerySubscriber.Strategy = ChooseStrategy.Favorite
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(-20).ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(spreadloginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
        'Dim arrContactList As String() = New String() {saveListName}
        If (count > 0) Then '表示创建联系人列表成功
            Dim helper As New EFHelper
            helper.InsertContactList(issueId, saveListName, "draft") 'waiting
        End If

        '创建联系人列表，begin
        'Dim helper As New EFHelper()
        'Dim count As Integer = -1
        'Try
        '    count = helper.CreateContactList(siteId, issueId, planType, saveListName, 30, ChooseStrategy.Favorite, "draft", spreadloginEmail, appId)
        'Catch ex As Exception
        '    Throw New Exception(ex.ToString())
        'End Try
        'If (count > 0) Then 'contactList创建成功
        '    helper.InsertContactList(issueId, saveListName, "draft") 'waiting
        '    'Else ''创建联系人失败后，就直接在SentLogs表中添加一条数据，表示不必再次进行一个自动化过程
        '    '    EFHelper.InsertSentLog(siteId, planType, splitCategory(0).Trim() & " nobody click", lastUpdate)
        'End If
        '创建联系人列表，end
    End Sub

    ''' <summary>
    ''' 从每个分类中拿取"Categorías Recomendación"块的一个产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="lastUpdate"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetCategoryProduct(ByVal siteId As Integer, ByVal issueId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim listmyCategory As List(Of Category) = helper.GetListCategory(siteId)
        Dim listLastProdUrl As New List(Of String)

        'Dim listLastAllProdUrl As New List(Of String)
        'listLastAllProdUrl = helper.GetTopNSectionProdUrl(issueId, 26, "", siteId)
        'Dim iTotalCounter As Integer = 36

        For Each li In listmyCategory '取'Categorías Recomendación块的产品
            If Not (String.IsNullOrEmpty(li.Url)) Then '如果该分类不是首页的分类，则不作处理
                Dim listProductId As New List(Of Integer)
                Dim doc As HtmlDocument = helper.GetHtmlDocument2(li.Url, 120000)
                Dim cateProdNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='hotitem bt_border mt10 Recommended_item']/div[2]/div[1]/div")
                Dim counter As Integer = 0
                For Each cateProdNode In cateProdNodes
                    Dim product As New Product
                    Dim productName As String = cateProdNode.SelectSingleNode("ul/li[2]/a").InnerText.Trim()
                    Dim productUrl As String = FocalPrice.ParseLink(cateProdNode.SelectSingleNode("ul/li[1]/a").GetAttributeValue("href", "").Trim())
                    If (counter >= 1) Then '插入一个产品
                        Exit For
                    Else
                        If Not (listNowProductUrl.Contains(productUrl)) Then '该产品在其他模块中没有出现
                            If Not (helper.IsProductSent(siteId, productUrl, Now.AddDays(-7).ToString, Now.ToString())) Then '产品在前7天中没有发送过
                                If (productName.Length >= 63) Then
                                    productName = productName.Substring(0, 63).Trim()
                                    productName = productName.Substring(0, productName.LastIndexOf(" ")) & "..."
                                End If
                                product.Prodouct = productName
                                product.Url = productUrl
                                product.Discount = Double.Parse(Regex.Split(cateProdNode.SelectSingleNode("ul/li[3]").InnerText.Trim(), "US\$")(1).Trim())
                                product.PictureUrl = cateProdNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                                product.LastUpdate = lastUpdate
                                product.SiteID = siteId
                                Dim productId As Integer = helper.InsertOrUpdateProduct(product, li.Url.Trim(), siteId, lastUpdate)
                                helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                                listNowProductUrl.Add(product.Url)
                                counter = counter + 1
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End Sub


End Class
