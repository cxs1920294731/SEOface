Imports HtmlAgilityPack

Public Class TenMonthTaoBao
    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, _
                     ByVal url As String, ByVal category As String, ByVal lastUpdate As DateTime)
        Dim searchPageUrl As String = "http://anshutong.taobao.com/search.htm"
        Dim helper As New EFHelper()
        If Not (String.IsNullOrEmpty(category)) Then '个性化邮件，星期五
            GetTopSellProducts(searchPageUrl, siteId, planType, IssueID, lastUpdate)
            '创建联系人列表
            Dim saveListName As String = "octlegend-" & category & "-" & DateTime.Now.ToString("yyyyMMdd")
            Dim categoryId As String = EFHelper.GetCategoryId(siteId, category.Trim())
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Favorite = categoryId
            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            QuerySubscriber.CountryList = New String() {}
            QuerySubscriber.StartDate = Date.Now.AddDays(-7).ToString("yyyy-MM-dd")
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()
            Dim count As Integer = -1
            Try
                count = mySpread.SearchContacts(spreadLogin, appId, CriteriaString, Integer.MaxValue, saveListName, True)
            Catch ex As Exception
                Throw New Exception(ex.ToString())
            End Try
            If (count > 0) Then 'contactList创建成功
                helper.InsertContactList(IssueID, saveListName, "draft") 'waiting
            End If
            Dim subject As String = "本周王牌热卖，十月传奇为你精心推荐：" & EFHelper.GetFirstProductSubject(IssueID) & "(AD)"
            helper.InsertIssueSubject(IssueID, subject)
            'helper.InsertIssueSubject(IssueID, "大女人小情怀，十月传奇春装上新(AD)")
        Else '新品推荐邮件，星期一
            GetCategory(siteId, lastUpdate)
            GetLatestProducts(searchPageUrl, siteId, planType, IssueID, lastUpdate)
            ''创建联系人列表
            'Dim saveListName As String = "TenMonthTaobao-" & category & "-" & DateTime.Now.ToString("yyyyMMdd")
            'Dim categoryId As String = EFHelper.GetCategoryId(siteId, category.Trim())
            'Dim helper As New EFHelper()
            'Dim count As Integer = -1
            'Try
            '    count = helper.CreateContactList(siteId, IssueID, planType, saveListName, 30, ChooseStrategy.Favorite, "draft", spreadLogin, appId)
            'Catch ex As Exception
            '    Throw New Exception(ex.ToString())
            'End Try
            'If (count > 0) Then 'contactList创建成功
            'helper.InsertContactList(IssueID, "Reasonables", "draft") '测试联系人列表
            'helper.InsertContactList(IssueID, "2014年3月7", "draft")
            helper.InsertContactList(IssueID, "opens", "draft")
            'End If
            Dim subject As String = "本周新品推荐：" & EFHelper.GetFirstProductSubject(IssueID) & "(AD)"
            helper.InsertIssueSubject(IssueID, subject)
        End If

    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "王牌热卖专区"
        myCategory.Url = "http://anshutong.taobao.com/p/rd000842.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Public Shared Sub GetLatestProducts(ByVal url As String, ByVal siteId As Integer, ByVal planType As String, _
                                        ByVal issueId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim doc As HtmlDocument = EFHelper.GetHtmlDocFromTaoBao(url)
        Dim categoryLis As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-item-cates']/div[2]/ul/li[3]/ul/li")
        Dim categoryUrl As String
        Dim counter As Integer = 1
        Dim productDivs As HtmlNodeCollection
        For Each liNode In categoryLis
            categoryUrl = doc.DocumentNode.SelectSingleNode("//div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-item-cates']/div[2]/ul/li[3]/ul/li[" & counter & "]/h4/a").GetAttributeValue("href", "")
            productDivs = EFHelper.GetHtmlDocFromTaoBao(categoryUrl).DocumentNode.SelectNodes("//div[@class='item3line1']")
            If Not (productDivs Is Nothing) Then
                Dim categoryName As String = doc.DocumentNode.SelectSingleNode("//div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-item-cates']/div[2]/ul/li[3]/ul/li[" & counter & "]/h4/a").InnerText
                If (categoryUrl.Contains("?")) Then
                    categoryUrl = categoryUrl.Substring(0, categoryUrl.IndexOf("?")).Trim()
                End If
                Dim myCategory As New Category
                myCategory.Category1 = "newon"
                myCategory.Description = "newon"
                myCategory.Url = "http://anshutong.taobao"
                myCategory.SiteID = siteId
                myCategory.LastUpdate = lastUpdate
                helper.InsertOrUpdateCate(myCategory, siteId)
                Exit For
            End If
            counter = counter + 1
        Next
        'Dim productDivs As HtmlNodeCollection = EFHelper.GetHtmlDocFromTaoBao(categoryUrl).DocumentNode.SelectNodes("//div[@class='item3line1']")
        For Each productDiv In productDivs
            Dim productNodes As HtmlNodeCollection = productDiv.SelectNodes("dl")
            For Each node In productNodes
                Dim pictureUrl As String = node.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                If (String.IsNullOrEmpty(pictureUrl)) Then
                    pictureUrl = node.SelectSingleNode("dt/a/img").GetAttributeValue("src", "")
                End If
                Dim productUrl As String = node.SelectSingleNode("dt/a").GetAttributeValue("href", "")
                Dim productName As String = node.SelectSingleNode("dd[1]/a").InnerText
                Dim strPrice As String = node.SelectSingleNode("dd[1]/div[1]/div[1]/span[2]").InnerText
                Dim product As New Product()
                product.Prodouct = productName
                product.Url = productUrl
                product.Discount = Double.Parse(strPrice)
                product.PictureUrl = pictureUrl
                product.LastUpdate = lastUpdate
                product.Description = productName
                product.SiteID = siteId
                Dim productId As Integer = helper.InsertOrUpdateProduct(product, "http://anshutong.taobao", siteId, lastUpdate)
                helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
            Next
        Next
    End Sub

    Public Shared Sub GetTopSellProducts(ByVal url As String, ByVal siteId As Integer, ByVal planType As String, _
                                        ByVal issueId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim doc As HtmlDocument = EFHelper.GetHtmlDocFromTaoBao(url)
        Dim categoryUrl As String = doc.DocumentNode.SelectSingleNode("//div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-item-cates']/div[2]/ul/li[2]/h4/a").GetAttributeValue("href", "")
        'Dim myCategory As New Category
        'Dim categoryName As String = doc.DocumentNode.SelectSingleNode("//div[@class='skin-box tb-module tshop-pbsm tshop-pbsm-shop-item-cates']/div[2]/ul/li[2]/h4/a").InnerText
        If (categoryUrl.Contains("?")) Then
            categoryUrl = categoryUrl.Substring(0, categoryUrl.IndexOf("?")).Trim()
        End If
        'myCategory.Category1 = categoryName
        'myCategory.SiteID = siteId
        'myCategory.LastUpdate = lastUpdate
        'myCategory.Description = categoryName
        'helper.InsertCategory(myCategory)
        Dim requestCategoryUrl As String = categoryUrl & "?orderType=hotsell_desc&pageNo=1"
        Dim productDivs As HtmlNodeCollection = EFHelper.GetHtmlDocFromTaoBao(requestCategoryUrl).DocumentNode.SelectNodes("//div[@class='item3line1']")
        Dim requestCategoryUrl2 As String = categoryUrl & "?orderType=hotsell_desc&pageNo=2"
        Dim productDivs2 As HtmlNodeCollection = EFHelper.GetHtmlDocFromTaoBao(requestCategoryUrl2).DocumentNode.SelectNodes("//div[@class='item3line1']")
        For Each div In productDivs2
            productDivs.Add(div)
        Next

        'Dim productDivs As HtmlNodeCollection = EFHelper.GetHtmlDocFromTaoBao(categoryUrl).DocumentNode.SelectNodes("//div[@class='item3line1']")
        Dim productCounter As Integer = 0
        For Each productDiv In productDivs
            Dim productNodes As HtmlNodeCollection = productDiv.SelectNodes("dl")
            For Each node In productNodes
                Dim pictureUrl As String = node.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                If (String.IsNullOrEmpty(pictureUrl)) Then
                    pictureUrl = node.SelectSingleNode("dt/a/img").GetAttributeValue("src", "")
                End If
                Dim productUrl As String = node.SelectSingleNode("dt/a").GetAttributeValue("href", "")
                If Not (productCounter >= 9) Then
                    If Not (helper.IsProductSent(siteId, productUrl, Now.AddDays(-30), Now)) Then '该产品在前30天内没有被发送
                        Dim productName As String = node.SelectSingleNode("dd[1]/a").InnerText
                        Dim strPrice As String = node.SelectSingleNode("dd[1]/div[1]/div[1]/span[2]").InnerText
                        Dim product As New Product()
                        product.Prodouct = productName
                        product.Url = productUrl
                        product.Discount = Double.Parse(strPrice)
                        Try
                            product.Price = Double.Parse(node.SelectSingleNode("dd[1]/div[1]/div[2]/span[2]").InnerText)
                        Catch ex As Exception  '只有一个价格
                            'Ignore
                        End Try
                        product.Currency = "¥"
                        product.PictureUrl = pictureUrl
                        product.LastUpdate = lastUpdate
                        product.Description = productName
                        product.SiteID = siteId
                        Dim hotsellCateUrl As String = "http://anshutong.taobao.com/p/rd000842.htm"
                        Dim productId As Integer = helper.InsertOrUpdateProduct(product, hotsellCateUrl, siteId, lastUpdate)
                        helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                        productCounter = productCounter + 1
                    End If
                Else
                    Exit For
                End If
            Next
            If (productCounter >= 9) Then
                Exit For
            End If
        Next
    End Sub
End Class
