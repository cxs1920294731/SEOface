Imports HtmlAgilityPack

Public Class TenMonthTmall

    Private Shared productCounter As Integer = 0

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, _
                     ByVal url As String, ByVal category As String, ByVal lastUpdate As DateTime)
        Dim tmallSearchUrl As String = "http://octlegendast.tmall.com/search.htm"
        Dim helper As New EFHelper()
        If Not (String.IsNullOrEmpty(category)) Then '个性化邮件
            Dim myCategory As Category = EFHelper.GetCategory(siteId, category.Trim())
            'GetLatestProducts(myCategory.Url, siteId, "CA", IssueID, lastUpdate)
            GetTopSellProducts(myCategory.Url, "", siteId, planType, IssueID, lastUpdate, 9)

            Dim saveListName As String = "octlegendTmall-" & category & "-" & DateTime.Now.ToString("yyyyMMdd")
            Dim counter As Integer = -1
            Try
                counter = helper.CreateContactList(siteId, IssueID, planType, myCategory.Category1, saveListName, 30, ChooseStrategy.Favorite, "draft", spreadLogin, appId)
            Catch ex As Exception
                Throw New Exception(ex.ToString())
            End Try
            If (counter > 0) Then
                helper.InsertContactList(IssueID, saveListName, "draft") 'waiting
            End If
            Dim subject As String = "本周王牌热卖，十月传奇为你精心推荐：" & EFHelper.GetFirstProductSubject(IssueID) & "(AD)"
            helper.InsertIssueSubject(IssueID, subject)
        Else '新品推荐邮件
            GetCategory(siteId, lastUpdate)
            GetLatestProducts(tmallSearchUrl, siteId, planType, IssueID, lastUpdate)
            If (productCounter < 9) Then
                Dim categoryUrl As String = "http://octlegendast.tmall.com/p/rd415809.htm"
                Dim requestUrl As String = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
                GetTopSellProducts(categoryUrl, requestUrl, siteId, planType, IssueID, lastUpdate, 9 - productCounter)
            End If
            helper.InsertContactList(IssueID, "opens_of_ 与C店不重复的名单", "draft")
            Dim subject As String = "本周新品推荐：" & EFHelper.GetFirstProductSubject(IssueID) & "(AD)"
            helper.InsertIssueSubject(IssueID, subject)  '大女人小情怀，十月传奇春装上新(AD)
        End If
    End Sub
    Private Shared Sub GetCategory(ByVal siteId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "王牌热卖专区"
        myCategory.Description = "王牌热卖专区"
        myCategory.Url = "http://octlegendast.tmall.com/category.htm"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "newon"
        myCategory2.Description = "newon"
        myCategory2.Url = "http://octlegendast.tmall.com/p/rd415809.htm"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "热卖折扣专区"
        myCategory3.Description = "热卖折扣专区"
        myCategory3.Url = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)
    End Sub
    ''' <summary>
    ''' 获取"新品发布"的产品
    ''' </summary>
    ''' <param name="searchUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="issueId"></param>
    ''' <param name="lastUpdate"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetLatestProducts(ByVal searchUrl As String, ByVal siteId As Integer, ByVal planType As String, _
                                        ByVal issueId As Integer, ByVal lastUpdate As DateTime)
        Dim helper As New EFHelper
        Dim searchDoc As HtmlDocument = EFHelper.GetHtmlDocByUrlTmall(searchUrl)
        Dim categoryLi As HtmlNodeCollection = searchDoc.DocumentNode.SelectNodes("//ul[@class='J_TAllCatsTree cats-tree']/li[2]/div[1]/div[1]/ul/li")
        'Dim categoryHtml As String = searchDoc.DocumentNode.SelectSingleNode("//ul[@class='J_TAllCatsTree cats-tree']/li[3]/div[1]/div[1]/ul/li").InnerHtml
        For Each li In categoryLi
            Dim categoryUrl As String = li.SelectSingleNode("h4/a").GetAttributeValue("href", "")
            Dim productDoc As HtmlDocument = EFHelper.GetHtmlDocByUrlTmall(categoryUrl)
            Dim productLines As HtmlNodeCollection = productDoc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")
            If (productLines Is Nothing) Then
                Continue For
            Else
                For Each productLine In productLines
                    Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
                    If Not (productDls Is Nothing) Then
                        For Each dl In productDls
                            Dim myProduct As New Product
                            Dim productName As String = ""
                            Try
                                productName = dl.SelectSingleNode("dd[2]/a").InnerText
                            Catch ex As Exception
                                productName = dl.SelectSingleNode("dd[1]/a").InnerText
                            End Try

                            myProduct.Prodouct = productName
                            Try
                                myProduct.Url = dl.SelectSingleNode("dd[2]/a").GetAttributeValue("href", "").Trim()
                                myProduct.Discount = dl.SelectSingleNode("dd[2]/div[1]/div[1]/span[2]").InnerText.Trim()
                            Catch ex As Exception
                                myProduct.Url = dl.SelectSingleNode("dd[1]/a").GetAttributeValue("href", "").Trim()
                                myProduct.Discount = dl.SelectSingleNode("dd[1]/div[1]/div[1]/span[2]").InnerText.Trim()
                            End Try
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                            myProduct.LastUpdate = lastUpdate
                            myProduct.Description = productName
                            myProduct.SiteID = siteId
                            Dim productId As Integer = helper.InsertOrUpdateProduct(myProduct, "http://octlegendast.tmall.com/p/rd415809.htm", siteId, lastUpdate)
                            helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                            productCounter = productCounter + 1
                        Next
                    Else
                        Continue For
                    End If
                Next
                Exit For
            End If
        Next
    End Sub
    ''' <summary>
    ''' 获取热卖/折扣的产品
    ''' </summary>
    ''' <param name="categoryUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="issueId"></param>
    ''' <param name="lastUpdate"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetTopSellProducts(ByVal categoryUrl As String, ByVal requestCateUrl As String, ByVal siteId As Integer, _
                                          ByVal planType As String, ByVal issueId As Integer, ByVal lastUpdate As DateTime, _
                                          ByVal counter As Integer)
        'Dim categoryUrl As String = "http://octlegendast.tmall.com/category.htm?orderType=hotsell_desc"
        Dim helper As New EFHelper
        Dim categoryDoc As HtmlDocument
        If (String.IsNullOrEmpty(requestCateUrl)) Then 'requestCateUrl为空，则使用categoryUrl作为请求的url
            categoryDoc = EFHelper.GetHtmlDocByUrlTmall(categoryUrl)
        Else
            categoryDoc = EFHelper.GetHtmlDocByUrlTmall(requestCateUrl)
        End If
        'Dim categoryDoc As HtmlDocument = EFHelper.GetHtmlDocByUrlTmall(categoryUrl)
        Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")
        Dim productCounter1 As Integer = 0
        For Each productLine In productLines
            Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
            If Not (productDls Is Nothing) Then
                For Each dl In productDls
                    If Not (productCounter1 >= counter) Then
                        Dim myProduct As New Product
                        Dim productName As String = dl.SelectSingleNode("dd[2]/a").InnerText
                        myProduct.Prodouct = productName
                        myProduct.Url = dl.SelectSingleNode("dd[2]/a").GetAttributeValue("href", "").Trim()
                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now)) Then
                            myProduct.Discount = dl.SelectSingleNode("dd[2]/div[1]/div[1]/span[2]").InnerText.Trim()
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                            myProduct.LastUpdate = lastUpdate
                            myProduct.Description = productName
                            myProduct.SiteID = siteId
                            Dim productId As Integer = helper.InsertOrUpdateProduct(myProduct, categoryUrl, siteId, lastUpdate)
                            helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                            productCounter1 = productCounter1 + 1
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
            If (productCounter >= 9) Then
                Exit For
            End If
        Next
    End Sub
End Class
