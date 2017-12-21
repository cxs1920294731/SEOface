Imports System.Configuration
Imports HtmlAgilityPack

Public Class EverBuyingPersonalization

    Private Shared listNowProductUrl As New List(Of String)  '保证每个模块的产品都不一样
    Private Shared listProductName As New List(Of String)  '保证相同型号的产品也不出现

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                    ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, _
                    ByVal url As String, ByVal categoryName As String)
        Dim updateTime As DateTime = DateTime.Now
        Dim arrCategory As String() = categoryName.Split("^")
        GetCatesProducts(siteId, IssueID, updateTime, planType, arrCategory)
        'categoryName传值过来：''Sexy Lingerie^Women's Dresses^Women's Handbags^Jewelry^Wedding & Events^Women's Shoes
        'categoryName = categoryName.Substring(0, categoryName.IndexOf("^"))
        EverBuyingDeals.InsertSubject(IssueID, siteId, planType, arrCategory(0))
        CreateContactList(siteId, IssueID, planType, arrCategory(0), spreadLogin, appId, updateTime)
    End Sub

    ''' <summary>
    ''' 获取主推分类的产品，产品页面按"hot"进行排序，如：http://www.everbuying.com/Wholesale-Cell-Phones-b-22.html?odr=reviews
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="updateTime"></param>
    ''' <param name="planType"></param>
    ''' <param name="arrCategorys"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetCatesProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal updateTime As DateTime, _
                                       ByVal planType As String, ByVal arrCategorys As String())
        Dim helper As New EFHelper
        'Dim cate1 As String = "http://www.everbuying.com/Wholesale-Cell-Phones-b-22.html?page_size=36&odr=reviews&display=g"
        'Dim cate2 As String = "http://www.everbuying.com/Wholesale-iPhone-iPad-iPod-b-493.html?page_size=36&odr=reviews&display=g"
        'Dim cate3 As String = "http://www.everbuying.com/Wholesale-Cell-Phone-Accessories-b-667.html?page_size=36&odr=reviews&display=g"
        'Dim cate4 As String = "http://www.everbuying.com/Wholesale-Notebook-UMPC-MID-b-815.html?page_size=36&odr=reviews&display=g"
        'Dim cate5 As String = "http://www.everbuying.com/smlclass1674.html?page_size=36&odr=reviews&display=g"
        'Dim cate6 As String = "http://www.everbuying.com/Wholesale-Computers-Networking-b-926.html?page_size=36&odr=reviews&display=g"
        'Dim cate7 As String = "http://www.everbuying.com/Wholesale-LED-Lights-b-669.html?page_size=36&odr=reviews&display=g"
        'Dim cate8 As String = "http://www.everbuying.com/Wholesale-Consumer-Electronics-b-853.html?page_size=36&odr=reviews&display=g"
        'Dim cate9 As String = "http://www.everbuying.com/Wholesale-Men-s-Clothing-c-98.html?page_size=36&odr=reviews&display=g"
        'Dim cate10 As String = "http://www.everbuying.com/smlclass1252.html?page_size=36&odr=reviews&display=g"
        'Dim arrCategorys As String() = New String() {cate1, cate2, cate3, cate4, cate5, cate6, cate7, cate8, cate9, cate10}
        Dim listCategory As List(Of Category) = EFHelper.GetListCategorys(siteId, arrCategorys)
        For Each cate In listCategory
            Dim listProductId As New List(Of Integer)
            Dim cateDoc As HtmlDocument = helper.GetHtmlDocument(cate.Url, 60000)
            Dim productNodes As HtmlNodeCollection = cateDoc.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            Dim counter As Integer = 0
            For Each productNode In productNodes
                Dim product As New Product
                Dim productName As String = productNode.SelectSingleNode("ul/li[2]").InnerText.Trim()
                product.Prodouct = productName
                product.Url = EverBuyingDeals.AddDomain("http://www.everbuying.com", productNode.SelectSingleNode("ul/li[2]/a").GetAttributeValue("href", "").Trim())
                If Not (listProductName.Contains(productName)) Then  '相同型号的产品不出现
                    If Not (helper.IsProductSent(siteId, product.Url, Now.AddDays(-7), Now, planType)) Then
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
                        'product.PictureUrl = productNode.SelectSingleNode("ul/li[1]/a/img").GetAttributeValue("src", "")
                        product.PictureUrl = pictureUrl
                        product.LastUpdate = updateTime
                        product.Description = productName
                        product.SiteID = siteId
                        product.Currency = "USD"
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
                        'Dim categoryUrl As String = cate.Url
                        Dim productId As Integer = helper.InsertOrUpdateProduct(product, cate.Url, siteId, updateTime)
                        helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
                        listNowProductUrl.Add(product.Url)
                        listProductName.Add(productName)
                        counter = counter + 1
                        If (counter >= 6) Then
                            Exit For
                        End If
                        'listProductId.Add(productId)
                    Else
                        Continue For
                    End If
                Else
                    Continue For
                End If
            Next
            'helper.InsertProductIssue(listProductId, siteId, issueId, "CA", 3, productNodes.Count)
        Next
    End Sub

    ''' <summary>
    ''' 获取ContactList
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="categoryName"></param>
    ''' <remarks></remarks>
    Public Shared Sub CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, _
                                        ByVal categoryName As String, ByVal loginEmail As String, ByVal appId As String, _
                                        ByVal sentTime As DateTime)
        Dim saveListName As String = "EverBuying-" & categoryName & "-" & DateTime.Now.ToString("yyyyMMdd")
        Dim categoryId As String = EFHelper.GetCategoryId(siteId, categoryName.Trim())
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Favorite = categoryId
        QuerySubscriber.Strategy = ChooseStrategy.Favorite
        QuerySubscriber.CountryList = New String() {}
        'QuerySubscriber.StartDate = Date.Now.AddDays(-7).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = DateTime.Now.ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.Service = New SpreadWebReference.Service()
        mySpread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
        'Dim arrContactList As String() = New String() {saveListName}
        If (count > 0) Then '表示创建联系人列表成功
            Dim helper As New EFHelper
            helper.InsertContactList(issueId, saveListName, "waiting") 'waiting
        Else '创建联系人失败后，就直接在SentLogs表中添加一条数据，表示不必再次进行一个自动化过程
            EFHelper.InsertSentLog(siteId, planType, categoryName & " nobody click", sentTime)
        End If
    End Sub
End Class
