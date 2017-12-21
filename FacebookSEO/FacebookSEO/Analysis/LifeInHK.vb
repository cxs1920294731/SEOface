Imports Analysis.IAnalysis

Public Class LifeInHK
    Implements AutoInterface

    Private helper As New EFHelper()
    Private efContext As New EmailAlerterEntities()
    Private listProductId As New List(Of Integer)
    Private listProducts As New List(Of Product)
    Public Sub Start(ByVal issueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, _
                     ByVal appId As String, ByVal shopUrl As String, ByVal nowTime As DateTime) Implements AutoInterface.IStart
        'Try
        Dim subject As String = ""

        '2013/06/14,change the subject rule
        'subject = GetPageInfo(shopUrl, siteId, nowTime)
        GetPageInfo(shopUrl, siteId, nowTime)
        If (listProducts.Count > 0) Then
            subject = "獨家優惠 - " & GetProductName()
            If (subject.Length > 250) Then
                subject = subject.Substring(0, 250)
                subject = subject.Substring(0, subject.LastIndexOf("/"))
            End If
            helper.InsertProductIssue(listProductId, siteId, issueID, "NEDM") 'NEDM：New Arrival Desktop and Mobile
            helper.InsertIssueSubject(issueID, subject)
            InsertContactList(issueID)
        End If
        'Catch ex As Exception
        '    EFHelper.Log(ex)
        '    EFHelper.LogText(ex.ToString())
        'End Try
    End Sub

    ''' <summary>
    ''' get products and category from rss url
    ''' </summary>
    ''' <param name="rssURL"></param>
    ''' <param name="siteId"></param>
    ''' <remarks></remarks>
    Private Sub GetPageInfo(ByVal rssURL As String, ByVal siteId As Integer, ByVal updateTime As DateTime)
        'Dim dfadftest As String = ""
        Try
            ''2013/06/14,change the subject rule
            'Dim productsName As New Text.StringBuilder

            'get all the categoryName by siteId
            Dim listCategoryName As New HashSet(Of String)
            Dim listCategory As List(Of Category) = helper.GetListCategory(siteId)
            For Each li In listCategory
                listCategoryName.Add(li.Category1)
            Next

            'get all the productURL by siteId
            Dim listProduct As List(Of Product) = helper.GetListProduct2(siteId)
            Dim productUrlDic As New Dictionary(Of String, Integer)
            For Each list In listProduct
                If Not (productUrlDic.ContainsKey(list.Url)) Then
                    productUrlDic.Add(list.Url, list.ProdouctID)
                End If
            Next

            Dim xmlDoc As New System.Xml.XmlDocument
            Dim pageString As String = EFHelper.GetRssPageString(rssURL)
            xmlDoc.LoadXml(pageString)
            Dim itemNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//item")
            Dim listUpdateProduct As New List(Of Product) '要更新的产品List
            'Dim pubDate As DateTime = Now
            For Each item As System.Xml.XmlNode In itemNodes '获取每个Item节点的产品信息
                If (Format(item.SelectSingleNode("endDate").InnerText, "yyyy-MM-dd") >= Format(Now, "yyyy-MM-dd")) Then

                    '2013/06/14,change the subject rule
                    'productsName.Append("/")
                    'productsName.Append(item.SelectSingleNode("title").InnerText)

                    Dim categoryName As String = item.SelectSingleNode("category").InnerText.Trim()
                    Dim pubDate As DateTime = DateTime.Parse(item.SelectSingleNode("pubDate").InnerText.Trim())
                    If Not (listCategoryName.Contains(categoryName)) Then '添加Categories表信息
                        Dim myCategory As New Category
                        myCategory.Category1 = categoryName
                        myCategory.SiteID = siteId
                        myCategory.ParentID = -1
                        myCategory.LastUpdate = updateTime
                        efContext.AddToCategories(myCategory)
                        efContext.SaveChanges()
                        listCategoryName.Add(categoryName)
                    End If
                    Dim productUrl As String = item.SelectSingleNode("link").InnerText.Trim()

                    Dim myProduct As New Product
                    myProduct.Prodouct = item.SelectSingleNode("title").InnerText
                    myProduct.Url = productUrl
                    Try
                        myProduct.Price = Decimal.Parse(item.SelectSingleNode("description/price").InnerText.Replace("HKD", "").Trim)
                    Catch ex As Exception
                        Try
                            myProduct.Price = Decimal.Parse(item.SelectSingleNode("description/price").InnerText.Replace("HK$", "").Trim)
                        Catch ex1 As Exception
                            'Ignore
                        End Try
                    End Try
                    'Dim discountPrice As Decimal = 0
                    Try
                        myProduct.Discount = Decimal.Parse(item.SelectSingleNode("description/discountPrice").InnerText.Replace("HKD", "").Trim)
                    Catch ex As Exception
                        Try
                            myProduct.Discount = Decimal.Parse(item.SelectSingleNode("description/discountPrice").InnerText.Replace("HK$", "").Trim)
                        Catch ex1 As Exception
                            'Ignore
                        End Try
                    End Try
                    'myProduct.Discount = discountPrice
                    myProduct.PictureUrl = item.SelectSingleNode("description/productImage").InnerText
                    myProduct.SiteID = siteId
                    myProduct.Currency = "HKD"
                    myProduct.LastUpdate = pubDate
                    If (productUrlDic.ContainsKey(productUrl)) Then
                        '2013/09/26 revised,begin
                        Dim productId As Integer = productUrlDic(productUrl)
                        If Not (listProductId.Contains(productId)) Then
                            listProductId.Add(productId)
                        End If

                        ''listUpdateProduct.Add(myProduct)
                        ''Dim helper As New EFHelper
                        ''helper.UpdateSingleProduct(myProduct, categoryName, siteId, pubDate)

                        'Dim listPIssue As List(Of Products_Issue) = efContext.Products_Issue.Where(Function(pIssue) pIssue.ProductId = productId).ToList()
                        'For Each li In listPIssue '删除product对应的Products_Issue表中的数据
                        '    efContext.DeleteObject(li)
                        'Next
                        'efContext.SaveChanges()
                        'Dim queryProduct As Product = efContext.Products.Where(Function(p) p.ProdouctID = productId).Single()
                        'For i As Integer = 0 To queryProduct.Categories.Count - 1 '删除product对应的ProductCategory表中数据
                        '    queryProduct.Categories.Remove(queryProduct.Categories(i))
                        'Next
                        'efContext.SaveChanges()
                        'efContext.DeleteObject(queryProduct) '删除Products表中对应数据
                        'efContext.SaveChanges()
                        '2013/09/26 revised,end
                        '2013/09/26 added,begin

                        Dim queryProduct As Product = efContext.Products.firstordefault(Function(p) p.Url = productUrl)

                        queryProduct.LastUpdate = pubDate
                        Try
                            queryProduct.Price = Decimal.Parse(item.SelectSingleNode("description/price").InnerText.Replace("HKD", "").Trim)
                        Catch ex As Exception
                            Try
                                queryProduct.Price = Decimal.Parse(item.SelectSingleNode("description/price").InnerText.Replace("HK$", "").Trim)
                            Catch ex1 As Exception
                                'Ignore
                            End Try
                        End Try
                        'Dim discountPrice As Decimal = 0
                        Try
                            queryProduct.Discount = Decimal.Parse(item.SelectSingleNode("description/discountPrice").InnerText.Replace("HKD", "").Trim)
                        Catch ex As Exception
                            Try
                                queryProduct.Discount = Decimal.Parse(item.SelectSingleNode("description/discountPrice").InnerText.Replace("HK$", "").Trim)
                            Catch ex1 As Exception
                                'Ignore
                            End Try
                        End Try
                        queryProduct.Prodouct = item.SelectSingleNode("title").InnerText
                        queryProduct.PictureUrl = item.SelectSingleNode("description/productImage").InnerText
                    Else
                        Dim category As String = item.SelectSingleNode("category").InnerText.Trim()
                        Dim queryCategory As Category = efContext.Categories.Where(Function(c) c.Category1 = category AndAlso c.SiteID = siteId).Single()
                        myProduct.Categories.Add(queryCategory)
                        efContext.AddToProducts(myProduct)
                        efContext.SaveChanges()
                        If Not (listProductId.Contains(myProduct.ProdouctID)) Then
                            listProductId.Add(myProduct.ProdouctID)
                        End If

                        '2013/09/26 added,end
                    End If
                    '2013/09/26 revised,begin
                    'Dim category As String = item.SelectSingleNode("category").InnerText.Trim()
                    'Dim queryCategory As Category = efContext.Categories.Where(Function(c) c.Category1 = category AndAlso c.SiteID = siteId).Single()
                    'myProduct.Categories.Add(queryCategory)
                    'efContext.AddToProducts(myProduct)
                    'efContext.SaveChanges()
                    'listProductId.Add(myProduct.ProdouctID)
                    '2013/09/26 revised,end
                    listProducts.Add(myProduct)
                End If
            Next
            efContext.SaveChanges()
            '2013/06/14,change the subject rule
            'Return productsName.ToString.Remove(productsName.ToString.IndexOf("/"), 1)

        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 插入联系人列表
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Private Sub InsertContactList(ByVal issueId As Integer)
        'Dim arrContactName() As String = {"Life Lover", "Web-Newsletter Archive", "lifeinhk Comments", _
        '                                  "Lifeinhk Test", "Auctioners", "Cars Lover", "BK Female", "BK Male", _
        '                                  "BK Named No Gender", "BK Unknown", "BK Not Yet Matched", "Test Group_Reasonable", "Opened", "All Opens and Clicks as of 2012/04/26", _
        '                                  "InboxTest", "Newsletter Subscription", "Top Deals Newsletter"}

        Dim arrContactName() As String = {"All Opens as of 23 Jun 2013", "InboxTest"}
        For i As Integer = 0 To arrContactName.Length - 1
            Dim contactList As New ContactLists_Issue()
            contactList.IssueID = issueId
            contactList.ContactList = arrContactName(i)
            'contactList.SendingStatus = "waiting" '2013/08/16修改
            contactList.SendingStatus = "draft"
            '------------------------------------
            efContext.AddToContactLists_Issue(contactList)
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' get the name of product order by pubDate,
    ''' 2013/06/14 added
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductName() As String
        Dim productsName As String = ""
        Dim arrProducts() As Product = listProducts.ToArray()
        arrProducts = arrProducts.Reverse.ToArray()
        listProducts = arrProducts.ToList()
        listProducts = listProducts.OrderByDescending(Function(p) p.LastUpdate).ToList()
        Dim counter As Integer = 0
        For Each li In listProducts
            If (counter >= 5) Then
                Exit For
            End If
            productsName = productsName & "[" & li.Prodouct & "]"
            counter = counter + 1
        Next
        Return productsName  'productsName.Remove(productsName.ToString.IndexOf("/"), 1)
    End Function
End Class
