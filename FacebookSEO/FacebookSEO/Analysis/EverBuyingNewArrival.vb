Imports HtmlAgilityPack

Public Class EverBuyingNewArrival
    Private Shared listProductUrl As New List(Of String)
    Private Shared listProductName As New List(Of String)
    Private Shared listProductSku As New List(Of String)

    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                      ByVal spreadLogin As String, ByVal appId As String, ByVal updateTime As DateTime)
        GetCatesProducts(siteId, IssueID, updateTime, planType)
        EverBuyingDeals.InsertSubject(IssueID, siteId, planType, "New Arrival")
        Dim helper As New EFHelper
        '2014/04/02 added
        'helper.InsertContactList(IssueID, "opens 20140214", "draft")
        'helper.InsertContactList(IssueID, "E_SPREAD_001", "draft")
        helper.InsertContactList(IssueID, "Opens20140402-1", "waiting")
    End Sub

    ''' <summary>
    ''' 获取主推分类的产品，产品页面按"hot"进行排序，如：http://www.everbuying.com/Wholesale-Cell-Phones-b-22.html?odr=new
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="updateTime"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetCatesProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Dim helper As New EFHelper
        Dim cate1 As String = "http://www.everbuying.net/Wholesale-Cell-Phones-b-22.html?page_size=48&odr=new&display=g"
        Dim cate2 As String = "http://www.everbuying.net/Wholesale-iPhone-iPad-iPod-b-493.html?page_size=48&odr=new&display=g"
        Dim cate3 As String = "http://www.everbuying.net/Wholesale-Cell-Phone-Accessories-b-667.html?page_size=48&odr=new&display=g"
        Dim cate4 As String = "http://www.everbuying.net/Wholesale-Notebook-UMPC-MID-b-815.html?page_size=48&odr=new&display=g"
        Dim cate5 As String = "http://www.everbuying.net/smlclass1674.html?page_size=48&odr=new&display=g"
        Dim cate6 As String = "http://www.everbuying.net/Wholesale-Computers-Networking-b-926.html?page_size=48&odr=new&display=g"
        Dim cate7 As String = "http://www.everbuying.net/Wholesale-LED-Lights-b-669.html?page_size=48&odr=new&display=g"
        Dim cate8 As String = "http://www.everbuying.net/Wholesale-Consumer-Electronics-b-853.html?page_size=48&odr=new&display=g"
        Dim cate9 As String = "http://www.everbuying.net/Wholesale-Men-s-Clothing-c-98.html?page_size=48&odr=new&display=g"
        Dim cate10 As String = "http://www.everbuying.net/smlclass1252.html?page_size=48&odr=new&display=g"
        Dim arrCategorys As String() = New String() {cate1, cate2, cate3, cate4, cate5, cate6, cate7, cate8, cate9, cate10}
        Dim arrCategoryNames As String() = New String() {"Cell Phones", "Apple Accessories", "Cell Phone Accessories", "Tablet PC", "Mobile Power Bank", "Computers & Networking",
                                                        "LED Lights", "Consumer Electronics"}
        For i = 0 To 7
            Dim categoryId As Integer = EFHelper.GetCategoryId(siteId, arrCategoryNames(i))
            Dim listProduct As New List(Of Product)
            listProduct = EverBuyingDeals.GetProductList(siteId)

            Dim cate As String = arrCategorys(i)
            Dim cateDoc As HtmlDocument = helper.GetHtmlDocument(cate, 60000)
            Dim productNodes As HtmlNodeCollection = cateDoc.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            Dim counter As Integer = 0
            Dim sku As String
            For Each productNode In productNodes
                Try
                    If (counter >= 3) Then
                        Exit For
                    Else
                        Dim product As New Product
                        Dim productName As String = productNode.SelectSingleNode("ul/li[2]").InnerText.Trim()
                        product.Prodouct = productName
                        product.Url = EverBuyingDeals.AddDomain("http://www.everbuying.net", productNode.SelectSingleNode("ul/li[2]/a").GetAttributeValue("href", "").Trim())
                        If (listProductUrl.Contains(product.Url) OrElse listProductName.Contains(product.Prodouct)) Then '该产品在本期的邮件中了
                            Continue For
                        Else
                            listProductUrl.Add(product.Url)
                            listProductName.Add(product.Prodouct)
                        End If
                        sku = EverBuyingDeals.getSKU(product.Url)
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
            'helper.InsertProductIssue(listProductId, siteId, issueId, "CA", 3, productNodes.Count)
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
                            product.Url = EverBuyingDeals.AddDomain("http://www.everbuying.net", productNode.SelectSingleNode("ul/li[2]/a").GetAttributeValue("href", "").Trim())
                            If (listProductUrl.Contains(product.Url) OrElse listProductName.Contains(product.Prodouct)) Then '该产品在本期的邮件中了
                                Continue For
                            Else
                                listProductUrl.Add(product.Url)
                                listProductName.Add(product.Prodouct)
                            End If
                            sku = EverBuyingDeals.getSKU(product.Url)
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
        Next
    End Sub
End Class
