

Imports HtmlAgilityPack
Imports System.Xml

Public Class Resolver
    Private efContext As New Analysis.EmailAlerterEntities()

    Private efHelper As New EFHelper()

    ''' <summary>
    ''' 根据productPath和category规则解析出产品
    ''' </summary>
    ''' <param name="category"></param>
    ''' <param name="path"></param>
    ''' <param name="domain"></param>
    ''' <returns></returns>
    Public Function ResolveHtmlProduct(ByVal category As Category, ByVal path As ProductPath, ByVal domain As String) As List(Of Product)

        Dim listProds As New List(Of Product)
        Dim myHtmlDom As HtmlDocument = New HtmlDocument
        Dim myProdDom As HtmlNodeCollection

        Dim requestUrl As String = Common.addParamForUrl(category.Url, IIf(path.cateParam Is Nothing, "", path.cateParam))

        Dim crawlerUrls As New List(Of String)
        crawlerUrls.Add(requestUrl)


        If Not String.IsNullOrEmpty(category.BackUpUrls) Then
            Dim backUpUrls As List(Of String) = category.BackUpUrls.Split("^").ToList()
            For Each element As String In backUpUrls
                If Common.IsUrl(element) Then
                    crawlerUrls.Add(element)
                End If
            Next
        End If

        For Each url As String In crawlerUrls

            Dim issueDate As Date
            Dim pageHelper As New HtmlPageUtility()

            myHtmlDom = pageHelper.GetHtmlDocument(url, path.cookie, "refer", path.pageEncoding)
            myProdDom = myHtmlDom.DocumentNode.SelectNodes(path.prodPath)

            If myProdDom Is Nothing Then
                Common.LogText("myHtmlDom.DocumentNode.SelectNodes(item.prodPath) is null where  requestUrl:" & requestUrl)
                NotificationEmail.SentNotificationEmail("autoedm@reasonable.cn", "Please update crawler rule", String.Format("The website structure may have changed, you have to update thr rule,Url={0},siteid={1}", url, category.SiteID))
                'Return listProds
                Continue For
            End If

            For Each inode As HtmlNode In myProdDom
                Try
                    Dim newProd As New Product()
                    Dim SourceString As String
                    '因为某些网站含有奇怪的转义字符比如<a href="https&#58;&#47;&#47;www.vipme.com&#47;clothing_c900027"
                    '因此需要htmldecode()   (-hoyho, 2016 - 3 - 18)
                    If (String.IsNullOrEmpty(path.urlAttri)) Then
                        SourceString = inode.SelectSingleNode(path.urlPath).InnerText.Trim
                        newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim
                    Else
                        If (String.IsNullOrEmpty(path.urlPath)) Then
                            SourceString = inode.GetAttributeValue(path.urlAttri, "").Trim()
                            newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                        Else
                            Dim hnode As HtmlNode = inode.SelectSingleNode(path.urlPath)
                            SourceString = inode.SelectSingleNode(path.urlPath).GetAttributeValue(path.urlAttri, "").Trim
                            newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim()

                        End If
                    End If
                    newProd.Url = Common.addDominForUrl(domain, newProd.Url)
                    If Not (efHelper.IsProductSent(path.siteId, newProd.Url, DateTime.Now.AddDays(0 - path.noRepeatSentDays), DateTime.Now)) Then
                        If Not (String.IsNullOrEmpty(path.productPath1)) Then
                            If (String.IsNullOrEmpty(path.productAttri)) Then
                                SourceString = inode.SelectSingleNode(path.productPath1).InnerText.Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.productPath1).GetAttributeValue(path.productAttri, "").Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.pricePath)) Then
                            If (String.IsNullOrEmpty(path.priceAttri)) Then
                                'priceStr = "US&#36; 51.99" getPriceNum(priceStr)于是price和discount全都变成只提取到36，因此解码
                                Dim priceStr As String = inode.SelectSingleNode(path.pricePath).InnerText.Trim
                                priceStr = Web.HttpUtility.HtmlDecode(priceStr).Trim()
                                newProd.Price = getPriceNum(priceStr)
                            Else
                                SourceString = getPriceNum(inode.SelectSingleNode(path.pricePath).GetAttributeValue(path.priceAttri, "").Trim)
                                newProd.Price = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.discountPath)) Then
                            If (String.IsNullOrEmpty(path.discountAttri)) Then
                                SourceString = inode.SelectSingleNode(path.discountPath).InnerText.Trim
                                SourceString = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                                newProd.Discount = getPriceNum(SourceString)
                            Else
                                SourceString = getPriceNum(inode.SelectSingleNode(path.discountPath).GetAttributeValue(path.discountAttri, "").Trim)
                                newProd.Discount = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.salesPath)) Then
                            If (String.IsNullOrEmpty(path.salesAttri)) Then
                                newProd.Sales = inode.SelectSingleNode(path.salesPath).InnerText.Trim
                            Else
                                SourceString = inode.SelectSingleNode(path.salesPath).GetAttributeValue(path.salesAttri, "").Trim
                                newProd.Sales = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.pictureUrlPath)) Then
                            If (String.IsNullOrEmpty(path.pictureUrlAttri)) Then
                                SourceString = inode.SelectSingleNode(path.pictureUrlPath).InnerText.Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.pictureUrlPath).GetAttributeValue(path.pictureUrlAttri, "").Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                            newProd.PictureUrl = Common.addDominForUrl(domain, newProd.PictureUrl)
                        End If

                        If Not (String.IsNullOrEmpty(path.descriptionPath)) Then
                            If (String.IsNullOrEmpty(path.descriptionAttri)) Then
                                SourceString = inode.SelectSingleNode(path.descriptionPath).InnerText.Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.descriptionPath).GetAttributeValue(path.descriptionAttri, "").Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        newProd.Currency = path.currencyChar
                        If Not (String.IsNullOrEmpty(path.pictureAltPath)) Then
                            If (String.IsNullOrEmpty(path.pictureAltAttri)) Then
                                SourceString = inode.SelectSingleNode(path.pictureAltPath).InnerText.Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.pictureAltPath).GetAttributeValue(path.pictureAltAttri, "").Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If

                        newProd.SiteID = path.siteId
                        newProd.LastUpdate = DateTime.Now
                        If Not (String.IsNullOrEmpty(path.issueDate.Trim)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                            If (String.IsNullOrEmpty(path.issueDateAttri)) Then
                                SourceString = inode.SelectSingleNode(path.issueDate).InnerText.Trim
                                issueDate = Date.Parse(SourceString)
                                newProd.PublishDate = issueDate
                            Else
                                SourceString = inode.SelectSingleNode(path.issueDate).GetAttributeValue(path.issueDateAttri, "").Trim
                                issueDate = Date.Parse(SourceString)
                                newProd.PublishDate = issueDate
                            End If
                            If (Now.Date - issueDate > TimeSpan.FromDays(path.validityPeriod)) Then  '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        Else '抓取数固定
                            If (listProds.Count >= (path.prodDisplayCount * 2)) Then
                                Exit For 'inner for 
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Common.LogText(String.Format("Domain:{0} , Error occurs, skip this product\n", domain))
                    Common.LogText(ex.Message) 'skip this product
                End Try

            Next



            If listProds.Count > path.prodDisplayCount * 2 Then
                Exit For 'outer for
            End If



        Next



        Return listProds

    End Function

    ''' <summary>
    ''' 根据productPath和category规则解析出产品
    ''' </summary>
    ''' <param name="category"></param>
    ''' <param name="path"></param>
    ''' <param name="domain"></param>
    ''' <returns></returns>
    Public Function ResolveHtmlProduct(ByVal category As Category, ByVal path As ProductPath, ByVal domain As String, ByVal getProductCode As Func(Of Product, String)) As List(Of Product)

        If getProductCode = Nothing Then
            Return ResolveHtmlProduct(category, path, domain)
        End If

        Dim listProds As New List(Of Product)
        Dim myHtmlDom As HtmlDocument = New HtmlDocument
        Dim myProdDom As HtmlNodeCollection

        Dim requestUrl As String = Common.addParamForUrl(category.Url, IIf(path.cateParam Is Nothing, "", path.cateParam))

        Dim crawlerUrls As New List(Of String)
        crawlerUrls.Add(requestUrl)

        If Not String.IsNullOrEmpty(category.BackUpUrls) Then
            Dim backUpUrls As List(Of String) = category.BackUpUrls.Split("^").ToList()
            For Each element As String In backUpUrls
                If Common.IsUrl(element) Then
                    crawlerUrls.Add(element)
                End If
            Next
        End If

        For Each url As String In crawlerUrls

            Dim issueDate As Date
            Dim pageHelper As New HtmlPageUtility()
            Dim listProductCode As New List(Of String)

            myHtmlDom = pageHelper.GetHtmlDocument(url, path.cookie, "refer", path.pageEncoding)
            myProdDom = myHtmlDom.DocumentNode.SelectNodes(path.prodPath)

            If myProdDom Is Nothing Then
                Common.LogText("myHtmlDom.DocumentNode.SelectNodes(item.prodPath) is null where  requestUrl:" & requestUrl)
                NotificationEmail.SentNotificationEmail("autoedm@reasonable.cn", "Please update crawler rule", String.Format("The website structure may have changed, you have to update thr rule,Url={0},siteid={1}", url, category.SiteID))
                'Return listProds
                Continue For
            End If

            For Each inode As HtmlNode In myProdDom
                Try
                    Dim newProd As New Product()
                    Dim SourceString As String
                    '因为某些网站含有奇怪的转义字符比如<a href="https&#58;&#47;&#47;www.vipme.com&#47;clothing_c900027"
                    '因此需要htmldecode()   (-hoyho, 2016 - 3 - 18)
                    If (String.IsNullOrEmpty(path.urlAttri)) Then
                        SourceString = inode.SelectSingleNode(path.urlPath).InnerText.Trim
                        newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim
                    Else
                        If (String.IsNullOrEmpty(path.urlPath)) Then
                            SourceString = inode.GetAttributeValue(path.urlAttri, "").Trim()
                            newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                        Else
                            Dim hnode As HtmlNode = inode.SelectSingleNode(path.urlPath)
                            SourceString = inode.SelectSingleNode(path.urlPath).GetAttributeValue(path.urlAttri, "").Trim
                            newProd.Url = Web.HttpUtility.HtmlDecode(SourceString).Trim()

                        End If
                    End If
                    newProd.Url = Common.addDominForUrl(domain, newProd.Url)
                    If Not (efHelper.IsProductSent(path.siteId, newProd.Url, DateTime.Now.AddDays(0 - path.noRepeatSentDays), DateTime.Now)) Then
                        If Not (String.IsNullOrEmpty(path.productPath1)) Then
                            If (String.IsNullOrEmpty(path.productAttri)) Then
                                SourceString = inode.SelectSingleNode(path.productPath1).InnerText.Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.productPath1).GetAttributeValue(path.productAttri, "").Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.pricePath)) Then
                            If (String.IsNullOrEmpty(path.priceAttri)) Then
                                'priceStr = "US&#36; 51.99" getPriceNum(priceStr)于是price和discount全都变成只提取到36，因此解码
                                Dim priceStr As String = inode.SelectSingleNode(path.pricePath).InnerText.Trim
                                priceStr = Web.HttpUtility.HtmlDecode(priceStr).Trim()
                                newProd.Price = getPriceNum(priceStr)
                            Else
                                SourceString = getPriceNum(inode.SelectSingleNode(path.pricePath).GetAttributeValue(path.priceAttri, "").Trim)
                                newProd.Price = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.discountPath)) Then
                            If (String.IsNullOrEmpty(path.discountAttri)) Then
                                SourceString = inode.SelectSingleNode(path.discountPath).InnerText.Trim
                                SourceString = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                                newProd.Discount = getPriceNum(SourceString)
                            Else
                                SourceString = getPriceNum(inode.SelectSingleNode(path.discountPath).GetAttributeValue(path.discountAttri, "").Trim)
                                newProd.Discount = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.salesPath)) Then
                            If (String.IsNullOrEmpty(path.salesAttri)) Then
                                newProd.Sales = inode.SelectSingleNode(path.salesPath).InnerText.Trim
                            Else
                                SourceString = inode.SelectSingleNode(path.salesPath).GetAttributeValue(path.salesAttri, "").Trim
                                newProd.Sales = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(path.pictureUrlPath)) Then
                            If (String.IsNullOrEmpty(path.pictureUrlAttri)) Then
                                SourceString = inode.SelectSingleNode(path.pictureUrlPath).InnerText.Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.pictureUrlPath).GetAttributeValue(path.pictureUrlAttri, "").Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                            newProd.PictureUrl = Common.addDominForUrl(domain, newProd.PictureUrl)
                        End If

                        If Not (String.IsNullOrEmpty(path.descriptionPath)) Then
                            If (String.IsNullOrEmpty(path.descriptionAttri)) Then
                                SourceString = inode.SelectSingleNode(path.descriptionPath).InnerText.Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.descriptionPath).GetAttributeValue(path.descriptionAttri, "").Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If
                        newProd.Currency = path.currencyChar
                        If Not (String.IsNullOrEmpty(path.pictureAltPath)) Then
                            If (String.IsNullOrEmpty(path.pictureAltAttri)) Then
                                SourceString = inode.SelectSingleNode(path.pictureAltPath).InnerText.Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            Else
                                SourceString = inode.SelectSingleNode(path.pictureAltPath).GetAttributeValue(path.pictureAltAttri, "").Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(SourceString).Trim()
                            End If
                        End If

                        newProd.SiteID = path.siteId
                        newProd.LastUpdate = DateTime.Now
                        If Not (String.IsNullOrEmpty(path.issueDate.Trim)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                            If (String.IsNullOrEmpty(path.issueDateAttri)) Then
                                SourceString = inode.SelectSingleNode(path.issueDate).InnerText.Trim
                                issueDate = Date.Parse(SourceString)
                                newProd.PublishDate = issueDate
                            Else
                                SourceString = inode.SelectSingleNode(path.issueDate).GetAttributeValue(path.issueDateAttri, "").Trim
                                issueDate = Date.Parse(SourceString)
                                newProd.PublishDate = issueDate
                            End If
                            If (Now.Date - issueDate > TimeSpan.FromDays(path.validityPeriod)) Then  '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                Dim code As String = getProductCode(newProd)
                                If String.IsNullOrEmpty(code) Then
                                    listProds.Add(newProd)
                                Else
                                    If Not listProductCode.Contains(code) Then
                                        listProds.Add(newProd)
                                        listProductCode.Add(code)
                                    End If
                                End If
                            End If
                        Else '抓取数固定
                            If (listProds.Count >= (path.prodDisplayCount * 2)) Then
                                Exit For 'inner for 
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                Dim code As String = getProductCode(newProd)
                                If String.IsNullOrEmpty(code) Then
                                    listProds.Add(newProd)
                                Else
                                    If Not listProductCode.Contains(code) Then
                                        listProds.Add(newProd)
                                        listProductCode.Add(code)
                                    End If
                                End If

                            End If
                        End If
                    End If
                Catch ex As Exception
                    Common.LogText(String.Format("Domain:{0} , Error occurs, skip this product\n", domain))
                    Common.LogText(ex.Message) 'skip this product
                End Try

            Next



            If listProds.Count > path.prodDisplayCount * 2 Then
                Exit For 'outer for
            End If



        Next



        Return listProds

    End Function


    Public Function ResolveXMLProduct(ByVal category As Category, ByVal path As ProductPath, ByVal domain As String)


        Dim listProds As New List(Of Product)

        Dim requestUrl As String = Common.addParamForUrl(category.Url, IIf(path.cateParam Is Nothing, "", path.cateParam))

        Dim issueDate As Date
        Dim pageHelper As New HtmlPageUtility()

        Dim xmlDoc As XmlDocument = pageHelper.LoadXmlDoc(requestUrl, path.pageEncoding)
        Dim prodXmlNodeList As XmlNodeList = xmlDoc.SelectNodes(path.prodPath)
        Dim SourceString As String
        For Each inode As XmlNode In prodXmlNodeList
            Try
                Dim newProd As New Product()
                newProd.Url = inode.SelectSingleNode(path.urlPath).InnerText.Trim
                newProd.Url = Common.addDominForUrl(domain, newProd.Url)
                If Not (efHelper.IsProductSent(path.siteId, newProd.Url, DateTime.Now.AddDays(0 - path.noRepeatSentDays), DateTime.Now)) Then
                    If Not (String.IsNullOrEmpty(path.productPath1)) Then
                        newProd.Prodouct = inode.SelectSingleNode(path.productPath1).InnerText.Trim
                    End If
                    If Not (String.IsNullOrEmpty(path.pricePath)) Then
                        Try
                            newProd.Price = getPriceNum(inode.SelectSingleNode(path.pricePath).InnerText.Trim)
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(path.discountPath)) Then
                        Try
                            newProd.Discount = getPriceNum(inode.SelectSingleNode(path.discountPath).InnerText.Trim)
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(path.salesPath)) Then
                        Try
                            newProd.Sales = inode.SelectSingleNode(path.salesPath).InnerText.Trim
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(path.pictureUrlPath)) Then
                        Try
                            newProd.PictureUrl = inode.SelectSingleNode(path.pictureUrlPath).InnerText.Trim
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(path.descriptionPath)) Then
                        Try
                            newProd.Description = inode.SelectSingleNode(path.descriptionPath).InnerText.Trim
                        Catch ex As Exception

                        End Try
                    End If

                    newProd.Currency = path.currencyChar
                    If Not (String.IsNullOrEmpty(path.pictureAltPath)) Then
                        Try
                            newProd.PictureAlt = inode.SelectSingleNode(path.pictureAltPath).InnerText.Trim
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(path.issueDate)) Then

                    End If

                    newProd.SiteID = path.siteId
                    newProd.LastUpdate = DateTime.Now

                    If Not (String.IsNullOrEmpty(path.issueDate)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                        If (String.IsNullOrEmpty(path.issueDateAttri)) Then
                            SourceString = inode.SelectSingleNode(path.issueDate).InnerText.Trim
                            issueDate = Date.Parse(SourceString)
                            newProd.PublishDate = issueDate
                        Else
                            SourceString = inode.SelectSingleNode(path.issueDate).Attributes(path.issueDateAttri, "").InnerXml.ToString()
                            issueDate = Date.Parse(SourceString)
                            newProd.PublishDate = issueDate
                        End If
                        If (Now.Date - issueDate > TimeSpan.FromDays(path.validityPeriod)) Then '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                            Exit For
                        Else
                            listProds.Add(newProd)
                        End If
                    Else '抓取数固定
                        If (listProds.Count >= (path.prodDisplayCount * 2)) Then
                            Exit For
                        Else
                            listProds.Add(newProd)
                        End If
                    End If
                End If
            Catch ex As Exception
                Common.LogText("Error occurs,skip this product\n")
                Common.LogText(ex.Message) 'skip this product
            End Try
        Next

        Return listProds

    End Function



    ''' <summary>
    ''' judge if this product duplicate (in the given list )
    ''' </summary>
    ''' <param name="productList"></param>
    ''' <param name="prodcut"></param>
    ''' <returns></returns>
    Private Function isProductDuplicate(ByVal productList As List(Of Product), ByVal prodcut As Product) As Boolean
        Dim isDuplicate As Boolean = False
        For Each p As Product In productList
            If p.Prodouct = prodcut.Prodouct Then
                isDuplicate = True
                Exit For
            End If
        Next
        Return isDuplicate
    End Function



    Private Function getPriceNum(ByVal priceStr As String) As Double
        Dim reg As String = "([0-9.]+)"
        Dim price As Double = Double.Parse(RegularExpressions.Regex.Matches(priceStr, reg)(0).Value())
        Return price
    End Function




End Class
