Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq
Imports System.Text.RegularExpressions

Public Class HKPL
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()


    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        EFHelper.UpdateFbToken()
        If planType.Trim = "HA" Then
            Dim categoryList = GetCategory(siteId, categories)

            Dim accessToken As String = efHelper.GetLongTimeToken(1)

            Dim fbPosts As List(Of Product) = FetchfbPosts("hkleague", 7, accessToken, siteId)

            Dim listOf360 As List(Of Product) = New List(Of Product)
            Dim listOfSchedule As List(Of Product) = New List(Of Product)
            listOfSchedule = GetSchedule("https://www.hkfa.com/ch/leagueyear?type=10&other=2&gpage=1", 5)
            Dim listOfTable As List(Of Product) = New List(Of Product) '积分榜
            Dim ls1result As List(Of Product) = New List(Of Product)
            'Dim ls2result As List(Of Product) = New List(Of Product)
            'Dim ls3result As List(Of Product) = New List(Of Product)
            'Dim ls4result As List(Of Product) = New List(Of Product)
            'Dim ls5result As List(Of Product) = New List(Of Product)
            Dim resultCount = 0

            Dim lsYoutubeActivity As List(Of Product) = GetYoutubeRecentAlbum("https://www.youtube.com/channel/UCiFY7cDoGAdc757dQNuqX-g")

            For Each Prod As Product In fbPosts 'GET Product and store in different list
                If Prod.ExpiredDate > Now.AddDays(-7) Then
                    If InStr(Prod.Description, "#足總360", CompareMethod.Text) AndAlso listOf360.Count < 1 Then '#港超360
                        listOf360.Add(Prod)
                        ShortenPostDescription(listOf360)
                    ElseIf InStr(Prod.Description, "#港超積分榜", CompareMethod.Text) AndAlso listOfTable.Count < 1 Then '#港超積分榜
                        listOfTable.Add(Prod)
                        ShortenPostDescription(listOfTable)
                    ElseIf InStr(Prod.Description, "#港超賽果", CompareMethod.Text) AndAlso resultCount < 1 Then '#港超賽果
                        ls1result.Add(Prod)
                        ShortenPostDescription(ls1result)
                        'Dim multiPictureUrl() As String = Prod.PictureUrl.Split(";")
                        'Select Case multiPictureUrl.Count()
                        '    Case 1
                        '        ls1result.Add(Prod)
                        '    Case 2
                        '        ls2result.Add(Prod)
                        '    Case 3
                        '        ls3result.Add(Prod)
                        '    Case 4
                        '        ls4result.Add(Prod)
                        '    Case 5
                        '        ls5result.Add(Prod)
                        'End Select
                        'resultCount += 1
                        ' ShortenPostDescription(listOfResult)
                        'ElseIf InStr(Prod.Description, "#港超賽程", CompareMethod.Text) AndAlso listOfSchedule.Count < 1 Then '360 (Prod.Description.Contains("#港超賽程") OrElse Prod.Description.IndexOf("#港超賽程") > 0) AndAlso listOfSchedule.Count < 1 
                        '    listOfSchedule.Add(Prod)
                        'ShortenPostDescription(listOfSchedule)
                    End If
                End If
            Next


            For Each cateStr As String In categoryList 'Save list to database
                Dim listProudctid As List(Of Integer)
                Select Case cateStr
                    Case "360"
                        listProudctid = efHelper.insertProducts(listOf360, cateStr, "CA", planType, 1, siteId, IssueID)
                        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                    Case "Schedule"
                        listProudctid = efHelper.insertProducts(listOfSchedule, cateStr, "CA", planType, 5, siteId, IssueID)
                        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 5)
                    Case "Table"
                        listProudctid = efHelper.insertProducts(listOfTable, cateStr, "CA", planType, 1, siteId, IssueID)
                        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                    Case "1Result"
                        listProudctid = efHelper.insertProducts(ls1result, cateStr, "CA", planType, 1, siteId, IssueID)
                        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                    Case "Youtube"
                        listProudctid = efHelper.insertProducts(lsYoutubeActivity, cateStr, "CA", planType, 1, siteId, IssueID)
                        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                        'Case "2Result"
                        '    listProudctid = efHelper.insertProducts(ls2result, cateStr, "CA", planType, 1, siteId, IssueID)
                        '    efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                        'Case "3Result"
                        '    listProudctid = efHelper.insertProducts(ls3result, cateStr, "CA", planType, 1, siteId, IssueID)
                        '    efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                        'Case "4Result"
                        '    listProudctid = efHelper.insertProducts(ls4result, cateStr, "CA", planType, 1, siteId, IssueID)
                        '    efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)
                        'Case "5Result"
                        '    listProudctid = efHelper.insertProducts(ls5result, cateStr, "CA", planType, 1, siteId, IssueID)
                        '    efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 1)

                End Select
            Next

            Dim querySubject As String = (From p In efContext.Products
                                          Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                                          Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                                          Select p.PictureAlt).FirstOrDefault()
            efHelper.InsertIssueSubject(IssueID, "2017-18香港超級聯賽 - 每週速遞")
        End If
    End Sub



    Private Function GetCategory(ByVal siteId As Integer, ByRef categories As String) As String()

        Dim helper As New EFHelper
        Dim category() As String = categories.Split("^")
        For Each cate As String In category
            Dim fbCategory As New Category
            'myCategory.EntityKey = Nothing
            fbCategory.Category1 = cate
            fbCategory.Description = cate + " added by program"
            fbCategory.Url = "https://www.facebook.com/hkleague/timeline?" & cate
            fbCategory.SiteID = siteId
            fbCategory.LastUpdate = Now
            helper.InsertOrUpdateCate(fbCategory, siteId)
        Next
        Return category
        'myCategory.Category1 = "posts"
        'myCategory.Description = "Schedule,Table adn Result"
        'myCategory.Url = "https://www.facebook.com/hkleague/timeline"
        'myCategory.SiteID = siteId
        'myCategory.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(myCategory, siteId)
    End Function

    Public Function ShortenPostDescription(ByRef listPost As List(Of Product))
        For Each item In listPost
            If (item.Description.Length > 150) Then
                item.Description = item.Description.Substring(0, 150) & "..."
            End If
        Next
    End Function


    Public Function insertProducts(ByVal listProducts As List(Of Product), ByVal categoryName As String, ByVal section As String, ByVal planType As String,
                     ByVal productCount As Integer, ByVal siteId As Integer,
                     ByVal issueId As Integer) As List(Of Integer)
        Dim listProduct As List(Of Product) = New List(Of Product)
        Dim query = From p In efContext.Products
                    Where p.SiteID = siteId
                    Select p
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next

        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).Single().CategoryID.ToString()
        For Each li In listProducts
            Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

            listProductId.Add(returnId)
            If (listProductId.Count = productCount) Then
                Exit For
            End If
        Next
        Return listProductId
    End Function

#Region "facebook"
    ''' <summary>
    ''' 获取fb的一个page上发布的post，并返回List(Of Product)
    ''' </summary>
    ''' <param name="fbPageName">fb的page名</param>
    ''' <param name="timeLimit">指定某个时间之后的所有post</param>
    ''' <param name="accessToken"></param>
    ''' <param name="siteID"></param>
    ''' <returns>List(Of Product)</returns>
    ''' <remarks></remarks>
    Public Function FetchfbPosts(ByVal fbPageName As String, ByVal timeLimit As Integer, ByVal accessToken As String, ByVal siteID As Integer) As List(Of Product)
        Dim listProduct As New List(Of Product)
        Dim timeStamp = DateDiff("s", "01/01/1970 00:00:00", DateTime.UtcNow.AddDays(-timeLimit)) '计算now-timeLimit（即x天前）对应的时间戳 ,具体到天或者时，不能是分秒，否则结果为空
        'Dim timeStamp = CLng(DateTime.Now.AddDays(-5).Subtract(New DateTime(1970, 1, 1)).TotalMilliseconds) '计算now-timeLimit（即x天前）对应的时间戳 ,具体到天或者时，不能是分秒，否则结果为空
        Dim requestUrl As String = "https://graph.facebook.com/v2.3/" & fbPageName & "/posts/" & "?access_token=" & accessToken & "&since=" & timeStamp & "&format=json"
        Dim postsStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")
        Dim postsJson As JObject = JObject.Parse(postsStr)
        Dim postsJArr As JArray = postsJson("data")
        For i As Integer = 0 To postsJArr.Count - 1
            Dim item As JObject = postsJArr(i)
            Dim myPro As New Product()
            Dim title As String
            If (item.TryGetValue("message", myPro.Description)) Then
                Try
                    If myPro.Description.Contains(vbLf) Then
                        title = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                    Else
                        Dim startIndex As Integer = myPro.Description.IndexOf("【")
                        Dim endIndex As Integer = myPro.Description.IndexOf("】")
                        title = myPro.Description.Substring(startIndex, endIndex - startIndex + 1)
                    End If

                Catch ex As Exception
                    title = "Fail to get Title, Missing '【' or '】'?"
                End Try
                myPro.Prodouct = title
            Else
                Continue For
            End If

            Dim photoid As String = ""

            If Not (item.TryGetValue("id", photoid)) Then
                Continue For
            End If


            Dim photoJson As JObject = GetfbPhotobyid(photoid, accessToken)
            If Not (photoJson.TryGetValue("full_picture", myPro.PictureUrl)) Then
                'Continue For
            End If
            photoJson.TryGetValue("permalink_url", myPro.Url)


            Dim id As String = ""
            If Not (item.TryGetValue("id", id)) Then
                Continue For
            End If

            Dim attachmentsJson As JObject = GetAttachmentsbyid(id, accessToken)
            If attachmentsJson("data").Count > 0 Then
                attachmentsJson = attachmentsJson("data")(0)
                If attachmentsJson.TryGetValue("subattachments", "") Then '包含键subattachments,多图
                    Dim imgStr As String = ""
                    Dim imgTarget As String = "" '图片跳转链接，既不是post也不是图源，而是一个弹出式图片链
                    Dim photosJArr As JArray = attachmentsJson("subattachments")("data")
                    For index As Integer = 0 To photosJArr.Count - 1
                        If index < 5 Then
                            Dim json As JObject = photosJArr(index)
                            imgStr = imgStr & json.SelectToken("media.image.src").ToString() & ";"
                            imgTarget = imgTarget & json.SelectToken("target.url").ToString() & ";"
                        End If
                    Next
                    myPro.PictureUrl = imgStr.TrimEnd(";")
                    myPro.FreeShippingImg = imgTarget.TrimEnd(";") '借FreeShippingImg字段存储target链接
                Else '只有一图
                    myPro.PictureUrl = attachmentsJson.SelectToken("media.image.src")
                    myPro.FreeShippingImg = attachmentsJson.SelectToken("target.url")
                End If
            End If

            item.TryGetValue("created_time", myPro.ExpiredDate)



            'Dim rmatch As Match = Regex.Match(myPro.Description, "http(?:|s):\S+\s?")
            'If (rmatch.Success) Then
            '    myPro.Prodouct = rmatch.Value.Trim 'post文中的超链接
            'Else
            '    myPro.Prodouct = myPro.Url
            'End If

            myPro.Description = efHelper.AddfbPosttextHyperlink(myPro.Description)
            If (myPro.Description.Contains(vbLf)) Then
                myPro.PictureAlt = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                myPro.Description = myPro.Description.Remove(0, myPro.Description.IndexOf(vbLf)).TrimStart(vbLf)
            End If
            myPro.Description = myPro.Description.Replace(vbLf, "</br>")
            myPro.Description = myPro.Description
            myPro.SiteID = siteID
            myPro.LastUpdate = DateTime.Now
            Dim temp() As String = {}
            If (Not String.IsNullOrEmpty(myPro.PictureUrl)) Then
                temp = myPro.PictureUrl.Split(";").ToArray()
            End If

            Dim shareLink As String = "http://www.facebook.com/sharer.php?s=100&p[url]=" & myPro.Url ' & "&p[images][0]=" & myPro.PictureUrl
            For c As Integer = 0 To temp.Count() - 1
                shareLink = shareLink & "&p[images][" & c.ToString & "]" & temp(c)
            Next
            ' Dim shareLink As String = "http://www.facebook.com/sharer.php?s=100&p[url]=" & myPro.Url & "&p[images][0]=" & myPro.PictureUrl
            myPro.ShipsImg = shareLink

            If Not (efHelper.IsProductSent(siteID, myPro.Url, Now.AddDays(-7), Now)) Then
                listProduct.Add(myPro)
            End If 'recovery when public
            listProduct.Add(myPro)
        Next
        Return listProduct
    End Function

    ''' <summary>
    ''' 根据id获取所有photo的数据
    ''' </summary>
    ''' <param name="id">相片的id</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以jObject返回相片数据</returns> 
    ''' <remarks>method=GET&path=ID/attachments&version=v2.7</remarks>
    Public Function GetfbPhotobyid(ByVal id As String, ByVal accessToken As String) As JObject
        'Dim requestUrl As String = "https://graph.facebook.com/v2.3/" & id & "/attachments/" & "?access_token=" & accessToken & "&format=json"
        Dim requestUrl As String = "https://graph.facebook.com/v2.10/" & id & "?access_token=" & accessToken & "&fields=full_picture,permalink_url&format=json"
        Dim imageStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")
        Dim imageJson As JObject = JObject.Parse(imageStr)
        Return imageJson
    End Function


    ''' <summary>
    ''' 根据id获取所有attachments的数据
    ''' </summary>
    ''' <param name="id">相片的id</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以jObject返回相片数据</returns> 
    ''' <remarks>method=GET&path=ID/attachments&version=v2.7</remarks>
    Public Function GetAttachmentsbyid(ByVal id As String, ByVal accessToken As String) As JObject
        Dim requestUrl As String = "https://graph.facebook.com/v2.3/" & id & "/attachments/" & "?access_token=" & accessToken & "&format=json"
        'Dim requestUrl As String = "https://graph.facebook.com/v2.10/" & id & "?access_token=" & accessToken & "&fields=full_picture,permalink_url&format=json"
        Dim imageStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")
        Dim imageJson As JObject = JObject.Parse(imageStr)
        Return imageJson
    End Function
#End Region


    Public Function GetYoutubeRecentAlbum(ByVal PageUrl) As List(Of Product)
        Dim listOfActivity As List(Of Product) = New List(Of Product)
        Dim youtubeActivity As Product = New Product

        Try
            Dim userPageHtml As String = EFHelper.GetHtmlStringByUrl(PageUrl, "", "", "")
            Dim userPageDoc As HtmlDocument = New HtmlDocument()
            userPageDoc.LoadHtml(userPageHtml)
            Dim targetArea As HtmlNodeCollection = userPageDoc.DocumentNode.SelectNodes("//*[@id='browse-items-primary']/li")

            For Each li As HtmlNode In targetArea

                Try
                    youtubeActivity.PictureUrl = li.SelectSingleNode("div[2]/div[2]/div[2]/div/div[1]/div[1]/span/a/span/span/span/img").GetAttributeValue("data-thumb", "")
                    youtubeActivity.Prodouct = li.SelectSingleNode("div[2]/div[2]/div[2]/div/div[1]/div[2]/h3/a").GetAttributeValue("title", "")
                    If Not String.IsNullOrEmpty(youtubeActivity.PictureUrl) AndAlso youtubeActivity.PictureUrl.IndexOf("?") > 0 Then
                        youtubeActivity.PictureUrl = youtubeActivity.PictureUrl.Substring(0, youtubeActivity.PictureUrl.IndexOf("?"))
                    End If
                    youtubeActivity.Url = "https://www.youtube.com/" & li.SelectSingleNode("div[2]/div[2]/div[2]/div/div[1]/div[1]/span/a").GetAttributeValue("href"， "")
                    listOfActivity.Add(youtubeActivity)
                    Exit For 'GET THE FIRST ONE
                Catch ex As Exception
                    Common.LogText("can not get youtube activity cover or anchor" & ex.ToString())
                End Try


            Next

        Catch ex As Exception
            Common.LogText("GetYoutubeRecentAlbum()-->" & ex.ToString())
        End Try

        Return listOfActivity


    End Function


    Public Function GetSchedule(ByVal PageUrl As String, ByVal count As Integer) As List(Of Product)
        Dim listOfSchedule As List(Of Product) = New List(Of Product)
        Try
            Dim userPageHtml As String = EFHelper.GetHtmlStringByUrl(PageUrl, "", "", "")
            Dim userPageDoc As HtmlDocument = New HtmlDocument()
            userPageDoc.LoadHtml(userPageHtml)
            Dim targetArea As HtmlNodeCollection = userPageDoc.DocumentNode.SelectNodes("//tr[@class='trs']")

            For Each li As HtmlNode In targetArea
                Dim product As Product = New Product
                Try
                    'user simily filed to store information
                    product.Prodouct = li.SelectSingleNode("td[4]").InnerText  '主 - 對賽隊伍 - 客
                    Dim iconUrl As String = ""
                    If li.SelectSingleNode("td[4]/img") IsNot Nothing Then
                        iconUrl = li.SelectSingleNode("td[4]/img").GetAttributeValue("src", "")
                        iconUrl = efHelper.addDominForUrl("https://www.hkfa.com/", iconUrl)
                        Dim isDisplay = li.SelectSingleNode("td[4]/img").GetAttributeValue("style", "")
                        If isDisplay.Contains("display:none") Then '隐藏直播icon
                            iconUrl = ""
                        End If
                    End If
                    If Not String.IsNullOrEmpty(iconUrl) Then '如果有直播icon，集成到文字后面
                        product.Prodouct = product.Prodouct & "<br>  <img src='" & iconUrl & "' height='16' border='0' alt='東網直播' />"
                    End If
                    product.Sales = li.SelectSingleNode("td[3]").InnerText '週	
                    product.Url = "https://www.hkfa.com/ch/leagueyear?type=10&other=2&gpage=1" & "&p=" & product.Prodouct 'to make sure url is different
                    product.Description = li.SelectSingleNode("td[1]").InnerText  '日期
                    product.PictureUrl = li.SelectSingleNode("td[7]").InnerText '时间
                    'product.PictureUrl = li.SelectSingleNode("td[5]").InnerText  '賽事
                    product.PictureAlt = li.SelectSingleNode("td[6]/a").InnerText  '場地
                    product.FreeShippingImg = li.SelectSingleNode("td[8]").InnerText '$


                    If listOfSchedule.Count < count Then
                        listOfSchedule.Add(product)
                    Else
                        Dim time As Date
                        If Date.TryParse(listOfSchedule(0).Description, time) Then
                            If time < Now.Date Then
                                listOfSchedule.RemoveAt(0)
                                listOfSchedule.Add(product)
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    End If
                Catch ex As Exception
                    Common.LogText("can not get youtube activity cover or anchor" & ex.ToString())
                End Try


            Next

        Catch ex As Exception
            Common.LogText("GetYoutubeRecentAlbum()-->" & ex.ToString())
        End Try

        Return listOfSchedule


    End Function


    Public Structure YouTuBeActivity
        Dim image As String
        Dim anchor As String
    End Structure



End Class
