Imports HtmlAgilityPack

Public Class Siluguoguo
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        If planType.Trim = "HO" Then '热卖
            GetCategory(siteId)
            
            GetHOTProducts(url, "siluguoguo", siteId, planType, IssueID)
            'If (productCounter < 9) Then
            '    Dim categoryUrl As String = "http://octlegendast.tmall.com/p/rd415809.htm"
            '    Dim requestUrl As String = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
            '    GetTopSellProducts(categoryUrl, requestUrl, siteId, planType, IssueID, lastUpdate, 9 - productCounter)
            'End If
            efHelper.InsertContactList(IssueID, "allexceptneteaseopen20140519", "waiting")
            Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct

            Dim chineseCategoryName As String
            Dim newAd As Ad = New Ad()
            Select Case url.Trim
                Case "http://siluguoguo.taobao.com/category-208648255.htm"
                    chineseCategoryName = "养颜红枣"
                    url = "http://siluguoguo.taobao.com/category-208648427.htm"
                    newAd.PictureUrl = "http://app.reasonablespread.com/SpreaderFiles/16577/files/upload/siluguoguo/c1.jpg"
                    newAd.Url = "http://siluguoguo.taobao.com/category-208648255.htm"
                    newAd.Ad1 = "养颜红枣"
                    newAd.Description() = "养颜红枣"
                    newAd.SiteID = siteId
                    newAd.Lastupdate = Now
                Case "http://siluguoguo.taobao.com/category-208648427.htm"
                    chineseCategoryName = "薄皮大核桃"
                    url = "http://siluguoguo.taobao.com/category-284818760.htm"
                    newAd.PictureUrl = "https://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/siluguoguo5.jpg"
                    newAd.Url = "http://siluguoguo.taobao.com/category-208648427.htm"
                    newAd.Ad1 = "薄皮大核桃"
                    newAd.Description() = "薄皮大核桃"
                    newAd.SiteID = siteId
                    newAd.Lastupdate = Now
                Case "http://siluguoguo.taobao.com/category-284818760.htm"
                    chineseCategoryName = "美味坚果"
                    url = "http://siluguoguo.taobao.com/category-208648426.htm"
                    newAd.PictureUrl = "http://app.reasonablespread.com/SpreaderFiles/16577/files/upload/siluguoguo/c2.jpg"
                    newAd.Url = "http://siluguoguo.taobao.com/category-284818760.htm"
                    newAd.Ad1 = "美味坚果"
                    newAd.Description() = "美味坚果"
                    newAd.SiteID = siteId
                    newAd.Lastupdate = Now
                Case "http://siluguoguo.taobao.com/category-208648426.htm"
                    chineseCategoryName = "风干葡萄"
                    url = "http://siluguoguo.taobao.com/category-208648429.htm"
                    newAd.PictureUrl = "http://app.reasonablespread.com/SpreaderFiles/16577/files/upload/siluguoguo/c3.jpg"
                    newAd.Url = "http://siluguoguo.taobao.com/category-208648426.htm"
                    newAd.Ad1 = "风干葡萄"
                    newAd.Description() = "风干葡萄"
                    newAd.SiteID = siteId
                    newAd.Lastupdate = Now
                Case "http://siluguoguo.taobao.com/category-208648429.htm"
                    chineseCategoryName = "绿色果脯"
                    url = "http://siluguoguo.taobao.com/category-208648255.htm"
                    newAd.PictureUrl = "http://app.reasonablespread.com/SpreaderFiles/16577/files/upload/siluguoguo/c4.jpg"
                    newAd.Url = "http://siluguoguo.taobao.com/category-208648429.htm"
                    newAd.Ad1 = "绿色果脯"
                    newAd.Description() = "绿色果脯"
                    newAd.SiteID = siteId
                    newAd.Lastupdate = Now
            End Select

            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId AndAlso c.Category1.Trim = "siluguoguo"
                           Select c
            Dim adid As Integer = efHelper.InsertAd(newAd, efHelper.GetListAd(siteId), queryURL.FirstOrDefault.CategoryID)
            Dim listAdId As New List(Of Integer)
            listAdId.add(adid)
            InsertAdsIssue(siteId, IssueID, listAdId, 1)

            efHelper.InsertIssueSubject(IssueID, "[FIRSTNAME] 你好，本周为你精选 " & chineseCategoryName & "：" & querySubject.ToList.First.Trim)

            '将url更新为下一期邮件要发送的category
            Dim queryCategory As AutomationPlan = efContext.AutomationPlans.Where(Function(c) c.PlanType = planType AndAlso c.SiteID = siteId).Single()
            queryCategory.URL = url
            efContext.SaveChanges()
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "siluguoguo"
        myCategory.Description = "siluguoguo"
        myCategory.Url = "http://siluguoguo.taobao.com"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        'Dim myCategory2 As New Category
        'myCategory2.Category1 = "hetao"
        'myCategory2.Description = "hetao"
        'myCategory2.Url = "http://siluguoguo.taobao.com/category-208648427.htm"
        'myCategory2.SiteID = siteId
        'myCategory2.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(myCategory2, siteId)

        'Dim myCategory3 As New Category
        'myCategory3.Category1 = "jianguo"
        'myCategory3.Description = "jianguo"
        'myCategory3.Url = "http://siluguoguo.taobao.com/category-284818760.htm"
        'myCategory3.SiteID = siteId
        'myCategory3.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(myCategory3, siteId)

        'Dim myCategory4 As New Category
        'myCategory4.Category1 = "putao"
        'myCategory4.Description = "putao"
        'myCategory4.Url = "http://siluguoguo.taobao.com/category-208648426.htm"
        'myCategory4.SiteID = siteId
        'myCategory4.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(myCategory4, siteId)

        'Dim myCategory5 As New Category
        'myCategory5.Category1 = "guopu"
        'myCategory5.Description = "guopu"
        'myCategory5.Url = "http://siluguoguo.taobao.com/category-208648429.htm"
        'myCategory5.SiteID = siteId
        'myCategory5.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(myCategory5, siteId)

    End Sub


    Private Sub GetHOTProducts(ByVal url As String, ByVal categoryName As String, ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer)

        Try
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now

            'Dim maxProductCount As Integer = 34
            iProIssueCount = 6 '在模板中要填充产品的个数
            GetProducts(siteID, issueID, url, categoryName, iProIssueCount, updateTime, planType)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    Private Function GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String,
        ByVal iProIssueCount As Integer, ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)

            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)

            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)
                listProductId.Add(returnId)

                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next

            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Function

    Private Function GetListProducts(ByVal categoryUrl As String, ByVal siteId As Integer, ByVal lastUpdate As DateTime, _
                                          ByVal categoryname As String, ByVal planType As String, ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper()
        Try
            Dim categoryDoc As HtmlDocument = efHelper.GetHtmlDocByUrlTmall(categoryUrl)
            Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div")

            For Each productLine In productLines
                Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
                If Not (productDls Is Nothing) Then
                    For Each dl In productDls

                        Dim myProduct As New Product
                        Dim productName As String = dl.SelectSingleNode("dd[@class='detail']/a").InnerText
                        If productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
                            Continue For
                        End If
                        myProduct.Prodouct = productName
                        myProduct.Url = dl.SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                        If (listProductUrl.Contains(myProduct.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(myProduct.Url)
                        End If

                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-60), DateTime.Now, planType)) Then
                            myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                            Try
                                myProduct.Price = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[2]/span[2]").InnerText.Trim()
                            Catch ex As Exception
                            End Try
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                            myProduct.Description = productName
                            myProduct.LastUpdate = lastUpdate
                            myProduct.SiteID = siteId
                            listProducts.Add(myProduct)
                        End If
                    Next
                End If
                If (listProducts.Count >= iProCount) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when categoryUrl:" & categoryUrl & " categoryName:" & categoryname & ex.StackTrace)
        End Try
        Return listProducts
    End Function

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
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

    Public Sub InsertAdsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal listAdId As List(Of Integer), ByVal iProIssueCount As Integer)
        Dim i As Integer = 0
        For Each li In listAdId
            If i < iProIssueCount Then
                Dim aIssue As New Ads_Issue
                aIssue.AdId = li
                aIssue.SiteId = siteId
                aIssue.IssueID = issueId
                efContext.AddToAds_Issue(aIssue)
                i = i + 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
