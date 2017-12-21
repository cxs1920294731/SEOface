Imports HtmlAgilityPack

Public Class MiaimiTmall
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        If planType.Trim = "HO" Then '热卖
            GetCategory(siteId)
            GetHOTProducts(siteId, planType, IssueID)
            'If (productCounter < 9) Then
            '    Dim categoryUrl As String = "http://octlegendast.tmall.com/p/rd415809.htm"
            '    Dim requestUrl As String = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
            '    GetTopSellProducts(categoryUrl, requestUrl, siteId, planType, IssueID, lastUpdate, 9 - productCounter)
            'End If
            efHelper.InsertContactList(IssueID, {"TaobaoList", "maiaimiqijiandian"}, "waiting")
            Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "[FIRSTNAME] 你好，产品升级，全店清仓，先下手为强啦：" & querySubject.ToList.First.Trim)
        Else '新品推荐邮件
            'GetCategory(siteId, lastUpdate)
            'GetLatestProducts(tmallSearchUrl, siteId, planType, IssueID, lastUpdate)
            'If (productCounter < 9) Then
            '    Dim categoryUrl As String = "http://octlegendast.tmall.com/p/rd415809.htm"
            '    Dim requestUrl As String = "http://octlegendast.tmall.com/category-836408374.htm?orderType=hotsell_desc"
            '    GetTopSellProducts(categoryUrl, requestUrl, siteId, planType, IssueID, lastUpdate, 9 - productCounter)
            'End If
            'helper.InsertContactList(IssueID, "opens_of_ 与C店不重复的名单", "draft")
            'helper.InsertIssueSubject(IssueID, "大女人小情怀，十月传奇春装上新(AD)")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "fengmi"
        myCategory.Description = "fengmi"
        myCategory.Url = "http://bee1.taobao.com/category.htm?viewType=grid&v=1"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Private Sub GetHOTProducts(ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer)
        Dim fengmiUrl As String = ""

        Try
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteID
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "fengmi"
                        fengmiUrl = q.Url
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            iProIssueCount = 6 '在模板中要填充产品的个数
            GetProducts(siteID, issueID, fengmiUrl, "fengmi", iProIssueCount, updateTime, planType)
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
            Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div[@class='item3line1']")
            Dim index As Integer
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
                        If (myProduct.Url.Contains("&")) Then
                            index = myProduct.Url.IndexOf("&")
                            myProduct.Url = myProduct.Url.Remove(index)
                        End If
                        If (listProductUrl.Contains(myProduct.Url)) Then
                            Continue For
                        Else
                            listProductUrl.Add(myProduct.Url)
                        End If

                        If Not (helper.IsProductSent(siteId, myProduct.Url, DateTime.Now.AddDays(-30), DateTime.Now, planType)) Then
                            myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
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

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class
