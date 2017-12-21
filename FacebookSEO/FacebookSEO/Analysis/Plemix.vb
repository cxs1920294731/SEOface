
Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports System.Web.Script.Serialization
Imports System.Xml

Public Class Plemix
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    'Private subjectString As String
    Private listProductUrl As New List(Of String)

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        EFHelper.UpdateFbToken()
        'GetCategory(siteId)
        GetCategoryProducts(siteId, IssueID, planType)

        Dim accessToken As String = efHelper.GetLongTimeToken(1)
        Dim listProduct As List(Of Product) = efHelper.FetchfbPosts("plemixcom", 5, accessToken, siteId)
        Dim listProudctid As New List(Of Integer)
        listProudctid = efHelper.insertProducts(listProduct, "smartPhone", "NE", planType, 3, siteId, IssueID)
        efHelper.InsertProductsIssue(siteId, IssueID, "NE", listProudctid, 3)

        'Dim subject As String = efHelper.GetFirstProductSubject(IssueID, "NE")

        Dim product = (From p In efContext.Products
                       Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                       Where (pIssue.IssueID = IssueID And pIssue.SectionID = "NE")
                       Select p.PictureAlt).FirstOrDefault()
        Dim subject As String = product
        If (String.IsNullOrEmpty(subject)) Then
            subject = EFHelper.GetFirstProductSubject(IssueID, "CA")
        End If

        efHelper.InsertIssueSubject(IssueID, "Hi [FIRSTNAME], Plemix New Arrivals For You:" & subject & " ...")
    End Sub


    Public Sub GetCategoryProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal planType As String)
        Dim phoneUrl As String = ""
        Dim tabletsUrl As String = ""
        Dim camerasUrl As String = ""

        Try
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "smartPhone"
                        phoneUrl = q.Url.Trim()
                    Case "Tablets"
                        tabletsUrl = q.Url.Trim()
                    Case "Cameras"
                        camerasUrl = q.Url.Trim()
                End Select
            Next

            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 3
            GetProducts(siteId, Issueid, phoneUrl, "smartPhone", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, tabletsUrl, "Tablets", iProIssueCount, updateTime, planType)
            GetProducts(siteId, Issueid, camerasUrl, "Cameras", iProIssueCount, updateTime, planType)

        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub

    

    

    Public Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                          ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product)
            listProducts = GetListProducts(categoryUrl, siteId, categoryName, planType, iProIssueCount)

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
    End Sub

    Private Function GetListProducts(ByVal categoryUrl As String, ByVal siteId As Integer, _
                                          ByVal categoryname As String, ByVal planType As String, ByVal iProCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper()

        Try
            Dim categoryDoc As New XmlDocument()
            categoryDoc.LoadXml(GetHtmlString(categoryUrl))
            Dim itemList As XmlNodeList = categoryDoc.SelectNodes("//channel/item")
            Dim description As String
            Dim myHtmlDocument As New HtmlDocument()
            For Each item As XmlNode In itemList
                Dim product As New Product
                product.Url = item.SelectSingleNode("link").InnerText.ToString.Trim
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If
                product.Prodouct = item.SelectSingleNode("title").InnerText.ToString.Trim
                product.Description = product.Prodouct

                description = item.SelectSingleNode("description").InnerText.ToString
                myHtmlDocument.LoadHtml(description)
                Try
                    product.Discount = myHtmlDocument.DocumentNode.SelectSingleNode("//div/p[@class='old-price']/span[@class='price']").InnerText.Replace("$", "")
                    product.Price = myHtmlDocument.DocumentNode.SelectSingleNode("//div/p[@class='special-price']/span[@class='price']").InnerText.Replace("$", "")
                Catch ex As Exception
                    product.Discount = myHtmlDocument.DocumentNode.SelectSingleNode("//div").InnerText.Replace("$", "")
                    product.Price = myHtmlDocument.DocumentNode.SelectSingleNode("//div").InnerText.Replace("$", "")
                End Try
                product.Currency = "$"
                product.PictureUrl = myHtmlDocument.DocumentNode.SelectSingleNode("//img").GetAttributeValue("src", "")
                product.LastUpdate = Now
                product.SiteID = siteId
                If (Not efHelper.IsProductSent(siteId, product.Url, Now.AddDays(-30), Now, planType)) Then '控制3个月内获取的产品不重复,分类最少的产品数为47个，够发3个月
                    listProducts.Add(product)
                End If
                If (listProducts.Count > iProCount) Then
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

    Public Function GetHtmlString(ByVal pageUrl As String) As String
        Dim ressting As String = ""
        Try
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "en-US,en;q=0.8")
            request.Referer = "https://www.facebook.com/plemixcom/info"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()

            Dim resStreamReader As StreamReader = New StreamReader(resStream, System.Text.Encoding.GetEncoding("utf-8"))
            ressting = resStreamReader.ReadToEnd()
        Catch ex As Exception
            LogText(ex.StackTrace)
        End Try
        Return ressting
    End Function

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub


End Class
