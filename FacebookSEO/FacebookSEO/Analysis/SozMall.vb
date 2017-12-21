Public Class SozMall

    Private efHelper As New EFHelper()
    Private efContext As New EmailAlerterEntities()

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
            ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String, ByVal preSubject As String)
        Dim category() As String = categories.Split("^")
        'efHelper.GetBanner(siteId, IssueID, Category(0).Trim, planType, url)
        GetProducts(siteId, planType, url, IssueID, preSubject)
        SetSubject(IssueID, siteId, category(1).Trim, planType, preSubject)


    End Sub


    Public Function GetProducts(ByVal siteid As Integer, ByVal planType As String, ByVal domain As String, ByVal issueId As Integer, ByRef preSubject As String)

        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath).ToList()
        For Each item As ProductPath In listProPath
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()
            Dim mylistProduct As List(Of Product) = efHelper.GetListProducts(cate, item, domain)

            For Each product As Product In mylistProduct     '产品格式化
                If product.Description = "0%" Then
                    product.Description = Nothing
                Else
                    product.Description = product.Description & " Off"
                End If
                If product.PictureAlt IsNot Nothing And product.PictureAlt = "$" Then  '借此字段来代替<originalPrice>$2400</originalPrice>
                    product.PictureAlt = Nothing
                End If
                If cate.Category1 = "subject" Then '获取到本期subject
                    preSubject = product.Prodouct
                End If
            Next
            mylistProduct = mylistProduct.OrderByDescending(Function(x) x.PublishDate).ToList()
            '邮件日期部分需要自动更新，当作第一个产品块处理
            If cate.Category1 = "deals" Then
                Dim edmDate As Product = New Product
                edmDate.ShipsImg = System.DateTime.Now.ToString("yyyy年MM月dd日")
                edmDate.Url = "https://www.sozmart.com/c/index.php"
                edmDate.LastUpdate = Now
                mylistProduct.Insert(0, edmDate)
            End If
            Dim listProductId As List(Of Integer) = insertProducts(mylistProduct, cate.Category1, "CA", planType, item.prodDisplayCount, siteid, issueId)
            efHelper.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
        Next
    End Function

    Public Function insertProducts(ByVal listProducts As List(Of Product), ByVal categoryName As String, ByVal section As String, ByVal planType As String,
                 ByVal productCount As Integer, ByVal siteId As Integer,
                 ByVal issueId As Integer) As List(Of Integer)
        Dim allIssuedProducts As List(Of Product) = New List(Of Product)
        Dim query = From p In efContext.Products
                    Where p.SiteID = siteId
                    Select p
        For Each q In query
            allIssuedProducts.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next

        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).Single().CategoryID.ToString()
        For Each li In listProducts
            Dim returnId As Integer = efHelper.InsertOrUpdateProduct(li, Now, categoryId, allIssuedProducts)
            If returnId > 0 Then
                listProductId.Add(returnId)
            End If
            If (listProductId.Count = productCount) Then
                Exit For
            End If
        Next
        Return listProductId
    End Function


    ''' <summary>
    ''' 从数据库Promotion产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Private Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal cateName As String, ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String
        Dim productList As List(Of String) = (From p In efContext.Products
                                              Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                                              Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                                              Order By p.LastUpdate
                                              Select p.Prodouct).ToList()
        If String.IsNullOrEmpty(preSubjdect) Then
            Dim query = From i In efContext.Issues
                        Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                        Select i

            If Not (String.IsNullOrEmpty(productList.Item(0))) Then
                subject = preSubjdect.Replace("[FIRST_PRODUCT]", productList.Item(0))
            Else
                subject = preSubjdect.Replace("[FIRST_PRODUCT]", "")
            End If
            If Not (String.IsNullOrEmpty(productList.Item(1))) Then
                subject = subject.Replace("[SECOND_PRODUCT]", productList.Item(1).Substring(0, 30)) & "..."
            Else
                subject = subject.Replace("[SECOND_PRODUCT]", "")
            End If


            subject = subject.Replace("[VOL_NUMBER]", (query.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
            subject = subject.Replace("[YYYY]", Now.Year.ToString)
            subject = subject.Replace("[MM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))

        Else
            subject = preSubjdect
        End If

        efHelper.InsertIssueSubject(issueId, subject)
    End Sub


End Class
