Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Xml

Public Class Couppie
    Dim efHelper As New EFHelper()
    Private efContext As New EmailAlerterEntities()

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
           ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String,
           ByVal url As String, ByRef categories As String, ByVal preSubject As String)
        Dim category() As String = categories.Split("^")

        If planType.Contains("HA") Then
            efHelper.GetBanner(siteId, IssueID, category(0).Trim, planType, url)
            'efHelper.GetProducts(siteId, planType, url, IssueID)
            'todo //combine 2 category into 1 list
            GetProducts(siteId, planType, url, IssueID, category)
            SetSubject(IssueID, siteId, category(0).Trim, planType, preSubject)
        ElseIf planType.Contains("HP") Then
            category = GetRssProducts(siteId, planType, url, IssueID, category, 7, 6)
            efHelper.GetBanner(siteId, IssueID, category(0).Trim, planType, url)
            SetSubject(IssueID, siteId, category(0).Trim, planType, preSubject)

            Dim helper As New EFHelper()
            Dim cate As String = category(0).Trim
            Dim contactlistCount As Integer
            Dim saveListName1 As String = siteId & cate & Now.ToString("yyyyMMdd")
            contactlistCount = helper.CreateContactList(siteId, IssueID, planType.Trim, cate, saveListName1, 15, ChooseStrategy.Favorite, "draft", spreadLogin, appId)

            'debug# set  contactlistCount = 1
            'contactlistCount = 1
            If (contactlistCount > 0) Then
                helper.InsertContactList(IssueID, saveListName1, "draft")
            End If
        Else
            category = GetRssProducts(siteId, planType, url, IssueID, New String() {}, 7, 3)
            efHelper.GetBanner(siteId, IssueID, category(0).Trim, planType, url)
            SetSubject(IssueID, siteId, category(0).Trim, planType, preSubject)
            categories = String.Join("^", category)
        End If
    End Sub

    ''' <summary>
    ''' 从数据库Promotion产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Private Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal cateName As String, ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String

        'Dim productList As List(Of String) = (From p In efContext.Products
        '                                      Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '                                      Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
        '                                      Order By p.LastUpdate
        '                                      Select p.Prodouct).ToList()


        Dim productList As List(Of String) = GetProductsByCateId(issueId, cateName, siteId)

        Dim query = From i In efContext.Issues
                    Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                    Select i

        If productList.Count > 0 AndAlso Not (String.IsNullOrEmpty(productList.Item(0))) Then
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", productList.Item(0))
        Else
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", "")
        End If
        If productList.Count > 1 AndAlso Not (String.IsNullOrEmpty(productList.Item(1))) Then
            subject = subject.Replace("[SECOND_PRODUCT]", productList.Item(1))
        Else
            subject = subject.Replace("[SECOND_PRODUCT]", "")
        End If

        If productList.Count > 2 AndAlso Not (String.IsNullOrEmpty(productList.Item(1))) Then
            subject = subject.Replace("[THIRD_PRODUCT]", productList.Item(2))
        Else
            subject = subject.Replace("[THIRD_PRODUCT]", "")
        End If
        If productList.Count > 3 AndAlso Not (String.IsNullOrEmpty(productList.Item(1))) Then
            subject = subject.Replace("[FOURTH_PRODUCT]", productList.Item(3))
        Else
            subject = subject.Replace("[FOURTH_PRODUCT]", "")
        End If


        subject = subject.Replace("[VOL_NUMBER]", (query.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
        subject = subject.Replace("[YYYY]", Now.Year.ToString)
        subject = subject.Replace("[MM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))

        efHelper.InsertIssueSubject(issueId, subject)
    End Sub

    Public Function GetProducts(ByVal siteid As Integer, ByVal planType As String, ByVal domain As String, ByVal issueId As Integer, ByVal category As String())
        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Join cate In efContext.Categories On proPath.prodcate Equals cate.CategoryID
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath
                                                    ).OrderBy(Function(x) x.prodcate).ToList()
        Dim allProducts As List(Of Product) = New List(Of Product)
        For Each item As ProductPath In listProPath 'get main category and sub category and combine to 1 list
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()
            Dim listOfProduct As List(Of Product) = New List(Of Product)
            listOfProduct = efHelper.GetListProducts(cate, item, domain)
            'listOfProduct = listOfProduct.OrderByDescending(Function(x) x.Discount).ToList()
            'For i As Integer = listOfProduct.Count - 1 To 0 Step -1
            '    listOfProduct(i).LastUpdate = Now ' set hight price with latest update time 
            '    Threading.Thread.Sleep(10)
            'Next
            If allProducts.Count < 8 Then
                allProducts.AddRange(listOfProduct)
                allProducts = allProducts.Union(listOfProduct).ToList()
            End If

            If InStr(domain, "www.couppie.com", CompareMethod.Text) Then 'couppie order by price
                allProducts = allProducts.OrderByDescending(Function(x) x.Discount).ToList()
                For i As Integer = allProducts.Count - 1 To 0 Step -1
                    allProducts(i).LastUpdate = Now ' set hight price with latest update time 
                    Threading.Thread.Sleep(10)
                Next
            End If

        Next

        If listProPath.Count > 0 Then
            Dim mainPath = listProPath(0) 'main path is 
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = mainPath.prodcate
                                    Select c).FirstOrDefault()
            allProducts(0).ShipsImg = System.DateTime.Now.ToString("yyyy年MM月dd日")
            Dim listProductId As List(Of Integer) = efHelper.insertProducts(allProducts, cate.Category1, "CA", planType, mainPath.prodDisplayCount, siteid, issueId)
            efHelper.InsertProductsIssue(siteid, issueId, "CA", listProductId, mainPath.prodDisplayCount)
        End If

    End Function

    Private Function GetProductsByCateId(ByVal issueid As Integer, ByVal categoryName As String, ByVal siteid As Integer) As List(Of String)
        categoryName = categoryName.Trim
        Dim productlist As List(Of String) = New List(Of String)
        Dim sql As String = "select p.Prodouct from Products p " &
                                            "inner join Products_Issue pi on p.ProdouctID=pi.ProductId " &
                                            "inner join ProductCategory pc on pc.ProductID=p.ProdouctID " &
                                            "inner join Categories c on pc.CategoryID=c.CategoryID " &
                                            "where pi.IssueID=@IssueID and c.Category=@Category and c.SiteID=@SiteID order by p.LastUpdate"
        Dim sqlParam As SqlParameter()
        sqlParam = New SqlParameter() {New SqlParameter With {.ParameterName = "IssueID", .Value = issueid.ToString()},
                                       New SqlParameter With {.ParameterName = "Category", .Value = categoryName.Trim()},
                                       New SqlParameter With {.ParameterName = "SiteID", .Value = siteid.ToString()}}
        Try
            productlist = efContext.ExecuteStoreQuery(Of String)(sql, sqlParam).ToList()
        Catch ex As Exception
            EFHelper.LogText("GetProductsByCateId()-->" & ex.ToString)
        End Try

        Return productlist
    End Function

    Private Function GetRssProducts(ByVal siteid As Integer, ByVal planType As String, ByVal URL As String, ByVal issueId As Integer, ByVal category As String(),
                               ByVal noRepeatSentDays As Integer, ByVal prodDisplayCount As Integer) As String()
        Try
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
            xmlDoc = efHelper.LoadXmlDoc(URL)

            Dim itemNodeList As XmlNodeList = xmlDoc.SelectNodes("//item")

            Dim ChannelTitle As String = System.Web.HttpUtility.HtmlDecode(xmlDoc.SelectSingleNode("//channel").SelectSingleNode("title").InnerText.Trim())
            Dim PubDateString As String = xmlDoc.SelectSingleNode("//channel").SelectSingleNode("pubDate").InnerText.Trim()

            Dim dicProds = GetProductDict(category)
            Dim picRegex = "\<img.+?src\s*=\s*[""'](?<src>.*?)[""']"

            Dim prodXmlNodeList As XmlNodeList = xmlDoc.SelectNodes("//item")
            Dim decodeStr As String = ""
            For Each inode As XmlNode In prodXmlNodeList
                Try
                    Dim cateName As String = inode.SelectSingleNode("category").InnerText.Trim
                    If category.Length = 0 Then
                        If Not dicProds.ContainsKey(cateName) Then
                            dicProds(cateName) = New List(Of Product)()
                        End If
                    ElseIf Not category.Contains(cateName) Then
                        Continue For
                    End If


                    Dim cate As Category = (From c In efContext.Categories
                                            Where c.SiteID = siteid AndAlso c.Category1 = cateName
                                            Select c).FirstOrDefault
                    If cate Is Nothing Then
                        Continue For
                    End If

                    Dim uri = New Uri(cate.Url)
                    Dim domain As String = uri.Scheme & "://" & uri.Host

                    Dim newProd As New Product()
                    newProd.Url = inode.SelectSingleNode("link").InnerText.Trim
                    newProd.Url = Me.efHelper.addDominForUrl(domain, newProd.Url)
                    If Not (Me.efHelper.IsProductSent(siteid, newProd.Url, Date.Now.AddDays(0 - noRepeatSentDays), Date.Now)) Then

                        newProd.Description = inode.SelectSingleNode("title").InnerText.Trim
                        newProd.Prodouct = CutProductName(newProd.Description)
                        Dim match = Regex.Match(inode.SelectSingleNode("description").InnerText, picRegex)
                        If match IsNot Nothing AndAlso match.Groups.Count > 1 Then
                            newProd.PictureUrl = match.Groups("src").Value.Trim
                        Else
                            newProd.PictureUrl = inode.SelectSingleNode("thumbnail").InnerText.Trim
                        End If
                        newProd.PictureUrl = Me.efHelper.addDominForUrl(domain, newProd.PictureUrl)
                        newProd.Price = getPriceNum(inode.SelectSingleNode("orginalPrice").InnerText.Trim)
                        newProd.Discount = getPriceNum(inode.SelectSingleNode("actualPrice").InnerText.Trim)
                        newProd.Sales = inode.SelectSingleNode("noofPurchased").InnerText.Trim

                        Dim newDate As Date
                        If Date.TryParse(inode.SelectSingleNode("pubDate").InnerText.Trim.Replace("HKT", ""),
                                             Globalization.CultureInfo.GetCultureInfo("en-us"), Globalization.DateTimeStyles.None, newDate) Then
                            newProd.PublishDate = newDate
                        End If
                        newProd.SiteID = siteid
                        newProd.LastUpdate = Date.Now

                        If Not isProductDuplicate(dicProds(cate.Category1), newProd) Then
                            dicProds(cate.Category1).Add(newProd)
                        End If

                    End If
                Catch ex As Exception
                    EFHelper.LogText("Error occurs,skip this product\n")
                    EFHelper.LogText(ex.Message) 'skip this product
                End Try
            Next
            Dim ienum = dicProds.OrderByDescending(Function(vp) vp.Value.Count).Take(5)
            For Each vp In ienum
                Dim listProds As List(Of Product) = vp.Value.OrderByDescending(Function(x) x.PublishDate).OrderByDescending(Function(x) x.Discount).ToList()
                Dim listProductId As List(Of Integer) = Me.efHelper.insertProducts(listProds, vp.Key, "CA", planType, prodDisplayCount, siteid, issueId)
                Me.efHelper.InsertProductsIssue(siteid, issueId, "CA", listProductId, prodDisplayCount)
            Next
            Return ienum.Select(Function(vp) vp.Key).ToArray
        Catch ex As Exception
            EFHelper.LogText(ex.Message) 'skip this product
            Throw ex
        End Try
    End Function

    Private Function CutProductName(ByVal name As String) As String
        'Dim chars() As String = {",", ".", "。", "!", "！", ";", "；"}
        If name.Length > 20 Then
            If name.IndexOf("，") > 0 OrElse name.IndexOf(",") > 0 Then
                Dim index = name.IndexOf("，")
                If name.IndexOf(",") > 0 Then
                    index = Math.Min(index, name.IndexOf(","))
                End If
                Return name.Substring(0, index)
            Else
                Return name.Substring(0, 20) & "..."
            End If
        End If
        Return name

    End Function
    Private Function GetProductDict(ByVal cates As String()) As Dictionary(Of String, List(Of Product))
        Dim dic As New Dictionary(Of String, List(Of Product))
        For Each c As String In cates
            dic(c) = New List(Of Product)
        Next
        Return dic
    End Function
    Private Function getPriceNum(ByVal priceStr As String) As Double
        Dim reg As String = "([0-9.]+)"
        Dim price As Double = Double.Parse(Regex.Matches(priceStr, reg)(0).Value())
        Return price
    End Function
    Private Function isProductDuplicate(ByVal productList As List(Of Product), ByVal prodcut As Product) As Boolean
        Dim isDuplicate As Boolean = False
        For Each p As Product In productList
            If p.Prodouct = prodcut.Prodouct AndAlso p.Price = prodcut.Price AndAlso
                IIf(p.Discount Is Nothing, -1, p.Discount) = IIf(prodcut.Discount Is Nothing, -1, prodcut.Discount) Then
                isDuplicate = True
                Exit For
            End If
        Next
        Return isDuplicate
    End Function
End Class
