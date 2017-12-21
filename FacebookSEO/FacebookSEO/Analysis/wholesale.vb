Imports System.Data.SqlClient

Public Class wholesale

    Dim efHelper As New EFHelper()
    Private efContext As New Analysis.EmailAlerterEntities()

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal spreadLogin As String, ByVal APIkey As String,
           ByVal siteUrl As String, ByVal categories As String, ByVal preSubject As String, ByVal utmCode As String)
        Dim category() As String = categories.Split("^")

        efHelper.GetBanner(siteId, IssueID, category(0).Trim, planType, siteUrl)
        '#debug#
        'efHelper.GetProducts(siteId, planType, siteUrl, IssueID)
        'GetProducts(siteId, planType, siteUrl, IssueID, preSubject)
        Dim productDictionary As Dictionary(Of String, List(Of Product)) = GetProductDict(siteId, planType, category, siteUrl, IssueID)


        Dim extraProducts As Dictionary(Of String, List(Of Product)) = New Dictionary(Of String, List(Of Product))
        Dim HotProductList As List(Of Product) = New List(Of Product)
        Dim IsSpecialCollection As Boolean = False

        'move item from productDictionary to extraProducts where item's name contain ""#ignore#"
        Dim keys As String() = productDictionary.Keys.ToArray()
        For Each categoryName As String In keys

            If categoryName.Contains("#ignore#") Then
                IsSpecialCollection = True
                extraProducts.Add(categoryName, productDictionary(categoryName))
                productDictionary.Remove(categoryName)

            End If
        Next



        'combine to one category
        For Each key As String In extraProducts.Keys
            Extract(extraProducts(key), HotProductList, 4)
        Next

        If IsSpecialCollection Then '标签包含#ignore#的为特殊用户，需要将所有#ignore#分类合并成一个#specialcollection#
            productDictionary.Add("#specialcollection#", HotProductList)

        End If

        'blouses--http://www.clothing-marketing.com/women-s-blouses-c2105
        'hair-accessories--http://www.clothing-marketing.com/hair-accessories-c2148
        'necklaces--http://www.clothing-marketing.com/necklaces-c2152
        'kids-tops-http://www.clothing-marketing.com/kids-tops-c2174
        'cosplay--http://www.clothing-marketing.com/cosplay-costumes-c2201
        'baby-dolls--http://www.clothing-marketing.com/baby-dolls-c2206
        'muslim-dresses--http://www.clothing-marketing.com/dresses-c2504
        'women-tshirts--http://www.clothing-marketing.com/women-s-tees-t-shirts-c2104
        'men-s-shirts--http://www.clothing-marketing.com/men-s-shirts-c2198
        'UnderFive--http://www.clothing-marketing.com/channel-UnderFiveListView-1
        'MomBabyListView---http://www.clothing-marketing.com/channel-MomBabyListView-right-474
        'women-s-sleepwear--http://www.clothing-marketing.com/women-s-sleepwear-c2114


        'Dim extraCategories As String = "#blouses^hair-accessories^necklaces^kids-tops^cosplay^baby-dolls^muslim-dresses^women-tshirts^men-s-shirts^UnderFive^MomBabyListView^women-s-sleepwea"
        'Dim extraCategory() As String = extraCategories.Split("^")
        'Dim extraProducts As Dictionary(Of String, List(Of Product)) = GetProductDict(siteId, planType, extraCategory, siteUrl, IssueID)

        'Dim HotProductList As List(Of Product) = New List(Of Product)

        'extraProducts(dictionary) --> HotProductList



        '把普通商品插入issue表
        For Each categoryName As String In productDictionary.Keys


            Dim productPath As ProductPath = (From proPath In efContext.ProductPaths
                                              Join cate In efContext.Categories
                                                           On proPath.prodcate Equals cate.CategoryID
                                              Where proPath.siteId = siteId AndAlso proPath.planType = planType AndAlso cate.Category1 = categoryName
                                              Select proPath).FirstOrDefault


            Dim productIds As List(Of Integer) = insertProducts(productDictionary(categoryName), categoryName, "CA", planType, productPath.prodDisplayCount, siteId, IssueID)
            'efHelper.InsertProductsIssue(siteId, IssueID, "CA", productIds, productPath.prodDisplayCount)
            ProductIssueDAL.InsertProductsIssue(siteId, IssueID, "CA", productIds, productPath.prodDisplayCount, productPath.prodcate)
        Next





        Dim hotIssueCount As Integer = 3 '热品从最近3期抓取
        Dim prodDisplayCount As Integer = 8 '每次8个往期热品
        Dim startTime As DateTime = Now.AddDays(-21)
        Dim endTime As DateTime = Now

        If category.Contains("hotclick") Then
            Dim recentHotProduct As List(Of Product) = GetRecentHotClickProduct(siteId, planType, spreadLogin, APIkey, hotIssueCount, prodDisplayCount, startTime, endTime, utmCode)
            Dim listProductId As List(Of Integer) = insertProducts(recentHotProduct, "hotclick", "CA", planType, prodDisplayCount, siteId, IssueID)

            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteId AndAlso c.Category1 = "hotclick"
                                    Select c).FirstOrDefault()

            ProductIssueDAL.InsertProductsIssue(siteId, IssueID, "CA", listProductId, prodDisplayCount, cate.CategoryID)
            'efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProductId, prodDisplayCount)
        End If




        Dim cateProductList As List(Of String) = GetProductsByCateId(IssueID, category(0).Trim, siteId)
        efHelper.SetSubject(IssueID, siteId, category(0).Trim, cateProductList, planType, preSubject)
        'SetSubject(IssueID, siteId, category(0).Trim, planType, preSubject)

        If (planType.Contains("HP")) Then
            Dim plan As AutomationPlan = (From autoplan In efContext.AutomationPlans
                                          Where autoplan.SiteID = siteId AndAlso autoplan.PlanType = planType
                                          Select autoplan).FirstOrDefault()
            Dim triggerIndex As Integer
            If plan.TriggerForNFC.HasValue Then
                triggerIndex = plan.TriggerForNFC
            Else
                triggerIndex = 0
            End If

            Dim cate As String = category(triggerIndex).Trim
            Dim contactlistCount As Integer
            Dim saveListName1 As String = cate & Now.ToString("yyyyMMdd-HHMMss")
            contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate, saveListName1, 90, ChooseStrategy.Favorite, "draft", spreadLogin, APIkey)
            '!debug# account return 0,so modify its values to enter debug mode
            'contactlistCount = 1
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssueID, saveListName1, "draft")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 从数据库产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Private Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal cateName As String, ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String

        Dim productList As List(Of String) = (From p In efContext.Products
                                              Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                                              Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                                              Order By p.LastUpdate Descending
                                              Select p.Prodouct).ToList()

        Dim query = From i In efContext.Issues
                    Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                    Select i

        If Not (String.IsNullOrEmpty(productList.Item(0))) Then
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", productList.Item(0))
        Else
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", "")
        End If
        If Not (String.IsNullOrEmpty(productList.Item(1))) Then
            subject = subject.Replace("[SECOND_PRODUCT]", productList.Item(1))
        Else
            subject = subject.Replace("[SECOND_PRODUCT]", "")
        End If


        subject = subject.Replace("[VOL_NUMBER]", (query.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
        subject = subject.Replace("[YYYY]", Now.Year.ToString)
        subject = subject.Replace("[MM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))
        Dim debug As String = Now.ToString("MMM", New Globalization.CultureInfo("en-US"))
        subject = subject.Replace("[MMM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))

        efHelper.InsertIssueSubject(issueId, subject)
    End Sub





    Private Function GetProductsByCateId(ByVal issueid As Integer, ByVal categoryName As String, ByVal siteid As Integer) As List(Of String)
        categoryName = categoryName.Trim
        Dim productlist As List(Of String) = New List(Of String)
        'Dim sql As String = "select p.Prodouct from Products p " &
        '                                    "inner join Products_Issue pi on p.ProdouctID=pi.ProductId " &
        '                                    "inner join ProductCategory pc on pc.ProductID=p.ProdouctID " &
        '                                    "inner join Categories c on pc.CategoryID=c.CategoryID " &
        '                                    "where pi.IssueID=@IssueID and c.Category=@Category and c.SiteID=@SiteID order by p.LastUpdate"
        Dim sql As String = "select p.Prodouct from Products p " &
                                                "inner join Products_Issue pi on p.ProdouctID=pi.ProductId " &
                                                "inner join Categories c on pi.CategoryID=c.CategoryID  " &
                                                "where pi.IssueID=@IssueID and c.Category=@Category and c.SiteID=@SiteID order by p.LastUpdate "

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


    Public Function GetProducts(ByVal siteid As Integer, ByVal planType As String, ByVal domain As String, ByVal issueId As Integer, ByRef preSubject As String)

        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath).ToList()
        For Each item As ProductPath In listProPath
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()
            Dim mylistProduct As List(Of Product) = efHelper.GetListProducts(cate, item, domain)
            mylistProduct = mylistProduct.OrderByDescending(Function(x) x.PublishDate).ToList()

            Dim listProductId As List(Of Integer) = insertProducts(mylistProduct, cate.Category1, "CA", planType, item.prodDisplayCount, siteid, issueId)

            'efHelper.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
            ProductIssueDAL.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount, cate.CategoryID)
        Next
    End Function




    ''' <summary>
    ''' 根据categories获取产品列表，返回字典格式(category,list of product)
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="planType"></param>
    ''' <param name="domain"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    Public Function GetProductDict(ByVal siteid As Integer, ByVal planType As String, ByVal categoryArray() As String, ByVal domain As String, ByVal issueId As Integer) As Dictionary(Of String, List(Of Product))

        Dim productDict As Dictionary(Of String, List(Of Product)) = New Dictionary(Of String, List(Of Product))

        For Each categoryName As String In categoryArray

            If categoryName.Contains("#specialcollection#") Then
                'don't crawl this category
                Continue For
            End If
            Try

                Dim productPath As ProductPath = (From proPath In efContext.ProductPaths
                                                  Join cate In efContext.Categories
                                                               On proPath.prodcate Equals cate.CategoryID
                                                  Where proPath.siteId = siteid AndAlso proPath.planType = planType AndAlso cate.Category1 = categoryName
                                                  Select proPath).FirstOrDefault()

                Dim category As Category = (From c In efContext.Categories
                                            Where c.SiteID = siteid AndAlso c.CategoryID = productPath.prodcate
                                            Select c).FirstOrDefault()

                Dim HtmlResolver As New Resolver()
                Dim pageHelper As New HtmlPageUtility()

                '对全麦-英文站的产品进行去重复（www.clothing-marketing.com）
                Dim getProductCode As Func(Of Product, String) = Function(prodcut)
                                                                     Try
                                                                         If Not prodcut.Url.Contains("www.clothing-marketing.com") Then
                                                                             Return ""
                                                                         End If
                                                                         Dim myHtmlDom = pageHelper.GetHtmlDocument(prodcut.Url, productPath.cookie, "refer", productPath.pageEncoding)
                                                                         Dim myProdDom = myHtmlDom.DocumentNode.SelectSingleNode("//*[@id=""view_goods_info""]/h1/span[2]")
                                                                         Return myProdDom.InnerText.Split("-")(0)
                                                                     Catch ex As Exception
                                                                         Return ""
                                                                     End Try
                                                                 End Function

                Dim productInCategory As List(Of Product) = HtmlResolver.ResolveHtmlProduct(category, productPath, domain, getProductCode)
                'Dim productInCategory As List(Of Product) = efHelper.GetListProducts(category, productPath, domain)

                productInCategory = productInCategory.OrderByDescending(Function(x) x.PublishDate).ToList()
                If Not productDict.ContainsKey(categoryName) Then
                    productDict.Add(categoryName, productInCategory)
                End If


            Catch ex As Exception
                Common.LogText("GetProductDict()-->" & ex.StackTrace & " where siteid=" & siteid & " plantype=" & planType)
            End Try
        Next

        Return productDict


    End Function


    Private Sub Extract(ByVal OldList As List(Of Product), ByRef NewList As List(Of Product), ByVal Amount As Integer)
        Dim counter As Integer = 0
        For Each P In OldList
            If Not NewList.Contains(P) And counter < Amount Then
                NewList.Add(P)
                counter += 1
            End If
        Next

    End Sub

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
            'Dim returnId As Integer = efHelper.InsertOrUpdateProduct(li, Now, categoryId, allIssuedProducts)
            Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, allIssuedProducts)
            If returnId > 0 Then
                listProductId.Add(returnId)
            End If
            If (listProductId.Count = productCount) Then
                Exit For
            End If
        Next
        Return listProductId
    End Function

    Private Function GetRecentHotClickProduct(ByVal siteId As String, ByVal planType As String, ByVal spreadLogin As String, ByVal APIKEY As String, ByVal issueCount As Integer, ByVal productCount As Integer, startTime As DateTime, endTime As DateTime, ByVal utmCode As String) As List(Of Product)

        '找出最近issueCount个记录
        Dim issueLogs As List(Of Issue) = (From i In efContext.Issues
                                           Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType AndAlso Not (i.SpreadCampId Is Nothing)
                                           Order By i.IssueID Descending
                                           Select i).Take(issueCount).ToList()

        '删除本siteidclick数据
        Dim queryClickData = (From click In efContext.Clicks_Issue
                              Where click.siteid = siteId
                              Select click).ToList()
        For Each click As Clicks_Issue In queryClickData
            efContext.Clicks_Issue.DeleteObject(click)
        Next


        Dim mySpread As New SpreadWebReference.Service()
        Dim productDictionary As Dictionary(Of Product, Integer) = New Dictionary(Of Product, Integer) '(product,clickCount)
        Dim clickList As List(Of Clicks_Issue) = New List(Of Clicks_Issue)

        For Each Issue As Issue In issueLogs

            Try
                'debug emulate data
                'spreadLogin = "B2CExport@reasonables.com"
                'APIKEY = "65AB80EF-F6B1-485B-BD4E-8A4A6611E1BE"
                'Issue.SpreadCampId = 495272
                'startTime = Now.AddDays(-30)


                'GetCampaignClicks data by API
                EFHelper.LogText("#debug# issue:" & Issue.IssueID & " plantype:" & Issue.PlanType & " siteid:" & siteId & " spreadLoginEmail:" & spreadLogin & " APIKEY:" & APIKEY & "issue.SpreadCampId:" & Issue.SpreadCampId & " clickStartTime:" & startTime & " clickEndTime:" & endTime)
                Dim myClickData As DataSet = mySpread.GetCampaignClicks(spreadLogin, APIKEY, Issue.SpreadCampId, startTime, endTime)
                EFHelper.LogText("campID:" & Issue.SpreadCampId & " clickstarttime:" & startTime & " clickEndTime:" & endTime & " totalClickCount:" & myClickData.Tables(0).Rows.Count)

                For Each dr As DataRow In myClickData.Tables(0).Rows
                    Dim click As New Clicks_Issue
                    click.siteid = siteId
                    click.IssueID = Issue.IssueID
                    click.SpreadCampId = Issue.SpreadCampId
                    click.SubEmail = dr(0)
                    click.ClickDate = dr(1)

                    'clear tag in utm like [yyyyMMdd]  ?utm_source=reasonable&utm_medium=edm&utm_campaign=[yyyyMMdd]
                    If utmCode.Contains("[") AndAlso utmCode.Contains("]") Then
                        utmCode = utmCode.Substring(0, utmCode.IndexOf("[") - 1)
                    End If

                    If Not String.IsNullOrEmpty(utmCode) AndAlso (dr(2).ToString.Trim.Contains(utmCode)) Then
                        click.ClickUrl = dr(2).ToString.Remove(dr(2).ToString.IndexOf(utmCode) - 1)
                    Else
                        click.ClickUrl = dr(2).ToString.Trim
                    End If
                    clickList.Add(click)
                    'End If
                    efContext.AddToClicks_Issue(click)
                Next


                '本期记录在智能化数据库中的issue product
                Dim issueProduct As List(Of Product) = (From prod In efContext.Products
                                                        Join product_issue In efContext.Products_Issue
                                                         On prod.ProdouctID Equals product_issue.ProductId
                                                        Where prod.SiteID = product_issue.SiteId AndAlso product_issue.IssueID = Issue.IssueID
                                                        Select prod
                                                         ).ToList()
                '将每个product与spread click report对比进行统计点击数量，结果存在productDictionary<product,clickCount>
                For Each product As Product In issueProduct
                    Dim clickCount As Integer = 0
                    For index = 0 To clickList.Count - 1
                        If clickList(index).ClickUrl.Contains(product.Url) Then
                            clickCount += 1
                        End If
                    Next

                    If Not productDictionary.ContainsKey(product) Then
                        productDictionary.Add(product, clickCount)
                    Else
                        productDictionary.Item(product) = Math.Max(productDictionary.Item(product), clickCount)
                    End If

                Next

            Catch ex As Exception
                EFHelper.LogText("GetRecentHotClickProduct()-->" & ex.ToString)
            End Try


        Next


        '按照产品点击数排序
        Dim hotClickProduct As List(Of Product) = (From tPair As KeyValuePair(Of Product, Integer) In productDictionary Order By tPair.Value Descending
                                                   Select tPair.Key).ToList
        'Try
        '    productDictionary = (From entry In productDictionary Order By entry.Value Descending Select entry).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        '    Dim hotClickProduct2 As List(Of Product) = productDictionary.Keys.ToList()
        'Catch ex As Exception

        'End Try

        Return hotClickProduct

    End Function





End Class
