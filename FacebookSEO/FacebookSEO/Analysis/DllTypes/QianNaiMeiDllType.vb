
Imports System.Data.SqlClient
Imports Analysis
Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq

<DllType("qiannaimei")>
Public Class QianNaiMeiDllType
    Inherits ProductStart

    Dim efHelper As New EFHelper()
    Private efContext As New EmailAlerterEntities()

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim category() As String = list.Categories.Split("^")
        Dim PlanType As String = list.PlanType.Trim

        efHelper.GetBanner(list.SiteId, IssuesId, category(0).Trim, PlanType, list.Url)
        GetProducts(list.SiteId, PlanType, list.Url, IssuesId)
        SetSubject(IssuesId, list.SiteId, category(0).Trim, PlanType, list.SubjectForNFC)
        If (PlanType.Contains("HP")) Then
            Dim plan As AutomationPlan = (From autoplan In efContext.AutomationPlans
                                          Where autoplan.SiteID = list.SiteId AndAlso autoplan.PlanType = PlanType
                                          Select autoplan).FirstOrDefault()
            Dim triggerIndex As Integer
            If plan.TriggerForNFC.HasValue Then
                triggerIndex = plan.TriggerForNFC
            Else
                triggerIndex = 0
            End If

            Dim cate As String = category(triggerIndex).Trim
            Dim contactlistCount As Integer
            Dim saveListName1 As String = String.Empty
            If plan.IsAssociateSite Then
                saveListName1 = "Default Associate Trigger List"
            Else
                saveListName1 = cate & Now.ToString("yyyyMMdd-HHMMss")
            End If

            'Try
            '    contactlistCount = efHelper.CreateContactList(siteId, IssueID, planType.Trim, cate, saveListName1, 90, ChooseStrategy.Favorite, "draft", spreadLogin, appId)
            'Catch ex As Exception
            '    '已存在或者其他错误
            '    Common.LogText(ex.ToString())
            'End Try
            contactlistCount = efHelper.CreateContactList(list.SiteId, IssuesId, PlanType.Trim, cate, saveListName1, 90, ChooseStrategy.Favorite, "draft", list.SpreadLoginEmail, list.AppId)
            '!debug# account return 0,so modify its values to enter debug mode
            'contactlistCount = 1
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssuesId, saveListName1, "draft")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 遍历本站点的本类型计划的productPaths表中的规则，根据其中的设定获取产品，并填充至products及products_issue表
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="planType"></param>
    ''' <param name="domain"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Sub GetProducts(ByVal siteid As Integer, ByVal planType As String, ByVal domain As String, ByVal issueId As Integer)
        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath).ToList()
        For Each item As ProductPath In listProPath
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()

            If cate.Category1.ToLower = "hotclick" Then
                '往期热门，后面从数据库中提取，不需要爬站点
                Continue For
            End If

            Dim mylistProduct As List(Of Product)
            If item.fileType.ToLower.Trim <> "html" Then

                Dim myHtmlDom As HtmlDocument = New HtmlDocument
                Dim myProdDom As HtmlNodeCollection = Nothing
                myHtmlDom = efHelper.GetHtmlDocument(cate.Url, item.cookie, "refer", item.pageEncoding)
                Dim jsonNode As HtmlNode = myHtmlDom.DocumentNode.SelectSingleNode(item.cateParam)
                Dim jobj = JObject.Parse(jsonNode.InnerHtml)
                mylistProduct = efHelper.GetJsonProducts(jobj, cate, item)
                mylistProduct = mylistProduct.Select(Function(p)
                                                         p.Discount = p.Price - p.Discount
                                                         Return p
                                                     End Function).ToList
            Else
                mylistProduct = efHelper.GetListProducts(cate, item, domain)
            End If
            Dim listProductId As List(Of Integer) = efHelper.insertProducts(mylistProduct, cate.Category1, "CA", planType, item.prodDisplayCount, siteid, issueId)
            'InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
            ProductIssueDAL.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount, cate.CategoryID)

        Next
    End Sub
    ''' <summary>
    ''' 从数据库Promotion产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Private Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal cateName As String, ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String
        'Dim firstProduct As String = (From p In efContext.Products
        '             Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '             Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
        '             Select p.Prodouct).FirstOrDefault()
        'Dim productList As List(Of String) = (From p In efContext.Products
        '                                      Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '                                      Where pi.IssueID = issueId AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
        '                                      Order By p.LastUpdate Descending
        '                                      Select p.Prodouct).ToList()

        Dim productList As List(Of String) = GetProductsByCateId(issueId, cateName, siteId)
        Dim allIssues = From i In efContext.Issues
                        Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                        Select i  '所有成功的issue记录

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


        subject = subject.Replace("[VOL_NUMBER]", (allIssues.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
        subject = subject.Replace("[YYYY]", Now.Year.ToString)
        subject = subject.Replace("[MM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))
        subject = subject.Replace("[MMM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))

        efHelper.InsertIssueSubject(issueId, subject)
    End Sub
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
End Class