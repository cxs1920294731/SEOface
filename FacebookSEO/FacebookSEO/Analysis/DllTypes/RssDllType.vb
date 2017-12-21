Imports Analysis
Imports System.Xml
Imports System.Text.RegularExpressions

<DllType("rss")>
Public Class RssDllType
    Inherits ProductStart


    Dim efHelper As New EFHelper()
    Private efContext As New EmailAlerterEntities()

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim category() As String = list.Categories.Split("^")

        Dim uri = New Uri(list.Url)
        Dim domain As String = uri.Scheme & "://" & uri.Host
        efHelper.GetBanner(list.SiteId, IssuesId, category(0).Trim, planType, domain)

        Dim doc As XmlDocument = GetXmlDoc(list.Url)
        GetCategoryProduct(doc, list.SiteId, planType, IssuesId)

        SetSubject(IssuesId, list.SiteId, category(0).Trim, planType, list.SubjectForNFC)

        If (planType.Contains("HP")) Then
            Dim plan As AutomationPlan = (From autoplan In efContext.AutomationPlans
                                          Where autoplan.SiteID = list.SiteId AndAlso autoplan.PlanType = planType
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
            contactlistCount = efHelper.CreateContactList(list.SiteId, IssuesId, planType.Trim, cate, saveListName1, 90, ChooseStrategy.Favorite, "draft", list.SpreadLoginEmail, list.AppId)
            '!debug# account return 0,so modify its values to enter debug mode
            'contactlistCount = 1
            If (contactlistCount > 0) Then
                efHelper.InsertContactList(IssuesId, saveListName1, "draft")
            End If
        End If
    End Sub

    Private Function GetXmlDoc(url As String) As XmlDocument
        Return efHelper.LoadXmlDoc(url)
    End Function
    Private Function GetCategoryProduct(doc As XmlDocument, siteid As Integer, PlanType As String, issueId As Integer) As List(Of Product)
        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Join cate In efContext.Categories On proPath.prodcate Equals cate.CategoryID
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = PlanType
                                                   Select proPath
                                                    ).OrderBy(Function(x) x.prodcate).ToList()
        Dim allProducts As List(Of Product) = New List(Of Product)
        For Each item As ProductPath In listProPath 'get main category and sub category and combine to 1 list
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()
            Dim mylistProduct As List(Of Product) = New List(Of Product)
            mylistProduct = efHelper.GetListProducts(doc, cate, item)
            allProducts.AddRange(mylistProduct)

            Dim listProductId As List(Of Integer) = efHelper.insertProducts(mylistProduct, cate.Category1, "CA", PlanType, item.prodDisplayCount, siteid, issueId)
            'InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
            ProductIssueDAL.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount, cate.CategoryID)
        Next
        Return allProducts
    End Function
    ''' <summary>
    ''' 从数据库Promotion产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Private Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal cateName As String, ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String
        Dim productList As List(Of String) = (From p In efContext.Products
                                              Join pi In efContext.Products_Issue
                                             On p.ProdouctID Equals pi.ProductId
                                              Where p.SiteID = siteId AndAlso pi.IssueID = issueId
                                              Select p.Prodouct).ToList()
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
End Class
