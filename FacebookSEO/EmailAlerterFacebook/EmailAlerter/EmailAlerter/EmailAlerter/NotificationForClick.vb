Imports System.Linq
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class NotificationForClick
    Private efContext As New FaceBookForSEOEntities

    Public Sub start(ByVal dllType As String, ByVal spreadLoginEmail As String, ByVal appId As String, ByVal siteId As Integer, ByVal Trigger As Integer, ByVal subject As String,
                      ByVal templateId As Integer, ByVal senderEmail As String, ByVal senderName As String, ByVal UrlSpecialCode As String)
        '获取Spread报告数据
        GetSpreadData(spreadLoginEmail, appId, siteId, Trigger, UrlSpecialCode)

        '转换数据
        Dim listConverData As List(Of Clicks_Issue) = (From c In efContext.Clicks_Issue
                                   Join co In efContext.Convertions_Issue On c.SubEmail Equals co.SubEmail And c.SpreadCampId Equals co.SpreadCampId And c.ClickUrl Equals co.ConvertUrl
                                   Where c.siteid = siteId
                                    Select c).ToList()
        '点击数据（包括转换）
        Dim listClickData As List(Of Clicks_Issue) = (From c In efContext.Clicks_Issue
                                                       Where c.siteid = siteId
                                                 Select c).Distinct().ToList()
        '将转换数据过滤掉
        Dim listClickWithoutConvertData As New List(Of Clicks_Issue)
        For Each item As Clicks_Issue In listClickData
            If Not (listConverData.Contains(item)) Then
                listClickWithoutConvertData.Add(item)
            End If
        Next
        Dim listCustomer As List(Of String) = listClickWithoutConvertData.Select(Function(data) data.SubEmail).Distinct().ToList()
        If (listCustomer.Count > 0) Then
            Dim dbTemplate As String = efContext.AutoTemplates.Where(Function(t) t.Tid = templateId).Single().Contents
            Dim mycustomer As String
            Dim logProductUrl As String = ""
            Dim allSubject As String = ""
            For Each acustomer As String In listCustomer
                mycustomer = acustomer.Trim
                logProductUrl = ""
                allSubject = ""
                Dim listClickUrl As List(Of String) = listClickWithoutConvertData.Where(Function(c) c.SubEmail = mycustomer).Select(Function(data) data.ClickUrl).Distinct().ToList()
                Dim listProducts As New List(Of AutoProduct)
                listProducts = (From p In efContext.AutoProducts
                Where (p.SiteID = siteId AndAlso listClickUrl.Contains(p.Url) AndAlso Not p.LastUpdate Is Nothing)
                                                            Select p).Distinct().ToList()
                If (listProducts.Count > 0) Then
                    Dim firstProduct As String = listProducts(0).Prodouct.ToString.Trim
                    allSubject = subject & firstProduct & "..."
                    Dim myTemplate As String = FillTemplateFroNFC(dbTemplate, listProducts, UrlSpecialCode)
                    Dim mySpread As New SpreadService.Service
                    Dim campaignName As String = "NotifyForClick" & Now.ToString("yyyyMMdd")

                    Try
                        'If (mycustomer.Trim() = "alan@reasonable.hk" OrElse mycustomer.Trim() = "gtang@reasonables.com" OrElse _
                        '    mycustomer.Trim() = "dao@reasonable.cn" OrElse mycustomer.Trim() = "gtang@reasonable.cn" OrElse _
                        '    mycustomer.Trim() = "dao@reasonables.com" OrElse mycustomer.Trim = "119293631@qq.com" OrElse mycustomer.Trim() = "") Then
                        '    'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, mycustomer, subject, myTemplate)
                        '    'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, "dao@reasonable.cn", subject, myTemplate)
                        'End If
                        LogText("begin send2")
                        Dim send2Result As String = mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, mycustomer, allSubject, myTemplate)
                        LogText("send2 result:" & send2Result)
                        For Each p As AutoProduct In listProducts
                            logProductUrl = logProductUrl & p.Url & vbCrLf
                        Next
                        LogText(dllType & "点击提醒--senderName:" & senderName & ",toEmail:" & mycustomer & ",subject:" & allSubject & ",productURls:" & vbCrLf & logProductUrl)
                    Catch ex As Exception
                        Throw New Exception()
                    End Try
                End If
            Next
        End If
    End Sub

    Private Function GetSpreadData(ByVal spreadLoginEmail As String, ByVal appId As String, ByVal siteId As Integer, ByVal trigger As Integer, ByVal urlSpecialCode As String) As List(Of AutoIssue)
        Dim listIssues As New List(Of AutoIssue)
        Dim startTime As DateTime = Now.AddDays(-(trigger + 7))
        Dim endTime As DateTime = Now.AddDays(-trigger)
        listIssues = (From issue In efContext.AutoIssues
                    Where issue.SiteID = siteId AndAlso issue.SentStatus = "ES" _
                       AndAlso Not String.IsNullOrEmpty(issue.Subject) _
                       AndAlso issue.IssueDate.CompareTo(startTime) > 0 AndAlso issue.IssueDate.CompareTo(endTime) < 0 _
                       AndAlso Not (issue.SpreadCampId Is Nothing)
                    Select issue).ToList()
        '删除本siteid拥有的open、click、conversion数据
        Dim queryOpenData = (From open In efContext.Opens_Issue
                            Where open.siteid = siteId
                            Select open).ToList()
        For Each open As Opens_Issue In queryOpenData
            efContext.Opens_Issue.DeleteObject(open)
        Next

        Dim queryClickData = (From click In efContext.Clicks_Issue
                            Where click.siteid = siteId
                            Select click).ToList()
        For Each click As Clicks_Issue In queryClickData
            efContext.Clicks_Issue.DeleteObject(click)
        Next

        Dim queryConvertData = (From convert In efContext.Convertions_Issue
                                  Where convert.siteid = siteId
                                  Select convert).ToList()
        For Each convert As Convertions_Issue In queryConvertData
            efContext.Convertions_Issue.DeleteObject(convert)
        Next
        '删除本siteid拥有的open、click、conversion数据
        efContext.SaveChanges()
        Dim mySpread As New SpreadService.Service
        Dim openStartTime As DateTime = Now.AddDays(-(trigger + 1))
        Dim openEndTime As DateTime = Now.AddDays(-(trigger))
        Dim clickStartTime As DateTime = Now.AddDays(-(trigger + 1))
        Dim clickEndTime As DateTime = Now.AddDays(-(trigger))
        Dim converStartTime As DateTime = startTime '转换数据的时间跨度要大
        Dim converEndTime As DateTime = Now
        LogText("urlSpecialCode:" & urlSpecialCode)
        Dim partUrlSpecialCode As String = urlSpecialCode.Trim.Replace("[yyyyMMdd]", "").Replace("?", "")
        For Each issue As AutoIssue In listIssues
            Dim myOpenData As DataSet = mySpread.GetCampaignOpens(spreadLoginEmail, appId, issue.SpreadCampId, openStartTime, openEndTime)
            For Each dr As DataRow In myOpenData.Tables(0).Rows
                Dim open As New Opens_Issue
                open.siteid = siteId
                open.IssueID = issue.IssueID
                open.SpreadCampId = issue.SpreadCampId
                open.SubEmail = dr(0).ToString()
                open.OpenDate = dr(1)
                efContext.AddToOpens_Issue(open)
            Next
            Dim myClickData As DataSet = mySpread.GetCampaignClicks(spreadLoginEmail, appId, issue.SpreadCampId, clickStartTime, clickEndTime)
            For Each dr As DataRow In myClickData.Tables(0).Rows
                Dim click As New Clicks_Issue
                click.siteid = siteId
                click.IssueID = issue.IssueID
                click.SpreadCampId = issue.SpreadCampId
                click.SubEmail = dr(0)
                click.ClickDate = dr(1)
                'LogText("productURL:" & dr(2).ToString.Trim)
                'LogText("index:" & dr(2).ToString.IndexOf(partUrlSpecialCode))
                If (dr(2).ToString.Contains(partUrlSpecialCode)) Then
                    click.ClickUrl = dr(2).ToString.Remove(dr(2).ToString.IndexOf(partUrlSpecialCode) - 1)
                Else
                    click.ClickUrl = dr(2).ToString.Trim
                End If
                efContext.AddToClicks_Issue(click)
            Next
            Dim myConversionData As DataSet = mySpread.GetCampaignConversions(spreadLoginEmail, appId, issue.SpreadCampId, converStartTime, converEndTime)
            For Each dr As DataRow In myConversionData.Tables(0).Rows
                Dim convert As New Convertions_Issue
                convert.siteid = siteId
                convert.IssueID = issue.IssueID
                convert.SpreadCampId = issue.SpreadCampId
                convert.SubEmail = dr(0)
                convert.ConvertTime = dr(1)
                convert.ConvertValue = dr(2)
                convert.ConvertType = dr(3)

                'convert.ConvertUrl = dr(4).ToString.Remove(dr(4).ToString.IndexOf(partUrlSpecialCode) - 1)
                If (dr(4).ToString.Contains(partUrlSpecialCode)) Then
                    convert.ConvertUrl = dr(4).ToString.Remove(dr(4).ToString.IndexOf(partUrlSpecialCode) - 1)
                Else
                    convert.ConvertUrl = dr(4).ToString.Trim
                End If

                efContext.AddToConvertions_Issue(convert)
            Next
            efContext.SaveChanges()
        Next
        Return listIssues
    End Function


    Private Function FillTemplateFroNFC(ByVal SpreadTemplate As String, ByVal listProducts As List(Of AutoProduct), ByVal UrlSpecialCode As String) As String
        If (SpreadTemplate.Contains("[BEGIN_PRODUCTS]")) Then
            Dim categoryLen = "[BEGIN_PRODUCTS]".Length
            Dim endCategoryLen = "[END_PRODUCTS]".Length
            Dim cateStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_PRODUCTS]")
            Dim cateEndIndex As Integer = SpreadTemplate.IndexOf("[END_PRODUCTS]")
            Dim oldCategory As String = SpreadTemplate.Substring(cateStartIndex, cateEndIndex - cateStartIndex + endCategoryLen)
            Dim categoryContent As String = SpreadTemplate.Substring(cateStartIndex + categoryLen, cateEndIndex - cateStartIndex - categoryLen)
            Dim productLen = "[BEGIN_PRODUCT]".Length
            Dim endProductLen = "[END_PRODUCT]".Length
            Dim productStartIndex As Integer = categoryContent.IndexOf("[BEGIN_PRODUCT]")
            Dim productEndIndex As Integer = categoryContent.IndexOf("[END_PRODUCT]")
            For Each q As AutoProduct In listProducts
                If (productStartIndex > -1) Then

                    Dim oldProduct As String = categoryContent.Substring(productStartIndex, productEndIndex - productStartIndex + endProductLen)
                    Dim newProduct As String = categoryContent.Substring(productStartIndex + productLen, productEndIndex - productStartIndex - productLen)
                    If (newProduct.Contains("[URL]")) Then
                        newProduct = newProduct.Replace("[URL]", q.Url)
                    End If
                    If (newProduct.Contains("[CATEGORY_ID]")) Then
                        newProduct = newProduct.Replace("[CATEGORY_ID]", q.Categories.First.CategoryID.ToString())
                    End If
                    If (newProduct.Contains("[PICTURE_SRC]")) Then
                        newProduct = newProduct.Replace("[PICTURE_SRC]", q.PictureUrl)
                    End If
                    If (newProduct.Contains("[PICTURE_ALT]")) Then
                        newProduct = newProduct.Replace("[PICTURE_ALT]", q.PictureAlt)
                    End If
                    If (newProduct.Contains("[PRODUCT_NAME]")) Then
                        newProduct = newProduct.Replace("[PRODUCT_NAME]", q.Prodouct)
                    End If
                    If (newProduct.Contains("[PRODUCT_DESCRIPTION]")) Then
                        newProduct = newProduct.Replace("[PRODUCT_DESCRIPTION]", q.Description)
                    End If
                    If (newProduct.Contains("[PRODUCT_PRICE]")) Then
                        If (q.Price Is Nothing) Then  '原价获取不到，只获取到折扣价，则把原价删除掉
                            newProduct = newProduct.Replace("[PRODUCT_PRICE]", "")
                        Else
                            newProduct = newProduct.Replace("[PRODUCT_PRICE]", q.Currency + " " + String.Format("{0:0.00}", q.Price))
                        End If
                    End If
                    If (newProduct.Contains("[PRODUCT_MONEY]")) Then
                        newProduct = newProduct.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", q.Discount))
                    End If
                    If (newProduct.Contains("[PRODUCT_SALES]")) Then
                        newProduct = newProduct.Replace("[PRODUCT_SALES]", q.Sales)
                    End If

                    '2013/05/21,add free shiping pictures and ships pictures in template,begin
                    If (newProduct.Contains("[FREESHIPPING")) Then
                        If (String.IsNullOrEmpty(q.FreeShippingImg)) Then
                            newProduct = newProduct.Replace("[FREESHIPPING]", "")
                        Else
                            Dim addStrPicture As String = "<tr><td align=""center""><img alt="""" src=""" & q.FreeShippingImg & """ style=""border-width: 0px; border-style: solid; display: block;"" /></td></tr>"
                            newProduct = newProduct.Replace("[FREESHIPPING]", addStrPicture)
                        End If
                    End If
                    If (newProduct.Contains("[SHIPS]")) Then
                        If (String.IsNullOrEmpty(q.ShipsImg)) Then
                            newProduct = newProduct.Replace("[SHIPS]", "")
                        Else
                            Dim addStrPicture As String = "<tr><td align=""center""><img alt="""" src=""" & q.ShipsImg & """ style=""border-width: 0px; border-style: solid; display: block;"" /></td></tr>"
                            newProduct = newProduct.Replace("[SHIPS]", addStrPicture)
                        End If
                    End If
                    categoryContent = categoryContent.Remove(productStartIndex, productEndIndex - productStartIndex + endProductLen)
                    categoryContent = categoryContent.Insert(productStartIndex, newProduct)
                    If (categoryContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增
                        productStartIndex = categoryContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                        productEndIndex = categoryContent.IndexOf("[END_PRODUCT]", productEndIndex)
                    Else
                        Exit For
                    End If
                End If
            Next
            '删除掉多余的[BEGIN_PRODUCT]/[END_PRODUCT]块+
            While (categoryContent.Contains("[BEGIN_PRODUCT]"))
                categoryContent = categoryContent.Remove(productStartIndex, productEndIndex - productStartIndex + endProductLen)
                If (categoryContent.Contains("[BEGIN_PRODUCT]")) Then
                    productStartIndex = categoryContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                    productEndIndex = categoryContent.IndexOf("[END_PRODUCT]", productEndIndex)
                End If
            End While
            SpreadTemplate = SpreadTemplate.Replace(oldCategory, categoryContent)
        End If
        SpreadTemplate = SpecialCode.AddSpecialCode(UrlSpecialCode, SpreadTemplate)
        Return SpreadTemplate
    End Function

    Private Sub LogText(ByVal Ex As String)
        Try
            '2013/08/08 added, 发送错误日志到制定的邮箱组
            If (Ex.Contains("Exception")) Then
                NotificationEmail.SentErrorEmail(Ex.ToString())
            End If
            '----------------------------------------------------
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & "-" & Now.Day & "NotificationForClick.log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
