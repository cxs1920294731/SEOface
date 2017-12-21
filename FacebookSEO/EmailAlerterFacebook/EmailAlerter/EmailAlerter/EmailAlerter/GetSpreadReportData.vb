Imports System.Linq
Imports System.Data.Common
Imports System.Data.SqlClient
Imports EmailAlerter.GroupBuyer2

Public Class GetSpreadReportData

    Private Shared efContext As New FaceBookForSEOEntities
    Private Shared listWillExpiredProdUrl As New List(Of String)
    Public Shared Sub Start(ByVal spreadLoginEmail As String, ByVal appId As String, ByVal siteId As Integer, ByVal planType As String, _
                            ByVal templateId As Integer, ByVal senderEmail As String, ByVal senderName As String, ByVal trigger As Integer, ByVal urlSpecialCode As String)
        If (planType = "NFE") Then '过期提醒功能，NFE：Notification For Expiration
            GetSpreadData(spreadLoginEmail, appId, siteId, planType)
            listWillExpiredProdUrl = GetListExpiredProductUrl(siteId)
            If (listWillExpiredProdUrl.Count >= 0) Then

                Dim campaignName As String = "NotifyForExpiration" & Now.ToString("yyyyMMdd")
                Dim dbTemplate As String = efContext.AutoTemplates.Where(Function(t) t.Tid = templateId).Single().Contents
                Dim listSentData As List(Of SentData) = GetSentData(siteId, spreadLoginEmail, appId)
                Dim listCustomer As List(Of String) = listSentData.Select(Function(data) data._subEmail).Distinct().ToList()
                For Each customer As String In listCustomer
                    Dim subject As String = "你所關注的優惠即將在兩天后過期:"
                    Dim toEmail As String = customer
                    Dim listClickUrl As List(Of String) = listSentData.Where(Function(data) data._subEmail = toEmail AndAlso listWillExpiredProdUrl.Contains(data._convertUrl)).Select(Function(data) data._convertUrl).ToList()
                    Dim listProducts As New List(Of AutoProduct)
                    listProducts = (From p In efContext.AutoProducts
                                    Where (p.SiteID = siteId AndAlso listClickUrl.Contains(p.Url) AndAlso listWillExpiredProdUrl.Contains(p.Url) _
                                    AndAlso Not String.IsNullOrEmpty(p.ExpiredDate))
                                    Select p).ToList()
                    'Dim strProductId As String = ""
                    If (listProducts.Count > 0) Then '如果该客户点击了过期了的产品，则发送，否则不发送
                        'If (listProducts.Count > 5) Then
                        '    Dim listProds As List(Of AutoProduct) = listProducts.Take(5).ToList()
                        '    For Each li As AutoProduct In listProds
                        '        If (li.Prodouct.Contains("【")) Then
                        '            subject = subject & li.Prodouct
                        '        Else
                        '            subject = subject & "【" & li.Prodouct & "】"
                        '        End If
                        '        'strProductId = strProductId & li.ProdouctID.ToString() & ","
                        '    Next
                        'Else
                        '    For Each li As AutoProduct In listProducts
                        '        If (li.Prodouct.Contains("【")) Then
                        '            subject = subject & li.Prodouct
                        '        Else
                        '            subject = subject & "【" & li.Prodouct & "】"
                        '        End If
                        '        'strProductId = strProductId & li.ProdouctID.ToString() & ","
                        '    Next
                        'End If
                        subject = GetProductSubject(listProducts)

                        Dim myTemplate As String = GetSpreadTemplate(dbTemplate, listProducts)
                        Dim mySpread As New SpreadService.Service
                        'newsletter@groupbuyer.com.hk换成gtang@reasonables.com
                        'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, toEmail, subject, myTemplate)
                        Try
                            LogText("GroupBuyer过期提醒--" & "senderName:" & senderName & ",toEmail:" & toEmail & ",subject:" & subject)
                        Catch ex As Exception
                            'Ignore
                        End Try
                    End If
                Next
            End If
        ElseIf (planType = "NFC") Then '点击提醒功能，NFC：Notification For Click
            GetSpreadData(spreadLoginEmail, appId, siteId, trigger, urlSpecialCode)
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

            'Dim clickDate As String = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd")
            'Dim sql As String = "select distinct c.SubEmail,c.ClickUrl from Clicks_Issue c " & _
            '                "left join Convertions_Issue ci on c.SubEmail=ci.SubEmail and c.SubEmail=ci.SubEmail" & _
            '                " where ci.ConvertID is null and c.ClickDate=@clickDate"
            'Dim args As System.Data.SqlClient.SqlParameter()
            'args = New System.Data.SqlClient.SqlParameter() {New SqlParameter With {.ParameterName = "clickDate", .Value = clickDate}}
            'Dim clickData As New List(Of ClickData)
            'Using efContext1 As New FaceBookForSEOEntities
            '    clickData = efContext1.ExecuteStoreQuery(Of ClickData)(sql, args).ToList()
            'End Using
            'If (clickData.Count > 0) Then
            'Dim listCustomer As New List(Of String)
            'listCustomer = ClickData.Select(Function(data) data.SubEmail).Distinct().ToList()
            Dim dbTemplate As String = efContext.AutoTemplates.Where(Function(t) t.Tid = templateId).Single().Contents
            Dim elementAndTitle As New GroupBuyer2.ElementAndTitle
            If (listCustomer.Count > 0) Then '如果今天有需要发送邮件，则从Rss中获取最热卖的6个产品
                Dim groupbuyer As New GroupBuyer2
                elementAndTitle = groupbuyer.ReadRSSDeal("http://www.groupbuyer.com.hk/rss/rss_reasonable.php", siteId)
                Dim logProductUrl As String = ""
                For Each customer As String In listCustomer
                    Dim mycustomer As String = customer
                    logProductUrl = ""
                    Dim listClickUrl As List(Of String) = listClickWithoutConvertData.Where(Function(c) c.SubEmail = mycustomer).Select(Function(data) data.ClickUrl).Distinct().ToList()
                    Dim listProducts As New List(Of AutoProduct)
                    listProducts = (From p In efContext.AutoProducts
                    Where (p.SiteID = siteId AndAlso listClickUrl.Contains(p.Url) AndAlso Not p.LastUpdate Is Nothing AndAlso Not p.ExpiredDate Is Nothing)
                                                                Select p).Distinct().ToList() 'distinct(）不能去除productID不相同但是URL相同的情况
                    '去除相同URL的产品
                    Dim listProductsResult As New List(Of AutoProduct) '
                    For Each item As AutoProduct In listProducts
                        If Not isProductExist(listProductsResult, item.Url) Then
                            listProductsResult.Add(item)
                        End If
                    Next
                    listProducts = listProductsResult
                    If (listProducts.Count > 0) Then
                        Dim subject As String = "你所關注的優惠​快將過期​：" & GetProductSubject(listProducts)
                        'Dim groupbuyer As New GroupBuyer2
                        'Dim elementAndTitle As GroupBuyer2.ElementAndTitle = groupbuyer.ReadRSSDeal("http://www.groupbuyer.com.hk/rss/rss_reasonable.php", siteId)
                        Dim arrEmailCampaignElementDeal() As EmailCampaignElementDeal = elementAndTitle.EmailCampaignElements.ToList().OrderByDescending(Function(e) e.noofPurchased).Take(6).ToArray()
                        Dim listTopSellProducts As New List(Of AutoProduct)
                        For Each arr As EmailCampaignElementDeal In arrEmailCampaignElementDeal
                            Dim product As New AutoProduct
                            product.Prodouct = arr.Subject
                            product.Url = arr.Link
                            product.Price = Double.Parse(arr.Price.Replace("HK$", ""))
                            product.Discount = Double.Parse(arr.DiscountPrice.Replace("HK$", ""))
                            product.Currency = "HK$"
                            product.Sales = arr.noofPurchased
                            product.PictureUrl = arr.ProductImage
                            listTopSellProducts.Add(product)
                        Next
                        Dim myTemplate As String = GetSpreadTemplate(dbTemplate, listProducts, listTopSellProducts)
                        Dim mySpread As New SpreadService.Service
                        Dim campaignName As String = "NotifyForClick" & Now.ToString("yyyyMMdd")
                        'Dim toEmail As String = customer
                        'newsletter@groupbuyer.com.hk换成gtang@reasonables.com
                        '真实环境,begin
                        'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, toEmail, subject, myTemplate)
                        'Try
                        '    LogText("GroupBuyer点击提醒--" & "senderName:" & senderName & ",toEmail:" & toEmail & ",subject:" & subject)
                        'Catch ex As Exception
                        '    'Ignore
                        'End Try
                        'end

                        '测试环境,begin
                        'campaignName = campaignName & ""
                        Try
                            'If (customer.Trim() = "alan@reasonable.hk" OrElse customer.Trim() = "gtang@reasonables.com" OrElse _
                            '    customer.Trim() = "pzhen@reasonables.com" OrElse customer.Trim() = "lyuen@reasonable.hk" OrElse _
                            '    customer.Trim() = "dao@reasonables.com" OrElse customer.Trim = "119293631@qq.com" OrElse customer.Trim() = "") Then
                            '    'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, toEmail, subject, myTemplate)
                            '    'mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, "dao@reasonable.cn", subject, myTemplate)
                            'End If
                            LogText("begin send2")
                            LogText("groupbuyer点击提醒--senderName:" & senderName & ",toEmail:" & mycustomer & ",subject:" & subject)
                            Dim send2Result As String = mySpread.Send2(spreadLoginEmail, appId, campaignName, senderEmail, senderName, mycustomer, subject, myTemplate)
                            LogText("send2 result:" & send2Result)
                            For Each p As AutoProduct In listProducts
                                logProductUrl = logProductUrl & p.Url & vbCrLf
                            Next
                            LogText("productURls:" & vbCrLf & logProductUrl)
                        Catch ex As Exception
                            LogText(ex.Message)
                            'Ignore
                        End Try
                        'end
                    End If
                Next
            End If
        End If
    End Sub

    Private Shared Function GetProductSubject(ByVal listProducts As List(Of AutoProduct)) As String
        Dim subject As String = ""
        If (listProducts.Count > 0) Then '如果该客户点击了过期了的产品，则发送，否则不发送
            If (listProducts.Count > 5) Then
                Dim listProds As List(Of AutoProduct) = listProducts.Take(5).ToList()
                For Each li As AutoProduct In listProds
                    li.Prodouct = li.Prodouct.Replace("【", "").Replace("】", "")
                    subject = subject & " " & li.Prodouct.Trim & " |"
                Next
                subject = subject.TrimEnd("|")
            Else
                For Each li As AutoProduct In listProducts
                    li.Prodouct = li.Prodouct.Replace("【", "").Replace("】", "")
                    subject = subject & " " & li.Prodouct.Trim & " |"
                Next
                subject = subject.TrimEnd("|")
            End If
        End If
        Return subject
    End Function
    ''' <summary>
    ''' 获取Spread的Open、Click、Conversion数据。旧的，本方法已废弃。
    ''' </summary>
    ''' <param name="spreadLoginEmail"></param>
    ''' <param name="appId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetSpreadData(ByVal spreadLoginEmail As String, ByVal appId As String, ByVal siteId As Integer, ByVal planType As String)
        Dim listIssues As New List(Of AutoIssue)
        Dim startTime As DateTime = Now.AddDays(-15) 'Test:-7,Real:-15，60天时间太长会出现内存溢出
        Dim endTime As DateTime = Now
        listIssues = (From issue In efContext.AutoIssues
                     Where issue.SiteID = siteId AndAlso issue.SentStatus = "ES" _
                        AndAlso Not String.IsNullOrEmpty(issue.Subject) _
                        AndAlso issue.IssueDate.CompareTo(startTime) > 0 AndAlso issue.IssueDate.CompareTo(endTime) < 0 _
                        AndAlso Not (issue.SpreadCampId Is Nothing)
                     Select issue).ToList()
        'efContext.AutoIssues.Where(Function(iss) iss.SiteID = siteId AndAlso iss.PlanType = planType _
        'AndAlso iss.IssueDate.CompareTo(Now) > 0 AndAlso iss.IssueDate.CompareTo(Now.AddDays(-7)) < 0).ToList()
        For Each issue As AutoIssue In listIssues
            Dim issueId As Integer = issue.IssueID
            Dim queryOpenData = (From open In efContext.Opens_Issue
                              Where open.IssueID = issueId
                              Select open).ToList()
            For Each open As Opens_Issue In queryOpenData
                efContext.Opens_Issue.DeleteObject(open)
            Next
            'efContext.SaveChanges()
            Dim queryClickData = (From click In efContext.Clicks_Issue
                                Where click.IssueID = issueId
                                Select click).ToList()
            For Each click As Clicks_Issue In queryClickData
                efContext.Clicks_Issue.DeleteObject(click)
            Next
            'efContext.SaveChanges()
            Dim queryConvertData = (From convert In efContext.Convertions_Issue
                                  Where convert.IssueID = issueId
                                  Select convert).ToList()
            For Each convert As Convertions_Issue In queryConvertData
                efContext.Convertions_Issue.DeleteObject(convert)
            Next
            'efContext.SaveChanges()
        Next
        efContext.SaveChanges()
        Dim mySpread As New SpreadService.Service
        For Each issue As AutoIssue In listIssues
            Dim myOpenData As DataSet = mySpread.GetCampaignOpens(spreadLoginEmail, appId, issue.SpreadCampId, startTime, endTime)
            For Each dr As DataRow In myOpenData.Tables(0).Rows
                Dim open As New Opens_Issue
                open.IssueID = issue.IssueID
                open.SpreadCampId = issue.SpreadCampId
                open.SubEmail = dr(0).ToString()
                open.OpenDate = dr(1)
                efContext.AddToOpens_Issue(open)
            Next
            Dim myClickData As DataSet = mySpread.GetCampaignClicks(spreadLoginEmail, appId, issue.SpreadCampId, startTime, endTime)
            For Each dr As DataRow In myClickData.Tables(0).Rows
                Dim click As New Clicks_Issue
                click.IssueID = issue.IssueID
                click.SpreadCampId = issue.SpreadCampId
                click.SubEmail = dr(0)
                click.ClickDate = dr(1)
                click.ClickUrl = dr(2)
                efContext.AddToClicks_Issue(click)
            Next
            Dim myConversionData As DataSet = mySpread.GetCampaignConversions(spreadLoginEmail, appId, issue.SpreadCampId, startTime, endTime)
            For Each dr As DataRow In myConversionData.Tables(0).Rows
                Dim convert As New Convertions_Issue
                convert.IssueID = issue.IssueID
                convert.SpreadCampId = issue.SpreadCampId
                convert.SubEmail = dr(0)
                convert.ConvertTime = dr(1)
                convert.ConvertValue = dr(2)
                convert.ConvertType = dr(3)
                convert.ConvertUrl = dr(4)
                efContext.AddToConvertions_Issue(convert)
            Next
            efContext.SaveChanges()
        Next
        'efContext.SaveChanges()
    End Sub

    Private Shared Sub GetSpreadData(ByVal spreadLoginEmail As String, ByVal appId As String, ByVal siteId As Integer, ByVal trigger As Integer, ByVal urlSpecialCode As String)
        Dim listIssues As New List(Of AutoIssue)
        Dim startTime As DateTime = Now.AddDays(-(trigger + 7))
        Dim endTime As DateTime = Now.AddDays(-trigger)
        listIssues = (From issue In efContext.AutoIssues
                    Where issue.SiteID = siteId AndAlso issue.SentStatus = "ES" _
                       AndAlso Not String.IsNullOrEmpty(issue.Subject) _
                       AndAlso issue.IssueDate.CompareTo(startTime) > 0 AndAlso issue.IssueDate.CompareTo(endTime) < 0 _
                       AndAlso Not (issue.SpreadCampId Is Nothing)
                    Select issue).ToList()
        LogText("issueStartTime:" & startTime & " issueEndTime:" & endTime & "issueCount:" & listIssues.Count)
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
            LogText("campID:" & issue.SpreadCampId & " openstarttime:" & openStartTime & " openEndTime:" & openEndTime & " totalOpenCount:" & myOpenData.Tables(0).Rows.Count)
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
            LogText("campID:" & issue.SpreadCampId & " clickstarttime:" & clickStartTime & " clickEndTime:" & clickEndTime & " totalClickCount:" & myClickData.Tables(0).Rows.Count)
            For Each dr As DataRow In myClickData.Tables(0).Rows
                Dim click As New Clicks_Issue
                click.siteid = siteId
                click.IssueID = issue.IssueID
                click.SpreadCampId = issue.SpreadCampId
                click.SubEmail = dr(0)
                click.ClickDate = dr(1)
                'LogText("productURL:" & dr(2).ToString.Trim)
                If (dr(2).ToString.Contains(partUrlSpecialCode)) Then
                    click.ClickUrl = dr(2).ToString.Remove(dr(2).ToString.IndexOf(partUrlSpecialCode) - 1)
                Else
                    click.ClickUrl = dr(2).ToString.Trim
                End If
                'End If
                efContext.AddToClicks_Issue(click)
            Next
            Dim myConversionData As DataSet = mySpread.GetCampaignConversions(spreadLoginEmail, appId, issue.SpreadCampId, converStartTime, converEndTime)
            LogText("campID:" & issue.SpreadCampId & " converstarttime:" & converStartTime & " converEndTime:" & converEndTime & " totalConverCount:" & myConversionData.Tables(0).Rows.Count)
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
    End Sub
    ''' <summary>
    ''' 获取发送数据，包括联系人email和点击url
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="spreadLoginEmail"></param>
    ''' <param name="appId"></param>
    ''' <remarks></remarks>
    Private Shared Function GetSentData(ByVal siteId As Integer, ByVal spreadLoginEmail As String, ByVal appId As String) As List(Of SentData)
        Dim listConvertData As New List(Of SentData)
        listConvertData = (From convert In efContext.Convertions_Issue
                        Select New SentData With {._subEmail = convert.SubEmail, ._convertUrl = convert.ConvertUrl}).Distinct().ToList()
        Dim listAllClickData As New List(Of SentData)
        listAllClickData = (From click In efContext.Clicks_Issue
                       Select New SentData With {._subEmail = click.SubEmail, ._convertUrl = click.ClickUrl}).Distinct().ToList()
        Dim listClickData As New List(Of SentData)
        Dim counter As Integer = listAllClickData.Count
        For Each click As SentData In listAllClickData
            If Not (listConvertData.Contains(click)) Then
                listClickData.Add(click)
            End If
        Next
        Return listClickData.Distinct.ToList()
    End Function

    Private Shared Function GetListExpiredProductUrl(ByVal siteId As Integer) As List(Of String)
        Dim listExpiredProdUrl As New List(Of String)
        Dim startTime As DateTime = DateTime.Parse(Format(Now.AddDays(2), "yyyy-MM-dd"))  'DateTime.Parse(Format(Now, "yyyy-MM-dd"))
        Dim endTime As DateTime = DateTime.Parse(Format(Now.AddDays(3), "yyyy-MM-dd")) '两天后过期，Now.AddDays(2)
        listExpiredProdUrl = (From p In efContext.AutoProducts
                           Where p.SiteID = siteId AndAlso Not String.IsNullOrEmpty(p.LastUpdate) AndAlso Not String.IsNullOrEmpty(p.ExpiredDate) _
                          AndAlso startTime < p.ExpiredDate AndAlso p.ExpiredDate <= endTime
                           Select p.Url).ToList()
        Return listExpiredProdUrl
    End Function

    Private Shared Function GetSpreadTemplate(ByVal template As String, ByVal listProducts As List(Of AutoProduct))
        Dim arrTemplate() As String = template.Split("^")
        Dim FTemplate As String = arrTemplate(0)
        Dim NLeftTemplate As String = arrTemplate(1)
        Dim NRightTemplate As String = arrTemplate(2)
        Dim sign As Integer = 0
        Dim MainDealsHtml As String = ""
        For Each li As AutoProduct In listProducts
            If (sign = 0) Then
                Dim NLdeal As String = ""
                If (li.Discount Is Nothing) Then
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                'NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                MainDealsHtml = MainDealsHtml & NLdeal
                sign = 1
            ElseIf (sign = 1) Then
                Dim NLdeal2 As String  '= String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                If (li.Discount Is Nothing) Then
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing AndAlso li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                MainDealsHtml = MainDealsHtml & NLdeal2
                sign = 0
            End If
        Next
        If (listProducts.Count Mod 2 = 1) Then
            MainDealsHtml = MainDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" class=""mobile_hidden"" width=""13""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table width=""315"" align=""left"" border=""0"" cellspacing=""0"" valign=""top"" class=""mobile_hidden""><tbody><tr><td><table align=""left"" width=""315"" height=""333"" border=""0"" cellspacing=""0"" cellpadding=""5"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><tbody><tr><td style=""width:317px;height:333px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""305"" height=""367px"" border=""0"" alt=""""></a></td></tr></tbody></table></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "</td></tr></table></td></tr>"
        End If
        Dim spreadTemplate As String = FTemplate.Replace("[ALL_NEW_DEALS]", MainDealsHtml)
        Return spreadTemplate
    End Function
    ''' <summary>
    ''' 根据用户点击的产品和前6个热销产品组成的邮件模板
    ''' </summary>
    ''' <param name="template"></param>
    ''' <param name="listClickProducts"></param>
    ''' <param name="listTopSellProducts"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSpreadTemplate(ByVal template As String, ByVal listClickProducts As List(Of AutoProduct), _
                                              ByVal listTopSellProducts As List(Of AutoProduct)) As String
        Dim arrTemplate As String() = template.Split("^")
        Dim FTemplate As String = arrTemplate(0)
        Dim NLeftTemplate As String = arrTemplate(1)
        Dim NRightTemplate As String = arrTemplate(2)
        Dim OLeftTemplate As String = arrTemplate(3)
        Dim OMiddleTemplate As String = arrTemplate(4)
        Dim ORightTemplate As String = arrTemplate(5)
        Dim sign As Integer = 0
        Dim MainDealsHtml As String = ""
        Dim OtherDealsHtml As String = ""
        For Each li As AutoProduct In listClickProducts
            If (sign = 0) Then
                Dim NLdeal As String = ""
                If (li.Discount Is Nothing) Then
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                'NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                MainDealsHtml = MainDealsHtml & NLdeal
                sign = 1
            ElseIf (sign = 1) Then
                Dim NLdeal2 As String  '= String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                If (li.Discount Is Nothing) Then
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing AndAlso li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    NLdeal2 = String.Format(NRightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                MainDealsHtml = MainDealsHtml & NLdeal2
                sign = 0
            End If
        Next
        If (listClickProducts.Count Mod 2 = 1) Then
            MainDealsHtml = MainDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" class=""mobile_hidden"" width=""13""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table width=""315"" align=""left"" border=""0"" cellspacing=""0"" valign=""top"" class=""mobile_hidden""><tbody><tr><td><table align=""left"" width=""315"" height=""333"" border=""0"" cellspacing=""0"" cellpadding=""5"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><tbody><tr><td style=""width:317px;height:333px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""305"" height=""367px"" border=""0"" alt=""""></a></td></tr></tbody></table></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "</td></tr></table></td></tr>"
        End If
        '更多優惠模板填充
        Dim middleDealsHtml As String = ""
        Dim sign3 As Integer = 0
        For Each li As AutoProduct In listTopSellProducts
            If (sign3 = 0) Then
                Dim OLdeal1 As String = "" 'String.Format(OLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                If (li.Discount Is Nothing) Then
                    OLdeal1 = String.Format(OLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    OLdeal1 = String.Format(OLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                    'Dim discount As Double = li.Discount
                    'OLdeal1 = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    OLdeal1 = String.Format(OLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                middleDealsHtml = middleDealsHtml & OLdeal1
                sign3 = sign3 + 1
            ElseIf (sign3 = 1) Then
                Dim OMdeal1 As String = "" 'String.Format(OMiddleTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Discount, li.Price, li.Url)
                If (li.Discount Is Nothing) Then
                    OMdeal1 = String.Format(OMiddleTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    OMdeal1 = String.Format(OMiddleTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                    'Dim discount As Double = li.Discount
                    'OLdeal1 = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    OMdeal1 = String.Format(OMiddleTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                middleDealsHtml = middleDealsHtml & OMdeal1
                sign3 = sign3 + 1
            ElseIf (sign3 = 2) Then
                Dim ORdeal1 As String = ""  'String.Format(ORightTemplate, element.Link, element.ProductImage, element.Subject, element.DiscountPrice, element.Price, element.Link)
                If (li.Discount Is Nothing) Then
                    ORdeal1 = String.Format(ORightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, "", "", li.Url)
                ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                    Dim discount As Double = li.Discount
                    ORdeal1 = String.Format(ORightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                    'Dim discount As Double = li.Discount
                    'OLdeal1 = String.Format(NLeftTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), "", li.Url)
                Else
                    Dim price As Double = li.Price
                    Dim discount As Double = li.Discount
                    ORdeal1 = String.Format(ORightTemplate, li.Url, li.PictureUrl.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), li.Prodouct, li.Currency & Format("{0:#.00}", discount), li.Currency & Format("{0:#.00}", price), li.Url)
                End If
                middleDealsHtml = middleDealsHtml & ORdeal1
                sign3 = 0
            End If
        Next
        If listTopSellProducts.Count Mod 3 = 1 Then
            'OtherDealsHtml = OtherDealsHtml & "<td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""16""></td><td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/special"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_2.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""25""></td></tr>"
            OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
            OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
            OtherDealsHtml = OtherDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            OtherDealsHtml = OtherDealsHtml & "</td></tr>"
        ElseIf listTopSellProducts.Count Mod 3 = 2 Then
            OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289"" border=""0"" alt="""" /></a></td></tr></table>"
            OtherDealsHtml = OtherDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            OtherDealsHtml = OtherDealsHtml & "</td></tr>"
        End If
        Dim Deals As String = FTemplate.Replace("[ALL_NEW_DEALS]", MainDealsHtml)
        Dim html As String = Deals.Replace("[ALL_OLD_DEALS]", OtherDealsHtml).Replace("[ALL_TOP_DEALS]", middleDealsHtml)
        Return html
    End Function

    Private Shared Sub LogText(ByVal Ex As String)
        Try
            '2013/08/08 added, 发送错误日志到制定的邮箱组
            If (Ex.Contains("Exception")) Then
                NotificationEmail.SentErrorEmail(Ex.ToString())
            End If
            '----------------------------------------------------
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & "-" & Now.Day & "GroupBuyer.log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

    ''' <summary>
    ''' 判断一个产品是否已存在于一个list中，判断条件为productURL相同；存在返回true，否则false
    ''' </summary>
    ''' <param name="listProduct"></param>
    ''' <param name="productUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function isProductExist(ByVal listProduct As List(Of AutoProduct), ByVal productUrl As String) As Boolean
        For Each item As AutoProduct In listProduct
            If (item.Url.Trim = productUrl.Trim) Then
                Return True
            End If
        Next
        Return False
    End Function
End Class

Public Class SentData
    Private subEmail As String
    Private convertUrl As String
    Public Property _subEmail() As String
        Get
            Return subEmail
        End Get
        Set(ByVal value As String)
            subEmail = value
        End Set
    End Property
    Public Property _convertUrl() As String
        Get
            Return convertUrl
        End Get
        Set(ByVal value As String)
            convertUrl = value
        End Set
    End Property
End Class
