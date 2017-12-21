Imports System.Data.Objects
Imports HtmlAgilityPack
Imports System.IO
Imports System.Linq
Imports System.Web.Services '2013/3/28新增
Imports System.Data.EntityClient '2013/3/28新增
Imports System.Text

Public Structure ProductElement '2013/05/27新增
    Dim productName As String
    Dim url As String
    Dim price As Decimal
    Dim discountPrice As Decimal
    Dim pictureUrl As String
    Dim pubDate As DateTime
    Dim category As Integer
    Dim currency As String
End Structure


Public Class AutoModel
    Public a As New alerter
    Private efContext As New FaceBookForSEOEntities()
    'Private listRssSubscriptionOfAdd As New List(Of Subscriptions)
    'Private listProductElements As New List(Of ProductElement) '所有的产品
    'Private listCategory As New List(Of String) '某个siteId下的所有categoryName
    Private listCategorys As New List(Of AutoCategory) '某个siteId下的所有Category, 2013/06/17 added

    ''' <summary>
    ''' 根据siteId获取RssSubscription表中的数据
    ''' </summary>
    ''' <param name="logList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRssSubscriptionAndInsertAndSentEmail(ByVal logList As List(Of AutoSentLog)) As List(Of AutoSentLog)

        Dim nowTime As DateTime = Now

        Try
            '2013/4/15修改，使用LINQ方法
            '功能：搜索所有满足条件的自动化项目
            '条件：StartAt<现在的时间，Status为"A"，Categories字段为空
            Dim listRssSubscriptionOfAdd As List(Of Subscriptions) = (From autoPlan In efContext.AutomationPlans
                                                                    Join autoSite In efContext.AutomationSites
                                                                    On autoPlan.SiteID Equals autoSite.siteid
                                                                    Where autoPlan.SiteID >= 21 AndAlso autoPlan.StartAt <= nowTime AndAlso autoPlan.Status = "A" _
                                                                    Select New Subscriptions With {.SiteId = autoPlan.SiteID, .PlanType = autoPlan.PlanType, _
                                                                    .StartAt = autoPlan.StartAt, .IntervalDay = autoPlan.IntervalDay, .WeekDays = autoPlan.WeekDays, _
                                                                    .SenderName = If(String.IsNullOrEmpty(autoPlan.SenderName), "", autoPlan.SenderName), .SenderEmail = If(String.IsNullOrEmpty(autoPlan.SenderEmail), "", autoPlan.SenderEmail), .TemplateId = If(autoPlan.TemplateID Is Nothing, -1, autoPlan.TemplateID), _
                                                                    .SplitContactList = autoPlan.SplitContactCount, .Url = If(String.IsNullOrEmpty(autoPlan.URL), "", autoPlan.URL), .Categories = If(String.IsNullOrEmpty(autoPlan.Categories), "", autoPlan.Categories), .SiteName = If(String.IsNullOrEmpty(autoSite.SiteName), "", autoSite.SiteName), _
                                                                    .SpreadLoginEmail = autoSite.SpreadLogin, .AppId = autoSite.AppID, .DllType = If(String.IsNullOrEmpty(autoSite.DllType), "", autoSite.DllType), .SiteUrl = If(String.IsNullOrEmpty(autoSite.SiteUrl), "", autoSite.SiteUrl), _
                                                                    .Volumn = If(autoSite.volumn Is Nothing, -1, autoSite.volumn), .ContactList = If(String.IsNullOrEmpty(autoPlan.ContactList), "", autoPlan.ContactList), _
                                                                    .CampaignStatus = If(String.IsNullOrEmpty(autoPlan.CampaignStatus), "", autoPlan.CampaignStatus), .ScheduleTimeInteval = If(autoPlan.ScheduleTimeInteval Is Nothing, -1, autoPlan.ScheduleTimeInteval), _
                                                                    .ListType = If(String.IsNullOrEmpty(autoPlan.ListType), "", autoPlan.ListType), .SearchAPIType = If(String.IsNullOrEmpty(autoPlan.SearchAPIType), "", autoPlan.SearchAPIType), _
                                                                    .LimitQuantity = If(autoPlan.LimitQuantity Is Nothing, -1, autoPlan.LimitQuantity), .TimeLimit = If(String.IsNullOrEmpty(autoPlan.TimeLimit), "", autoPlan.TimeLimit), _
                                                                    .SellerEmail = If(String.IsNullOrEmpty(autoPlan.SellerEmail), "", autoPlan.SellerEmail), .UrlSpecialCode = autoPlan.UrlSpecialCode, _
                                                                    .TriggerForNFC = If(autoPlan.TriggerForNFC Is Nothing, -1, autoPlan.TriggerForNFC), _
                                                                    .SubjectForNFC = If(String.IsNullOrEmpty(autoPlan.SubjectForNFC), "", autoPlan.SubjectForNFC)}).ToList()
            For Each index As Subscriptions In listRssSubscriptionOfAdd
                Common.LogText(index.Url.ToString())
            Next
            For Each list As Subscriptions In listRssSubscriptionOfAdd

                Dim lastSent As DateTime
                Dim lastSent2 As String
                Dim planType As String = list.PlanType.Trim()
                If (list.WeekDays.Contains(nowTime.DayOfWeek)) Then 'Hour(Now) = Hour(list.startTime) AndAlso Minute(Now) = Minute(list.startTime) _

                    Dim mySiteId As Integer = list.SiteId
                    lastSent2 = GetLastSentTime(mySiteId, planType, logList)
                    Dim flag As Boolean = False '2013/4/2添加
                    If (lastSent2 = "") Then
                        flag = False
                    Else
                        lastSent = Format(DateTime.Parse(lastSent2), "yyyy-MM-dd")
                        If (nowTime - lastSent >= New TimeSpan(list.IntervalDay * 24, 0, 0)) Then
                            flag = True
                        End If
                    End If

                    Dim startAtTime As DateTime = list.StartAt
                    If (lastSent2 = "" OrElse flag) Then
                        If (Hour(startAtTime) < Hour(nowTime) OrElse (Hour(startAtTime) = Hour(nowTime) AndAlso Minute(startAtTime) <= Minute(nowTime))) Then
                            If Not (planType = "NFE" OrElse planType = "NFC") Then '非Notification For Expiration/Notification For Click

                                Dim issueId As Integer = InsertIssue(nowTime, list.SiteId, planType)
                                Common.LogText(issueId.ToString)
                                If (issueId > 0) Then
                                    Dim dllType As String = list.DllType.Trim().ToLower()
                                    Common.LogText("dlltype:" & dllType)
                                    If Not String.IsNullOrEmpty(dllType) Then
                                        Try
                                            Select Case dllType
                                                Case "facebook"
                                                    Dim mainStart As New Analysis.MainStart
                                                    mainStart.Start(dllType, issueId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, "", 0, "") ''modified by alex.
                                                    Threading.Thread.Sleep(100000)
                                            End Select
                                        Catch ex As Exception
                                            UpdateIssueStatus(issueId, "EU") 'EU(get data from Url error)
                                            Throw New Exception(ex.ToString())
                                        End Try
                                        Common.LogText("循环执行完")
                                        Dim mySubject As String = ""
                                        Using newContext As New FaceBookForSEOEntities() '获取邮件的subject
                                            mySubject = (From issue In newContext.AutoIssues
                                                         Where issue.IssueID = issueId
                                                         Select issue.Subject).Single()
                                        End Using
                                        Common.LogText("循环执行" + dllType.ToString)
                                        '填充模板，创建campaign，发送
                                        If (dllType = "lifeinhk") Then '2013/05/27新增需求，要求兼容手机版和电脑版模板
                                            Dim campaignName As String = "LIFEINHK" & list.PlanType.Trim() & DateTime.Now.ToString("yyyyMMdd")
                                            Dim listProductElements As List(Of ProductElement) = GetProductElements(list.SiteId, issueId)
                                            If (listProductElements.Count > 0) Then
                                                Dim SpreadTemplate() As String = GetSpreadTemplate2(list.SiteId, issueId, list.TemplateId, list.SenderName, list.PlanType.Trim(), campaignName, listProductElements, "", list.SpreadLoginEmail, list.AppId)
                                                Try
                                                    MyRender.Start(issueId, mySiteId, list.SpreadLoginEmail, list.AppId, list.SenderName,
                                                               list.SenderEmail, SpreadTemplate(0), mySubject, list.PlanType.Trim(), campaignName,
                                                               list.Volumn, dllType, list.CampaignStatus, list.ScheduleTimeInteval,
                                                               list.ListType, list.ContactList, list.SearchAPIType,
                                                               list.SearchStartDayInterval, list.SearchEndDayInterval, list.UrlSpecialCode,
                                                               list.Categories, list.LimitQuantity, list.TimeLimit, list.SellerEmail)
                                                Catch ex As Exception
                                                    UpdateIssueStatus(issueId, "EU") 'EU(其他)
                                                    Throw New Exception(ex.ToString)
                                                End Try
                                            End If
                                            logList = InsertSentLog(mySiteId, mySubject, list.PlanType.Trim(), logList)
                                            Common.LogText("SiteID:" & mySiteId & ",PlanType:" & list.PlanType.Trim() & ",Subject:" & mySubject)
                                            UpdateIssueStatus(issueId, "ES") 'ES(Success)
                                        ElseIf (dllType = "groupbuyer") Then
                                            Dim template As String = CreateGroupBuyerTemplate.GetTemplate(list.TemplateId, issueId, list.UrlSpecialCode)
                                            Dim campaignName As String = Format(Now.Date, "dd/MM/yyyy").ToString
                                            Dim CID As String = mySubject.Substring(mySubject.Length - 1)
                                            If (CID = 5) Then
                                                campaignName = campaignName & "(favorite:food-5-美食專享)"
                                            ElseIf (CID = 2) Then
                                                campaignName = campaignName & "(favorite:beauty-2-beauty deals)"
                                            End If
                                            mySubject = mySubject.Substring(0, mySubject.Length - 1)
                                            Try
                                                'Render(issueId, mySiteId, list.SpreadLoginEmail, list.AppId, list.SenderName, list.SenderEmail, template, mySubject, list.PlanType.Trim(), campaignName, list.Volumn, dllType)
                                                MyRender.Start(issueId, mySiteId, list.SpreadLoginEmail, list.AppId, list.SenderName,
                                                               list.SenderEmail, template, mySubject, list.PlanType.Trim(), campaignName,
                                                               list.Volumn, dllType, list.CampaignStatus, list.ScheduleTimeInteval,
                                                               list.ListType, list.ContactList, list.SearchAPIType,
                                                               list.SearchStartDayInterval, list.SearchEndDayInterval, list.UrlSpecialCode,
                                                               list.Categories, list.LimitQuantity, list.TimeLimit, list.SellerEmail)

                                                logList = InsertSentLog(mySiteId, mySubject, list.PlanType.Trim(), logList)
                                                Common.LogText("SiteID:" & mySiteId & ",PlanType:" & list.PlanType.Trim() & ",Subject:" & mySubject)
                                                UpdateIssueStatus(issueId, "ES") 'ES(Success)
                                            Catch ex As Exception
                                                UpdateIssueStatus(issueId, "EU") 'EU(其他)
                                                Throw New Exception(ex.ToString)
                                            End Try
                                        Else
                                            Common.LogText("结果记录测试点1")
                                            '如本期邮件无任何符合条件的产品，则不进入到创建campaign环节，但同样出提醒邮件,added by Dora 2015-02-02
                                            If (IsContentNull(issueId)) Then
                                                Common.LogText("结果记录测试点2")
                                                Dim subject As String = list.SiteName & "_" & list.SenderName & DateTime.Now.ToString("yyyy-MM-dd") & " 智能化邮件创建详情"
                                                Dim body As String = "<span style=""font-size: 14px; line-height: 16px; font-family: 微软雅黑, arial, helvetica, sans-serif; "">"
                                                body = body & "您好：<br/>"
                                                body = body & "本期智能化邮件内容为空，导致无相应邮件创建出。内容为空的原因可能为：<br/>"
                                                body = body & "&nbsp;1. 产品一月内不重复限定<br/>"
                                                body = body & "&nbsp;2. 您的定制方案中的特殊限定规则"
                                                body = body & "<br/>如对以上信息有任何疑问，请咨询您的客户经理或发邮件给我们<a href=""mailto:spread@reasonables.com"" target=""_blank"">spread@reasonables.com</a>"
                                                body = body & "</span><br/>"
                                                Common.LogText("结果记录测试点10")
                                                Common.LogText(subject.ToString)
                                                Common.LogText(body.ToString)
                                                NotificationEmail.SentStartEmail(subject, body)
                                                Common.LogText("结果记录测试点9")
                                                If Not (String.IsNullOrEmpty(list.SellerEmail)) Then '如已配置销售人员邮箱,则发送提醒邮件到销售人员邮箱中
                                                    If (list.SellerEmail.Contains(",")) Then
                                                        Dim arrSellerMail As String() = list.SellerEmail.Split(",")
                                                        For Each mail As String In arrSellerMail
                                                            NotificationEmail.SentStartEmail(subject, body, mail)
                                                        Next
                                                    Else
                                                        NotificationEmail.SentStartEmail(subject, body, list.SellerEmail)
                                                    End If
                                                End If
                                                logList = InsertSentLog(mySiteId, mySubject, list.PlanType.Trim(), logList)
                                                Common.LogText("SiteID:" & mySiteId & ",PlanType:" & list.PlanType.Trim() & ",Subject:" & mySubject)
                                                UpdateIssueStatus(issueId, "ES") 'ES(Success)
                                                Continue For
                                            End If
                                            '如本期邮件无任何符合条件的产品，则不进入到创建campaign环节，但同样出提醒邮件,added by Dora 2015-02-02
                                            Dim compaignName As String
                                            If (list.PlanType.Trim() = "HN") Then
                                                compaignName = list.SenderName.ToUpper & "NotOpen" & DateTime.Now.ToString("yyyyMMdd")
                                            Else
                                                compaignName = list.SenderName.ToUpper() & "_" & list.SiteId & list.PlanType.Trim() & "_Auto" & DateTime.Now.ToString("yyyyMMdd")
                                            End If
                                            Common.LogText("结果记录测试点3")
                                            '2014/02/17 added，分类促发，begin
                                            Dim SpreadTemplate As String = ""
                                            If (String.IsNullOrEmpty(list.Categories)) Then '非分类促发，模板填充
                                                SpreadTemplate = GetSpreadTemplate(mySiteId, issueId, list.TemplateId, list.SenderName, list.PlanType.Trim(), compaignName, dllType, list.Url, list.UrlSpecialCode)
                                            Else  '分类促发模板填充
                                                SpreadTemplate = SortCategory.GetSpreadTemplate(mySiteId, issueId, list.TemplateId, list.SiteName, list.PlanType.Trim(), compaignName, dllType, list.Url, list.Categories, list.UrlSpecialCode, list.LogoUrl)
                                                If (Not String.IsNullOrEmpty(list.LogoUrl) AndAlso SpreadTemplate.Contains("[LOGO_URL]")) Then
                                                    SpreadTemplate = SpreadTemplate.Replace("[LOGO_URL]", list.LogoUrl)
                                                End If
                                                '填充[BEGIN_ADS]  [BEGIN_DAILY_DEALS]/[BEGIN_NEW_ARRIVALS] 部分
                                                SpreadTemplate = FillSpreadTemplate(SpreadTemplate, mySiteId, issueId, list.PlanType)
                                                If (dllType = "everbuying") Then
                                                    listCategorys = efContext.AutoCategories.Where(Function(c) c.SiteID = mySiteId).ToList()
                                                    SpreadTemplate = AddSpecialCode(SpreadTemplate, listCategorys, list.SenderName, compaignName, list.PlanType)
                                                End If
                                            End If
                                            Common.LogText("结果记录测试点4")
                                            Try
                                                ''2014/04/22,begin adding,EmailAlerter主程序优化，减少对dll的修改次数
                                                MyRender.Start(issueId, mySiteId, list.SpreadLoginEmail, list.AppId, list.SenderName,
                                                               list.SenderEmail, SpreadTemplate, mySubject, list.PlanType.Trim(), compaignName,
                                                               list.Volumn, dllType, list.CampaignStatus, list.ScheduleTimeInteval,
                                                               list.ListType, list.ContactList, list.SearchAPIType,
                                                               list.SearchStartDayInterval, list.SearchEndDayInterval, list.UrlSpecialCode,
                                                               list.Categories, list.LimitQuantity, list.TimeLimit, list.SellerEmail)
                                                ''2014/04/22,end
                                                Common.LogText("结果记录测试点5")
                                                logList = InsertSentLog(mySiteId, mySubject, list.PlanType.Trim(), logList)
                                                Common.LogText("SiteID:" & mySiteId & ",PlanType:" & list.PlanType.Trim() & ",Subject:" & mySubject)
                                                UpdateIssueStatus(issueId, "ES") 'ES(Success)
                                            Catch ex As Exception
                                                UpdateIssueStatus(issueId, "EU") 'EU(其他)
                                                Throw New Exception(ex.ToString)
                                            End Try
                                        End If

                                        '填充模板，创建campaign，发送

                                    End If
                                End If
                            Else 'NFE,Notification For Expiration
                                If list.DllType.Trim().ToLower() = "groupbuyer" Then
                                    'Groupbuyer专用点击提醒
                                    Try
                                        GetSpreadReportData.Start(list.SpreadLoginEmail, list.AppId, list.SiteId, planType, list.TemplateId, list.SenderEmail, list.SenderName, list.TriggerForNFC, list.UrlSpecialCode)
                                        logList = InsertSentLog(mySiteId, "Notification For Expiration", planType, logList)
                                    Catch ex As Exception
                                        logList = InsertSentLog(mySiteId, "Error:Notification For Expiration", planType, logList)
                                        Common.LogText(ex.ToString())
                                    End Try
                                Else
                                    Try
                                        Dim NFC As New NotificationForClick()
                                        Common.LogText("NFC Start.")
                                        Common.LogText("siteid:" & mySiteId & list.SenderName)
                                        NFC.start(list.DllType, list.SpreadLoginEmail, list.AppId, list.SiteId, list.TriggerForNFC, list.SubjectForNFC, list.TemplateId, list.SenderEmail,
                                            list.SenderName, list.UrlSpecialCode)
                                        logList = InsertSentLog(mySiteId, list.DllType & " " & list.SenderName & " Notification For Expiration", planType, logList)
                                    Catch ex As Exception
                                        logList = InsertSentLog(mySiteId, "Error:Notification For Expiration", planType, logList)
                                        Common.LogText(ex.ToString())
                                    End Try
                                    '通用的点击提醒方法
                                    '获取Spread报告数据，并且插入到Opens/Clicks/Convertion，注意需要根据Trigger字段的大小，判断该获取哪些Campaignid（默认是Trigger + 7天内的CampaigneID），同时注意转换数据的获取时间范围要加大
                                    '分析数据：根据Trigger时间,整理对应的发送联系人和邮件内容+标题
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try
        'Try
        '    SortCategory.Start(nowTime, logList) '分类促发
        'Catch ex As Exception
        '    LogText(ex.ToString())
        '    'UpdateIssueStatus(issueId, "EU") 'EU(其他)
        'End Try
        Return logList
    End Function

    ''' <summary>
    ''' 判断本期邮件中是否有内容，如ads_issue及product_issue表均无数据则认定为content null;if null ,return true ,else false
    ''' </summary>
    ''' <param name="issueid"></param>
    ''' <returns></returns>
    ''' <remarks>author：Dora，Created in 2015-01-31</remarks>
    Private Function IsContentNull(ByVal issueid As Integer) As Boolean
        Dim adCount As Integer = efContext.AutoAds_Issue.Count(Function(m) m.IssueID = issueid)
        If (adCount > 0) Then
            Return False
        End If
        Dim productCount As Integer = efContext.AutoProducts_Issue.Count(Function(m) m.IssueID = issueid)
        If (productCount > 0) Then
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 将数据写入到Issues表中,
    ''' 2013/06/24新增标识状态位
    ''' </summary>
    ''' <param name="nowTime"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <returns>返回插入单条数据的IssueId</returns>
    ''' <remarks></remarks>
    Public Function InsertIssue(ByVal nowTime As DateTime, ByVal siteId As Integer, ByVal planType As String) As Integer
        '2013/06/24修改
        Dim queryIssue As AutoIssue = efContext.AutoIssues.Where(Function(iss) iss.SiteID = siteId AndAlso iss.PlanType = planType).OrderByDescending(Function(i) i.IssueID).FirstOrDefault
        Dim issueId As Integer = 0

        'SentStatus的值：EV(Error Volmn),ET(Error Template),ES(Success),EU(其他),EM(Error while Men Revised)
        If (queryIssue Is Nothing OrElse String.IsNullOrEmpty(queryIssue.SentStatus) OrElse queryIssue.SentStatus = "ES" _
            OrElse queryIssue.SentStatus = "EM") Then
            Dim issue As New AutoIssue()
            issue.IssueDate = nowTime
            issue.Subject = ""
            issue.SiteID = siteId
            issue.PlanType = planType
            efContext.AddToAutoIssues(issue)
            Try
                efContext.SaveChanges()
            Catch ex As Exception
                Common.LogText(ex.ToString())
            End Try
            issueId = issue.IssueID
        End If
        Return issueId
    End Function

    ''' <summary>
    ''' 根据siteId和planType获取SentLogs表中上一次发送时间，如果返回Nothing，则说明上一次没发送
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="logList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLastSentTime(ByVal siteId As Integer, ByVal planType As String, ByVal logList As List(Of AutoSentLog)) As String
        Dim lastSentTime As DateTime
        Dim mylastTime As DateTime
        Try
            lastSentTime = (From issue In efContext.AutoSentLogs
                  Where issue.SiteID = siteId AndAlso issue.PlanType = planType
                  Select issue.LastSentAt).Max()
        Catch ex As Exception
        End Try
        For Each li As AutoSentLog In logList
            If (li.SiteID = siteId AndAlso li.PlanType = planType) Then
                If (lastSentTime < li.LastSentAt) Then
                    lastSentTime = li.LastSentAt
                End If
            End If
        Next
        'If Not (lastSentTime = "") Then 
        If Not (lastSentTime = mylastTime) Then
            Return lastSentTime.ToString() '2013/4/2添加，防止接受参数报错
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 创建邮件模板，并创建Campaign后发送
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <remarks></remarks>
    Public Sub Render(ByVal issueId As Integer, ByVal siteId As Integer, _
                       ByVal spreadName As String, ByVal password As String, ByVal spreadSenderName As String, _
                       ByVal spreadSenderMail As String, ByVal AutoSpreadTemplate As String, _
                       ByVal subject As String, ByVal planType As String, ByVal compaignName As String, _
                       ByVal volumn As Integer, ByVal dllType As String)
        Dim SpreadTemplate As String = AutoSpreadTemplate

        While SpreadTemplate.Contains("[CAMPAIGN_NAME]")
            SpreadTemplate = SpreadTemplate.Replace("[CAMPAIGN_NAME]", compaignName)
        End While

        Dim loginName As String = spreadName
        Dim loginPwd As String = password
        Dim senderName As String = spreadSenderName
        Dim senderMail As String = spreadSenderMail

        '2013/4/15修改，使用LINQ方法减少查询次数
        Dim listSender As List(Of AutoContactLists_Issue) = (From c In efContext.AutoContactLists_Issue _
                         Where c.IssueID = issueId AndAlso (c.SendTime <= Now OrElse String.IsNullOrEmpty(c.SendTime)) _
                         AndAlso (String.IsNullOrEmpty(c.SendingStatus) OrElse c.SendingStatus.ToLower = "waiting") _
                         Select c).ToList()
        Dim listSenderDraft As List(Of AutoContactLists_Issue) = (From contact In efContext.AutoContactLists_Issue _
                                                                Where contact.SendingStatus.ToLower = "draft" AndAlso contact.IssueID = issueId
                                                                Select contact).ToList()
        Dim counter1 As Integer = listSender.Count
        Dim counterDraft As Integer = listSenderDraft.Count
        Dim lists() As String = New String(counter1 - 1) {} '立即发送的联系人列表
        Dim listsDraft() As String = New String(counterDraft - 1) {} '草稿状态的联系人列表
        Dim subscriptionCounter As Integer = 0  '2013/06/24新增，制定list下的订阅者人数
        Dim volumnFlag As Boolean = True  '2013/06/24新增，流量标志，默认状态下流量足够
        Dim AutoSpread As New SpreadService.Service
        AutoSpread.Timeout = 120000 '2013/08/08 added, Spread response time out
        '没有发送计划的list
        If (listSender.Count > 0) Then
            Dim counter As Integer = 0
            Dim subscritionDS As DataSet = AutoSpread.getSubscription(loginName, loginPwd, EmailAlerter.SpreadService.SubscriptionStatus.Any) '2013/06/24新增
            For Each sender As AutoContactLists_Issue In listSender
                Dim targetSubscription As String = sender.ContactList
                lists(counter) = targetSubscription
                counter = counter + 1
                If (volumn > 0) Then '需要统计流量账户的流程，volumn=-1则不需要统计账户流量
                    '2013/06/24新增，统计发送lists的subscription个数
                    For i As Integer = 0 To subscritionDS.Tables(0).Rows.Count - 1
                        If (subscritionDS.Tables(0).Rows(i).Item(0).ToString.Trim.ToUpper = sender.ContactList.Trim.ToUpper) Then
                            'subscritionDS.Tables(0).Rows(i).Item(3)表示某个list下的订阅者个数
                            subscriptionCounter = subscriptionCounter + subscritionDS.Tables(0).Rows(i).Item(3)
                        End If
                    Next
                End If
            Next
            If (subscriptionCounter > volumn AndAlso (Not volumn = -1)) Then '流量不够，流量标志位置为False状态
                volumnFlag = False
                UpdateIssueStatus(issueId, "EV")  '更新Issues表的发送状态，EV(Error Volumn)
            End If
        End If
        '有发送计划的list
        If (listSenderDraft.Count > 0) Then
            Dim myCounter As Integer = 0
            For Each contact As AutoContactLists_Issue In listSenderDraft
                listsDraft(myCounter) = contact.ContactList
                myCounter = myCounter + 1
            Next
        End If

        'If (lists.Count > 0) Then '2013/4/11新增，立即发送的联系人列表策略
        If (lists.Count > 0 AndAlso volumnFlag) Then '2013/06/24新增，新增流量判断
            '没有发送计划
            Dim AutoCampaign As New EmailAlerter.SpreadService.Campaign()
            AutoCampaign.subject = subject
            AutoCampaign.from = senderName
            AutoCampaign.fromEmail = senderMail
            AutoCampaign.content = SpreadTemplate
            AutoCampaign.campaignName = compaignName
            AutoCampaign.schedule = Now
            Dim campaignId As Integer = AutoSpread.createCampaign(loginName, loginPwd, AutoCampaign, lists, -1)
            AutoSpread.ChangePublishStatus(loginName, loginPwd, campaignId, True) 'public邮件到邮件档案馆

            'add by gary start 20130730
            Try
                Dim subject2 = senderName & " 开始发送邮件"
                Dim myList As String = ""
                For Each list As String In lists
                    myList = myList & list & ","
                Next
                'Dim SpreadTemplate2 = subject2 & "<br />状态：正在发送<br />" & " 标题:" & subject & "<br />内容：<br />" & SpreadTemplate
                Dim SpreadTemplate2 As String = subject2 & "<br />状态：正在发送<br />" & " 标题:" & subject _
                                                & "<br />发送对象：" & myList & "<br />发件人Email：" & senderMail & "<br />发件人：" & senderName _
                                                & "<br />Campaign名：" & compaignName & "<br />是否Publish to Newsletter Archive：" & "是" _
                                                & "<br />内容(*部分Spread系统按钮在提醒邮件里会失效，但不影响正常的发送)：<br />" & SpreadTemplate
                '2013/08/08 added
                'AutoSpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "自动化发送", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
                NotificationEmail.SentStartEmail(subject2, SpreadTemplate2)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'add by gary end 20130730

            '2013/06/24新增，Campaign开始时，改变账户下的总流量
            If (volumn >= 0) Then '需要统计账户流量的流程
                Dim queryAutoSite As AutomationSite = efContext.AutomationSites.Where(Function(autoSite) autoSite.siteid = siteId).SingleOrDefault
                queryAutoSite.volumn = volumn - subscriptionCounter
                efContext.SaveChanges()
            End If
        End If
        If listsDraft.Count > 0 Then
            Dim campaignCreative(0) As EmailAlerter.SpreadService.CampaignCreatives
            Dim c As EmailAlerter.SpreadService.CampaignCreatives = New EmailAlerter.SpreadService.CampaignCreatives
            With c
                .creativeContent = SpreadTemplate
                .displayName = senderName
                .fromAddress = senderMail
                .isCampaignDefault = True
                .replyTo = ""
                .subject = subject
                .target = "D"
            End With
            campaignCreative(0) = c
            Dim interval As Integer = -1

            ''2014/02/10 added,设置为Schedule发送，Schedule时间为创建Campaign时间延后两个小时,begin
            'Dim campaignId As Integer=AutoSpread.createCampaign2(loginName, loginPwd, compaignName, campaignCreative, listsDraft, interval, Now, "Spread", EmailAlerter.SpreadService.CampaignStatus.Draft)
            Dim campaignId As Integer = 0  'AutoSpread.createCampaign2(loginName, loginPwd, compaignName, campaignCreative, listsDraft, interval, Now, "Spread", EmailAlerter.SpreadService.CampaignStatus.Draft)
            If (dllType.ToLower = "sammydress" OrElse dllType.ToLower = "ahappydeal" OrElse dllType.ToLower = "eachbuyer" OrElse dllType.ToLower = "utsource") Then
                campaignId = AutoSpread.createCampaign2(loginName, loginPwd, compaignName, campaignCreative, listsDraft, interval, Now.AddHours(2), "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
            ElseIf (dllType.ToLower = "oasap") Then
                campaignId = AutoSpread.createCampaign2(loginName, loginPwd, compaignName, campaignCreative, listsDraft, interval, Now.AddHours(8), "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
            Else
                campaignId = AutoSpread.createCampaign2(loginName, loginPwd, compaignName, campaignCreative, listsDraft, interval, Now, "Spread", EmailAlerter.SpreadService.CampaignStatus.Draft)
            End If
            ''2014/02/10 added,设置为Schedule发送，Schedule时间为创建Campaign时间延后两个小时,end

            '2014/03/25 added,产品过期提醒使用
            ' AutoSpread.ChangePublishStatus(loginName, loginPwd, campaignId, True)
            If (campaignId > 0) Then
                AutoSpread.ChangePublishStatus(loginName, loginPwd, campaignId, True)
                Dim myIssue As AutoIssue = efContext.AutoIssues.Where(Function(issue) issue.IssueID = issueId).Single()
                myIssue.SpreadCampId = campaignId
                efContext.SaveChanges()
            End If

            'add by gary start 20130730
            Try
                Dim subject2 = senderName & " 草稿创建成功"
                'Dim SpreadTemplate2 = subject2 & "<br />状态：草稿<br />标题:" & subject & "<br />内容：<br />" & SpreadTemplate
                'AutoSpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "自动化草稿", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
                Dim myList As String = ""
                For Each list As String In listsDraft
                    myList = myList & list & ","
                Next
                Dim SpreadTemplate2 As String = subject2 & "<br />状态：正在发送<br />" & " 标题:" & subject _
                                                & "<br />发送对象：" & myList & "<br />发件人Email：" & senderMail & "<br />发件人：" & senderName _
                                                & "<br />Campaign名：" & compaignName & "<br />是否Publish to Newsletter Archive：" & "是" & "<br />内容：<br />" & SpreadTemplate
                NotificationEmail.SentStartEmail(subject2, SpreadTemplate2)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'add by gary end 20130730
        End If
        '2013/4/11新增，对有发送计划的联系人列表创建新的
        Dim querySchedule = From c In efContext.AutoContactLists_Issue _
                          Where c.IssueID = issueId AndAlso c.SendTime > Now _
                          Select c
        For Each schedule As AutoContactLists_Issue In querySchedule
            Dim AutoCampaign2 As New EmailAlerter.SpreadService.Campaign()
            AutoCampaign2.subject = subject
            AutoCampaign2.from = senderName
            AutoCampaign2.fromEmail = senderMail
            AutoCampaign2.content = SpreadTemplate
            AutoCampaign2.schedule = schedule.SendTime
            AutoCampaign2.campaignName = compaignName
            Dim lists2() As String = New String(0) {}
            lists2(0) = schedule.ContactList
            Dim campaignId As Integer = AutoSpread.createCampaign(loginName, loginPwd, AutoCampaign2, lists2, -1)
            AutoSpread.ChangePublishStatus(loginName, loginPwd, campaignId, True)
            'Using efContext2 As New FaceBookForSEOEntities '防止缓存
            '    Dim queryIssue As AutoIssue = efContext2.AutoIssues.Where(Function(iss) iss.IssueID = issueId).SingleOrDefault()
            '    queryIssue.SentStatus = "ES"
            '    efContext2.SaveChanges()
            'End Using

            'add by gary start 20130730
            Try
                Dim subject2 = senderName & " 开始发送邮件"
                'Dim SpreadTemplate2 = subject2 & "<br />状态：计划发送<br />发送时间：" & schedule.SendTime.ToString & "<br /> 标题:" & subject & "<br />内容：<br />" & SpreadTemplate
                'AutoSpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "自动化发送", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
                Dim myList As String = schedule.ContactList
                Dim SpreadTemplate2 As String = subject2 & "<br />状态：正在发送<br />" & " 标题:" & subject _
                                                & "<br />发送对象：" & myList & "<br />发件人Email：" & senderMail & "<br />发件人：" & senderName _
                                                & "<br />Campaign名：" & compaignName & "<br />是否Publish to Newsletter Archive：" & "是" & "<br />内容：<br />" & SpreadTemplate
                NotificationEmail.SentStartEmail(subject2, SpreadTemplate2)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'add by gary end 20130730

        Next
    End Sub

    ''' <summary>
    ''' 创建邮件模板,包括获取填充好数据的模板、添加link_category、添加追踪代码
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSpreadTemplate(ByVal siteId As Integer, ByVal issueId As Integer, ByVal templateId As Integer, _
                                       ByVal senderName As String, ByVal planType As String, _
                                       ByVal compaignName As String, ByVal dllType As String, ByVal Url As String, ByVal UrlSpecialCode As String) As String
        Try
            Dim queryCate = From c In efContext.AutoCategories
                        Where c.SiteID = siteId
                        Select c
            Dim listCategory As New List(Of AutoCategory)
            For Each q As AutoCategory In queryCate
                Dim cate As New AutoCategory()
                cate.Category1 = q.Category1
                cate.CategoryID = q.CategoryID
                listCategory.Add(cate)
            Next
            Dim SpreadTemplate As String = efContext.AutoTemplates.Single(Function(t) t.Tid = templateId).Contents

            '填充模板，返回填充好的模板
            SpreadTemplate = FillSpreadTemplate(SpreadTemplate, siteId, issueId, planType)
            '添加link_category
            SpreadTemplate = HandleCategory(SpreadTemplate, siteId)
            SpreadTemplate = SpreadTemplate.Replace("[CURRENT_MONTH_MMM]", Date.Now.ToString("MMM"))
            SpreadTemplate = SpreadTemplate.Replace("[CURRENT_MONTH_MM]", Date.Now.ToString("MM"))
            '2014/04/24 added,begin
            If Not (String.IsNullOrEmpty(UrlSpecialCode)) Then  '半固定追踪代码，即不是很特殊的追踪代码
                SpreadTemplate = SpecialCode.AddSpecialCode(UrlSpecialCode, SpreadTemplate)
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", senderName)
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", Url)
                Return SpreadTemplate
            End If
            'end

            '2013/05/07新增，客户在没有要求加追踪代码的时候，不加追踪代码
            Select Case dllType
                Case "dresslily"
                    SpreadTemplate = AddSpecialCode(SpreadTemplate, listCategory, senderName, compaignName, planType)
                Case "everbuying"
                    SpreadTemplate = AddSpecialCode(SpreadTemplate, listCategory, senderName, compaignName, planType)
                Case "dressvenus"
                    Dim codeType As String = "?utm_source=EDM&utm_medium=email&utm_campaign=auto"
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "focalprice"
                    '请在此处添加FocalPrice的追踪代码
                    'URL?utm_source=RSpread&utm_medium=EM&utm_campaign=DM_1350ES_MH0953W
                    SpreadTemplate = SpecialCode.AddSpecialCodeForFocalPrice(SpreadTemplate, "ES")
                Case "focalpricede"
                    '请在此处添加FocalPrice的追踪代码
                    'URL?utm_source=RSpread&utm_medium=EM&utm_campaign=DM_1350ES_MH0953W
                    Dim ci As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture
                    Dim cal As System.Globalization.Calendar = ci.Calendar
                    Dim cwr As System.Globalization.CalendarWeekRule = ci.DateTimeFormat.CalendarWeekRule
                    Dim dow As DayOfWeek = ci.DateTimeFormat.FirstDayOfWeek
                    Dim nowWeek As String = cal.GetWeekOfYear(DateTime.Now, cwr, dow).ToString() '当前时间是第几周
                    SpreadTemplate = SpreadTemplate.Replace("[WEEK]", nowWeek)
                Case "sammydress"
                    '?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular.20131209    
                    Dim codeType As String = "?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular." & DateTime.Now.ToString("yyyyMMdd")
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "ahappydeal"
                    '?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular.20131209   
                    Dim codeType As String = "?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular." & DateTime.Now.ToString("yyyyMMdd")
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "oasap", "dresslilynew"
                    '?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular.20140219
                    Dim codeType As String = "?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular." & DateTime.Now.ToString("yyyyMMdd")
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "eachbuyer"
                    Dim codeType As String = "?utm_source=affiliate&utm_medium=EDM&utm_campaign=spead_mail_" & DateTime.Now.ToString("yyyyMMdd")
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "utsource"
                    Dim codeType As String = "?utm_source=siqi&utm_medium=siqi&utm_campaign=siqi"
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                Case "lightake"
                    Dim codeType As String = "?utm_source=Rspread&utm_medium=email&utm_campaign=AutoRspread"
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
            End Select
            'For taobao/tmall add shopName
            SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", senderName)
            SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", Url)
            Return SpreadTemplate
        Catch ex As Exception
            Using efContext2 As New FaceBookForSEOEntities '防止缓存,Issues表的发送状态位Error Template
                Dim queryIssue As AutoIssue = efContext2.AutoIssues.Where(Function(iss) iss.IssueID = issueId).SingleOrDefault()
                queryIssue.SentStatus = "ET"  'ET-Error Template
                efContext2.SaveChanges()
            End Using
            'a.LogText(ex.ToString())
            'a.LogText(ex.InnerException.Message)
            Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
        End Try

    End Function

    ''' <summary>
    ''' 创建邮件模板，兼容电脑版和手机版
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="templateId"></param>
    ''' <param name="senderName"></param>
    ''' <param name="planType"></param>
    ''' <param name="compaignName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSpreadTemplate2(ByVal siteId As Integer, ByVal issueId As Integer, ByVal templateId As Integer, _
                                       ByVal senderName As String, ByVal planType As String, ByVal compaignName As String, _
                                       ByVal listProductItems As List(Of ProductElement), ByVal categoryName As String, _
                                       ByVal loginEmail As String, ByVal appId As String) As String()
        Try
            Dim listTemplate As List(Of AutoTemplate) = (From template In efContext.AutoTemplates
                                                    Where template.SiteID = siteId AndAlso template.PlanType = planType
                                                    Select template).ToList()

            '根据点击量最大的category，排序products List，并把最受欢迎的产品插入到List中
            If Not String.IsNullOrEmpty(categoryName) Then
                Dim arrProductItems As ProductElement() = listProductItems.ToArray()
                Dim comparer As New MyComparer
                comparer.Favorite = categoryName
                Try
                    Array.Sort(arrProductItems, comparer)
                Catch ex As Exception
                    'Ignore
                End Try
                listProductItems = arrProductItems.ToList()
                'Try
                '    listProductItems = InsertPopularElements(loginEmail, appId, listProductItems)
                'Catch ex As Exception
                '    LogText(ex.ToString())
                'End Try
            End If

            Dim listCounter As Integer = listTemplate.Count
            Dim myTemplate(listCounter - 1) As String
            If (listCounter = 0) Then
                Try
                    Dim SpreadTemplate As String = efContext.AutoTemplates.Where(Function(t) t.Tid = templateId).SingleOrDefault().Contents
                    ReDim myTemplate(0)
                    SpreadTemplate = FillSpreadTemplate(SpreadTemplate, listProductItems, siteId, planType, senderName, loginEmail, appId)
                    myTemplate(0) = SpreadTemplate
                Catch ex As Exception
                    Using efContext2 As New FaceBookForSEOEntities '防止缓存,Issues表的发送状态位Error Template
                        Dim queryIssue As AutoIssue = efContext2.AutoIssues.Where(Function(iss) iss.IssueID = issueId).SingleOrDefault()
                        queryIssue.SentStatus = "ET"  'ET-Error Template
                        efContext2.SaveChanges()
                    End Using
                    Common.LogText(ex.ToString())
                    Throw New Exception(ex.ToString)
                End Try
            Else
                '根据siteId和planType找到模板
            End If
            Return myTemplate
        Catch ex As Exception
            Common.LogText(ex.ToString())
            Throw New Exception(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' 使用数据库中的数据填充模板，返回填充好的模板
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FillSpreadTemplate(ByVal SpreadTemplate As String, ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String) As String
        Try
            Try
                '2013/4/15之后的版本
                If (SpreadTemplate.Contains("[BEGIN_PROMOTION]")) Then
                    'Dim promotionCount As Integer = (efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                    Dim promotionCount As Integer = (From ad In efContext.AutoAds
                                                    Join adIssue In efContext.AutoAds_Issue On ad.AdID Equals adIssue.AdId
                                                    Where adIssue.SiteId = siteId AndAlso adIssue.IssueID = issueId AndAlso ad.Type = "P"
                                                    Select adIssue).Count()
                    'Dim promotionCount As Integer =(efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                    If (promotionCount = 0) Then  '如果没有Promotion产品，则删除Promotion块
                        Dim endLen As Integer = "[END_PROMOTION]".Length
                        Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_PROMOTION]")
                        Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_PROMOTION]")
                        SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                    Else
                        For i As Integer = 0 To promotionCount - 1
                            If (SpreadTemplate.Contains("[BEGIN_PROMOTION]")) Then
                                Dim beginLen As Integer = "[BEGIN_PROMOTION]".Length
                                Dim endLen As Integer = "[END_PROMOTION]".Length
                                Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_PROMOTION]")
                                Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_PROMOTION]")
                                Dim oldPromotion As String = SpreadTemplate.Substring(startProIndex, endProIndex - startProIndex + endLen)
                                Dim promotion As String = SpreadTemplate.Substring(startProIndex + beginLen, endProIndex - startProIndex - beginLen)
                                Dim newPromotion As String = InsertPromotionAndAd(promotion, issueId, siteId, i)
                                'SpreadTemplate = SpreadTemplate.Replace(oldPromotion, newPromotion)
                                SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                                SpreadTemplate = SpreadTemplate.Insert(startProIndex, newPromotion)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
            Catch ex1 As Exception
                Common.LogText(ex1.ToString())
                Common.LogText(ex1.InnerException.Message.ToString())
                Throw New Exception(ex1.ToString()) 'throw the exception to the upper layer
            End Try

            Try
                '2013/07/04 added,ads块填充
                If (SpreadTemplate.Contains("[BEGIN_ADS]")) Then
                    If (SpreadTemplate.Contains("[BEGIN_ADS]")) Then
                        'Dim promotionCount As Integer = (efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                        '2013/07/04 added
                        'Dim adsCount As Integer = (efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                        Dim adsCount As Integer = (From adIssue In efContext.AutoAds_Issue
                                                Join ad In efContext.AutoAds On adIssue.AdId Equals ad.AdID
                                                Where adIssue.SiteId = siteId AndAlso adIssue.IssueID = issueId AndAlso (String.IsNullOrEmpty(ad.Type) OrElse Not (ad.Type = "P" Or ad.Type = "B"))
                                                Select adIssue).Count

                        '如果在Ads表中没查询到数据，则把[BEGIN_ADS]...[END_ADS]删除
                        If (adsCount = 0) Then  '没查到Ads表中有数据，则删除[BEGIN_ADS]块
                            While (SpreadTemplate.Contains("[BEGIN_ADS]"))
                                Dim endLen As Integer = "[END_ADS]".Length
                                Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_ADS]")
                                Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_ADS]")
                                SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                            End While
                        Else
                            Dim icounter As Integer = 0
                            While (SpreadTemplate.Contains("[BEGIN_ADS]"))
                                Dim beginLen As Integer = "[BEGIN_ADS]".Length
                                Dim endLen As Integer = "[END_ADS]".Length
                                Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_ADS]")
                                Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_ADS]")
                                Dim oldPromotion As String = SpreadTemplate.Substring(startProIndex, endProIndex - startProIndex + endLen)
                                Dim adsTemplate As String = SpreadTemplate.Substring(startProIndex + beginLen, endProIndex - startProIndex - beginLen)
                                Dim newPromotion As String = InsertAds(adsTemplate, issueId, siteId, icounter, "")
                                'SpreadTemplate = SpreadTemplate.Replace(oldPromotion, newPromotion)
                                SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                                SpreadTemplate = SpreadTemplate.Insert(startProIndex, newPromotion)
                                icounter += 1
                                If (icounter = adsCount) Then
                                    While (SpreadTemplate.Contains("[BEGIN_ADS]"))
                                        Dim endLen1 As Integer = "[END_ADS]".Length
                                        Dim startProIndex1 As Integer = SpreadTemplate.IndexOf("[BEGIN_ADS]")
                                        Dim endProIndex1 As Integer = SpreadTemplate.IndexOf("[END_ADS]")
                                        SpreadTemplate = SpreadTemplate.Remove(startProIndex1, endProIndex1 - startProIndex1 + endLen1)
                                    End While
                                    Exit While
                                End If
                            End While
                        End If
                    End If
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try

            Try
                '2013/07/04 added,ads块填充
                If (SpreadTemplate.Contains("[BEGIN_LITTLE_ADS]")) Then
                    'Dim promotionCount As Integer = (efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                    '2013/07/04 added
                    'Dim adsCount As Integer = (efContext.AutoAds_Issue.Where(Function(adIss) adIss.SiteId = siteId AndAlso adIss.IssueID = issueId)).Count
                    Dim adsCount As Integer = (From adIssue In efContext.AutoAds_Issue
                                            Join ad In efContext.AutoAds On adIssue.AdId Equals ad.AdID
                                            Where adIssue.SiteId = siteId AndAlso adIssue.IssueID = issueId AndAlso ad.Type = "B"
                                            Select adIssue).Count

                    '如果在Ads表中没查询到数据，则把[BEGIN_ADS]...[END_ADS]删除
                    If (adsCount = 0) Then  '没查到Ads表中有数据，则删除[BEGIN_ADS]块
                        While (SpreadTemplate.Contains("[BEGIN_LITTLE_ADS]"))
                            Dim endLen As Integer = "[END_LITTLE_ADS]".Length
                            Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_LITTLE_ADS]")
                            Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_LITTLE_ADS]")
                            SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                        End While
                    Else
                        Dim icounter As Integer = 0
                        While (SpreadTemplate.Contains("[BEGIN_LITTLE_ADS]"))
                            Dim beginLen As Integer = "[BEGIN_LITTLE_ADS]".Length
                            Dim endLen As Integer = "[END_LITTLE_ADS]".Length
                            Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_LITTLE_ADS]")
                            Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_LITTLE_ADS]")
                            Dim oldPromotion As String = SpreadTemplate.Substring(startProIndex, endProIndex - startProIndex + endLen)
                            Dim adsTemplate As String = SpreadTemplate.Substring(startProIndex + beginLen, endProIndex - startProIndex - beginLen)
                            Dim newPromotion As String = InsertAds(adsTemplate, issueId, siteId, icounter, "B")
                            'SpreadTemplate = SpreadTemplate.Replace(oldPromotion, newPromotion)
                            SpreadTemplate = SpreadTemplate.Remove(startProIndex, endProIndex - startProIndex + endLen)
                            SpreadTemplate = SpreadTemplate.Insert(startProIndex, newPromotion)
                            icounter += 1
                            If (icounter = adsCount) Then
                                While (SpreadTemplate.Contains("[BEGIN_LITTLE_ADS]"))
                                    Dim endLen1 As Integer = "[END_LITTLE_ADS]".Length
                                    Dim startProIndex1 As Integer = SpreadTemplate.IndexOf("[BEGIN_LITTLE_ADS]")
                                    Dim endProIndex1 As Integer = SpreadTemplate.IndexOf("[END_LITTLE_ADS]")
                                    SpreadTemplate = SpreadTemplate.Remove(startProIndex1, endProIndex1 - startProIndex1 + endLen1)
                                End While
                                Exit While
                            End If
                        End While
                    End If
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try

            Try
                'Promotion大分类数据填充,使用Ads表中Type="P"的产品
                If (SpreadTemplate.Contains("[BEGIN_PROMOTION_PRODUCT]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_PROMOTION_PRODUCT]", "[END_PROMOTION_PRODUCT]", SpreadTemplate, "PO")
                End If
            Catch ex2 As Exception
                'LogText(ex2.ToString())
                Throw New Exception(ex2.ToString()) 'throw the exception to the upper layer
            End Try

            Try
                'NewArrival大分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_NEW_ARRIVALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_NEW_ARRIVALS]", "[END_NEW_ARRIVALS]", SpreadTemplate, "NE")
                End If
            Catch ex3 As Exception
                'LogText(ex3.ToString())
                Throw New Exception(ex3.ToString()) 'throw the exception to the upper layer
            End Try

            '2014/02/11 added,Weekly Deals
            Try
                'Weekly Deals大分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_WEEKLY_DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_WEEKLY_DEALS]", "[END_WEEKLY_DEALS]", SpreadTemplate, "WE")
                End If
            Catch ex3 As Exception
                'LogText(ex3.ToString())
                Throw New Exception(ex3.ToString()) 'throw the exception to the upper layer
            End Try

            Try
                'Daily Deals大分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_DAILY_DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_DAILY_DEALS]", "[END_DAILY_DEALS]", SpreadTemplate, "DA")
                End If
            Catch ex4 As Exception
                'LogText(ex4.ToString())
                Throw New Exception(ex4.ToString())
            End Try

            Try
                '2013/07/04 added,Group Deals分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_GROUPDEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_GROUPDEALS]", "[END_GROUPDEALS]", SpreadTemplate, "GD")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            Try
                '2013/07/04 added,Price Deals分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_PRICE_DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_PRICE_DEALS]", "[END_PRICE_DEALS]", SpreadTemplate, "PD")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            Try
                '2013/07/04 added,Discount Center分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_DISCOUNT]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_DISCOUNT]", "[END_DISCOUNT]", SpreadTemplate, "DI")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try

            'Deals标签说明，begin
            Try
                '2013/07/10 added,1.99 DEALS
                If (SpreadTemplate.Contains("[BEGIN_1.99DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_1.99DEALS]", "[END_1.99DEALS]", SpreadTemplate, "PD1")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            Try
                '2013/07/10 added,2.99 DEALS
                If (SpreadTemplate.Contains("[BEGIN_2.99DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_2.99DEALS]", "[END_2.99DEALS]", SpreadTemplate, "PD2")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            Try
                '2013/07/10 added,3.99 DEALS
                If (SpreadTemplate.Contains("[BEGIN_3.99DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_3.99DEALS]", "[END_3.99DEALS]", SpreadTemplate, "PD3")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            Try
                '2013/07/10 added,4.99 DEALS
                If (SpreadTemplate.Contains("[BEGIN_4.99DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_4.99DEALS]", "[END_4.99DEALS]", SpreadTemplate, "PD4")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString())
            End Try
            'Deals标签说明，end

            Try
                '大Category分类数据填充
                If (SpreadTemplate.Contains("[BEGIN_CATEGORIES]")) Then
                    Dim categoryLen = "[BEGIN_CATEGORIES]".Length
                    Dim endCategoryLen = "[END_CATEGORIES]".Length
                    Dim cateStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_CATEGORIES]")
                    Dim cateEndIndex As Integer = SpreadTemplate.IndexOf("[END_CATEGORIES]")
                    Dim oldCategory As String = SpreadTemplate.Substring(cateStartIndex, cateEndIndex - cateStartIndex + endCategoryLen)
                    Dim categoryContent As String = SpreadTemplate.Substring(cateStartIndex + categoryLen, cateEndIndex - cateStartIndex - categoryLen)
                    Dim childStartIndex As Integer = categoryContent.IndexOf("[BEGIN_CATEGORY_")
                    Dim childEndIndex As Integer = categoryContent.IndexOf("[END_CATEGORY_")
                    Dim childCategory As String = categoryContent.Substring(childStartIndex, childEndIndex - childStartIndex)
                    While (categoryContent.Contains("[BEGIN_CATEGORY_") AndAlso childCategory.Contains("[BEGIN_CATEGORY_"))
                        Dim beginIndex As Integer = childCategory.IndexOf("[")
                        Dim endIndex As Integer = childCategory.IndexOf("]")
                        Dim labelName As String = childCategory.Substring(beginIndex + 1, endIndex - beginIndex - 1)
                        'BEGIN_CATEGORY_WOMEN'S_CLOTHING
                        Dim splitStr = System.Text.RegularExpressions.Regex.Split(labelName, "BEGIN_CATEGORY_", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        Dim searchStr = splitStr(1).Replace("_", " ") 'splitStr(1)="WOMEN'S_CLOTHING"
                        Dim beginStr = "[BEGIN_CATEGORY_" + splitStr(1) + "]"
                        Dim endStr = "[END_CATEGORY_" + splitStr(1) + "]"
                        categoryContent = FillRootCategory(siteId, issueId, beginStr, endStr, categoryContent, searchStr)
                        If (categoryContent.Contains("[BEGIN_CATEGORY_")) Then
                            childStartIndex = categoryContent.IndexOf("[BEGIN_CATEGORY_")
                            childEndIndex = categoryContent.IndexOf("[END_CATEGORY_")
                            childCategory = categoryContent.Substring(childStartIndex, childEndIndex - childStartIndex)
                        End If
                    End While
                    SpreadTemplate = SpreadTemplate.Replace(oldCategory, categoryContent)
                End If
            Catch ex5 As Exception
                'LogText(ex5.ToString())
                Common.LogText(ex5.InnerException.Message.ToString())
                Throw New Exception(ex5.ToString()) 'throw the exception to the upper layer
            End Try

            '2013/4/12新增，fever-print
            Try
                'Best Arrival(精品推荐)
                If (SpreadTemplate.Contains("[BEGIN_BEST_ARRIVAL]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_BEST_ARRIVAL]", "[END_BEST_ARRIVAL]", SpreadTemplate, "DA")
                End If
            Catch ex6 As Exception
                Common.LogText(ex6.ToString())
                Throw New Exception(ex6.ToString()) 'throw the exception to the upper layer
            End Try

            'taobaoCategory-2013/3/26
            Try
                If (SpreadTemplate.Contains("[BEGIN_BESTSELLER]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_BESTSELLER]", "[END_BESTSELLER]", SpreadTemplate, "BE")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try
            Try
                If (SpreadTemplate.Contains("[BEGIN_DEALS]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_DEALS]", "[END_DEALS]", SpreadTemplate, "DE")
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try
            Try
                If (SpreadTemplate.Contains("[BEGIN_HOT_KEEP]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_HOT_KEEP]", "[END_HOT_KEEP]", SpreadTemplate, "HK")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try
            Try
                If (SpreadTemplate.Contains("[BEGIN_MOST_POPULAR]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_MOST_POPULAR]", "[END_MOST_POPULAR]", SpreadTemplate, "MP")
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try
            Try
                If (SpreadTemplate.Contains("[BEGIN_PROMOTION_PRODUCT]")) Then
                    SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_PROMOTION_PRODUCT]", "[END_PROMOTION_PRODUCT]", SpreadTemplate, "PO")
                End If
            Catch ex As Exception
                'LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try

            '2013/06/17 added,intelligentize the template,begin
            Try
                If (SpreadTemplate.Contains("[BEGIN_CATEGORIES_P]")) Then
                    Dim tCategoryName As String = "" '存放TemplateInfo表中的StoredCates
                    Dim querySortedCate As String = efContext.TemplateInfoes.Where(Function(t) t.SiteId = siteId AndAlso t.PlanType = planType).SingleOrDefault.SortedCates
                    Dim splitCategory() As String = querySortedCate.Split(",")
                    For i As Integer = 0 To splitCategory.Count - 1
                        If Not (String.IsNullOrEmpty(splitCategory(i))) Then
                            tCategoryName = tCategoryName & "," & splitCategory(i).Split(":")(1)
                        End If
                    Next
                    tCategoryName = tCategoryName.Remove(0, 1)

                    Dim categoryLen = "[BEGIN_CATEGORIES_P]".Length
                    Dim endCategoryLen = "[END_CATEGORIES_P]".Length
                    Dim cateStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_CATEGORIES_P]")
                    Dim cateEndIndex As Integer = SpreadTemplate.IndexOf("[END_CATEGORIES_P]")
                    Dim oldCategory As String = SpreadTemplate.Substring(cateStartIndex, cateEndIndex - cateStartIndex + endCategoryLen)
                    Dim categoryContent As String = SpreadTemplate.Substring(cateStartIndex + categoryLen, cateEndIndex - cateStartIndex - categoryLen)
                    Dim fillCategories() As String = tCategoryName.Split(",")
                    For i As Integer = 0 To fillCategories.Count - 1 '遍历分割后的categoryId，如：1  2  3  4  5
                        If categoryContent.Contains("[BEGIN_CATEGORY_") Then
                            Dim searchStr As String = fillCategories(i).Trim
                            Dim beginStr As String = "[BEGIN_CATEGORY_" + (i + 1).ToString() + "]" '如：[BEGIN_CATEGORY_1]
                            Dim endStr As String = "[END_CATEGORY_" + (i + 1).ToString() + "]" '如：[END_CATEGORY_1]
                            categoryContent = FillRootCategory(siteId, issueId, beginStr, endStr, categoryContent, searchStr)
                        End If
                    Next
                    SpreadTemplate = SpreadTemplate.Replace(oldCategory, categoryContent)
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString())
                Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
            End Try
            '2013/06/17 added,intelligentize the template,end

            '2014/01/10 added,自动排序的Category,begin
            '支持关联销售
            If (SpreadTemplate.Contains("[BEGIN_SORT_CATEGORY]")) Then

            End If
            '2014/01/10 added,自动排序的Category,end
        Catch ex As Exception
            Common.LogText(ex.ToString())
            Throw New Exception(ex.ToString())
        End Try
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 使用排序号的产品信息填充电脑版模板
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="listProductElements"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FillSpreadTemplate(ByVal SpreadTemplate As String, ByVal listProductElements As List(Of ProductElement), _
                                        ByVal siteId As Integer, ByVal planType As String, ByVal senderName As String, _
                                        ByVal loginEmail As String, ByVal appId As String) As String
        If (SpreadTemplate.Contains("[BEGIN_GROUPER_NEW_ARRIVAL]")) Then
            Dim beginStr As String = "[BEGIN_GROUPER_NEW_ARRIVAL]"
            Dim endStr As String = "[END_GROUPER_NEW_ARRIVAL]"
            Dim iGrouperLen As Integer = beginStr.Length
            Dim iGrouperEndLen As Integer = endStr.Length
            Dim iGrouperStartIndex As Integer = SpreadTemplate.IndexOf(beginStr)
            Dim iGrouperEndIndex As Integer = SpreadTemplate.IndexOf(endStr)
            Dim grouperBlockTemplate As String = SpreadTemplate.Substring(iGrouperStartIndex, iGrouperEndIndex - iGrouperStartIndex + iGrouperEndLen)
            Dim grouperBlockContent As String = SpreadTemplate.Substring(iGrouperStartIndex + iGrouperLen, iGrouperEndIndex - iGrouperStartIndex - iGrouperLen)

            Dim beginStr_2 As String = "[BEGIN_PRODUCT]"
            Dim endStr_2 As String = "[END_PRODUCT]"
            Dim iProductLen As Integer = beginStr_2.Length
            Dim iProductEndLen As Integer = endStr_2.Length
            Dim iProductStartIndex As Integer = grouperBlockContent.IndexOf(beginStr_2)
            Dim iProductEndIndex As Integer = grouperBlockContent.IndexOf(endStr_2)
            Dim productBlockTemplate As String = grouperBlockContent.Substring(iProductStartIndex, iProductEndIndex - iProductStartIndex + iProductEndLen)
            Dim productBlockContent As String = grouperBlockContent.Substring(iProductStartIndex + iProductLen, iProductEndIndex - iProductStartIndex - iProductLen)

            Dim flagGroupBuyer As Boolean = False
            Dim flagBeeCrazy As Boolean = False
            Dim productGroupBuyer As New ProductElement
            Dim productBeeCrazy As New ProductElement

            '2013/09/06 修改
            'Dim arrProductElements() As ProductElement = listProductElements.ToArray()
            ''arrProductElements = arrProductElements.Reverse.ToArray()
            'listProductElements = arrProductElements.ToList()
            listProductElements = listProductElements.OrderByDescending(Function(o) o.pubDate).ToList()

            '2013/09/06 修改
            '修改原因：Since more and more business promotions on Lifein.HK with monthly charge or CPA charge, the top 2 position of smart email always needed to be revised manually
            '修改前：
            'For i As Integer = 0 To listProductElements.Count - 1 Step 1 '移除第一个GroupBuyer产品
            '    If listProductElements(i).url.Contains("http://www.groupbuyer.com.hk") Then
            '        productGroupBuyer = listProductElements(i)
            '        flagGroupBuyer = True
            '        listProductElements.Remove(listProductElements(i))
            '        Exit For
            '    End If
            'Next
            'For j As Integer = 0 To listProductElements.Count - 1 Step 1 '移除第一个BeeCrazy产品
            '    If listProductElements(j).url.Contains("http://www.beecrazy.hk") Then
            '        productBeeCrazy = listProductElements(j)
            '        flagBeeCrazy = True
            '        listProductElements.Remove(listProductElements(j))
            '        Exit For
            '    End If
            'Next
            '修改后
            productGroupBuyer = listProductElements(0)
            listProductElements.Remove(listProductElements(0))
            flagGroupBuyer = True
            productBeeCrazy = listProductElements(1)
            listProductElements.Remove(listProductElements(1))
            flagBeeCrazy = True
            'Over----------------------------------------------------------------------------------------------

            If Not flagGroupBuyer Then '第一行的GroupBuyer产品未找到，使用第一个产品填充
                productGroupBuyer = listProductElements(0)
                flagGroupBuyer = True
            ElseIf Not flagBeeCrazy Then '第一行的GroupBuyer产品找到,BeeCrazy产品未找到，使用第一个产品填充BeeCrazy
                productBeeCrazy = listProductElements(0)
                flagBeeCrazy = True
            End If
            If Not flagBeeCrazy Then '第一行的GroupBuyer产品找到,BeeCrazy产品未找到，使用第二个产品填充BeeCrazy
                productBeeCrazy = listProductElements(1)
            End If
            Dim mySection1 As StringBuilder = New StringBuilder(productBlockContent) '产品块的内容
            mySection1.Replace("[URL]", productGroupBuyer.url)
            mySection1.Replace("[PICTURE_SRC]", productGroupBuyer.pictureUrl)
            mySection1.Replace("[PRODUCT_NAME]", productGroupBuyer.productName)
            '原价,2013/08/29 added
            If Not (IsDBNull(productGroupBuyer.price)) Then
                mySection1.Replace("[PRODUCT_PRICE]", productGroupBuyer.currency & DoFormat(productGroupBuyer.price))
            Else
                mySection1.Replace("[PRODUCT_PRICE]", "")
            End If
            'mySection1.Replace("[PRODUCT_PRICE]", productGroupBuyer.currency & DoFormat(productGroupBuyer.price))
            '-------------------------------------------------------------------------------------------------------
            '2013/07/16 added，折扣价前面不用加上HKD
            'mySection1.Replace("[PRODUCT_MONEY]", productGroupBuyer.currency & DoFormat(productGroupBuyer.discountPrice))
            'mySection1.Replace("[PRODUCT_MONEY]", DoFormat(productGroupBuyer.discountPrice)) '2013/08/29 added
            If Not (IsDBNull(productGroupBuyer.price)) Then
                mySection1.Replace("[PRODUCT_MONEY]", DoFormat(productGroupBuyer.discountPrice))
            Else
                mySection1.Replace("[PRODUCT_MONEY]", "")
            End If
            '----------------------------------------------------------------------------------------------------------
            mySection1.Replace("[CATEGORY_ID]", productGroupBuyer.category)
            Dim mySection2 As StringBuilder = New StringBuilder(productBlockContent)
            mySection2.Replace("[URL]", productBeeCrazy.url)
            mySection2.Replace("[PICTURE_SRC]", productBeeCrazy.pictureUrl)
            mySection2.Replace("[PRODUCT_NAME]", productBeeCrazy.productName)
            '原价
            mySection2.Replace("[PRODUCT_PRICE]", productBeeCrazy.currency & DoFormat(productBeeCrazy.price))
            '2013/07/16 added，折扣价前面不用加上HKD
            'mySection2.Replace("[PRODUCT_MONEY]", productBeeCrazy.currency & DoFormat(productBeeCrazy.discountPrice))
            mySection2.Replace("[PRODUCT_MONEY]", DoFormat(productBeeCrazy.discountPrice))
            mySection2.Replace("[CATEGORY_ID]", productBeeCrazy.category)

            Dim addString As String = "<table align=""left"" cellpadding=""0"" cellspacing=""0"" order=""0"" class=""mobile_hidden""><tbody><tr><td style=""height:1px;""><img alt="""" style=""display: block; margin:0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""16"" height=""1"" /></td></tr></tbody></table>"
            'Dim addString As String = "<td style=""width: 16px;""> </td>" 'desktop版本

            productBlockContent = mySection1.ToString() & addString & mySection2.ToString()
            grouperBlockContent = grouperBlockContent.Replace(productBlockTemplate, productBlockContent)
            SpreadTemplate = SpreadTemplate.Replace(grouperBlockTemplate, grouperBlockContent)
        End If
        If SpreadTemplate.Contains("[BEGIN_GROUPER_OTHER]") Then

            '获取点击量最大的两个产品
            'listProductElements = GetTopNProducts(listProductElements, loginEmail, appId, siteId, senderName, planType, 2)

            Dim beginStr As String = "[BEGIN_GROUPER_OTHER]"
            Dim endStr As String = "[END_GROUPER_OTHER]"
            Dim iGrouperLen As Integer = beginStr.Length
            Dim iGrouperEndLen As Integer = endStr.Length
            Dim iGrouperStartIndex As Integer = SpreadTemplate.IndexOf(beginStr)
            Dim iGrouperEndIndex As Integer = SpreadTemplate.IndexOf(endStr)
            Dim grouperBlockTemplate As String = SpreadTemplate.Substring(iGrouperStartIndex, iGrouperEndIndex - iGrouperStartIndex + iGrouperEndLen)
            Dim grouperBlockContent As String = SpreadTemplate.Substring(iGrouperStartIndex + iGrouperLen, iGrouperEndIndex - iGrouperStartIndex - iGrouperLen)

            Dim beginStr_2 As String = "[BEGIN_PRODUCT]"
            Dim endStr_2 As String = "[END_PRODUCT]"
            Dim iProductLen As Integer = beginStr_2.Length
            Dim iProductEndLen As Integer = endStr_2.Length
            Dim iProductStartIndex As Integer = grouperBlockContent.IndexOf(beginStr_2)
            Dim iProductEndIndex As Integer = grouperBlockContent.IndexOf(endStr_2)
            Dim productBlockTemplate As String = grouperBlockContent.Substring(iProductStartIndex, iProductEndIndex - iProductStartIndex + iProductEndLen)
            Dim productBlockContent As String = grouperBlockContent.Substring(iProductStartIndex + iProductLen, iProductEndIndex - iProductStartIndex - iProductLen)

            Dim productSection As String = ""
            Dim productCounter As Integer = 1
            For Each li As ProductElement In listProductElements
                Dim mySection As StringBuilder = New StringBuilder(productBlockContent)
                mySection.Replace("[URL]", li.url)
                mySection.Replace("[PICTURE_SRC]", li.pictureUrl)
                mySection.Replace("[PRODUCT_NAME]", li.productName)
                '原价
                mySection.Replace("[PRODUCT_PRICE]", li.currency & DoFormat(li.price))
                '2013/07/16 added，折扣价前面不用加上HKD
                'mySection.Replace("[PRODUCT_MONEY]", li.currency & DoFormat(li.discountPrice))
                mySection.Replace("[PRODUCT_MONEY]", DoFormat(li.discountPrice))
                mySection.Replace("[CATEGORY_ID]", li.category)

                If Not (productCounter Mod 3 = 0) AndAlso Not (productCounter = listProductElements.Count) Then
                    '每一行的第一个、第二个产品填充，而非最后产品
                    Dim addString As String = "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""16"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""16"" height=""1"" /></td></tr></tbody></table>"

                    'Dim addString As String = "<td style=""width: 16px;""> </td>"  'desktop版本

                    productSection = productSection & mySection.ToString() & addString
                ElseIf (productCounter Mod 3 = 0) AndAlso Not (productCounter = listProductElements.Count) Then
                    '每一行的第三个产品填充，而非最后一个产品
                    Dim addString As String = "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""16"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""16"" height=""1"" /></td></tr></tbody></table>" & _
                        "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 5px;""><img alt="""" style=""display: block; margin: 0px;"" id=""Img4"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""5"" /></td></tr></tbody></table>" & _
                        "<table align=""left"" width=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""mobile_hidden""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""Img1"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""25"" height=""1"" /></td></tr></tbody></table>"

                    'desktop版本
                    'Dim addString As String = "<td style=""width: 25px;""> </td></tr></tbody></table></td></tr><tr><td style=""background-color: #ffffff; width: 700px; height: 15px;""> </td></tr><tr><td><table width=""700"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td style=""width: 25px;""> </td>"

                    productSection = productSection & mySection.ToString() & addString
                ElseIf productCounter = listProductElements.Count Then
                    '最后一个产品填充
                    productSection = productSection & mySection.ToString
                    If Not (listProductElements.Count Mod 3 = 0) Then '（产品的数量-2)不是3的整数倍
                        For i As Integer = 1 To 3 - (listProductElements.Count Mod 3) Step 1

                            'Dim addString As String = "<td style=""width: 16px;""> </td>" 'desktop版本
                            'Dim moreStr As String = "<td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td>" 'desktop版本

                            Dim addString As String = "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""16"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""16"" height=""1"" /></td></tr></tbody></table>"
                            Dim moreStr As String = "<table width=""204"" align=""left"" border=""0"" cellspacing=""0"" valign=""top"" class=""mobile_hidden"" ><tbody><tr><td   style=""height: 5px;""><img alt="""" style=""display: block; margin-top: 0px; margin-bottom: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""5"" /></td></tr><tr><td><table align=""left"" width=""204"" border=""0"" cellspacing=""0"" cellpadding=""5""   style=""border: 1px solid #25a9e0; border-collapse: collapse;"" valign=""top""><tbody><tr><td><table width=""194"" border=""0"" cellspacing=""0"" cellpadding=""0"" >" & _
                            "<tbody><tr><td align=""left""><a href=""http://www.lifein.hk/Deals/TopDeals.aspx"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""225"" border=""0"" alt="""" /></a></td></tr><tr><td   style=""height: 5px;""></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td   style=""height: 5px;""><img alt="""" style=""display: block; margin-top: 0px; margin-bottom: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""5"" /></td></tr></tbody></table>"
                            productSection = productSection & addString & moreStr
                        Next
                    End If
                End If
                productCounter = productCounter + 1

            Next
            grouperBlockContent = grouperBlockContent.Replace(productBlockTemplate, productSection)
            SpreadTemplate = SpreadTemplate.Replace(grouperBlockTemplate, grouperBlockContent)
        End If
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 读取Ads表中Type="P"的一条记录，并填充相应的Promotion模板
    ''' </summary>
    ''' <param name="promotion"></param>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <remarks></remarks>
    Private Function InsertPromotionAndAd(ByVal promotion As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal counter As Integer) As String
        Dim queryPromotion = (From ad In efContext.AutoAds _
                             Join issue In efContext.AutoAds_Issue _
                             On ad.AdID Equals issue.AdId
                             Order By issue.IssueID Descending
                             Where issue.IssueID = issueId And siteId = siteId And ad.Type = "P"
                             Select ad).Skip(counter).Take(1)
        Dim AutoAd As New AutoAd ' = queryPromotion.Single() '2013/3/28修改
        For Each p As AutoAd In queryPromotion
            If (promotion.Contains("[URL]")) Then
                promotion = promotion.Replace("[URL]", p.Url)
            End If
            If (promotion.Contains("[PICTURE_SRC]")) Then
                promotion = promotion.Replace("[PICTURE_SRC]", p.PictureUrl)
            End If
            If (promotion.Contains("[PICTURE_ALT]")) Then
                promotion = promotion.Replace("[PICTURE_ALT]", p.Description)
            End If

            '2013/4/26添加，为Promotion添加link_category
            'If (promotion.Contains("[CATEGORY_ID]")) Then
            '    promotion = promotion.Replace("[CATEGORY_ID]", p.Categories.First.CategoryID.ToString())
            'End If
            If Not (p.Categories.FirstOrDefault Is Nothing) Then
                promotion = promotion.Replace("[CATEGORY_ID]", p.Categories.First.CategoryID.ToString())
                promotion = promotion.Replace("[CATEGORY_NAME]", p.Categories.FirstOrDefault.Category1.ToString())
            End If
            promotion = promotion.Replace("[PRODUCT_PRICE]", String.Format("{0:0.00}", p.Price)) 'String.Format("{0:0.00}", q.Price)
            promotion = promotion.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", p.Discount))
        Next
        Return promotion
    End Function

    ''' <summary>
    ''' 2013/07/04 added,读取Ads表中Type不为P的一条记录，并填充相应的Ads模板
    ''' </summary>
    ''' <param name="adsTemplate"></param>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="counter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertAds(ByVal adsTemplate As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal counter As Integer, ByVal type As String) As String
        If (type = "B") Then
            Dim queryadsTemplate = (From ad In efContext.AutoAds _
                             Join issue In efContext.AutoAds_Issue _
                             On ad.AdID Equals issue.AdId
                             Order By issue.IssueID Descending
                             Where issue.IssueID = issueId And siteId = siteId And ad.Type = "B"
                             Select ad).Skip(counter).Take(1)
            Dim AutoAd As New AutoAd ' = queryadsTemplate.Single() '2013/3/28修改
            For Each p As AutoAd In queryadsTemplate
                adsTemplate = adsTemplate.Replace("[URL]", p.Url)
                adsTemplate = adsTemplate.Replace("[PICTURE_SRC]", p.PictureUrl)
                adsTemplate = adsTemplate.Replace("[PICTURE_ALT]", p.Description)
                adsTemplate = adsTemplate.Replace("[PRODUCT_NAME]", p.Ad1)


                '2013/4/26添加，为adsTemplate添加link_category
                If Not (p.Categories.FirstOrDefault Is Nothing) Then
                    adsTemplate = adsTemplate.Replace("[CATEGORY_ID]", p.Categories.First.CategoryID.ToString())
                    adsTemplate = adsTemplate.Replace("[CATEGORY_NAME]", p.Categories.FirstOrDefault.Category1.ToString())
                End If


                adsTemplate = adsTemplate.Replace("[PRODUCT_PRICE]", String.Format("{0:0.00}", p.Price)) 'String.Format("{0:0.00}", q.Price)
                adsTemplate = adsTemplate.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", p.Discount))
            Next
        Else

            Dim queryadsTemplate = (From ad In efContext.AutoAds _
                             Join issue In efContext.AutoAds_Issue _
                             On ad.AdID Equals issue.AdId
                             Order By issue.IssueID Descending
                             Where issue.IssueID = issueId And siteId = siteId And (String.IsNullOrEmpty(ad.Type) OrElse Not (ad.Type = "P" Or ad.Type = "B"))
                             Select ad).Skip(counter).Take(1)
            Dim AutoAd As New AutoAd ' = queryadsTemplate.Single() '2013/3/28修改
            For Each p As AutoAd In queryadsTemplate
                'adsTemplate = adsTemplate.Replace("[URL]", p.Url)
                If (siteId <> 68) Then 'siteid= 68 为focalDE站点，因de站点在添加GA代码loadhtml时导致德语乱码，所以用此特殊方法添加追踪代码
                    adsTemplate = adsTemplate.Replace("[URL]", p.Url)
                Else
                    Dim firstIndex As Integer = 0 ' href.LastIndexOf("/", href.LastIndexOf("/") - 1) + 1
                    Dim lastIndex As Integer = 0 ' href.LastIndexOf("/")
                    Dim skuName As String = "" ' href.Substring(firstIndex, lastIndex - firstIndex)
                    Try
                        firstIndex = p.Url.IndexOf("/")
                        lastIndex = p.Url.LastIndexOf("/")
                        skuName = p.Url.Substring(firstIndex, lastIndex - firstIndex - 1)
                    Catch ex As Exception
                        '待确定处理
                    End Try
                    adsTemplate = adsTemplate.Replace("[URL]", p.Url & "?utm_source=RSpread&utm_media=EM&utm_campaign=DM_[WEEK]DE_" & skuName & "&source=rspread")
                End If
                adsTemplate = adsTemplate.Replace("[PICTURE_SRC]", p.PictureUrl)
                adsTemplate = adsTemplate.Replace("[PICTURE_ALT]", p.Description)
                adsTemplate = adsTemplate.Replace("[PRODUCT_NAME]", p.Ad1)


                '2013/4/26添加，为adsTemplate添加link_category
                If Not (p.Categories.FirstOrDefault Is Nothing) Then
                    adsTemplate = adsTemplate.Replace("[CATEGORY_ID]", p.Categories.First.CategoryID.ToString())
                    adsTemplate = adsTemplate.Replace("[CATEGORY_NAME]", p.Categories.FirstOrDefault.Category1.ToString())
                End If


                adsTemplate = adsTemplate.Replace("[PRODUCT_PRICE]", String.Format("{0:0.00}", p.Price)) 'String.Format("{0:0.00}", q.Price)
                adsTemplate = adsTemplate.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", p.Discount))
            Next
        End If

        Return adsTemplate
    End Function

    ''' <summary>
    ''' 处理含有产品的大分类，返回填充好产品的模板
    ''' </summary>
    ''' <param name="beginStr">大分类的开始标签，[NEW_...]</param>
    ''' <param name="endStr">大分类的结束标签，[END_...]</param>
    ''' <param name="SpreadTemplate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FillRootCategory(ByVal siteId As Integer, ByVal issueId As Integer, ByVal beginStr As String, ByVal endStr As String, ByVal SpreadTemplate As String, ByVal searchStr As String) As String
        Dim strLenBegin As Integer = beginStr.Length
        Dim strLenEnd As Integer = endStr.Length '2013/06/26 added
        Dim startIndex As Integer = SpreadTemplate.IndexOf(beginStr)
        Dim endIndex As Integer = SpreadTemplate.IndexOf(endStr)
        Dim oldAutoTemplatestr As String = SpreadTemplate.Substring(startIndex, endIndex - startIndex + strLenEnd)
        Dim AutoTemplatestr As String = SpreadTemplate.Substring(startIndex + strLenBegin, endIndex - startIndex - strLenBegin)
        Dim newAutoTemplatestr As String = InsertProducts(AutoTemplatestr, issueId, siteId, beginStr, searchStr)
        SpreadTemplate = SpreadTemplate.Replace(oldAutoTemplatestr, newAutoTemplatestr)
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 在大分类中填充产品信息
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertProducts(ByVal oldContent As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal beginStr As String, ByVal searchStr As String) As String
        Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
        'NewArrival分类产品填充
        If (beginStr.Contains("[BEGIN_NEW_ARRIVALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "NE", lengthProduct, siteId)

        ElseIf (beginStr.Contains("[BEGIN_DAILY_DEALS]")) Then   'Daily Deals分类产品填充
            oldContent = InsertProductOfOthers(oldContent, issueId, "DA", lengthProduct, siteId)

        ElseIf (beginStr.Contains("[BEGIN_WEEKLY_DEALS]")) Then '2014/02/11 added, Weekly Deals分类产品填充
            oldContent = InsertProductOfOthers(oldContent, issueId, "WE", lengthProduct, siteId)

        ElseIf (beginStr.Contains("[BEGIN_CATEGORY_")) Then '2013/3/21新加，Category的产品
            oldContent = InsertProductOfCategory(oldContent, searchStr, lengthProduct, siteId, issueId)

        ElseIf (beginStr.Contains("[BEGIN_PROMOTION_PRODUCT]")) Then '2013/4/15新加，Promotion使用产品填充
            oldContent = InsertProductOfOthers(oldContent, issueId, "PO", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_BEST_ARRIVAL]")) Then '2013/4/15新加，精品推荐(New Arrival)
            oldContent = InsertProductOfOthers(oldContent, issueId, "BR", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_BEST_SELLER]")) Then '2013/4/15新加，热卖商品(Hot Arrival)
            oldContent = InsertProductOfOthers(oldContent, issueId, "BE", lengthProduct, siteId)
            '2013/07/04,EverBuying added,begin
        ElseIf (beginStr.Contains("[BEGIN_GROUPDEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "GD", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_PRICE_DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PD", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_DISCOUNT]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "DI", lengthProduct, siteId)
            '2013/07/04,EverBuying added,end
            'taobaoProduct-2013/4/15
        ElseIf (beginStr.Contains("[BEGIN_BESTSELLER]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "BE", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "DE", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_HOT_KEEP]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "HK", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_MOST_POPULAR]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "MP", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_PROMOTION_PRODUCT]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PO", lengthProduct, siteId)

            '2013/07/10 added,Deals,begin
        ElseIf (beginStr.Contains("[BEGIN_1.99DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PD1", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_2.99DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PD2", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_3.99DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PD3", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_4.99DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PD4", lengthProduct, siteId)
            '2013/07/10 added,Deals,end
        End If





        Return oldContent
    End Function

    ''' <summary>
    ''' 所有Section="Category"（Sections表）的产品填充
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="searchString">大类的名称(AutoCategories表)</param>
    ''' <param name="lengthProduct">"[BEGIN_PRODUCT]"字符串的长度</param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertProductOfCategory(ByVal oldContent As String, ByVal searchString As String, ByVal lengthProduct As Integer, ByVal siteId As Integer, ByVal issueId As Integer) As String
        Dim queryProduct = From p In efContext.AutoProducts
                         Join pi In efContext.AutoProducts_Issue
                         On p.ProdouctID Equals pi.ProductId
                         Where pi.SectionID = "CA" AndAlso pi.SiteId = siteId AndAlso pi.IssueID = issueId
                         Select p

        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")
        Dim lenProductEnd As Integer = "[END_PRODUCT]".Length '2013/06/26 added

        For Each q As AutoProduct In queryProduct

            Dim categoryName = q.Categories.ToList() '2013/4/15
            Dim cName As New HashSet(Of String)
            For Each c As AutoCategory In categoryName
                cName.Add(c.Category1.ToUpper)
                Common.LogText(q.ProdouctID & "  " & q.Url & "  " & c.Category1 & "  " & c.CategoryID)
            Next

            'If (productStartIndex > -1 And q.Categories.First.Category1.ToUpper = searchString) Then
            If (productStartIndex > -1 And cName.Contains(searchString.ToUpper)) Then
                'Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + length)
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                Common.LogText(q.ProdouctID & "  searchString:" & searchString)
                '2013/3/25 新加
                If (newProduct.Contains("[URL]")) Then
                    'newProduct = newProduct.Replace("[URL]", q.Url)
                    'siteid= 68 为focalDE站点，因de站点在添加GA代码loadhtml时导致德语乱码，所以用此特殊方法添加追踪代码,以下为追踪代码规则
                    'utm_source=RSpread&utm_media=EM&utm_campaign=DM_14DE_MF0284B&source=rspread，其中参数utm_source/utm_media/ source每期固定不变，
                    '参数 utm_campaign的变化规则如下：utm_campaign=DM_[星期数]DE_[产品SKU/默认]，高亮部分为每期变化，[星期数]为一年中的第几个星期，[产品SKU/默认]指链接为产品链接时，
                    '则添加产品SKU, 非产品链接时, 则添加默认添加de.focalprice.com
                    If (siteId <> 68) Then 'siteid= 68 为focalDE站点，因de站点在添加GA代码loadhtml时导致德语乱码，所以用此特殊方法添加追踪代码
                        newProduct = newProduct.Replace("[URL]", q.Url)
                    Else
                        Dim firstIndex As Integer = 0 ' href.LastIndexOf("/", href.LastIndexOf("/") - 1) + 1
                        Dim lastIndex As Integer = 0 ' href.LastIndexOf("/")
                        Dim skuName As String = "" ' href.Substring(firstIndex, lastIndex - firstIndex)
                        Try
                            firstIndex = q.Url.LastIndexOf("/", q.Url.LastIndexOf("/") - 1) + 1
                            lastIndex = q.Url.LastIndexOf("/")
                            skuName = q.Url.Substring(firstIndex, lastIndex - firstIndex)
                        Catch ex As Exception
                            '待确定处理
                        End Try
                        newProduct = newProduct.Replace("[URL]", q.Url & "?utm_source=RSpread&utm_media=EM&utm_campaign=DM_[WEEK]DE_" & skuName & "&source=rspread")
                    End If
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
                '2013/05/21,add free shiping pictures and ships pictures in template,end

                oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                oldContent = oldContent.Insert(productStartIndex, newProduct)
                If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增
                    productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                    productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
                Else
                    Exit For
                End If
            End If
        Next
        While (oldContent.Contains("[BEGIN_PRODUCT]"))
            oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            End If
        End While
        Return oldContent
    End Function

    ''' <summary>
    ''' 返回所有Section != "Category"（Sections表）的产品填充的模板
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="lengthProduct"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertProductOfOthers(ByVal oldContent As String, ByVal issueId As Integer, ByVal sectionId As String, ByVal lengthProduct As Integer, ByVal siteId As Integer)
        'LogText(sectionId)
        Dim queryNewArrival As New List(Of AutoProduct)
        queryNewArrival = (From issue In efContext.AutoProducts_Issue _
                            Join product In efContext.AutoProducts
                            On issue.ProductId Equals product.ProdouctID
                            Where (issue.SectionID = sectionId And issue.SiteId = siteId And issue.IssueID = issueId)
                            Select product).ToList()
        'LogText(queryNewArrival.Count)
        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")
        Dim lenProductEnd As Integer = "[END_PRODUCT]".Length
        '2013/06/26 added'2014/01/13 添加，如果没查询到产品，则删除该块的内容
        If (queryNewArrival.Count = 0) Then
            oldContent = ""
        End If
        For Each q As AutoProduct In queryNewArrival
            If (productStartIndex > -1) Then
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                If (newProduct.Contains("[URL]")) Then
                    If (siteId <> 68) Then 'siteid= 68 为focalDE站点，因de站点在添加GA代码loadhtml时导致德语乱码，所以用此特殊方法添加追踪代码
                        newProduct = newProduct.Replace("[URL]", q.Url)
                    Else
                        Dim firstIndex As Integer = 0 ' href.LastIndexOf("/", href.LastIndexOf("/") - 1) + 1
                        Dim lastIndex As Integer = 0 ' href.LastIndexOf("/")
                        Dim skuName As String = "" ' href.Substring(firstIndex, lastIndex - firstIndex)
                        Try
                            firstIndex = q.Url.LastIndexOf("/", q.Url.LastIndexOf("/") - 1) + 1
                            lastIndex = q.Url.LastIndexOf("/")
                            skuName = q.Url.Substring(firstIndex, lastIndex - firstIndex)
                        Catch ex As Exception
                            '待确定处理
                        End Try
                        newProduct = newProduct.Replace("[URL]", q.Url & "?utm_source=RSpread&utm_media=EM&utm_campaign=DM_[WEEK]DE_" & skuName & "&source=rspread")
                    End If
                End If
                Common.LogText("Before First")
                If Not (q.Categories.FirstOrDefault Is Nothing) Then
                    newProduct = newProduct.Replace("[CATEGORY_ID]", q.Categories.First.CategoryID.ToString())
                End If
                Common.LogText("First")
                If (newProduct.Contains("[PICTURE_SRC]")) Then
                    newProduct = newProduct.Replace("[PICTURE_SRC]", q.PictureUrl)
                End If
                If (newProduct.Contains("[PICTURE_ALT]")) Then
                    newProduct = newProduct.Replace("[PICTURE_ALT]", q.PictureAlt)
                End If
                If (newProduct.Contains("[PRODUCT_DESCRIPTION]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_DESCRIPTION]", q.Description)
                End If
                If (newProduct.Contains("[PRODUCT_NAME]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_NAME]", q.Prodouct)
                End If
                If (newProduct.Contains("[PRODUCT_PRICE]")) Then
                    If (q.Price Is Nothing) Then  '原价获取不到，只获取到折扣价，则把原价删除掉
                        newProduct = newProduct.Replace("[PRODUCT_PRICE]", "")
                    Else
                        'newProduct = newProduct.Replace("[PRODUCT_PRICE]", q.Currency + " " + String.Format("{0:0.00}", q.Price))
                        If (q.Price Is Nothing) Then '产品的原价为空
                            newProduct = newProduct.Replace("[PRODUCT_PRICE]", "")
                        Else
                            newProduct = newProduct.Replace("[PRODUCT_PRICE]", q.Currency + " " + String.Format("{0:0.00}", q.Price))
                        End If
                    End If
                End If
                If (newProduct.Contains("[PRODUCT_MONEY]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", q.Discount))
                End If
                If (newProduct.Contains("[PRODUCT_SALES]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_SALES]", q.Sales)
                End If
                If (newProduct.Contains("[PRODUCT_TBSCORE]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_TBSCORE]", String.Format("{0:0.0}", q.TbScore))
                End If
                If (newProduct.Contains("[PRODUCT_TBCOMMENT]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_TBCOMMENT]", q.TbComment)
                End If
                Common.LogText("Before Second")
                '2013/07/09 added, add [CATEGORY_NAME] tag
                If Not (q.Categories.FirstOrDefault Is Nothing) Then
                    newProduct = newProduct.Replace("[CATEGORY_NAME]", q.Categories.First.Category1)
                End If
                Common.LogText("Second")
                '2013/05/21,add free shiping pictures and ships pictures in template,begin
                If (newProduct.Contains("[FREESHIPPING]")) Then
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
                '2013/05/21,add free shiping pictures and ships pictures in template,end

                oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                oldContent = oldContent.Insert(productStartIndex, newProduct)
            End If
            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            Else '2013/4/28新增
                Exit For
            End If
        Next
        While (oldContent.Contains("[BEGIN_PRODUCT]"))
            oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            End If
        End While
        Return oldContent
    End Function

    ''' <summary>
    ''' 用CategoryID(表：AutoCategories)替换模板中的link_category的值标签
    ''' （如,[CATEGORY_ID_BAGS]用553替换）
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function HandleCategory(ByVal SpreadTemplate As String, ByVal siteId As Integer) As String
        Dim queryCategory = From c In efContext.AutoCategories Where c.SiteID = siteId Select c
        For Each q As AutoCategory In queryCategory
            Dim cateModifyStr As String = q.Category1.ToUpper().Replace(" ", "_") 'Women's Dresses替换成WOMEN'S_DRESSES
            Dim replaceStr As String = "[CATEGORY_ID_" & cateModifyStr & "]"
            SpreadTemplate = SpreadTemplate.Replace(replaceStr, q.CategoryID.ToString())
        Next
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 添加日期：2013/05/07
    ''' 功能实现：给SpreadTemplate添加追踪代码
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="listCategory"></param>
    ''' <param name="senderName"></param>
    ''' <param name="compaignName"></param>
    ''' <param name="planType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddSpecialCode(ByVal SpreadTemplate As String, ByVal listCategory As List(Of AutoCategory), ByVal senderName As String, ByVal compaignName As String, ByVal planType As String) As String
        Dim byteArray As Byte() = Encoding.GetEncoding("gb2312").GetBytes(SpreadTemplate) 'System.Text.Encoding.UTF8.GetBytes(SpreadTemplate)
        Dim stream As New System.IO.MemoryStream(byteArray)
        Dim fromSource As String = "mail_spread" 'email来源
        Dim utmMedium As String = "mail"
        Dim utmCampaign As String = "regular."
        If (planType.Trim().ToUpper = "ME" OrElse planType.Trim().ToUpper = "WO") Then
            utmCampaign = "special."
        End If
        If (planType.Trim().ToUpper.Contains("HP")) Then
            utmCampaign = "special."
        End If
        '给模板最后加工
        Dim doc As New HtmlDocument() '2013/3/28新加
        'Dim encode As Encoding = Encoding.GetEncoding("gb2312")
        doc.Load(stream) 'Encoding.GetEncodings("gb18030")
        Dim linkNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a")
        For Each link As HtmlNode In linkNodes
            Dim hrefStr As String = link.GetAttributeValue("href", "")
            Dim linkCategory As String = link.GetAttributeValue("link_category", "")
            If Not String.IsNullOrEmpty(linkCategory) Then
                For Each li As AutoCategory In listCategory
                    '2013/07/09 revised,everbuying use the [CATEGORY_NAME] tag
                    'Dim categoryId As Integer = Convert.ToInt32(linkCategory.Trim())
                    'Dim AutoCategory As String
                    'If (li.CategoryID = categoryId) Then
                    '    If (li.Category1.Contains("'")) Then
                    '        AutoCategory = li.Category1.Replace("'", "-")
                    '    Else
                    '        AutoCategory = li.Category1
                    '    End If
                    '    Dim oldUrl As String = hrefStr & "?utm_source=spread&utm_medium=email&utm_term=" & AutoCategory & "&utm_campaign=" & compaignName
                    '    Dim newUrl As String = oldUrl
                    '    link.SetAttributeValue("href", newUrl)
                    'End If

                    Try
                        Dim categoryId As Integer = Convert.ToInt32(linkCategory.Trim())
                        Dim AutoCategory As String = ""
                        If (li.CategoryID = categoryId) Then
                            If (li.Category1.Contains("'")) Then
                                AutoCategory = li.Category1.Replace("'", "-")
                            Else
                                AutoCategory = li.Category1
                            End If
                            '2013/08/07修改，防止出现多个问号（***.html?...?...）
                            'Dim oldUrl As String=hrefStr & "?utm_source=spread&utm_medium=email&utm_term=" & AutoCategory & "&utm_campaign=" & compaignName
                            Dim oldUrl As String = ""
                            If (hrefStr.Contains("?") OrElse hrefStr.Contains("[NEWSLETTER_FORWARD_URL]") OrElse hrefStr.Contains("[UPDATE_PROFILE_URL]") OrElse hrefStr.Contains("[UNSUBSCRIBE_URL]")) Then
                                oldUrl = hrefStr & "&utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                            Else
                                oldUrl = hrefStr & "?utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                            End If
                            '-------------------------------------------------------------------------------------------------------------
                            Dim newUrl As String = oldUrl
                            link.SetAttributeValue("href", newUrl)
                        End If
                    Catch ex As Exception
                        '2013/08/07修改，防止出现多个问号，参数相互影响
                        ' Dim oldUrl As String = hrefStr & "?utm_source=spread&utm_medium=email&utm_term=" & linkCategory & "&utm_campaign=" & compaignName
                        Dim oldUrl As String = ""
                        If (hrefStr.Contains("?") OrElse hrefStr.Contains("[NEWSLETTER_FORWARD_URL]") OrElse hrefStr.Contains("[UPDATE_PROFILE_URL]") OrElse hrefStr.Contains("[UNSUBSCRIBE_URL]")) Then
                            oldUrl = hrefStr & "&utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                        Else
                            oldUrl = hrefStr & "?utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                        End If
                        '-----------------------------------------------------------------------------------------------------------------
                        Dim newUrl As String = oldUrl
                        link.SetAttributeValue("href", newUrl)
                    End Try
                Next
            Else
                '2013/08/07修改，防止出现多个问号，参数相互影响
                'Dim newUrl As String = hrefStr & "?utm_source=spread&utm_medium=email&utm_term=" & senderName & "&utm_campaign=" & compaignName
                Dim newUrl As String = ""
                If (hrefStr.Contains("?") OrElse hrefStr.Contains("[NEWSLETTER_FORWARD_URL]") OrElse hrefStr.Contains("[UPDATE_PROFILE_URL]") OrElse hrefStr.Contains("[UNSUBSCRIBE_URL]")) Then
                    newUrl = hrefStr & "&utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                Else
                    newUrl = hrefStr & "?utm_source=" & fromSource & "&utm_medium=" & utmMedium & "&utm_campaign=" & utmCampaign & DateTime.Now.ToString("yyyyMMdd")
                End If
                '-----------------------------------------------------------------------------------------------------------------------------
                link.SetAttributeValue("href", newUrl)
            End If
        Next
        SpreadTemplate = doc.DocumentNode.OuterHtml.ToString()
        Dim htmlBeginIndex As Integer = SpreadTemplate.IndexOf("<meta")
        SpreadTemplate = SpreadTemplate.Remove(0, htmlBeginIndex)
        Return SpreadTemplate
    End Function



    ''' <summary>
    ''' 获取所有Mobile和Desktop版本的产品信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductElements(ByVal siteId As Integer, ByVal issueId As Integer) As List(Of ProductElement)
        Try
            Dim productElements As New List(Of ProductElement)
            Dim queryProduct = From product In efContext.AutoProducts
                             Join pIssue In efContext.AutoProducts_Issue On product.ProdouctID Equals pIssue.ProductId
                             Where pIssue.SectionID = "NEDM" AndAlso pIssue.IssueID = issueId AndAlso product.SiteID = siteId
                             Select product
            For Each q As AutoProduct In queryProduct
                Dim element As New ProductElement
                element.productName = q.Prodouct
                element.url = q.Url
                If Not (q.Price Is Nothing) Then
                    element.price = q.Price
                End If
                If Not (q.Discount Is Nothing) Then
                    element.discountPrice = q.Discount
                End If
                element.pictureUrl = q.PictureUrl
                element.pubDate = q.LastUpdate
                element.currency = q.Currency
                'Dim category As String = ""
                Dim category As Integer
                For Each c As AutoCategory In q.Categories
                    'category = category & "-" & c.Category1
                    category = c.CategoryID
                Next
                element.category = category
                productElements.Add(element)
            Next
            Return productElements
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' 把最受欢迎的产品的插入到产品list中
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="listProductElements"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertPopularElements(ByVal LoginEmail As String, ByVal AppID As String, ByVal listProductElements As List(Of ProductElement)) As List(Of ProductElement)
        Try
            Dim popularProductList As New List(Of ProductElement)
            Dim NumNew As Integer = 0 'Rss中今日团购的产品个数，根据时间判断
            Dim NumOld As Integer = 0 'Rss中往日团购的产品个数，根据时间判断
            For i As Integer = 0 To listProductElements.Count - 1
                Dim d1 As DateTime = Format(listProductElements(i).pubDate, "yyyy-MM-dd")
                Dim d2 As DateTime = Format(listProductElements(i).pubDate, "yyyy-MM-dd")
                If (d1 = d2) Then
                    NumNew = NumNew + 1
                Else
                    NumOld = NumOld + 1
                End If
            Next
            Dim NeedNumNew As Integer = 2 - (NumNew Mod 2) '需要插入的产品的个数
            Dim CampaignName As String = "LifeInHK".ToUpper & "HO" & Format(Now.Date.AddDays(-1), "dd/MM/yyyy")
            Dim spreadiws As New EmailAlerter.IntrawWebService.Service
            Dim ds As DataSet = spreadiws.dsGetURLTrackings(LoginEmail, AppID, CampaignName)
            Dim dt As DataTable = ds.Tables(0)
            If dt.Rows.Count > 0 Then
                Dim n As Integer = 0
                'For m As Integer = 0 To NeedNumNew - 1
                Dim m As Integer = 0
                Dim foundFlag As Boolean = False
                For t As Integer = n To dt.Rows.Count - 1
                    Dim favouriteUrl As String = dt.Rows(t).Item(0).ToString()
                    For Each li As ProductElement In listProductElements
                        Dim productUrl As String = li.url
                        If (productUrl.LastIndexOf("?") > -1) Then
                            productUrl = productUrl.Substring(0, productUrl.LastIndexOf("?"))
                        End If
                        If String.Compare(productUrl, favouriteUrl, True) = 0 Then
                            Dim PubDate As String = Format(DateTime.Parse(li.pubDate), "yyyy-MM-dd")
                            Dim NowDate As String = Format(Now, "yyyy-MM-dd")
                            If PubDate <> NowDate Then
                                Dim productElement As New ProductElement
                                productElement.productName = li.productName
                                productElement.url = li.url
                                productElement.price = li.price
                                productElement.discountPrice = li.discountPrice
                                productElement.pictureUrl = li.pictureUrl
                                productElement.pubDate = NowDate
                                productElement.category = li.category
                                productElement.currency = li.currency
                                popularProductList.Add(productElement)
                                listProductElements.Remove(li)
                                m = m + 1
                                Exit For
                            End If
                        End If
                    Next
                    If m >= 2 Then '找到了一个就继续下一次循环
                        Exit For
                    End If
                Next
            End If
            Try
                '把点击量最高的产品信息添加到第二、三个位置
                For i As Integer = 0 To popularProductList.Count - 1
                    listProductElements.Insert(i + 1, popularProductList(i))
                Next
            Catch ex As Exception
                Common.LogText(ex.ToString())
            End Try
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try
        Return listProductElements
    End Function

    ''' <summary>
    ''' 整数不加.00，而float数字添加.00，并返回转化后的字符串
    ''' </summary>
    ''' <param name="myNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DoFormat(ByVal myNumber As Double) As String
        Dim s As String = String.Format("{0:0.00}", myNumber)
        If s.EndsWith("00") Then
            Return Integer.Parse(myNumber).ToString()
        Else
            Return s
        End If
    End Function

    ''' <summary>
    ''' 写日志到日志表(SentLogs)中,
    ''' 2013/06/04添加
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="subject"></param>
    ''' <param name="planType"></param>
    ''' <param name="logList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertSentLog(ByVal siteId As Integer, ByVal subject As String, _
                              ByVal planType As String, ByVal logList As List(Of AutoSentLog)) As List(Of AutoSentLog)
        Try
            '发送日志处理
            Dim log As New AutoSentLog()
            log.Subject = subject
            log.LastSentAt = Now
            log.SiteID = siteId
            log.PlanType = planType
            efContext.AddToAutoSentLogs(log)
            efContext.SaveChanges()
        Catch ex As Exception
            '写数据到内存中
            Dim listOfLog As New AutoSentLog()
            listOfLog.SiteID = siteId
            listOfLog.PlanType = planType
            listOfLog.LastSentAt = Now
            logList.Add(listOfLog)
        End Try
        Return logList
    End Function

    ''' <summary>
    ''' 获取点击量最大的产品
    ''' </summary>
    ''' <param name="listProductElements"></param>
    ''' <param name="loginEmail"></param>
    ''' <param name="appId"></param>
    ''' <param name="counter">点击量最大的产品个数</param>
    ''' <returns>点击量最大的产品URL</returns>
    ''' <remarks></remarks>
    Private Function GetTopNProducts(ByVal listProductElements As List(Of ProductElement), ByVal loginEmail As String, _
                                     ByVal appId As String, ByVal siteId As Integer, ByVal senderName As String, ByVal planType As String, ByVal counter As Integer) As List(Of ProductElement)
        Dim queryLastSentTime As DateTime = (From log In efContext.AutoSentLogs
                              Where log.SiteID = siteId AndAlso log.PlanType = planType AndAlso log.LastSentAt < Now AndAlso Not log.Subject = ""
                              Order By log.LastSentAt Descending
                              Select log.LastSentAt).FirstOrDefault
        Dim campaignName As String = senderName.ToUpper() & planType & queryLastSentTime.ToString("yyyyMMdd")
        Dim spreadiws As New EmailAlerter.IntrawWebService.Service
        Dim ds As New DataSet
        Try
            ds = spreadiws.dsGetURLTrackings(loginEmail, appId, campaignName)
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try
        Dim dt As DataTable = ds.Tables(0)
        If (counter > listProductElements.Count) Then  '需求调整的产品个数<现有产品的个数
            counter = listProductElements.Count
        End If
        Dim productCounter As Integer = 0
        For i As Integer = 0 To dt.Rows.Count - 1
            If productCounter < counter Then
                For j As Integer = 0 To listProductElements.Count - 1
                    If listProductElements(j).url = dt.Rows(i).Item(0) Then
                        Dim element As ProductElement = listProductElements(j)
                        listProductElements.RemoveAt(j) '移除产品信息
                        listProductElements.Insert(productCounter, element) '在list的前几个索引处添加该产品信息
                        productCounter = productCounter + 1
                        Exit For
                    End If
                Next
            End If
        Next
        Return listProductElements
    End Function

    ''' <summary>
    ''' update the table TemplateLibrary record according to TemplateId,
    ''' or add a new record if not exist the record
    ''' </summary>
    ''' <param name="templateId"></param>
    ''' <remarks></remarks>
    Private Sub InsertTemplateLibrary(ByVal templateId As Integer, ByVal SpreadTemplate As String)
        Using efContext2 As New FaceBookForSEOEntities
            Dim queryTempLib As TemplateLibrary = efContext2.TemplateLibraries.Where(Function(t) t.Tid = templateId).SingleOrDefault
            If (queryTempLib Is Nothing) Then 'Database has no record
                Dim tempLib As New TemplateLibrary
                tempLib.Tid = templateId
                tempLib.Contents = SpreadTemplate
                efContext2.AddToTemplateLibraries(tempLib)
            Else
                queryTempLib.Contents = SpreadTemplate
            End If
            efContext2.SaveChanges()
        End Using
    End Sub

    ''' <summary>
    ''' update the SentStatus of table TemplateLibrary
    ''' </summary>
    ''' <param name="IssueId"></param>
    ''' <param name="sentStatus"></param>
    ''' <remarks></remarks>
    Public Sub UpdateIssueStatus(ByVal IssueId As Integer, ByVal sentStatus As String)
        Using efContext2 As New FaceBookForSEOEntities '防止缓存,Issues表的发送状态位Success
            Dim queryIssue As AutoIssue = efContext2.AutoIssues.Where(Function(iss) iss.IssueID = IssueId).SingleOrDefault()
            queryIssue.SentStatus = sentStatus
            efContext2.SaveChanges()
        End Using
    End Sub

    ' ''' <summary>
    ' ''' 写错误日志到日志文件中
    ' ''' </summary>
    ' ''' <param name="Ex"></param>
    ' ''' <remarks></remarks>
    'Sub Log(ByVal Ex As Exception)
    '    Try
    '        LogText(Now & ", " & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub

    'Sub LogText(ByVal Ex As String)
    '    Try
    '        System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())

    '        '2013/08/08 added, 发送错误日志到制定的邮箱组
    '        If (Ex.Contains("Exception")) Then
    '            NotificationEmail.SentErrorEmail(Ex.ToString())
    '        End If
    '        '----------------------------------------------------

    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub
End Class
