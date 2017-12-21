Imports System.Linq

Public Class MyRender
    Private Shared efContext As New FaceBookForSEOEntities()
    ''' <summary>
    ''' EmailAlerter主程序优化：减少对dll的更新，可在AutomationPlan表中直接配置部分dll的内容 
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="spreadEmail"></param>
    ''' <param name="appId"></param>
    ''' <param name="SenderName"></param>
    ''' <param name="SenderMail"></param>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="subject"></param>
    ''' <param name="planType"></param>
    ''' <param name="campaignName"></param>
    ''' <param name="volumn"></param>
    ''' <param name="dllType"></param>
    ''' <param name="CampaignStatus"></param>
    ''' <param name="ScheduleTimeInteval"></param>
    ''' <param name="ListType"></param>
    ''' <param name="ContactList"></param>
    ''' <param name="SearchAPIType"></param>
    ''' <param name="SearchStartDayInterval"></param>
    ''' <param name="SearchEndDayInterval"></param>
    ''' <param name="UrlSpecialCode"></param>
    ''' <param name="categories">各个分类之间用"^"分隔，如：Dresses^Apparel^Shoes^Accessories</param>
    ''' <param name="LimitQuantity">每个发送计划的一次发送量，即一次Campaign的发送量</param>
    ''' <param name="strTimeLimit">时间限制，可以设置多个时段</param>
    ''' <param name="sellerEmail">销售人员邮箱，多个邮箱之间用","隔开</param>
    ''' <remarks></remarks>
    Public Shared Sub Start(ByVal issueId As Integer, ByVal siteId As Integer, _
                       ByVal spreadEmail As String, ByVal appId As String, ByVal SenderName As String, _
                       ByVal SenderMail As String, ByVal SpreadTemplate As String, ByVal subject As String, _
                       ByVal planType As String, ByVal campaignName As String, ByVal volumn As Integer, ByVal dllType As String, _
                       ByVal CampaignStatus As String, ByVal ScheduleTimeInteval As Double, ByVal ListType As String, _
                       ByVal ContactList As String, ByVal SearchAPIType As String, ByVal SearchStartDayInterval As Integer, _
                       ByVal SearchEndDayInterval As Integer, ByVal UrlSpecialCode As String, ByVal categories As String, _
                       ByVal LimitQuantity As Long, ByVal strTimeLimit As String, ByVal SellerEmail As String)

        While SpreadTemplate.Contains("[CAMPAIGN_NAME]")
            SpreadTemplate = SpreadTemplate.Replace("[CAMPAIGN_NAME]", campaignName)
        End While
        Dim arrContactList As String()
        Dim campaignId As Integer = 0
        If Not (String.IsNullOrEmpty(ListType)) Then '配置AutomationPlan表中的ListType字段，值：NOR代表Normal，SCT代表SearchContact
            Dim searchStrategy As ChooseStrategy
            Select Case ListType
                Case "NOR"
                    arrContactList = GetContactList(ContactList, issueId)
                Case "SCT"
                    arrContactList = {categories.Split("^")(0) & planType & DateTime.Now.ToString("yyyyMMdd")}
                    Select Case SearchAPIType
                        Case "OPE"  'Open
                            searchStrategy = ChooseStrategy.Open
                            CreateContactList(siteId, issueId, planType, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0), searchStrategy, spreadEmail, appId)
                        Case "RAN" 'Random
                            searchStrategy = ChooseStrategy.Random
                            CreateContactList(siteId, issueId, planType, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0), searchStrategy, spreadEmail, appId)
                        Case "ORD" 'OpenRandom
                            searchStrategy = ChooseStrategy.OpenRandom
                            CreateContactList(siteId, issueId, planType, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0), searchStrategy, spreadEmail, appId)
                        Case "FAV" 'Favorite
                            searchStrategy = ChooseStrategy.Favorite
                            CreateContactList(siteId, issueId, planType, categories.Split("^")(0).Trim, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0),
                                              searchStrategy, spreadEmail, appId)
                        Case "OEC" 'OpenExcludeCategory
                            searchStrategy = ChooseStrategy.OpenExcludeCategory
                            CreateContactList(siteId, issueId, planType, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0), searchStrategy, spreadEmail, appId)
                        Case "NOP" 'NotOpen
                            searchStrategy = ChooseStrategy.NotOpen
                            CreateContactList(siteId, issueId, planType, SearchStartDayInterval, SearchEndDayInterval, arrContactList(0), searchStrategy, spreadEmail, appId)
                    End Select
            End Select
            campaignId = CreateCampaign(spreadEmail, appId, campaignName, arrContactList, SpreadTemplate, SenderName, SenderMail, subject, CampaignStatus, ScheduleTimeInteval, LimitQuantity, SellerEmail, strTimeLimit)
            '2014/08/01 added by dora,NFC/NFE使用
            If (campaignId > 0) Then
                Dim myIssue As AutoIssue = efContext.AutoIssues.Where(Function(issue) issue.IssueID = issueId).Single()
                myIssue.SpreadCampId = campaignId
                efContext.SaveChanges()
            End If
        Else  'AutomationPlan表ListType字段为空
            If Not (String.IsNullOrEmpty(CampaignStatus)) Then
                arrContactList = GetContactList(ContactList, issueId)
                campaignId = CreateCampaign(spreadEmail, appId, campaignName, arrContactList, SpreadTemplate, SenderName, SenderMail, subject, CampaignStatus, ScheduleTimeInteval, LimitQuantity, SellerEmail, strTimeLimit)
                '2014/08/01 added by dora,NFC/NFE使用
                If (campaignId > 0) Then
                    Dim myIssue As AutoIssue = efContext.AutoIssues.Where(Function(issue) issue.IssueID = issueId).Single()
                    myIssue.SpreadCampId = campaignId
                    efContext.SaveChanges()
                End If
            Else
                Dim auto As New AutoModel
                auto.Render(issueId, siteId, spreadEmail, appId, SenderName, SenderMail, SpreadTemplate, subject, planType, campaignName, volumn, dllType)
            End If
        End If
    End Sub

    Public Shared Sub CreateNotification(ByVal senderName As String, ByVal senderMail As String, ByVal lists As String(), _
                                         ByVal subject As String, ByVal campaignName As String, ByVal SpreadTemplate As String)
        Try
            Dim subject2 = senderName & " 开始发送邮件"
            Dim myList As String = ""
            For Each list As String In lists
                myList = myList & list & ","
            Next
            Dim SpreadTemplate2 As String = subject2 & "<br />状态：正在发送<br />" & " 标题:" & subject _
                                            & "<br />发送对象：" & myList & "<br />发件人Email：" & senderMail & "<br />发件人：" & senderName _
                                            & "<br />Campaign名：" & campaignName & "<br />是否Publish to Newsletter Archive：" & _
                                            "是" & "<br />内容(*部分Spread系统按钮在提醒邮件里会失效，但不影响正常的发送)：<br />" & SpreadTemplate
            NotificationEmail.SentStartEmail(subject2, SpreadTemplate2)
        Catch ex As Exception
            'Ignore
            'LogText(ex.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' 创建Campaign，并设定Campaign的发送状态：Draft,Schedule,Right Now
    ''' </summary>
    ''' <param name="spreadEmail"></param>
    ''' <param name="appId"></param>
    ''' <param name="campaignName"></param>
    ''' <param name="mylist"></param>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="senderName"></param>
    ''' <param name="senderMail"></param>
    ''' <param name="subject"></param>
    ''' <param name="campaignStatus"></param>
    ''' <param name="scheduleTimeInteval"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CreateCampaign(ByVal spreadEmail As String, ByVal appId As String, ByVal campaignName As String, _
                                           ByVal mylist As String(), ByVal SpreadTemplate As String, ByVal senderName As String, _
                                           ByVal senderMail As String, ByVal subject As String, ByVal campaignStatus As String, _
                                           ByVal scheduleTimeInteval As Double, ByVal sendingQuantity As Long, _
                                           ByVal sellerEmail As String, ByVal strTimeLimit As String) As Integer
        Dim mySpread As SpreadService.Service = New SpreadService.Service()
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
        Dim campaignId As Integer = 0
        Dim myCampaignStatusStr As String = ""
        Dim myCampaignStatus As SpreadService.CampaignStatus
        Dim scheduleTime As Date = Date.Now
        Select Case campaignStatus
            Case "DFT"  'Draft
                myCampaignStatus = SpreadService.CampaignStatus.Draft
                myCampaignStatusStr = "草稿"
            Case "SDE"  'Schedule
                myCampaignStatus = SpreadService.CampaignStatus.Waiting
                scheduleTime = scheduleTime.AddHours(scheduleTimeInteval)
                myCampaignStatusStr = "待发" & "<br/>预定发出时间：" & Date.Now.AddHours(scheduleTimeInteval).ToString("yyyy-MM-dd HH:mm:ss")
            Case "SRN"  'Send Right Now
                myCampaignStatus = SpreadService.CampaignStatus.Waiting
                myCampaignStatusStr = "发送中"
        End Select
        campaignId = mySpread.createCampaign2(spreadEmail, appId, campaignName, campaignCreative, mylist, interval, scheduleTime, "Spread", myCampaignStatus)
        If (campaignId > 0) Then
            Try
                If (sendingQuantity > 0 AndAlso Not String.IsNullOrEmpty(strTimeLimit)) Then '如果没配置该字段，则表示不限制每天的发送量；最大发送量限制和时间限制必须是同时限制
                    mySpread.SetCampaignDailyLimit(spreadEmail, appId, campaignId, sendingQuantity)
                    myCampaignStatusStr = myCampaignStatusStr & "<br/> " & "每天最大发送量：" & sendingQuantity
                    'End If
                    'If Not (String.IsNullOrEmpty(strTimeLimit)) Then
                    mySpread.SetCampaignTimeLimit(spreadEmail, appId, campaignId, strTimeLimit)
                    myCampaignStatusStr = myCampaignStatusStr & "<br/> " & "发送时间段设置：" & strTimeLimit & "<span style=""font-size: 12px; line-height: 15px; font-family: 微软雅黑, arial, helvetica, sans-serif;""> (注：'0'代表每天,'1': 星期一,'7': 星期天，依此类推)</span>"
                End If

                Dim sellersubject As String = senderName & Date.Now.ToString("yyyy-MM-dd") & "智能化邮件创建详情"
                Dim body As String
                Dim allList As String
                For Each li As String In mylist
                    allList = allList & li & ","
                Next
                allList = allList.TrimEnd(",")
                body = "<span style=""font-size: 14px; line-height: 16px; font-family: 微软雅黑, arial, helvetica, sans-serif; "">"
                body = body & "以下信息为系统自动发出，请勿回复，如需修改，请登录Spread:  https://app.rspread.com/login.aspx <br/>" 'https://app.rspread.com/login.aspx  http://globaledm.asia/login.aspx
                body = body & "活动名称： " & campaignName & "<br/> " & "邮件标题： " & subject & "<br/> " & "创建时间： " & Date.Now.ToString("yyyy-MM-dd HH:mm:ss") & "<br/> " &
                       "邮件状态： " & myCampaignStatusStr & "<br/> " & "发送名单： " & allList & "<br/> " & "发件人名称： " & senderName & "<br/> " & "发件人地址： " & senderMail & "<br/> " &
                       "邮件内容(*部分Spread系统按钮在提醒邮件里会失效，但不影响正常的发送)： " & "<br/> " & SpreadTemplate
                body = body & "</span>"
                If Not (String.IsNullOrEmpty(sellerEmail)) Then '如已配置销售人员邮箱,则发送提醒邮件到销售人员邮箱中
                    If (sellerEmail.Contains(",")) Then
                        Dim arrSellerMail As String() = sellerEmail.Split(",")
                        For Each mail As String In arrSellerMail
                            NotificationEmail.SentStartEmail(sellersubject, body, mail)
                        Next
                    Else
                        NotificationEmail.SentStartEmail(sellersubject, body, sellerEmail)
                    End If
                End If
                '发送一封提醒邮件到Emailalerter组别中，以供技术人员查看,
                NotificationEmail.SentStartEmail(sellersubject, body)
            Catch ex As Exception
                NotificationEmail.SentErrorEmail(ex.ToString())
            End Try
        End If
        mySpread.ChangePublishStatus(spreadEmail, appId, campaignId, True) ''public邮件到邮件档案馆
        'End If
        Return campaignId
        'Dim campaignId As Integer = mySpread.createCampaign2(loginEmail, appId, campaignName, campaignCreative, mylist, interval, Now.AddHours(2), "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
    End Function

    ''' <summary>
    ''' 从AutomationPlan表或者ContactLists_Issue表中获取联系人列表数组
    ''' </summary>
    ''' <param name="ContactList1"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetContactList(ByVal ContactList1 As String, ByVal issueId As Integer) As String()
        Dim arrContactList As String()
        If (String.IsNullOrEmpty(ContactList1)) Then '没有配置AutomationPlan表中的ContactList字段
            Using efContext As New FaceBookForSEOEntities
                arrContactList = (From c In efContext.AutoContactLists_Issue _
                                         Where c.IssueID = issueId
                                         Select c.ContactList).ToArray()
                'arrContactList = (From c In efContext.AutoContactLists_Issue _
                'Where(c.IssueID = issueId AndAlso String.IsNullOrEmpty(c.SendingStatus))
                '                         Select c.ContactList).ToArray()
            End Using
        Else
            If (ContactList1.Contains(",")) Then
                arrContactList = ContactList1.Split(",")
            Else
                arrContactList = New String() {ContactList1}
            End If
        End If
        Return arrContactList
    End Function

#Region "SeachContacts"
    ''' <summary>
    ''' 调用spreadApi SearchContacts（） 函数，创建一个未点击任何分类的联系人列表 
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="saveListName"></param>
    ''' <param name="daySpan"></param>
    ''' <param name="strategy"></param>
    ''' <param name="sendStatus"></param>
    ''' <param name="loginEmail"></param>
    ''' <param name="appId"></param>
    ''' <remarks></remarks>
    Public Shared Function CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal startDayInterval As Integer, ByVal endDayInterval As Integer, _
                                        ByVal saveListName As String, ByVal strategy As ChooseStrategy, ByVal loginEmail As String, ByVal appId As String) As Integer
        'Dim saveListName As String = planType & DateTime.Now.ToString("yyyyMMdd")
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(startDayInterval).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = Date.Now.AddDays(endDayInterval).ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadService.Service = New SpreadService.Service()
        mySpread.Timeout = 2400000
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
            Return count
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' 调用spreadApi SearchContacts（） 函数，创建一个点击了指定分类的联系人列表 
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="saveListName"></param>
    ''' <param name="daySpan"></param>
    ''' <param name="strategy"></param>
    ''' <param name="sendStatus"></param>
    ''' <param name="loginEmail"></param>
    ''' <param name="appId"></param>
    ''' <remarks></remarks>
    Public Shared Function CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal categoryName As String, ByVal startDayInterval As Integer,
                                      ByVal endDayInterval As Integer, ByVal saveListName As String, ByVal strategy As ChooseStrategy, ByVal loginEmail As String, ByVal appId As String) As Integer
        'Dim saveListName As String = categoryName & planType & DateTime.Now.ToString("yyyyMMdd")
        Dim categoryId As Integer = -1
        Using efContext As New FaceBookForSEOEntities
            If (categoryName.Contains("^")) Then
                categoryName = categoryName.Split("^")(0).Trim()
            Else
                categoryName = categoryName.Trim()
            End If
            categoryId = efContext.AutoCategories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).FirstOrDefault().CategoryID
        End Using
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Favorite = categoryId
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(startDayInterval).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = Date.Now.AddDays(endDayInterval).ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadService.Service = New SpreadService.Service()
        mySpread.Timeout = 2400000
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
            Return count
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' 将某个特定的List名加入到ContactLists_Issue表中
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="contactList"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertContactList(ByVal IssueID As Integer, ByVal contactList As String, ByVal sendingStatus As String)
        Dim myContactList As New AutoContactLists_Issue
        myContactList.IssueID = IssueID
        myContactList.ContactList = contactList
        myContactList.SendingStatus = sendingStatus
        efContext.AddToAutoContactLists_Issue(myContactList)
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 将一个或多个List名插入到发送表ContactLists_Issue表中
    ''' </summary>
    ''' <param name="IssueId"></param>
    ''' <param name="ContactListArr"></param>
    ''' <param name="sendingStatus"></param>
    ''' <remarks></remarks>
    Public Sub InsertContactList(ByVal IssueId As Integer, ByVal ContactListArr As String(), ByVal sendingStatus As String)
        For i As Integer = 0 To ContactListArr.Length - 1
            Dim myContactList As New AutoContactLists_Issue
            myContactList.IssueID = IssueId
            myContactList.ContactList = ContactListArr(i)
            myContactList.SendingStatus = sendingStatus
            efContext.AddToAutoContactLists_Issue(myContactList)
        Next
        efContext.SaveChanges()
    End Sub

#End Region
End Class
