Public Class alerter
    'Create campaign and send to all list under the account
    Function CreateCampaign(ByVal Subject As String, ByVal LoginName As String, ByVal AppID As String,
                            ByVal senderName As String, ByVal SenderEmail As String, ByVal EmailBody As String) _
        As String
        Try

            'get data from local database
            'Dim emailAlerterData As EmailTableAdapters.RssSubscriptionsTableAdapter = New EmailTableAdapters.RssSubscriptionsTableAdapter()
            'Dim dt As System.Data.DataTable = emailAlerterData.GetDataById(1)
            'Dim rssId As String = dt.Rows(0)("rssid").ToString()
            'Dim url As String = dt.Rows(0)("url").ToString()
            Try
                If Subject.Length >= 121 Then
                    Dim Sub1 As String = Subject.Substring(0, 120)
                    Dim Index As Integer = Sub1.LastIndexOf("/")
                    Subject = Sub1.Substring(0, Index)
                    'Subject.
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

            'get contact lists from spread service
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            Dim subscriptionData As System.Data.DataSet = mySpread.getAllSubscription(LoginName, AppID)

            'Dim temp As String = ""
            Dim ContactLists(subscriptionData.Tables(0).Select("status='Invisible'").Length - 1) As String
            Dim i As Integer
            For Each row As System.Data.DataRow In subscriptionData.Tables(0).Rows
                If row(2).ToString() = "Invisible" Then
                    ContactLists(i) = row("Subscription Name")
                    'temp = temp & "," & row(0).ToString()
                    i += 1
                End If
            Next

            'create campaign
            'Dim targetSubscription As String() = temp.Split(",")
            Dim intervalTime As Integer = - 1
            Dim campaignId As Integer
            Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

            myCampaign.campaignName = (Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year).ToString

            myCampaign.from = senderName
            myCampaign.fromEmail = SenderEmail
            myCampaign.subject = Subject
            myCampaign.content = EmailBody
            Dim de As Date = Now
            myCampaign.schedule = de
            'send campaign
            Try
                campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            'add by gary start 20130801
            Try
                Dim subject2 = senderName & " 开始发送邮件"
                Dim SpreadTemplate2 = subject2 & "<br />状态：正在发送<br />" & " 标题:" & Subject & "<br />内容：<br />" &
                                      EmailBody
                mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com",
                              "自动化发送", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'add by gary end 20130801
            Return myCampaign.subject
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' 创建Campaign，并设置Schedule时间
    ''' </summary>
    ''' <param name="Subject"></param>
    ''' <param name="LoginName"></param>
    ''' <param name="AppID"></param>
    ''' <param name="senderName"></param>
    ''' <param name="SenderEmail"></param>
    ''' <param name="EmailBody"></param>
    ''' <param name="ScheduleTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CreateCampaign(ByVal Subject As String, ByVal LoginName As String, ByVal AppID As String,
                            ByVal senderName As String,
                            ByVal SenderEmail As String, ByVal EmailBody As String, ByVal ScheduleTime As DateTime) _
        As String
        Try

            'get data from local database
            'Dim emailAlerterData As EmailTableAdapters.RssSubscriptionsTableAdapter = New EmailTableAdapters.RssSubscriptionsTableAdapter()
            'Dim dt As System.Data.DataTable = emailAlerterData.GetDataById(1)
            'Dim rssId As String = dt.Rows(0)("rssid").ToString()
            'Dim url As String = dt.Rows(0)("url").ToString()
            Try
                If Subject.Length >= 121 Then
                    Dim Sub1 As String = Subject.Substring(0, 120)
                    Dim Index As Integer = Sub1.LastIndexOf("/")
                    Subject = Sub1.Substring(0, Index)
                    'Subject.
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

            'get contact lists from spread service
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            Dim subscriptionData As System.Data.DataSet = mySpread.getAllSubscription(LoginName, AppID)

            'Dim temp As String = ""
            Dim ContactLists(subscriptionData.Tables(0).Select("status='Invisible'").Length - 1) As String
            Dim i As Integer
            For Each row As System.Data.DataRow In subscriptionData.Tables(0).Rows
                If row("status").ToString() = "Invisible" Then
                    ContactLists(i) = row("Subscription Name")
                    'temp = temp & "," & row(0).ToString()
                    i += 1
                End If
            Next

            'create campaign
            'Dim targetSubscription As String() = temp.Split(",")
            Dim intervalTime As Integer = - 1
            Dim campaignId As Integer
            Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

            myCampaign.campaignName = (Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year).ToString

            myCampaign.from = senderName
            myCampaign.fromEmail = SenderEmail
            myCampaign.subject = Subject
            myCampaign.content = EmailBody
            Dim de As Date = Now
            myCampaign.schedule = ScheduleTime
            'send campaign
            Try
                campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            'add by gary start 20130801
            Try
                Dim subject2 = senderName & " 开始发送邮件"
                Dim SpreadTemplate2 = subject2 & "<br />状态：正在发送<br />" & " 标题:" & Subject & "<br />内容：<br />" &
                                      EmailBody
                mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com",
                              "自动化发送", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'add by gary end 20130801
            Return myCampaign.subject
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    Sub UpdateSentLog(ByVal RSSId As Integer, ByVal Subject As String)
        Try
            Dim sentlog As SentLog = New SentLog()
            sentlog.RssId = RSSId
            sentlog.Subject = Subject
            sentlog.LastSentAt = Now
            SentLogs.UpdateSentLog(sentlog)
            'emailAlerterData.UpdateQuery(1, 1)
        Catch ex As Exception

            Common.LogText(ex.ToString)
        End Try
    End Sub

    'Sub Log(ByVal Ex As Exception)
    '    Try
    '        LogText(Now & ", " & Ex.InnerException.ToString & ";" & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub

    'Sub LogText(ByVal Ex As String)
    '    Try
    '        '2013/08/08 added, 发送错误日志到制定的邮箱组
    '        If (Ex.Contains("Exception")) Then
    '            NotificationEmail.SentErrorEmail(Ex.ToString())
    '        End If
    '        '-------------------------------------------------------
    '        System.IO.File.AppendAllText(
    '            System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log",
    '            Now & ", " & Ex & Environment.NewLine())

    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub
End Class
