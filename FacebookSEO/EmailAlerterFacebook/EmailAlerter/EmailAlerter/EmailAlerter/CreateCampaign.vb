Public Class CreateCampaign
    'send out emails with the category(Open & Clicked)
    Public Shared Function CreateCampaignByCategory(ByVal Favorite As String, ByVal CampaignName As String,
                                                    ByVal Subject As String, ByVal LoginName As String,
                                                    ByVal AppID As String, ByVal senderName As String,
                                                    ByVal SenderEmail As String, ByVal EmailBody As String) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Favorite = Favorite
            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0
            Try
                Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList,
                                                True)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
            '    Catch ex As Exception
            '        A.LogText("retry to create f-0 failed" & ex.ToString)

            '    End Try
            'End If
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613760312261" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
            '        Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
            '        request.GetResponse()
            '    Catch ex As Exception
            '        A.LogText(ex.ToString)
            '    End Try
            'End If


            If Count > 0 Then
                Dim ContactLists() As String = New String() {FavoriteContactsList}
                Dim intervalTime As Integer = - 1
                Dim campaignId As Integer
                Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

                myCampaign.campaignName = CampaignName

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
                Common.LogText(Subject)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Else
                Common.LogText("No contact in this favorite:" & Favorite)
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function

    'Add By Gary Start 2013/06/09
    '根据时间段判断用户点击的分类发送个性化邮件/分类邮件
    Public Shared Function CreateCampaignByCategoryByClickDateRange(ByVal Favorite As String,
                                                                    ByVal CampaignName As String,
                                                                    ByVal Subject As String, ByVal LoginName As String,
                                                                    ByVal AppID As String, ByVal senderName As String,
                                                                    ByVal SenderEmail As String,
                                                                    ByVal EmailBody As String,
                                                                    Optional ByVal ClickStartDate As String = "",
                                                                    Optional ByVal ClickEndDate As String = "") _
        As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Favorite = Favorite
            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            If ClickStartDate IsNot Nothing AndAlso ClickStartDate > "" Then
                QuerySubscriber.StartDate = ClickStartDate
            End If
            If ClickEndDate IsNot Nothing AndAlso ClickEndDate > "" Then
                QuerySubscriber.EndDate = ClickEndDate
            End If
            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0
            Try
                Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList,
                                                True)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
            '    Catch ex As Exception
            '        A.LogText("retry to create f-0 failed" & ex.ToString)

            '    End Try
            'End If
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613760312261" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
            '        Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
            '        request.GetResponse()
            '    Catch ex As Exception
            '        A.LogText(ex.ToString)
            '    End Try
            'End If


            If Count > 0 Then
                Dim ContactLists() As String = New String() {FavoriteContactsList}
                Dim intervalTime As Integer = - 1
                Dim campaignId As Integer
                Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

                myCampaign.campaignName = CampaignName

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
                Common.LogText(Subject)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Else
                Common.LogText("No contact in this favorite:" & Favorite)
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function

    'Add By Gary End

    'Add By Gary Start 2013/07/19
    '根据时间段判断用户点击的分类发送个性化邮件/分类邮件
    Public Shared Function CreateCampaignByCampaignStatusAndCategoryByClickDateRange(ByVal Favorite As String,
                                                                                     ByVal CampaignName As String,
                                                                                     ByVal Subject As String,
                                                                                     ByVal LoginName As String,
                                                                                     ByVal AppID As String,
                                                                                     ByVal senderName As String,
                                                                                     ByVal SenderEmail As String,
                                                                                     ByVal EmailBody As String,
                                                                                     ByVal status As String,
                                                                                     Optional ByVal ClickStartDate As _
                                                                                        String = "",
                                                                                     Optional ByVal ClickEndDate As _
                                                                                        String = "") As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Favorite = Favorite
            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            If ClickStartDate IsNot Nothing AndAlso ClickStartDate > "" Then
                QuerySubscriber.StartDate = ClickStartDate
            End If
            If ClickEndDate IsNot Nothing AndAlso ClickEndDate > "" Then
                QuerySubscriber.EndDate = ClickEndDate
            End If
            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0
            Try
                Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList,
                                                True)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
            '    Catch ex As Exception
            '        A.LogText("retry to create f-0 failed" & ex.ToString)

            '    End Try
            'End If
            'If Count = 0 And Favorite = "0" Then
            '    Try
            '        Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613760312261" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
            '        Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
            '        request.GetResponse()
            '    Catch ex As Exception
            '        A.LogText(ex.ToString)
            '    End Try
            'End If


            If Count > 0 Then
                Dim ContactLists() As String = New String() {FavoriteContactsList}
                Dim intervalTime As Integer = - 1
                Dim campaignId As Integer
                Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()
                'SpreadService.CampaignStatus.Waiting
                Dim CampaignSatus As SpreadService.CampaignStatus = SpreadService.CampaignStatus.Waiting
                'SpreadService.CampaignStatus.Draft '2013/10/31 added，所有的分类触发立即发送，而不是创建草稿状态

                myCampaign.campaignName = CampaignName

                myCampaign.from = senderName
                myCampaign.fromEmail = SenderEmail
                myCampaign.subject = Subject
                myCampaign.content = EmailBody
                Dim de As Date = Now
                myCampaign.schedule = de


                Dim campaignCreative(0) As EmailAlerter.SpreadService.CampaignCreatives
                Dim c As EmailAlerter.SpreadService.CampaignCreatives = New EmailAlerter.SpreadService.CampaignCreatives
                With c
                    .creativeContent = EmailBody
                    .displayName = senderName
                    .fromAddress = SenderEmail
                    .isCampaignDefault = True
                    .replyTo = ""
                    .subject = Subject
                    .target = "D"
                End With
                campaignCreative(0) = c
                If status.ToLower = "draft" Then
                    CampaignSatus = SpreadService.CampaignStatus.Draft
                End If
                Dim interval As Integer = - 1
                'send campaign
                Try
                    'campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
                    campaignId = mySpread.createCampaign2(LoginName, AppID, CampaignName, campaignCreative, ContactLists,
                                                          interval, Now, "Spread", CampaignSatus)
                    'add by gary start 20130801
                    Try
                        Dim subject2 = senderName & "自动化邮件开始发送..."
                        '创建草稿成功
                        Dim SpreadTemplate2 = subject2 & "<br />状态：草稿<br />" & " 标题:" & Subject & "<br />内容：<br />" &
                                              EmailBody
                        mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4",
                                      "autoedm@reasonables.com", "自动化草稿", "emailalerter@reasonables.com", subject2,
                                      SpreadTemplate2)
                    Catch ex As Exception
                        Common.LogText(ex.ToString)
                    End Try
                    'add by gary end 20130801
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
                Common.LogText(Subject)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Else
                Common.LogText("No contact in this favorite:" & Favorite)
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function

    'Add By Gary End 2013/07/19

    'send out campaign to the lists stored in table rsscontacts(All)
    Public Shared Function CreateCampaignEverbuying(ByVal RssID As Integer, ByVal Subject As String,
                                                    ByVal LoginName As String, ByVal AppID As String,
                                                    ByVal senderName As String, ByVal SenderEmail As String,
                                                    ByVal EmailBody As String,
                                                    ByVal Contactlist As List(Of RssContactList)) As String
        Dim A As New alerter
        Try

            Dim rssContactLists As New List(Of String)
            For Each rssContactList As RssContactList In Contactlist
                rssContactLists.Add(rssContactList.ContactListName.Trim())
            Next

            'get contact lists from spread service
            Dim mySpread As SpreadService.Service = New SpreadService.Service()

            'create campaign
            'Dim targetSubscription As String() = temp.Split(",")
            Dim intervalTime As Integer = - 1
            Dim campaignId As Integer
            Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

            myCampaign.campaignName = "EverbuyingHomePage" + Format(Now.Date, "yyyyMMdd")

            myCampaign.from = senderName
            myCampaign.fromEmail = SenderEmail
            myCampaign.subject = Subject
            myCampaign.content = EmailBody
            myCampaign.schedule = Now
            'send campaign
            Try
                campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, rssContactLists.ToArray(),
                                                     intervalTime)
                A.UpdateSentLog(RssID, Subject)
                Common.LogText(Subject)
            Catch ex As Exception
                Common.LogText(ex.ToString)
                Common.LogText("retry to createcampaign for everbuying")
                campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, rssContactLists.ToArray(),
                                                     intervalTime)
                A.UpdateSentLog(RssID, Subject)
                Common.LogText(Subject)
            End Try
            mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Return myCampaign.subject
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    'send out campaign which is open but no category.(Open but not Cliked)
    Public Shared Function CreateCampaignForOpenButNoCategory(ByVal CampaignName As String, ByVal Subject As String,
                                                              ByVal LoginName As String, ByVal AppID As String,
                                                              ByVal senderName As String, ByVal SenderEmail As String,
                                                              ByVal EmailBody As String) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Strategy = ChooseStrategy.OpenExcludeCategory
            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0
            'Try
            '    Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
            'Catch ex As Exception
            '    A.LogText(ex.ToString)
            'End Try

            'If Count > 0 Then
            Dim ContactLists() As String = New String() {FavoriteContactsList}
            Dim intervalTime As Integer = - 1
            Dim campaignId As Integer
            Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()
            myCampaign.campaignName = CampaignName
            myCampaign.from = senderName
            myCampaign.fromEmail = SenderEmail
            myCampaign.subject = Subject
            myCampaign.content = EmailBody
            myCampaign.schedule = Now

            Try
                campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            Common.LogText(Subject)
            SentResult = True
            mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            'Else
            'A.LogText("No contact list found")
            'End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function

    'Not open
    Public Shared Function CreateCampaignForNotOpen(ByVal CampaignName As String, ByVal Subject As String,
                                                    ByVal LoginName As String, ByVal AppID As String,
                                                    ByVal senderName As String, ByVal SenderEmail As String,
                                                    ByVal EmailBody As String) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000

            Dim intervalTime As Integer = - 1
            Dim ContactLists0() As String = New String() {CampaignName}
            Dim Campaign0 As SpreadService.Campaign = New SpreadService.Campaign()
            Campaign0.campaignName = CampaignName
            Campaign0.from = senderName
            Campaign0.fromEmail = SenderEmail
            Campaign0.subject = Subject
            Campaign0.content = EmailBody
            Campaign0.schedule = Now

            Dim c1 As Integer
            Try
                c1 = mySpread.createCampaign(LoginName, AppID, Campaign0, ContactLists0, intervalTime)
                Common.LogText(Subject)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, c1, True)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function
End Class
