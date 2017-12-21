Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying

Public Class Everbuying
    Private ReadOnly e As New EDM4Everbuying.EverbuyDemo
    Private ReadOnly A As New alerter

    '获取maindemo 并且发送给rsscontactlists 表里的lists-- 发送所有
    Public Function SendEmailForAll(ByVal r As RssSubscription)

        Dim Contactlist As List(Of RssContactList) = RssContactLists.GetRssContactListByRssId(r.RssID)
        If Contactlist IsNot Nothing AndAlso Contactlist.Count > 0 Then
            Dim html As String = ""
            Dim subject As String = ""

            html = e.GetMainPageDemo()
            subject = "Hi [FIRSTNAME]," & e.GetMainPageTitle()

            CreateCampaign.CreateCampaignEverbuying(r.RssID, subject, r.SpreadLogin, r.AppID, r.SenderName,
                                                    r.SenderEmail, html, Contactlist)
        End If
    End Function

    '按照rsssubscriptions表的category字段创建campaign发送相应的category邮件
    Public Function SendEmailByCategories(ByVal r As RssSubscription)

        Dim categories As String() = r.Categories.Split(",")
        For Each category As String In categories
            Try
                Dim html As String = ""
                Dim subject As String = ""
                html = e.GetCategoryPageDemo(category)
                'GetEmailDemo()
                subject = category.ToUpper & " NEW ARRIVALS:" & e.GetCategoryPageTitle(category)
                'GetTitle()

                Dim CampaignName As String = "EverbuyingCategoryNewArrival" +
                                             category.Replace("'", "").Replace(" ", "").Replace("&", "") +
                                             Format(Now.Date, "yyyyMMdd").ToString()

                'Add By Gary Start 2013/06/09
                '调整为根据上一周的点击记录发送分类邮件
                Dim startDate As String = Date.Now.AddDays(- Date.Now.DayOfWeek - 14).ToString("yyyy-MM-dd")
                Dim endDate As String = Date.Now.AddDays(- Date.Now.DayOfWeek).ToString("yyyy-MM-dd")

                'modified by gary 2013/07/19
                'Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignByCategoryByClickDateRange(category, CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, html, startDate, endDate)
                '2013/10/31 modified,change the status of campaign
                Dim status As String = "draft"
                'draft 'waiting
                Dim SentSuccess As Boolean =
                        CreateCampaign.CreateCampaignByCampaignStatusAndCategoryByClickDateRange(category, CampaignName,
                                                                                                 subject, r.SpreadLogin,
                                                                                                 r.AppID, r.SenderName,
                                                                                                 r.SenderEmail, html,
                                                                                                 status, startDate,
                                                                                                 endDate)
                'Add By Gary End 2013/06/09
                'Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignByCategory(category, CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, html)
                If SentSuccess = True Then
                    A.UpdateSentLog(r.RssID, subject)
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
        Next
    End Function

    '发送个性化主页模板
    Public Function SendPersonalDefaultPageByCategory(ByVal r As RssSubscription)
        Dim categories As String() = r.Categories.Split(",")
        For Each category As String In categories
            Try
                Dim html As String = ""
                Dim subject As String = ""
                html = e.GetPersonalDefaultPageDemo(category)
                'GetEmailDemo()
                subject = "Hi [FIRSTNAME]," & e.GetPersonalDefaultPageTitle(category)
                'GetTitle()

                Dim CampaignName As String = "EverbuyingPersonality" +
                                             category.Replace("'", "").Replace(" ", "").Replace("&", "") +
                                             Format(Now.Date, "yyyyMMdd").ToString()

                'Add By Gary Start 2013/06/09
                '调整为根据上一周的点击记录发送分类邮件
                Dim startDate As String = Date.Now.AddDays(- Date.Now.DayOfWeek - 14).ToString("yyyy-MM-dd")
                Dim endDate As String = Date.Now.AddDays(- Date.Now.DayOfWeek).ToString("yyyy-MM-dd")
                Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignByCategoryByClickDateRange(category,
                                                                                                     CampaignName,
                                                                                                     subject,
                                                                                                     r.SpreadLogin,
                                                                                                     r.AppID,
                                                                                                     r.SenderName,
                                                                                                     r.SenderEmail, html,
                                                                                                     startDate, endDate)
                'Add By Gary End 2013/06/09

                'Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignByCategory(category, CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, html)
                If SentSuccess = True Then
                    A.UpdateSentLog(r.RssID, subject)
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
        Next
    End Function

    '发送打开过但没有类别记录的联系名单
    Public Function SendEmailByNoCategory(ByVal r As RssSubscription)

        Try
            '获得点击率最高的分类名称
            Dim iws As New EmailAlerter.IntrawWebService.Service
            Dim category As String =
                    iws.dsGetCategoryTrackings(r.SpreadLogin, r.AppID, Now.Date.AddDays(- 30), Now.Date).Tables(0).Rows(
                        0)(0).ToString()


            Dim html As String = ""
            Dim subject As String = ""
            html = e.GetPersonalDefaultPageDemo(category)
            'GetEmailDemo()
            subject = "Hi [FIRSTNAME]," & e.GetPersonalDefaultPageTitle(category)
            'GetTitle()


            '分list处理逻辑，将open list分成四个list发送
            Dim splitlists As List(Of SplitContactList) = SplitContactLists.GetActiveSplitContactLists(r.SenderName,
                                                                                                       r.rssType.Trim())
            If splitlists.Count > 0 Then
                Dim campaignName As String = splitlists(0).ContactListName
                Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignForOpenButNoCategory(campaignName, subject,
                                                                                               r.SpreadLogin, r.AppID,
                                                                                               r.SenderName,
                                                                                               r.SenderEmail, html)
                If SentSuccess = True Then
                    A.UpdateSentLog(r.RssID, subject)
                    SplitContactLists.UpdateFlagByName(campaignName, 1)
                End If
            Else
                'delete all inactive contact list
                SplitContactLists.DeleteSplitContactLists(r.SenderName, r.rssType.Trim())

                'create new contact list
                Dim mySpread As SpreadService.Service = New SpreadService.Service()
                mySpread.Timeout = 1200000

                Dim SplitlistNames As List(Of String) = New List(Of String)()
                'Format CampaignName as SenderName + RssType + DateTime + SeqNo
                Dim prefix = r.SenderName + "OpenButNoCate" + r.rssType.Trim() +
                             Format(Now, "yyyyMMddHHmmss").ToString()
                SplitlistNames.Add(prefix + "1")
                SplitlistNames.Add(prefix + "2")
                SplitlistNames.Add(prefix + "3")
                SplitlistNames.Add(prefix + "4")
                Dim FavoriteContactsList As String = SplitlistNames(0) + ";" + SplitlistNames(1) + ";" +
                                                     SplitlistNames(2) + ";" + SplitlistNames(3)
                'Dim FavoriteContactsList As String = SplitlistNames(0) + ";" + SplitlistNames(1)
                Dim QuerySubscriber As New QuerySubscriber
                QuerySubscriber.Strategy = ChooseStrategy.OpenExcludeCategory
                'add by Gary Start,调整为只发送最后半年Open no click的list
                QuerySubscriber.StartDate = Date.Now.AddMonths(- 6).ToString("yyyy-MM-dd")
                'add by Gary End
                QuerySubscriber.CountryList = New String() {}
                Dim CriteriaString As String = QuerySubscriber.ToJsonString
                Dim Count As Integer = 0
                Try
                    Count = mySpread.SearchContacts(r.SpreadLogin, r.AppID, CriteriaString, Integer.MaxValue,
                                                    FavoriteContactsList, True)
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
                If Count > 0 Then
                    'insert split contact lists into table
                    For Each splitlistName As String In SplitlistNames
                        Dim sc As New SplitContactList
                        sc.ShopName = r.SenderName
                        sc.ContactListName = splitlistName
                        sc.RssType = r.rssType.Trim()
                        sc.Flag = 0
                        '0 stands for active, 1 stands for inactive
                        SplitContactLists.AddContactList(sc)
                    Next
                    Dim campaignName As String = SplitlistNames(0)
                    Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignForOpenButNoCategory(campaignName, subject,
                                                                                                   r.SpreadLogin,
                                                                                                   r.AppID, r.SenderName,
                                                                                                   r.SenderEmail, html)
                    If SentSuccess = True Then
                        A.UpdateSentLog(r.RssID, subject)
                        SplitContactLists.UpdateFlagByName(campaignName, 1)
                    End If
                End If
            End If

            'update by jake 2013/5/24
            'Format CampaignName by SenderName + RssType + DateTime
            'Dim CampaignName As String = r.SenderName + "OpenButNoCate" + r.rssType.Trim() + Format(Now, "yyyyMMddHHmmss")

            'Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignForOpenButNoCategory(CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, html)
            'If SentSuccess = True Then
            '    A.UpdateSentLog(r.RssID, subject)
            'End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    '发送Not open list
    Public Function SendNotOpen(ByVal r As RssSubscription)
        Try

            '获得点击率最高的分类名称
            Dim iws As New EmailAlerter.IntrawWebService.Service
            Dim category As String =
                    iws.dsGetCategoryTrackings(r.SpreadLogin, r.AppID, Now.Date.AddDays(- 30), Now.Date).Tables(0).Rows(
                        0)(0).ToString()

            '获取模板Html
            Dim html As String = ""
            Dim subject As String = ""
            html = e.GetPersonalDefaultPageDemo(category)
            'GetEmailDemo()
            subject = "Hi [FIRSTNAME]," & e.GetPersonalDefaultPageTitle(category)
            'GetTitle()

            '分list处理逻辑，将Not open list分成四个list发送
            Dim splitlists As List(Of SplitContactList) = SplitContactLists.GetActiveSplitContactLists(r.SenderName,
                                                                                                       r.rssType.Trim())
            If splitlists.Count > 0 Then
                Dim campaignName As String = splitlists(0).ContactListName
                Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignForNotOpen(campaignName, subject,
                                                                                     r.SpreadLogin, r.AppID,
                                                                                     r.SenderName, r.SenderEmail, html)
                If SentSuccess = True Then
                    A.UpdateSentLog(r.RssID, subject)
                    SplitContactLists.UpdateFlagByName(campaignName, 1)
                End If
            Else
                'delete all inactive contact list
                SplitContactLists.DeleteSplitContactLists(r.SenderName, r.rssType.Trim())

                'create new contact list
                Dim mySpread As SpreadService.Service = New SpreadService.Service()
                mySpread.Timeout = 1200000

                Dim SplitlistNames As List(Of String) = New List(Of String)()
                'Format CampaignName as SenderName + RssType + DateTime + SeqNo
                Dim prefix = r.SenderName + "NotOpen" + r.rssType.Trim() + Format(Now, "yyyyMMddHHmmss").ToString()
                SplitlistNames.Add(prefix + "1")
                SplitlistNames.Add(prefix + "2")
                SplitlistNames.Add(prefix + "3")
                SplitlistNames.Add(prefix + "4")
                Dim FavoriteContactsList As String = SplitlistNames(0) + ";" + SplitlistNames(1) + ";" +
                                                     SplitlistNames(2) + ";" + SplitlistNames(3)
                Dim QuerySubscriber As New QuerySubscriber
                QuerySubscriber.Strategy = ChooseStrategy.NotOpen
                QuerySubscriber.CountryList = New String() {}
                Dim CriteriaString As String = QuerySubscriber.ToJsonString
                Dim Count As Integer = 0
                Try
                    Count = mySpread.SearchContacts(r.SpreadLogin, r.AppID, CriteriaString, Integer.MaxValue,
                                                    FavoriteContactsList, True)
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
                If Count > 0 Then
                    'insert split contact lists into table
                    For Each splitlistName As String In SplitlistNames
                        Dim sc As New SplitContactList
                        sc.ShopName = r.SenderName
                        sc.ContactListName = splitlistName
                        sc.RssType = r.rssType.Trim()
                        sc.Flag = 0
                        '0 stands for active, 1 stands for inactive
                        SplitContactLists.AddContactList(sc)
                    Next
                    Dim campaignName As String = SplitlistNames(0)
                    Dim SentSuccess As Boolean = CreateCampaign.CreateCampaignForNotOpen(campaignName, subject,
                                                                                         r.SpreadLogin, r.AppID,
                                                                                         r.SenderName, r.SenderEmail,
                                                                                         html)
                    If SentSuccess = True Then
                        A.UpdateSentLog(r.RssID, subject)
                        SplitContactLists.UpdateFlagByName(campaignName, 1)
                    End If
                End If
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log",
                                         ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function
End Class
