Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying
Imports System.Linq
Public Class Service1

    Private listLog As New List(Of AutoSentLog)
    Public efContext1 As New FaceBookForSEOEntities()
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' 请在此处添加代码以启动您的服务。此方法应完成设置工作，
        ' 以使您的服务开始工作。
        EventLog.WriteEntry("Start service")
        Common.LogText("Service Start")
        'CheckSchedule()
        '  MainTimer.Interval = ConfigurationManager.AppSettings("Interval")
    End Sub

    Protected Overrides Sub OnStop()
        ' 在此处添加代码以执行任何必要的拆解操作，从而停止您的服务。
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>


    Public Sub New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    Private Sub MainTimer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MainTimer.Elapsed
        Me.MainTimer.Enabled = False
        CheckSchedule()
        Me.MainTimer.Enabled = True
    End Sub

    ''' <summary>
    ''' Check if any new email campaigns to send 
    ''' </summary>
    ''' <remarks></remarks>
    Sub CheckSchedule()

        'Common.LogText("Schedule Start")
        'System.Threading.Thread.Sleep(1000)


        Dim A As New alerter
        Try
            Dim autoTest As New AutoModel()
            autoTest.GetRssSubscriptionAndInsertAndSentEmail(listLog)

            Dim hour As Integer = ConfigurationManager.AppSettings("hourforSitemap").ToString().Trim()
            Dim minute As Integer = ConfigurationManager.AppSettings("minuteforSitemap").ToString().Trim()

            If (DateTime.Now.Hour = hour And DateTime.Now.Minute > minute And DateTime.Now.Minute <= minute + 2) Then
                Common.LogText("sitemap start")
                GetSiteMap.SiteMap()
                Common.LogText("sitemap end")
            End If
        Catch ex As Exception
            Common.LogException(ex)
        End Try
    End Sub


    '''从表AutomationSites复制元素到Categories和AutomationPlan中
    Sub dataMove()
        Try
            Dim ListSites As List(Of AutomationSite) = (From autoSites In efContext1.AutomationSites Select autoSites).ToList
            Common.LogText("开始1")
            For Each item As AutomationSite In ListSites
                Common.LogText(item.SiteUrl.ToString)
                '如果已经加了，则跳过
                If (is_exis(1, item.siteid)) Then
                    Dim issue As New AutomationPlan()
                    issue.SiteID = item.siteid
                    issue.PlanType = "HO"
                    issue.StartAt = Now
                    issue.IntervalDay = 1
                    issue.WeekDays = "1234560"
                    issue.SenderName = item.SiteName
                    issue.SenderEmail = "autoedm@reasonables.com"
                    issue.Status = "A"
                    issue.SplitContactCount = 1
                    issue.URL = item.SiteUrl
                    efContext1.AddToAutomationPlans(issue)
                    efContext1.SaveChanges()
                End If
                If (is_exis(2, item.siteid)) Then
                    Dim cate As New AutoCategory
                    cate.Category1 = "groupbuyer"
                    cate.SiteID = item.siteid
                    cate.LastUpdate = Now
                    cate.Description = "groupbuyer"
                    cate.Url = item.SiteUrl
                    cate.Gender = "N"
                    efContext1.AddToAutoCategories(cate)
                    efContext1.SaveChanges()
                End If
            Next
        Catch ex As Exception
            Common.LogText("出错")
            Common.LogText(ex.ToString)
        End Try
    End Sub
    '检查是否已经存在
    'tableIndex为1是代表为AutomationPlan表，为2时代表Categories表

    Sub AddSiteUrl()
        Try
            Dim listSite As List(Of AutomationSite) = (From a In efContext1.AutomationSites Select a).ToList
            For Each item As AutomationSite In listSite
                Dim queryIssue As AutomationSite = efContext1.AutomationSites.Where(Function(iss) iss.siteid = item.siteid).SingleOrDefault()
                queryIssue.k11SiteUrl = "/" + item.SiteName + "/" + item.siteid.ToString
                Common.LogText(queryIssue.k11SiteUrl)
                efContext1.SaveChanges()
            Next
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Sub


    Public Function is_exis(ByVal tableIndex As Integer, ByVal siteid As Integer) As Boolean
        If (tableIndex = 1) Then
            Dim myAutoplan As AutomationPlan = (From a In efContext1.AutomationPlans
                                                Where a.SiteID = siteid
                                                Select a).FirstOrDefault()
            If (myAutoplan Is Nothing) Then
                Return True
            Else
                Return False
            End If
        End If
        If (tableIndex = 2) Then
            Dim myCate As AutoCategory = (From b In efContext1.AutoCategories Where b.SiteID = siteid Select b).FirstOrDefault()
            If (myCate Is Nothing) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function
    ''' <summary>
    ''' 
    ''' 
    ''' </summary>
    ''' <param name="r"></param>
    ''' <remarks></remarks>
    Sub SendEmail(ByVal r As RssSubscription)
        Dim A As New alerter
        Try
            '龙震天博客
            If r.rssType.Trim() = "B" Then
                Try
                    Dim B As New Blog
                    Dim EmailCampaignElements = B.ReadRSS(r.Url)
                    Dim CustomizedTemplate As String = B.CustomizeTempalate(r.Template, EmailCampaignElements)
                    Dim Subject As String = B.CreateCampaignBlog(r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, CustomizedTemplate, EmailCampaignElements)
                    A.UpdateSentLog(r.RssID, Subject)
                    'A.LogText(Subject)
                    NotificationEmail.SentStartEmail(Subject, CustomizedTemplate)
                Catch ex As Exception
                    'A.LogText(ex.ToString)
                    NotificationEmail.SentErrorEmail(ex.ToString())
                End Try

                EventLog.WriteEntry("succeed to send email")

            ElseIf r.rssType.Trim() = "M" Then
                Dim M As New Ladies
                M.SendEmailMagazine(r)

                'GroupBuyer项目入口
            ElseIf r.rssType.Trim() = "D" Then
                Dim D As New GroupBuyer2
                D.SendEmailDeal(r)

                'E stands for Everbuying
            ElseIf r.rssType.Trim().StartsWith("E") Then
                Dim E As New Everbuying
                Select Case r.rssType.Trim()
                    Case "EA" : E.SendEmailForAll(r)
                    Case "EP" : E.SendPersonalDefaultPageByCategory(r)
                    Case "EC" : E.SendEmailByCategories(r)
                    Case "ED" : E.SendEmailByNoCategory(r)
                    Case "EN" : E.SendNotOpen(r)
                End Select
            End If

        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Sub










End Class
