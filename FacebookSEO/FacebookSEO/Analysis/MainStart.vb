''' <summary>
''' dll总入口，各个不同的自动化在此开始分流
''' author：DoraAO
''' create time;20140628
''' 自动化站点逐渐增多，如果全部在emailAlerter中分流有诸多弊端：1每新增一个站点都需同步更新emailAlerter ;2站点增多emailAlerter中的分流琐碎代码会逐渐庞大而影响简洁性可读性
''' 为dll提供一个总入口，可以解决上述弊端
''' </summary>
''' <remarks></remarks>

Public Class MainStart
    Public Sub Start(ByVal dllType As String, ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String,
                ByVal subject As String, ByVal productCount As Integer, ByVal utmCode As String)

        'EFHelper.UpdateFbToken()

        Select Case dllType
            Case "ladykingdom"
                Dim ladykingdom As New LadyKingdom
                ladykingdom.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "seasonwind"
                Dim seasonwind As New SeasonWind
                seasonwind.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "michaelkorsplazas"
                Dim MK As New michaelkors
                MK.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories, dllType.Trim)
            Case "louisvuitton"
                Dim louisvuitton As New louisvuitton
                louisvuitton.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories, dllType)
            Case "tbdress"
                Dim tbdress As New Tbdress
                tbdress.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId)
            Case "nollmet" 'here
                Dim nollmet As New Nollmet
                nollmet.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "plemix"
                Dim plemix As New Plemix
                plemix.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "travelexpert"
                Dim travelExpert As New TravelExpert
                travelExpert.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "lining"
                Dim lining As New LiNing
                lining.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "facebook"
                If (planType.Contains("HO")) Then
                    '获取facebook的内容
                    Dim k11artmart As New K11forFBSeo()
                    k11artmart.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
                ElseIf (planType.Contains("HA")) Then
                    Dim weibo As New Weibo()
                    weibo.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
                End If
            Case "jiacheng"
                Dim jiacheng As New jiacheng()
                jiacheng.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "taofen8"
                Dim taofen As New Taofen8()
                taofen.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
                'Case "mailego"
                '    Dim mailego As New MaileGo()
                '    mailego.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "ailiti"
                Dim ailiti As New Ailiti()
                ailiti.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories)
            Case "aotu"
                Dim aotu As New Aotu()
                aotu.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories, subject, productCount)
            Case "sozmall"
                Dim sozmall As New SozMall()
                sozmall.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories, subject)
            Case "couppie"
                Dim coup As New Couppie()
                coup.Start(IssueID, siteId, planType, splitContactCount, spreadLogin, appId, url, categories, subject)
            Case "wholesale"
                Dim wholesale As New wholesale()
                wholesale.Start(IssueID, siteId, planType, spreadLogin, appId, url, categories, subject, utmCode)
        End Select
    End Sub

End Class
