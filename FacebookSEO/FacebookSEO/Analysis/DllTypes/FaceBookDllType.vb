Imports Analysis

<DllType("facebook")>
Public Class FaceBookDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        If (planType.Contains("HO")) Then
            '获取facebook的内容
            Dim k11artmart As New K11forFBSeo()
            k11artmart.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories)
        ElseIf (planType.Contains("HA")) Then
            Dim weibo As New Weibo()
            weibo.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories)
        End If
    End Sub
End Class
