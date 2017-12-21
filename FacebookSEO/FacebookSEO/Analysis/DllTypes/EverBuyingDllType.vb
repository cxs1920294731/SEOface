Imports Analysis

<DllType("everbuying")>
Public Class EverBuyingDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        If (String.IsNullOrEmpty(list.Categories)) Then '非分类促发
            Analysis.Everbuying.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url)
        Else
            Analysis.EverBuyingPersonalization.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories)
        End If
    End Sub
End Class
