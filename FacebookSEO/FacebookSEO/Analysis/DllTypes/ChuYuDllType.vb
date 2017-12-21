Imports Analysis

<DllType("chuyu")>
Public Class ChuYuDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim chuyu As New Analysis.ToYouth
        chuyu.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories)
    End Sub
End Class
