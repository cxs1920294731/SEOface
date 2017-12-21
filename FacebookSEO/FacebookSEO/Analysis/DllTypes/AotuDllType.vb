Imports Analysis

<DllType("aotu")>
Public Class AotuDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim aotu As New Aotu()
        aotu.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.SubjectForNFC, list.DisplayCount)
    End Sub
End Class