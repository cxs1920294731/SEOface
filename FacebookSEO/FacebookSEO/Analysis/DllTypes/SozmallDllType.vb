Imports Analysis

<DllType("sozmall")>
Public Class SozmallDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim sozmall As New SozMall()
        sozmall.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.SubjectForNFC)
    End Sub
End Class
