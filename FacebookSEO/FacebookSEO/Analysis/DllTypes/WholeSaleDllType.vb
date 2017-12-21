Imports Analysis

<DllType("wholesale")>
Public Class WholeSaleDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim wholesale As New wholesale()
        wholesale.Start(IssuesId, list.SiteId, planType, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.SubjectForNFC, list.UrlSpecialCode)
    End Sub
End Class
