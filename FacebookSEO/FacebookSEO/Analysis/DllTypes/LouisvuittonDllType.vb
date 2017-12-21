Imports Analysis

<DllType("louisvuitton")>
Public Class LouisvuittonDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim louisvuitton As New louisvuitton
        louisvuitton.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.DllType)
    End Sub
End Class
