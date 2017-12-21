Imports Analysis

<DllType("michaelkorsplazas")>
Public Class MichaelKorsplazasDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim MK As New michaelkors
        MK.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.DllType.Trim)

    End Sub
End Class