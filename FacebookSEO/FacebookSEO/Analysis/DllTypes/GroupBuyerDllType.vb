Imports Analysis

<DllType("groupbuyer")>
Public Class GroupBuyerDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        If (list.PlanType.Trim = "HO2" OrElse list.PlanType.Trim = "HO3" OrElse list.PlanType.Trim = "HO4") Then
            Dim indep As New Analysis.IndependentWebSite()
            indep.Start(IssuesId, list.SiteId, list.SiteName, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.SubjectForNFC)
        Else
            Analysis.GroupBuyer.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, Now)
        End If
    End Sub
End Class
