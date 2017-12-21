Imports Analysis

<DllType("indepwebsite")>
Public Class IndepwebsiteDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim indep As New Analysis.IndependentWebSite()
        indep.Start(IssuesId, list.SiteId, list.SiteName, list.PlanType.Trim, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories, list.SubjectForNFC)
    End Sub
End Class
