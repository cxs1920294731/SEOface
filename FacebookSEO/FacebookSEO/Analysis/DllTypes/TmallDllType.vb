Imports Analysis

<DllType("tmall", "taobao")>
Public Class TmallDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Analysis.Tmall.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.SiteUrl, list.Categories, list.SiteName,
                                                                         list.DllType, list.DisplayCount, list.BannerFromUrl, list.BannerRegex, list.BannerIndex, list.SubjectForNFC)
    End Sub
End Class
