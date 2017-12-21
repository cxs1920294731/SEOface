Imports Analysis

<DllType("lifeinhk")>
Public Class LifeInHKDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim lifeHK As New Analysis.LifeInHK
        lifeHK.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, Now)
    End Sub
End Class
