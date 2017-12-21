Imports Analysis

<DllType("tbdress")>
Public Class TbDressDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim tbdress As New Tbdress
        tbdress.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId)
    End Sub
End Class
