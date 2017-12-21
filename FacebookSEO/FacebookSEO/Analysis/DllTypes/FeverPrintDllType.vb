Imports Analysis

<DllType("feverprint")>
Public Class FeverPrintDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Analysis.FeverPrint.Start(IssuesId, list.SiteId, list.PlanType.Trim, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url)
    End Sub
End Class
