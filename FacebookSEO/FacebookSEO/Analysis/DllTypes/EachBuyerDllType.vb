﻿Imports Analysis

<DllType("eachbuyer")>
Public Class EachBuyerDllType
    Inherits ProductStart

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)
        Dim planType As String = list.PlanType.Trim()
        Dim eachbuyer As New Analysis.Eachbuyer
        eachbuyer.Start(IssuesId, list.SiteId, planType, list.SplitContactList, list.SpreadLoginEmail, list.AppId, list.Url, list.Categories)
    End Sub
End Class