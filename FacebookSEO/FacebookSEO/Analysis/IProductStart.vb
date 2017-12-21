Imports Analysis

Public Interface IProductStart

    Sub Start(list As Subscriptions)

End Interface

Public Class DllTypeAttribute
    Inherits Attribute

    Public Sub New()

    End Sub

    Public Sub New(ParamArray dllType As String())
        Me.DllType = dllType
    End Sub

    Public Property DllType As String()

End Class

Public MustInherit Class ProductStart
    Implements IProductStart

    Public Property IssuesId As Integer
    Protected Sub New(issues As Integer)
        Me.IssuesId = issues
    End Sub

    Public MustOverride Sub Start(list As Subscriptions) Implements IProductStart.Start
End Class
