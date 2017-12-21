Public Class ClickData
    Private _subEmail As String
    Private _clickUrl As String

    Public Property SubEmail() As String
        Get
            Return _subEmail
        End Get
        Set(ByVal value As String)
            _subEmail = value
        End Set
    End Property

    Public Property ClickUrl() As String
        Get
            Return _clickUrl
        End Get
        Set(ByVal value As String)
            _clickUrl = value
        End Set
    End Property
End Class
