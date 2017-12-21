<Serializable()>
Public Class RssSubscription
    Public Sub New()
    End Sub

    Public Property RssID As Int32
    Public Property Url As String
    Public Property SpreadLogin As String
    Public Property StartAt As DateTime
    Public Property EndAt As DateTime
    Public Property IntervalDay As Double

    Public Property WeekDays As String
    Public Property SenderName As String
    Public Property SenderEmail As String
    Public Property Template As String
    Public Property Status As String

    Public Property rssType As String
    Public Property AppID As String
    Public Property Categories As String
    Public Property ExcludeSubject As String
End Class
