Public Class SentLogs
    Public Shared Sub UpdateSentLog(sentLog As SentLog)
        If sentLog IsNot Nothing Then
            SqlDataProvider.AddSentLog(sentLog)
        End If
    End Sub

    Public Shared Function GetLastSentLogByRssId(rssId As Int32) As SentLog
        If String.IsNullOrEmpty(rssId) Then
            Return Nothing
        End If
        Return CBO.FillObject (Of SentLog)(SqlDataProvider.GetLastSentLogByRssId(rssId))
    End Function
End Class
