Public Class RssSubscriptions
    Public Shared Function GetActiveRss() As List(Of RssSubscription)
        Return CBO.FillCollection(Of RssSubscription)(SqlDataProvider.GetActiveRss())
    End Function

    Public Shared Function GetRssSubscriptionByRssId(rssId As Int32) As RssSubscription
        If String.IsNullOrEmpty(rssId) Then
            Return Nothing
        End If
        Return CBO.FillObject(Of RssSubscription)(SqlDataProvider.GetRssSubscriptionByRssId(rssId))
    End Function
End Class
