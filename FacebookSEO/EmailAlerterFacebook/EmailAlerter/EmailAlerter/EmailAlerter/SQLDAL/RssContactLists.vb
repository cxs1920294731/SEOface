Public Class RssContactLists
    Public Shared Function GetRssContactListByRssId(rssId As String)
        If String.IsNullOrEmpty(rssId) Then
            Return Nothing
        End If
        Return CBO.FillCollection (Of RssContactList)(SqlDataProvider.GetRssContactListByRssId(rssId))
    End Function
End Class
