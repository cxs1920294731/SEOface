Public Class SplitContactLists
    Public Shared Sub AddContactList(splitContactList As SplitContactList)
        SqlDataProvider.AddContactList(splitContactList)
    End Sub

    Public Shared Sub DeleteSplitContactLists(shopName As String, rssType As String)
        SqlDataProvider.DeleteSplitContactLists(shopName, rssType)
    End Sub

    Public Shared Sub UpdateFlagByName(contactListName As String, flag As Int32)
        SqlDataProvider.UpdateFlagByName(contactListName, flag)
    End Sub

    Public Shared Function GetActiveSplitContactLists(shopName As String, rssType As String) _
        As List(Of SplitContactList)
        Return CBO.FillCollection (Of SplitContactList)(SqlDataProvider.GetActiveSplitContactLists(shopName, rssType))
    End Function
End Class
