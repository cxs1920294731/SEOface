Imports Microsoft.ApplicationBlocks.Data

Public Class SqlDataProvider
    Private Shared ReadOnly ConnString As String = Config.GetConnectionString("EmailAlerterConnectionString")

#Region "RssSubscription"

    Public Shared Function GetRssSubscriptionByRssId(rssId As Int32) As IDataReader
        Return SqlHelper.ExecuteReader(ConnString, "GetRssSubscriptionByRssId", rssId)
    End Function

    Public Shared Function GetActiveRss() As IDataReader
        Return SqlHelper.ExecuteReader(ConnString, "GetActiveRss")
    End Function

#End Region

#Region "RssContactList"

    Public Shared Sub AddRssContactList(rssSub As RssSubscription)
        SqlHelper.ExecuteNonQuery(ConnString, "RssContactListADD", rssSub.RssID)
    End Sub

    Public Shared Function GetRssContactListByRssId(rssId As Int32) As IDataReader
        Return SqlHelper.ExecuteReader(ConnString, "GetRssContactListByRssId", rssId)
    End Function

#End Region

#Region "SentLog"

    Public Shared Sub AddSentLog(sentLog As SentLog)
        SqlHelper.ExecuteNonQuery(ConnString, "AddSentLog", sentLog.RssId, sentLog.Subject, sentLog.LastSentAt)
    End Sub

    Public Shared Function GetLastSentLogByRssId(rssId As String) As IDataReader
        Return SqlHelper.ExecuteReader(ConnString, "GetLastSentLogByRssId", rssId)
    End Function

#End Region

#Region "SplitContactList"

    Public Shared Sub AddContactList(splitContactList As SplitContactList)
        SqlHelper.ExecuteNonQuery(ConnString, "AddSplitContactList", splitContactList.ShopName,
                                  splitContactList.ContactListName, splitContactList.RssType, splitContactList.Flag)
    End Sub

    Public Shared Sub DeleteSplitContactLists(shopName As String, rssType As String)
        SqlHelper.ExecuteNonQuery(ConnString, "DeleteListByNameNType", shopName, rssType)
    End Sub

    Public Shared Sub UpdateFlagByName(contactListName As String, flag As Int32)
        SqlHelper.ExecuteNonQuery(ConnString, "UpdateFlagByName", contactListName, flag)
    End Sub

    Public Shared Function GetActiveSplitContactLists(shopName As String, rssType As String)
        Return SqlHelper.ExecuteReader(ConnString, "GetActiveSplitLists", shopName, rssType)
    End Function

#End Region
End Class
