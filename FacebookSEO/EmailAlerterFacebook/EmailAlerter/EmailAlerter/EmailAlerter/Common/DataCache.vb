Imports System
Imports System.Web
Imports System.Web.Caching


Public Class DataCache
    Private Shared _cache As Cache
    Private Shared ReadOnly _cachePrefix As String = "{0}"

    Public Shared ReadOnly Property CacheObj() As Cache
        Get
            If _cache Is Nothing Then
                _cache = HttpRuntime.Cache
            End If
            Return _cache
        End Get
    End Property

    Private Shared Function GetCacheKey(ByVal CacheKey As String) As String
        If String.IsNullOrEmpty(CacheKey) Then
            Throw New ArgumentException("Argument cannot be null or an empty string", CacheKey)
        End If
        Return String.Format(_cachePrefix, CacheKey)
    End Function

    Public Shared Function GetCache(ByVal CacheKey As String) As Object
        Return CacheObj.Get(GetCacheKey(CacheKey))
    End Function

    Public Shared Sub SetCache(scope As CacheScope, ByVal CacheKey As String, data As Object)
        Dim ExpirationTime As DateTime

        Select Case scope
            Case CacheScope.Forever
                ExpirationTime = DateTime.MaxValue
            Case CacheScope.OneHour
                ExpirationTime = DateTime.Now.AddHours(1)
        End Select

        CacheObj.Insert(GetCacheKey(CacheKey), data, Nothing, ExpirationTime, Nothing)
    End Sub

    Public Shared Sub SetCache(ByVal CacheKey As String, objMap As ObjectMappingInfo)
        ' 1 Hour
        SetCache(CacheScope.Forever, CacheKey, objMap)
    End Sub

    Public Shared Sub ClearCache(key As String)
        'Dim toRemove As New List(Of String)()
        'For Each cacheItem As DictionaryEntry In HttpRuntime.Cache
        '    toRemove.Add(cacheItem.Key.ToString())
        'Next

        'For Each key As String In toRemove
        '    HttpRuntime.Cache.Remove(key)
        'Next
        HttpRuntime.Cache.Remove(GetCacheKey(key))
    End Sub
End Class
