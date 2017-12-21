Imports System.Text.RegularExpressions


Public NotInheritable Class UrlUtils
    Public Shared Function AddHttpPrefix(url As String, enableSSL As Boolean) As String
        If String.IsNullOrEmpty(url) Then Return url
        If Not url.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
            url = CStr(IIf(enableSSL, "https://", "http://")) & url
        End If
        Return url
    End Function

    Public Shared Function AddEndSlash(url As String) As String
        If String.IsNullOrEmpty(url) Then Return url
        If Not url.EndsWith("/") Then
            Return url & "/"
        End If
        Return url
    End Function

    Public Shared Function StandardizeUrl(url As String) As String
        Return AddEndSlash(AddHttpPrefix(url, False))
    End Function
End Class

