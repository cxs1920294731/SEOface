Public Class UrlValid
    Public Shared Function getUrlValid(ByVal param As String) As String
        If (String.IsNullOrEmpty(param)) Then
            Return ""
        End If
        param = param.Replace("&", "").Replace("'", "").Replace("+", "").Replace("<", "").Replace(">", "")
        param = param.Replace("*", "").Replace("%", "").Replace(":", "").Replace("\", "").Replace("?", "").Replace("#", "")
        param = param.Trim
        If (param.Length > 100) Then
            param = param.Substring(0, 100)
        End If
        If (param.EndsWith(".")) Then
            param = param.Substring(0, param.Length - 1)
        End If
        param = param.Replace(" ", "-")
        Return param
    End Function
End Class
