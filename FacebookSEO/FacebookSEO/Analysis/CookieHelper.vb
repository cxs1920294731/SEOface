Imports System.Web

Public Class CookieHelper
    ''' <summary>
    ''' create a cookie and set value
    ''' </summary>
    ''' <param name="cookieName"></param>
    ''' <param name="cookieVal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SetCookies(ByVal cookieName As String, ByVal cookieVal As String) As Boolean
        Try
            Dim newCookie As New HttpCookie(cookieName)
            newCookie.Value = cookieVal
            newCookie.Expires = DateTime.Now.AddHours(1)
            System.Web.HttpContext.Current.Response.Cookies.Add(newCookie)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取某个特定Cookie的值
    ''' </summary>
    ''' <param name="cookieName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCookies(ByVal cookieName As String) As String
        Dim cookie As HttpCookie = HttpContext.Current.Request.Cookies(cookieName)
        If (cookie Is Nothing) Then
            Return Nothing
        Else
            Return cookie.Value.ToString()
        End If
    End Function

    ''' <summary>
    ''' 清除某个cookie
    ''' </summary>
    ''' <param name="cookieName"></param>
    ''' <remarks></remarks>
    Public Shared Sub ClearCookies(ByVal cookieName As String)
        If Not (HttpContext.Current.Request.Cookies(cookieName) Is Nothing) Then
            Dim myCookie As HttpCookie = New HttpCookie(cookieName)
            myCookie.Expires = DateTime.Now.AddDays(-1)
            HttpContext.Current.Response.Cookies.Add(myCookie)
        End If
    End Sub
End Class
