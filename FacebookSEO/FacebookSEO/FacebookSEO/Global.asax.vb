Imports System.Web.SessionState

Public Class Global_asax
    Inherits System.Web.HttpApplication

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application is started
    End Sub
    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        Dim url As String = Request.RawUrl
        If (url.Contains("post")) Then
            Dim s() As String
            Dim id As String
            Dim arrUrl As String
            arrUrl = Mid(url, 1, Len(url) - 4)
            s = Split(arrUrl, "/")
            For Each x In s
                id = x.ToString
            Next
            Server.Transfer("/Post.aspx?id=" + id)
        ElseIf (url.Contains("AllShops")) Then
            Dim s() As String
            Dim index As String
            s = Split(url, "/")
            For Each x In s
                index = x.ToString
            Next
            Server.Transfer("/Default.aspx?sitename=AllShops&pageIndex=" + index)
        ElseIf (url.Contains("kw")) Then
            ' /kw/exhibition/32
            Dim s() As String
            Dim kw As String
            s = Split(url, "/")
            For Each x In s
                kw = x.ToString
            Next
            Server.Transfer("/Default.aspx?keyWordId=" + kw)
        ElseIf (url.Contains("ctag")) Then
            Dim s() As String
            Dim cate As String
            s = Split(url, "/")
            For Each x In s
                cate = x.ToString
            Next
            Server.Transfer("/Default.aspx?cateId=" + cate)
        ElseIf (url.Contains("Default")) Then
            Server.Transfer("/Default.aspx")
        Else
            Dim s() As String
            Dim site As String
            s = Split(url, "/")
            For Each x In s
                site = x.ToString
            Next
            Server.Transfer("/Default.aspx?siteid=" + site)
        End If
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

End Class