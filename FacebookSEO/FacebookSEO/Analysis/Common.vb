Imports System.Threading
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Reflection

''' <summary>
''' 调用AliexpressApi过程中需要使用到的公共方法
''' author：DoraAo
''' create time：20140703
''' 可以剥离出来的公共方法请放置入此类中；
''' </summary>
''' <remarks></remarks>

Public Class Common
    Shared LogCache As New ArrayList
    Const N As Integer = 1


    Public Function GetHtmlString(ByVal pageUrl As String, ByVal timeOut As Integer) As String
        Dim resultString As String = ""
        Try
            resultString = LodeWeb(pageUrl, timeOut)
        Catch ex As Exception
            Thread.Sleep(10000)
            resultString = LodeWeb(pageUrl, timeOut)
        End Try
        Return resultString
    End Function

    Public Function GetHtmlString(ByVal pageUrl As String, ByVal timeOut As Integer, ByVal postData As String) As String
        Dim resultString As String = ""
        Try
            resultString = LodeWeb(pageUrl, timeOut, postData)
        Catch ex As Exception
            Thread.Sleep(10000)
            resultString = LodeWeb(pageUrl, timeOut, postData)
        End Try
        Return resultString
    End Function

    Private Function LodeWeb(ByVal pageUrl As String, ByVal timeOut As Integer) As String
        Dim ressting As String = ""

        Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        request.Timeout = timeOut
        request.Headers.Add("Accept-Language", "en-US,en;q=0.8")
        'request.Referer = "https://www.facebook.com/LadyKingdomHK/info"
        request.UserAgent =
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
        request.Method = "post"
        request.AllowAutoRedirect = True
        'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
        Dim response As WebResponse = request.GetResponse()
        'WebRequest.GetResponse方法，返回对 Internet 请求的响应
        Dim resStream As Stream = response.GetResponseStream()

        Dim resStreamReader As StreamReader = New StreamReader(resStream, Encoding.GetEncoding("utf-8"))
        ressting = resStreamReader.ReadToEnd()
        Return ressting
    End Function

    ''' <summary>
    ''' 用post方法实现web request。可设定postdata
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="timeOut"></param>
    ''' <param name="PostData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LodeWeb(ByVal URL As String, ByVal timeOut As Integer, ByVal PostData As String) As String
        Dim result As String = ""
        Dim request As HttpWebRequest
        request = WebRequest.Create(URL)
        request.AllowAutoRedirect = True
        request.Method = "POST"
        request.UserAgent =
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36"
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = Encoding.UTF8.GetByteCount(PostData)
        request.Timeout = timeOut

        Dim request_stream As Stream = request.GetRequestStream()
        Dim requestWriter As StreamWriter = New StreamWriter(request_stream, Encoding.GetEncoding("gb2312"))
        requestWriter.Write(PostData)
        requestWriter.Close()

        Dim httpWebResponse As HttpWebResponse = request.GetResponse()
        Dim responseStream As Stream = httpWebResponse.GetResponseStream()
        Dim responseReader As StreamReader = New StreamReader(responseStream, Encoding.GetEncoding("utf-8"))
        result = responseReader.ReadToEnd()
        responseReader.Close()
        responseStream.Close()
        Return result
    End Function

    ''' <summary>
    ''' 根据提供的url生成API签名或参数签名，如url中包含“param2”，则生成的是API签名，否则是参数签名
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Signature(ByVal url As String) As String

        Dim para As String = url.Split("?")(1).Trim
        Dim paras() As String = para.Replace("=", "").Split("&")
        Array.Sort(paras)
        para = ""
        For Each p In paras
            para = para & p.Trim()
        Next
        If (url.Contains("param2")) Then
            para = url.Substring(url.IndexOf("param2"), url.IndexOf("?") - url.IndexOf("param2")) & para
        End If
        Dim appSecret = "nJWM_z~POdR"
        Dim key() As Byte = Encoding.UTF8.GetBytes(appSecret)
        Dim myhmacsha As New HMACSHA1(key)
        Dim myHashCode As Byte() = myhmacsha.ComputeHash(Encoding.UTF8.GetBytes(para))
        Dim hmac As String = ""
        For Each b In myHashCode
            hmac = hmac & b.ToString("X2")
        Next
        Return hmac
    End Function

    Public Shared Sub LogText(ByVal LogText As String)
        ' Cached write log
        LogText = String.Format("At {0} ({1}), {2}", Now.ToString, LogCache.Count, LogText & ControlChars.NewLine)
        LogCache.Add(LogText)
        If LogCache.Count >= N Then 'cannot load configuration manager
            Dim Folder As String = ""
            Folder = System.AppDomain.CurrentDomain.BaseDirectory.ToString + "App_Data\Send"
            If Not IO.Directory.Exists(Folder) Then IO.Directory.CreateDirectory(Folder)
            Dim LogFile As String = Folder & "\" & Now.Year & "-" & Now.Month & "-" & Now.Day & ".log"
            SyncLock (LogCache.SyncRoot)
                Try
                    System.IO.File.AppendAllText(LogFile, String.Join("", CType(LogCache.ToArray(System.Type.GetType("System.String")), String())))
                    LogCache.Clear()
                Catch ex As Exception 'ignore
                    System.IO.File.AppendAllText(LogFile, "Error, cache size = " & LogCache.Count)
                End Try
            End SyncLock
        End If
    End Sub

    Public Shared Sub LogException(ByVal ex As Exception)
        Dim LogMsg As String = ex.Message & ControlChars.NewLine & ex.StackTrace & ControlChars.NewLine
        LogText(LogMsg)
    End Sub


    Public Shared Function IsUrl(str_url As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(str_url, "http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?")
    End Function



    ''' <summary>
    ''' 补全链接 ，产品herf属性的域名部分
    ''' </summary>
    ''' <param name="domin">应去除末尾/,以防重复</param>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function addDominForUrl(ByVal domin As String, ByVal url As String) As String
        If Not (url.ToLower.Contains("http")) Then
            While (domin.Trim.EndsWith("/"))
                domin = domin.Substring(0, domin.Length - 1)
            End While
            If (url.Trim().Trim.StartsWith("/")) Then
                url = domin & url
            Else
                url = domin & "/" & url
            End If
        End If
        Return url
    End Function


    ''' <summary>
    ''' 为请求的url添加query参数
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="param"></param>
    ''' <returns></returns>
    Public Shared Function addParamForUrl(ByVal url As String, ByVal param As String) As String
        If Not (String.IsNullOrEmpty(param)) Then
            If (param.StartsWith("?") OrElse param.StartsWith("&")) Then
                param = param.Substring(1)
            End If
            If (param.StartsWith("/")) Then '直接添加
                url = url & param
            ElseIf (url.Contains("?")) Then
                url = url & "&" & param
            Else
                url = url & "?" & param
            End If
        End If
        Return url
    End Function

End Class
