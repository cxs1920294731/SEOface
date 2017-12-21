Imports System.IO
Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Reflection
'Imports System.Windows.Forms
'Imports NHtmlUnit


Public Class EFHelper

    Private Shared efContext As New EmailAlerterEntities()

#Region "页面信息"
    ''' <summary>
    ''' 根据URL获取网页的HtmlDocument
    ''' </summary>
    ''' <param name="url">网页的URL</param>
    ''' <param name="iTimeOut">设置网页加载的时间</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetHtmlDocument(ByVal url As String, ByVal iTimeOut As Integer, Optional ByVal dg As Integer = 3) As HtmlDocument
        Try

            Dim myWebRequest As HttpWebRequest = System.Net.WebRequest.Create(url)
            If (url.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
                ServicePointManager.Expect100Continue = True
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls '; //SSL3协议替换成TLS协议
                ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
                myWebRequest.ProtocolVersion = HttpVersion.Version10
            End If
            Dim myWebResponse As System.Net.WebResponse
            '如果网站速度加载慢，则设置较长的timeout，并设置加载2次
            Try
                myWebRequest.Timeout = iTimeOut
                myWebResponse = myWebRequest.GetResponse()
            Catch ex As Exception
                myWebRequest.Timeout = 120000
                myWebResponse = myWebRequest.GetResponse()
            End Try
            Dim receiveStream As Stream = myWebResponse.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
            Dim hd As New HtmlDocument()
            hd.Load(receiveStream, encode)
            LogText("递归次数dg=" & dg)
            Return hd
        Catch ex As Exception
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlDocument(url, iTimeOut, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function



    Public Function GetHtmlDocument(ByVal pageUrl As String, ByVal cookie As String, ByVal refer As String, ByVal pageEncoding As String, Optional ByVal dg As Integer = 3) As HtmlDocument
        Try
            cookie = cookie.Trim
            pageEncoding = pageEncoding.Trim
            refer = refer.Trim
            Dim myhtmlDoc As New HtmlDocument()
            Dim resultSting As String
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If

#If DEBUG Then
            '+debug# proxy
            'Dim myProxy As New WebProxy("127.0.0.1:1080", True)
            'request.Proxy = myProxy
#End If

            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = refer
            request.Headers.Add("Cookie", cookie)
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim myencoding As Encoding
            If (String.IsNullOrEmpty(pageEncoding)) Then
                myencoding = Encoding.UTF8
            Else
                myencoding = Encoding.GetEncoding(pageEncoding)
            End If
            Dim resStreamReader As StreamReader = New StreamReader(resStream, myencoding)
            resultSting = resStreamReader.ReadToEnd()
            myhtmlDoc.LoadHtml(resultSting)
            LogText("递归次数dg=" & dg)
            Return myhtmlDoc
        Catch ex As Exception
            LogText("GetHtmlDocument()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlDocument(pageUrl, cookie, refer, pageEncoding, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function

    Public Function GetHtmlDocument2(ByVal url As String, ByVal iTimeOut As Integer) As HtmlDocument
        Try
            Dim myWebRequest As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            Dim myWebResponse As HttpWebResponse
            '如果网站速度加载慢，则设置较长的timeout，并设置加载2次
            Try
                myWebRequest.Timeout = iTimeOut
                myWebRequest.Method = "GET"
                myWebRequest.AutomaticDecompression = DecompressionMethods.GZip
                myWebRequest.KeepAlive = False
                myWebRequest.ConnectionGroupName = Guid.NewGuid().ToString()
                myWebRequest.ServicePoint.Expect100Continue = False
                myWebRequest.Pipelined = False
                myWebRequest.MaximumResponseHeadersLength = 4
                myWebResponse = myWebRequest.GetResponse()
            Catch ex As Exception
                Try
                    myWebRequest.Timeout = 120000
                    myWebRequest.Method = "GET"
                    myWebRequest.AutomaticDecompression = DecompressionMethods.GZip
                    myWebRequest.KeepAlive = False
                    myWebRequest.ConnectionGroupName = Guid.NewGuid().ToString()
                    myWebRequest.ServicePoint.Expect100Continue = False
                    myWebRequest.Pipelined = False
                    myWebRequest.MaximumResponseHeadersLength = 4
                    myWebResponse = myWebRequest.GetResponse()
                Catch ex1 As Exception
                    Throw New Exception(ex.ToString())
                End Try
            End Try
            Dim receiveStream As Stream = myWebResponse.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
            Dim reader As StreamReader = New StreamReader(receiveStream, encode)
            Dim allText As String = reader.ReadToEnd()
            Dim hd As New HtmlDocument()
            hd.LoadHtml(allText)
            Return hd
        Catch ex As Exception
            Throw New Exception()
        End Try
    End Function

    Public Shared Function GetHtmlDocByUrlTmall(ByVal pageUrl As String, Optional ByVal dg As Integer = 3) As HtmlDocument
        Dim document As New HtmlDocument()
        Try
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = "aaa"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            Dim cookie As String = ConfigurationManager.AppSettings("aliTmallCookie").ToString.Trim
            request.Headers.Add("Cookie", cookie)
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")  'gb2312
            document.Load(resStream, pageEncoding)
            LogText("递归次数dg=" & dg)
            Return document
        Catch ex As Exception
            LogText("GetHtmlDocByUrlTmall-RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlDocByUrlTmall(pageUrl, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function

    Public Shared Function GetAsynHtmlDocByUrlTmall(ByVal pageUrl As String) As HtmlDocument
        Dim htmlDoc As New HtmlDocument()
        Dim categoryDocStr As String
        Try
            categoryDocStr = GetAsynHtmlStrTmall(pageUrl)
        Catch ex As Exception
            Threading.Thread.Sleep(5 * (60 + 5) * 1000)
            categoryDocStr = GetAsynHtmlStrTmall(pageUrl)
        End Try
        htmlDoc.LoadHtml(categoryDocStr.Replace("\", ""))
        Return htmlDoc
    End Function

    ''' <summary>
    ''' 获取ali的异步请求asynSearch.htm页面的请求string
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAsynHtmlStrTmall(ByVal pageUrl As String, Optional ByVal dg As Integer = 3) As String

        Dim pathIndex As Integer = pageUrl.IndexOf(".com")
        Dim aliUrl As String = pageUrl.Substring(0, pathIndex) & ".com"
        Dim aliAsynRandomParam As String
        Dim aliAsynDocument As HtmlDocument = GetHtmlDocByUrlTmall(pageUrl)
        aliAsynRandomParam = aliAsynDocument.DocumentNode.SelectSingleNode("//input[@id='J_ShopAsynSearchURL']").GetAttributeValue("value", "").Replace("&amp;", "&")
        aliUrl = aliUrl & aliAsynRandomParam
        Try
            Dim request As HttpWebRequest = HttpWebRequest.Create(aliUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = "aaa"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            Dim cookie As String = ConfigurationManager.AppSettings("aliTmallCookie").ToString.Trim
            request.Headers.Add("Cookie", cookie)
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")  'gb2312
            Dim resStreamReader As StreamReader = New StreamReader(resStream, pageEncoding)
            Dim resultSting As String = resStreamReader.ReadToEnd()
            LogText("递归次数dg=" & dg)
            Return resultSting
        Catch ex As Exception
            LogText("GetAsynHtmlStrTmall()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetAsynHtmlStrTmall(pageUrl, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function


    ''' <summary>
    ''' 请求tmall或baobao，以txt类型返回请求的html
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlStringByUrlAli(ByVal pageUrl As String, Optional ByVal dg As Integer = 3) As String
        Dim resultSting As String
        Try
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = "aaa"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            Dim cookie As String = ConfigurationManager.AppSettings("aliTmallCookie").ToString.Trim
            request.Headers.Add("Cookie", cookie)
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")  'gb2312
            Dim resStreamReader As StreamReader = New StreamReader(resStream, pageEncoding)
            resultSting = resStreamReader.ReadToEnd()
            LogText("递归次数dg=" & dg)
            Return resultSting
        Catch ex As Exception
            LogText("GetHtmlStringByUrlAli()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlStringByUrlAli(pageUrl, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function

    ''' <summary>
    ''' 以txt形式返回请求页面的内容。推荐使用其重载函数GetHtmlStringByUrl（pageUrl,cookie,refer)
    ''' </summary>
    ''' <param name="pageUrl">请求页面url</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlStringByUrl(ByVal pageUrl As String, Optional ByVal dg As Integer = 3) As String
        Try
            Dim resultSting As String
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
#If DEBUG Then
            '+debug# proxy
            'Dim myProxy As New WebProxy("127.0.0.1:1080", True) '54.249.57.179:8686
            'request.Proxy = myProxy
#End If
            request.Timeout = 120000
            Dim myProxy As New WebProxy("127.0.0.1:1080", True)
            request.Proxy = myProxy
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = "aaa"
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")  'gb2312
            Dim resStreamReader As StreamReader = New StreamReader(resStream, pageEncoding)
            resultSting = resStreamReader.ReadToEnd()
            LogText("递归次数dg=" & dg)
            Return resultSting
        Catch ex As Exception
            LogText("GetHtmlStringByUrl()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlStringByUrl(pageUrl, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function

    ''' <summary>
    ''' 以txt形式返回请求页面的内容。推荐使用。
    ''' </summary>
    ''' <param name="pageUrl">请求页面Url</param>
    ''' <param name="cookie">cookie</param>
    ''' <param name="refer">请求refer</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlStringByUrl(ByVal pageUrl As String, ByVal cookie As String, ByVal refer As String, Optional ByVal dg As Integer = 3) As String
        Try
            cookie = cookie.Trim
            refer = refer.Trim
            Dim resultSting As String
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = refer
            request.Headers.Add("Cookie", cookie)
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")  'gb2312
            Dim resStreamReader As StreamReader = New StreamReader(resStream, pageEncoding)
            resultSting = resStreamReader.ReadToEnd()
            LogText("递归次数dg=" & dg)
            Return resultSting
        Catch ex As Exception
            LogText("GetHtmlStringByUrl()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlStringByUrl(pageUrl, cookie, refer, dg)
            Else
                LogText("dg error occured.")
                Throw ex
            End If
        End Try
    End Function

    ''' <summary>
    ''' 以txt形式返回请求页面的内容。
    ''' </summary>
    ''' <param name="pageUrl">请求页面url</param>
    ''' <param name="cookie"cookie></param>
    ''' <param name="refer">参照页面</param>
    ''' <param name="pageEncoding">页面编码</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlStringByUrl(ByVal pageUrl As String, ByVal cookie As String, ByVal refer As String, ByVal pageEncoding As String, Optional ByVal dg As Integer = 3) As String
        Try
            cookie = cookie.Trim
            pageEncoding = pageEncoding.Trim
            Dim resultSting As String
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
#If DEBUG Then
            '+debug# proxy
            'Dim myProxy As New WebProxy("127.0.0.1:1080", True)
            'request.Proxy = myProxy
#End If
            '模拟https请求，这个方法对tmall失败
            'If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            '    ServicePointManager.Expect100Continue = True
            '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3  '; //SSL3协议替换成TLS协议
            '    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            '    request = HttpWebRequest.Create(pageUrl)
            '    request.ProtocolVersion = HttpVersion.Version10
            'End If
            request.Proxy = New WebProxy("127.0.0.1:1080", True)
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = refer
            request.Headers.Add("Cookie", cookie)
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            '设置代理

            'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
            Dim response As WebResponse = request.GetResponse()
            'WebRequest.GetResponse方法，返回对 Internet 请求的响应
            Dim resStream As Stream = response.GetResponseStream()
            'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
            Dim myencoding As Encoding
            If (pageEncoding Is Nothing Or String.IsNullOrEmpty(pageEncoding)) Then
                myencoding = Encoding.UTF8
            Else
                myencoding = Encoding.GetEncoding(pageEncoding)
            End If
            Dim resStreamReader As StreamReader = New StreamReader(resStream, myencoding)
            resultSting = resStreamReader.ReadToEnd()
            LogText("递归次数dg=" & dg)
            Common.LogText("递归次数dg=" & dg)
            Return resultSting
        Catch ex As Exception
            LogText("GetHtmlStringByUrl()->RequestURL:" & pageUrl)
            If dg > 0 Then
                dg = dg - 1
                Return GetHtmlStringByUrl(pageUrl, cookie, refer, pageEncoding, dg)
            Else
                Common.LogText("dg error occured.")
                LogText("dg error occured.")

                Throw ex
            End If
        End Try
    End Function

    Public Shared Function GetHtmlStringByUrlSinaWB(ByVal pageUrl As String, ByVal cookie As String) As String
        Dim resultSting As String
        Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
        If (pageUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase)) Then
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls '; //SSL3协议替换成TLS协议
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
            request.ProtocolVersion = HttpVersion.Version10
        End If
        request.Timeout = 120000
        request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

        'request.Headers.Add("Accept-Encoding", "gzip, deflate, sdch")'导致请求回来的数据是乱码
        request.Headers.Add("Accept-Language", "zh-CN")

        request.Referer = "https://www.google.com.hk/"
        request.Host = "weibo.com"
        request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.65 Safari/537.36"
        request.Method = "GET"

        request.Headers.Add("Cookie", cookie)
        request.AllowAutoRedirect = True
        'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
        Dim response As WebResponse = request.GetResponse()
        'WebRequest.GetResponse方法，返回对 Internet 请求的响应
        Dim resStream As Stream = response.GetResponseStream()
        'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。

        Dim resStreamReader As StreamReader = New StreamReader(resStream, Encoding.UTF8)
        resultSting = resStreamReader.ReadToEnd()
        Return resultSting
    End Function


    ''' <summary>
    ''' 下载图片保存至本地
    ''' </summary>
    ''' <param name="imageUrl"></param>
    ''' <param name="filePath"></param>
    ''' <param name="fileName"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DownloadImage(ByVal imageUrl As String, ByVal filePath As String, ByVal fileName As String, ByVal siteid As Integer) As String
        Dim client As New WebClient()

        Dim index As Integer = imageUrl.LastIndexOf("/")
        Dim imageName As String 'imageurl中第一个参数值+name作为imagename
        imageName = imageUrl.Substring(index + 1)
        Dim paramIndex As Integer = imageName.IndexOf("?")
        If (paramIndex > 0) Then
            'Dim params As String() = imageName.Split("=")
            imageName = imageName.Substring(0, paramIndex)
        End If
        imageName = Now.Day & "_" & Now.Minute & "_" & Now.Second & "_" & Now.Millisecond & "_" & imageName

        If Not (filePath.EndsWith("\")) Then
            filePath = filePath & "\"
        End If
        Dim datetime As String = Now.Year & "_" & Now.Month
        '以店铺名年月划分文件夹
        filePath = filePath & fileName & "\" & siteid & "_" & datetime
        If Not IO.Directory.Exists(filePath) Then IO.Directory.CreateDirectory(filePath)

        filePath = filePath & "\" & imageName
        Try
            client.DownloadFile(imageUrl, filePath)
            Return "\" & fileName & "\" & siteid & "_" & datetime & "\" & imageName
        Catch ex As Exception
            EFHelper.Log(ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' get the page content according to the rssUrl of website
    ''' </summary>
    ''' <param name="rssUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRssPageString(ByVal rssUrl As String) As String
        Try
            Dim myWebRequest As WebRequest = WebRequest.Create(rssUrl)
            Dim myWebResponse As WebResponse
            Try
                myWebRequest.Timeout = 60000
                myWebResponse = myWebRequest.GetResponse()
            Catch ex As Exception
                myWebRequest.Timeout = 120000
                myWebResponse = myWebRequest.GetResponse()
            End Try
            Dim receiveStream As Stream = myWebResponse.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(receiveStream, System.Text.Encoding.UTF8)
            Dim pageString As String = reader.ReadToEnd()
            Return pageString
        Catch ex As Exception
            'LogText(ex.ToString())
            Throw New Exception(ex.ToString())
        End Try
    End Function

    ''' <summary>
    ''' 补全图片imgSRC ，产品herf属性的域名部分
    ''' </summary>
    ''' <param name="domain">应去除末尾/,以防重复</param>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function addDominForUrl(ByVal domain As String, ByVal url As String) As String
        If Not (url.ToLower.StartsWith("http")) Then
            If url.StartsWith("//") OrElse url.StartsWith("://") Then
                Try
                    Dim uri As New Uri(domain)
                    Return uri.Scheme & ":" & url.TrimStart(":")
                Catch ex As Exception

                End Try
            End If
            domain = domain.TrimEnd("/", "\\")
            If (url.Trim().Trim.StartsWith("/")) Then
                url = domain & url
            Else
                url = domain & "/" & url
            End If
        End If
        Return url
    End Function

    ''' <summary>
    ''' 为阿里的url及imgUrl增加省略的“http”头部
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddHttpForAli(ByVal url As String) As String
        If Not (url.StartsWith("http")) Then
            If (url.StartsWith("//")) Then
                url = "http:" & url
            Else
                url = "http://" & url
            End If
        End If
        Return url
    End Function

    ''' <summary>
    ''' 将一个rss Url内容装载进xmldocument，并将其返回
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="pageEncoding">请求url时的编码方式</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadXmlDoc(ByVal url As String, ByVal pageEncoding As String) As Xml.XmlDocument
        Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
        Dim result As String = GetHtmlStringByUrl(url, "", "", pageEncoding)

        Try
            Dim Collection As MatchCollection = Regex.Matches(result, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", RegexOptions.IgnoreCase)

            If Collection.Count > 0 Then
                '获取channel title，并为channel title加上![CDATA，
                '防止获取到的channel title没加上![CDATA,使用.net方法读取XML文件时出错
                Dim myTitle As String = Collection.Item(0).Groups(1).Value
                Dim myTitleNode As String = "<title><![CDATA[" & myTitle & "]]></title>"
                result = System.Text.RegularExpressions.Regex.Replace(result, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", myTitleNode)
            End If
        Catch ex As Exception
            LogText("LoadXmlDoc()->RequestURL:" & url)
            LogText(ex.ToString)
        End Try

        xmlDoc.LoadXml(result)
        Return xmlDoc
    End Function
    ''' <summary>
    ''' 将一个rss Url内容装载进xmldocument，并将其返回
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="pageEncoding">请求url时的编码方式</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadXmlDoc(ByVal url As String) As Xml.XmlDocument
        Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
        Dim result As String = GetHtmlStringByUrl(url, "", "", "utf-8")

        Try
            'Dim para() As String = {"title", "description", "limitation", "feature"}
            'For Each p As String In para
            '    Dim Collection As MatchCollection = Regex.Matches(result, "<" & p & ">((?!<\!\[CDATA|\<" & p & ")[\s\S])*</" & p & ">", RegexOptions.IgnoreCase)

            '    If Collection.Count > 0 Then
            '        '获取channel title，并为channel title加上![CDATA，
            '        '防止获取到的channel title没加上![CDATA,使用.net方法读取XML文件时出错
            '        Dim myTitle As String = Collection.Item(0).Groups(1).Value
            '        Dim myTitleNode As String = "<" & p & "><![CDATA[" & myTitle & "]]></" & p & ">"
            '        result = System.Text.RegularExpressions.Regex.Replace(result, "<" & p & ">((?!<\!\[CDATA|\<" & p & ")[\s\S])*</" & p & ">", myTitleNode)
            '    End If
            'Next
        Catch ex As Exception
            LogText("LoadXmlDoc()->RequestURL:" & url)
            LogText(ex.ToString)
        End Try

        Dim b As Boolean = True
        Dim watch = System.Diagnostics.Stopwatch.StartNew()
        watch.Start()

        While (b)
            Try
                If watch.ElapsedMilliseconds > 120000 Then
                    watch.Stop()
                    Throw New TimeoutException()
                End If

                xmlDoc.LoadXml(result)
                b = False
            Catch ex As Exception
                Dim errorType As Type = ex.GetType()
                Dim field As FieldInfo = errorType.GetField("args", BindingFlags.NonPublic Or BindingFlags.Instance)
                If field IsNot Nothing Then
                    Dim args() As String = field.GetValue(ex)
                    result = result.Replace(args(0), "")
                Else
                    Throw ex
                End If
            End Try
        End While

        Return xmlDoc
    End Function

    ''' <summary>
    ''' 验证证书
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="certificate"></param>
    ''' <param name="chain"></param>
    ''' <param name="sslPolicyErrors"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckValidationResult(ByVal sender As Object, ByVal certificate As X509Certificate,
                                                               ByVal chain As X509Chain, ByVal sslPolicyErrors As SslPolicyErrors) As Boolean
        ''Return True to force the certificate to be accepted.
        Return True
    End Function
#End Region

#Region "Categories表"
    ''' <summary>
    ''' 根据siteId获取某个商家的产品分类
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListCategory(ByVal siteId As Integer) As List(Of Category)
        Dim queryCategory = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c

        Dim listCategory As New List(Of Category)
        'Try
        '    listCategory = queryCategory.ToList()
        'Catch ex As Exception
        '    'Ignore
        'End Try
        If (queryCategory.Count > 0) Then
            listCategory = queryCategory.ToList()
        End If
        Return listCategory
    End Function

    ''' <summary>
    ''' 将多个Category插入到Categories表中
    ''' </summary>
    ''' <param name="listCategory"></param>
    ''' <remarks></remarks>
    Public Sub InsertListCategory(ByVal listCategory As List(Of Category))
        Try
            For Each li As Category In listCategory
                efContext.AddToCategories(li)
            Next
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertListCategory()-->" & ex.ToString)
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 将一个Category插入到Categories表中
    ''' </summary>
    ''' <param name="myCategory"></param>
    ''' <remarks></remarks>
    Public Sub InsertCategory(ByVal myCategory As Category)
        Try
            efContext.AddToCategories(myCategory)
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertCategory()-->" & ex.ToString)
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' 插入或者更新Category的内容
    ''' </summary>
    ''' <param name="myCategory"></param>
    ''' <remarks></remarks>
    Public Sub InsertOrUpdateCate(ByVal myCategory As Category, ByVal siteId As Integer)
        Try
            efContext.Refresh(Objects.RefreshMode.ClientWins, efContext.Categories)
            Dim queryCategory As Category = efContext.Categories.Where(Function(c) c.Url = myCategory.Url AndAlso c.SiteID = siteId).Single()
            queryCategory.Category1 = myCategory.Category1
            queryCategory.LastUpdate = Now
        Catch ex As Exception
            efContext.AddToCategories(myCategory)
        End Try
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 更新多个Category，如果Url相同的数据，则更新Categories表中数据，2013/4/24新加
    ''' </summary>
    ''' <param name="listCategory"></param>
    ''' <remarks></remarks>
    Public Sub UpdateListCategory(ByVal listCategory As List(Of Category))
        For Each li In listCategory
            Dim updateCate = efContext.Categories.Single(Function(c) c.Url = li.Url)
            If Not (String.IsNullOrEmpty(li.Category1)) Then
                updateCate.Category1 = li.Category1
            End If
            If Not (String.IsNullOrEmpty(li.LastUpdate)) Then
                updateCate.LastUpdate = li.LastUpdate
            End If
            If Not (String.IsNullOrEmpty(li.Description)) Then
                updateCate.Description = li.Description
            End If
            If Not (String.IsNullOrEmpty(li.PictureUrl)) Then
                updateCate.PictureUrl = li.PictureUrl
            End If
            If Not (String.IsNullOrEmpty(li.PictureAlt)) Then
                updateCate.PictureAlt = li.PictureAlt
            End If
            If Not (li.SizeHeight Is Nothing) Then
                updateCate.SizeHeight = li.SizeHeight
            End If
            If Not (li.SizeWidth Is Nothing) Then
                updateCate.SizeWidth = li.SizeWidth
            End If
            If Not (String.IsNullOrEmpty(li.Gender)) Then
                updateCate.Gender = li.Gender
            End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 更新一个Category，如果Url相同的数据，则更新Categories表中数据，2013/4/27新加
    ''' </summary>
    ''' <param name="myCategory"></param>
    ''' <remarks></remarks>
    Public Sub UpdateCategory(ByVal myCategory As Category)
        Dim updateCate = efContext.Categories.Single(Function(c) c.Url = myCategory.Url)
        If Not (String.IsNullOrEmpty(myCategory.Category1)) Then
            updateCate.Category1 = myCategory.Category1
        End If
        If Not (String.IsNullOrEmpty(myCategory.LastUpdate)) Then
            updateCate.LastUpdate = myCategory.LastUpdate
        End If
        If Not (String.IsNullOrEmpty(myCategory.Description)) Then
            updateCate.Description = myCategory.Description
        End If
        If Not (String.IsNullOrEmpty(myCategory.PictureUrl)) Then
            updateCate.PictureUrl = myCategory.PictureUrl
        End If
        If Not (String.IsNullOrEmpty(myCategory.PictureAlt)) Then
            updateCate.PictureAlt = myCategory.PictureAlt
        End If
        If Not (myCategory.SizeHeight Is Nothing) Then
            updateCate.SizeHeight = myCategory.SizeHeight
        End If
        If Not (myCategory.SizeWidth Is Nothing) Then
            updateCate.SizeWidth = myCategory.SizeWidth
        End If
        If Not (String.IsNullOrEmpty(myCategory.Gender)) Then
            updateCate.Gender = myCategory.Gender
        End If
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 获取某个账号下的所有分类URL
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetListCateUrl(ByVal siteId As Integer) As List(Of String)
        Dim listCateUrl As New List(Of String)
        listCateUrl = (From c In efContext.Categories
                       Where c.SiteID = siteId
                       Select c.Url).ToList()
        Return listCateUrl
    End Function

    ''' <summary>
    ''' 根据名字获取某个账号下对应的分类List
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="arrCategoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetListCategorys(ByVal siteId As Integer, ByVal arrCategoryName As String()) As List(Of Category)
        Dim listCategorys As New List(Of Category)
        listCategorys = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso arrCategoryName.Contains(c.Category1)).ToList()
        Return listCategorys
    End Function

    ''' <summary>
    ''' 根据某个分类的名字获取某个账户下面的分类Id
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCategoryId(ByVal siteId As Integer, ByVal categoryName As String) As String
        Dim categoryId As Integer = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).Single().CategoryID
        Return categoryId.ToString()
    End Function

    'Public Sub InsertProductCategory(ByVal productid As Long, ByVal categoryid As Integer)
    '    Dim query = (From pc In efContext.productcategory
    '                 Where pc.Productid = productid And pc.categoryid = GetCategoryId()).count

    'End Sub
    ''' <summary>
    ''' 根据某个分类的名字获取某个账户下面的分类对象
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCategory(ByVal siteId As Integer, ByVal categoryName As String) As Category
        Dim myCategory As New Category
        myCategory = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).SingleOrDefault()
        Return myCategory
    End Function

    ''' <summary>
    ''' 根据cateid获取一个分类对象
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="cateId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetOneCategory(ByVal siteid As Integer, ByVal cateId As Integer) As Category
        Dim myCate As Category = (From c In efContext.Categories
                                  Where c.SiteID = siteid AndAlso c.CategoryID = cateId
                                  Select c).FirstOrDefault()
        Return myCate
    End Function

    ''' <summary>
    ''' cates的参数规范cate1^cate2...
    ''' </summary>
    ''' <param name="cates"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCategories(ByVal cates As String, ByVal siteid As Integer) As List(Of Category)
        Dim arrCate As String() = cates.Split("^")
        Dim listCate As List(Of Category) = New List(Of Category)
        For Each item In arrCate
            Dim acate As Category = (From c In efContext.Categories
                                     Where c.SiteID = siteid AndAlso c.Category1 = item
                                     Select c).FirstOrDefault()
            If Not (acate Is Nothing) Then
                listCate.Add(acate)
            End If
        Next
        Return listCate
    End Function

    ''' <summary>
    ''' cates的参数规范cate1^cate2...返回cate1，care1_url;格式的string
    ''' </summary>
    ''' <param name="cates"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCategoryItems(ByVal cates As String, ByVal siteid As Integer) As String
        Dim arrCate As String() = cates.Split("^")
        Dim cateItems As String = ""
        For Each item In arrCate
            Dim acate As Category = (From c In efContext.Categories
                                     Where c.SiteID = siteid AndAlso c.Category1 = item
                                     Select c).FirstOrDefault()
            If Not (acate Is Nothing) Then
                cateItems = cateItems & acate.Category1.Trim & "," & acate.Url.Trim & ";"
            End If
        Next
        Return cateItems
    End Function

    ''' <summary>
    ''' 将category数据整理进数据库。如果cateName存在，则更新此category。如果cateName不存在则添加至数据库
    ''' </summary>
    ''' <param name="listCategory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrUpdateCategory(ByVal listCategory As List(Of Category)) As String
        If (listCategory.Count > 0) Then

            Dim siteid As Integer = listCategory(0).SiteID
            Dim listExistCategoryName As List(Of String) = (From c In efContext.Categories
                                                            Where c.SiteID = siteid
                                                            Select c.Category1).ToList()
            For Each item As Category In listCategory
                If Not (listExistCategoryName.Contains(item.Category1)) Then
                    efContext.Categories.AddObject(item)
                Else
                    Dim updateCate As Category = (From c In efContext.Categories
                                                  Where c.SiteID = item.SiteID AndAlso c.Category1 = item.Category1
                                                  Select c).FirstOrDefault()
                    updateCate.Url = item.Url
                    updateCate.LastUpdate = DateTime.Now
                    efContext.SaveChanges()
                End If
            Next
            efContext.SaveChanges()
        End If
    End Function
#End Region

#Region "Products表"
    ''' <summary>
    ''' 根据siteId获取某个商家的产品链表URL
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns>返回产品的Url链表</returns>
    ''' <remarks></remarks>
    Public Function GetListProduct(ByVal siteId As Integer) As List(Of String)
        Dim queryProduct = From p In efContext.Products
                           Where p.SiteID = siteId
                           Select p.Url
        Dim listProduct As New List(Of String)
        Try
            listProduct = queryProduct.ToList()
        Catch ex As Exception
            'Ignore
        End Try
        Return listProduct
    End Function

    ''' <summary>
    ''' 根据siteId获取某个商家的产品链表
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListProduct2(ByVal siteId As Integer) As List(Of Product)
        Dim queryProduct = From p In efContext.Products
                           Where p.SiteID = siteId
                           Select p
        Dim listProduct As New List(Of Product)
        listProduct = queryProduct.ToList()
        Return listProduct
    End Function

    ''' <summary>
    ''' 将多个Product插入到Products表中
    ''' </summary>
    ''' <param name="listProduct"></param>
    ''' <remarks></remarks>
    Public Function InsertListProduct(ByVal listProduct As List(Of Product), ByVal strCategory As String, ByVal siteId As Integer) As List(Of Integer)
        Dim listProductId As New List(Of Integer)
        Dim category = efContext.Categories.FirstOrDefault(Function(c) c.Category1 = strCategory.Trim() AndAlso c.SiteID = siteId)
        For Each li As Product In listProduct
            Try
                li.Categories.Add(category)
                efContext.AddToProducts(li)
                efContext.SaveChanges()
                listProductId.Add(li.ProdouctID)
            Catch ex As Exception
                LogText("InsertListProduct()-->" & ex.ToString)

            End Try


        Next
        Return listProductId
    End Function

    ''' <summary>
    ''' 将一个产品插入到Products表中
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="strCategory"></param>
    ''' <param name="siteId"></param>
    ''' <returns>返回插入的产品的Id</returns>
    ''' <remarks></remarks>
    Public Function InsertSingleProduct(ByVal myProduct As Product, ByVal strCategory As String, ByVal siteId As Integer) As Integer
        'Dim ef As New EmailAlerterEntities
        Try
            Dim category = efContext.Categories.Single(Function(c) c.Category1 = strCategory AndAlso c.SiteID = siteId)
            'myProduct.Categories.Add(category)
            efContext.SaveChanges()
            efContext.Products.AddObject(myProduct)
            ' efContext.AddToProducts(myProduct)
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertSingleProduct()-->" & ex.ToString)
        End Try

        Return myProduct.ProdouctID
    End Function

    ''' <summary>
    ''' 将一个产品插入到Products表中，
    ''' 改进InsertSingleProduct()方法
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="categoryUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns>返回插入的产品的Id</returns>
    ''' <remarks></remarks>
    Public Function InsertSingleProduct2(ByVal myProduct As Product, ByVal categoryUrl As String, ByVal siteId As Integer) As Integer
        Dim category = efContext.Categories.Single(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId)
        myProduct.Categories.Add(category)
        efContext.AddToProducts(myProduct)
        efContext.SaveChanges()
        Return myProduct.ProdouctID
    End Function

    ''' <summary>
    ''' 获取或者更新单个产品信息到Products表中
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="categoryUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrUpdateProduct(ByVal myProduct As Product, ByVal categoryUrl As String, ByVal siteId As Integer,
                                          ByVal updateTime As DateTime) As Integer

        Using EF As New EmailAlerterEntities
            Try
                Dim updateProduct As Product = EF.Products.FirstOrDefault(Function(p) p.Url = myProduct.Url)
                If Not (String.IsNullOrEmpty(myProduct.Prodouct)) Then
                    updateProduct.Prodouct = myProduct.Prodouct
                End If
                If Not (myProduct.Price Is Nothing) Then
                    updateProduct.Price = myProduct.Price
                End If
                If Not (myProduct.Discount Is Nothing) Then
                    updateProduct.Discount = myProduct.Discount
                End If
                If Not (myProduct.Sales Is Nothing) Then
                    updateProduct.Sales = myProduct.Sales
                End If
                If Not (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                    updateProduct.PictureUrl = myProduct.PictureUrl.Trim()
                End If
                updateProduct.LastUpdate = updateTime

                If Not (String.IsNullOrEmpty(myProduct.Description)) Then
                    updateProduct.Description = myProduct.Description
                End If

                If Not (String.IsNullOrEmpty(myProduct.Currency)) Then
                    updateProduct.Currency = myProduct.Currency
                End If

                If Not (String.IsNullOrEmpty(myProduct.PictureAlt)) Then
                    updateProduct.PictureAlt = myProduct.PictureAlt
                End If

                If Not (myProduct.SizeWidth Is Nothing) Then
                    updateProduct.SizeWidth = myProduct.SizeWidth
                End If

                If Not (myProduct.SizeHeight Is Nothing) Then
                    updateProduct.SizeHeight = myProduct.SizeHeight
                End If
                If Not (myProduct.TbScore Is Nothing) Then
                    updateProduct.TbScore = myProduct.TbScore
                End If
                If Not (myProduct.TbComment Is Nothing) Then
                    updateProduct.TbComment = myProduct.TbComment
                End If

                '2013/06/08 added
                'FreeShipping picture and ships24 picture update,begin
                If Not (String.IsNullOrEmpty(myProduct.FreeShippingImg)) Then
                    updateProduct.FreeShippingImg = myProduct.FreeShippingImg
                End If
                If Not (String.IsNullOrEmpty(myProduct.ShipsImg)) Then
                    updateProduct.ShipsImg = myProduct.ShipsImg
                End If

                'li.Categories.Add(category) '出错的地方，此处虽然没有添加product但是，它会自动添加，而我们只需要更新表数据即可，导致错误
                EF.SaveChanges()
                'listProductId.Add(updateProduct.ProdouctID)

                '2013/4/7新增-begin
                Dim queryProduct = From p In EF.Products Where p.ProdouctID = updateProduct.ProdouctID Select p
                Dim cate = queryProduct.FirstOrDefault.Categories '关系表存在的Category
                Dim queryCategory = EF.Categories.FirstOrDefault(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId) '新增记录的Category
                Dim counter As Integer = 0
                For Each c In cate.ToList
                    If (c.CategoryID = queryCategory.CategoryID) Then
                        Exit For
                    Else
                        counter = counter + 1
                    End If
                    If (counter = cate.Count) Then
                        queryCategory.Products.Add(updateProduct)
                    End If
                Next
                If (cate.Count = 0) Then
                    queryCategory.Products.Add(updateProduct)
                End If
                EF.SaveChanges()
                Return updateProduct.ProdouctID
            Catch ex As Exception
                Dim category As New Category
                Try
                    category = EF.Categories.Single(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId)
                Catch ex1 As Exception
                    If (categoryUrl.Contains("http: //es.focalprice.com")) Then
                        categoryUrl = categoryUrl.Replace("http://es.focalprice.com", "http://dynamic.es.focalprice.com")
                    ElseIf (categoryUrl.Contains("http://dynamic.es.focalprice.com")) Then
                        categoryUrl = categoryUrl.Replace("http://dynamic.es.focalprice.com", "http://es.focalprice.com")
                    End If
                    category = EF.Categories.Single(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId)
                End Try
                myProduct.Categories.Add(category)
                EF.AddToProducts(myProduct)
                EF.SaveChanges()
                Return myProduct.ProdouctID
            End Try
        End Using

    End Function

    ''' <summary>
    ''' 根据categoryUrl或者categoryName获取或者更新单个产品信息到Products表中
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="categoryUrl">如果categoryUrl不为空，则categoryName为空；</param>
    ''' <param name="categoryName">如果categoryName不为空，则categoryUrl为空；</param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrUpdateProduct(ByVal myProduct As Product, ByVal categoryUrl As String, ByVal categoryName As String,
                                          ByVal siteId As Integer) As Integer
        Try
            Dim category As New Category
            If (String.IsNullOrEmpty(categoryUrl)) Then 'Update Products according to categoryName
                category = efContext.Categories.Single(Function(c) c.Category1 = categoryName AndAlso c.SiteID = siteId)
            Else 'Update Products according to categoryUrl
                category = efContext.Categories.Single(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId)
            End If
            myProduct.Categories.Add(category)
            efContext.AddToProducts(myProduct)
            efContext.SaveChanges()
            Return myProduct.ProdouctID
        Catch ex As Exception
            Try
                Dim updateProduct = efContext.Products.FirstOrDefault(Function(p) p.Url = myProduct.Url)
                If Not (String.IsNullOrEmpty(myProduct.Prodouct)) Then
                    updateProduct.Prodouct = myProduct.Prodouct
                End If
                If Not (myProduct.Price Is Nothing) Then
                    updateProduct.Price = myProduct.Price
                End If
                If Not (myProduct.Discount Is Nothing) Then
                    updateProduct.Discount = myProduct.Discount
                End If
                If Not (myProduct.Sales Is Nothing) Then
                    updateProduct.Sales = myProduct.Sales
                End If
                If Not (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                    updateProduct.PictureUrl = myProduct.PictureUrl.Trim()
                End If
                If Not (String.IsNullOrEmpty(myProduct.LastUpdate)) Then
                    updateProduct.LastUpdate = myProduct.LastUpdate
                End If

                If Not (String.IsNullOrEmpty(myProduct.ExpiredDate)) Then
                    updateProduct.ExpiredDate = myProduct.ExpiredDate
                End If

                If Not (String.IsNullOrEmpty(myProduct.Description)) Then
                    updateProduct.Description = myProduct.Description
                End If

                If Not (String.IsNullOrEmpty(myProduct.Currency)) Then
                    updateProduct.Currency = myProduct.Currency
                End If

                If Not (String.IsNullOrEmpty(myProduct.PictureAlt)) Then
                    updateProduct.PictureAlt = myProduct.PictureAlt
                End If

                If Not (myProduct.SizeWidth Is Nothing) Then
                    updateProduct.SizeWidth = myProduct.SizeWidth
                End If

                If Not (myProduct.SizeHeight Is Nothing) Then
                    updateProduct.SizeHeight = myProduct.SizeHeight
                End If
                If Not (myProduct.TbScore Is Nothing) Then
                    updateProduct.TbScore = myProduct.TbScore
                End If
                If Not (myProduct.TbComment Is Nothing) Then
                    updateProduct.TbComment = myProduct.TbComment
                End If

                '2013/06/08 added
                'FreeShipping picture and ships24 picture update,begin
                If Not (String.IsNullOrEmpty(myProduct.FreeShippingImg)) Then
                    updateProduct.FreeShippingImg = myProduct.FreeShippingImg
                End If
                If Not (String.IsNullOrEmpty(myProduct.ShipsImg)) Then
                    updateProduct.ShipsImg = myProduct.ShipsImg
                End If

                'li.Categories.Add(category) '出错的地方，此处虽然没有添加product但是，它会自动添加，而我们只需要更新表数据即可，导致错误
                efContext.SaveChanges()
                Dim queryProduct = From p In efContext.Products Where p.ProdouctID = updateProduct.ProdouctID Select p
                Dim cate = queryProduct.FirstOrDefault.Categories '关系表存在的Category
                Dim queryCategory As New Category
                If (String.IsNullOrEmpty(categoryUrl)) Then 'Update Products according to categoryName
                    queryCategory = efContext.Categories.Single(Function(c) c.Category1 = categoryName AndAlso c.SiteID = siteId)
                Else 'Update Products according to categoryUrl
                    queryCategory = efContext.Categories.Single(Function(c) c.Url = categoryUrl AndAlso c.SiteID = siteId)
                End If
                Dim counter As Integer = 0
                For Each c In cate.ToList
                    If (c.CategoryID = queryCategory.CategoryID) Then
                        Exit For
                    Else
                        counter = counter + 1
                    End If
                    If (counter = cate.Count) Then
                        queryCategory.Products.Add(updateProduct)
                    End If
                Next
                efContext.SaveChanges()
                Return updateProduct.ProdouctID
            Catch ex2 As Exception
                LogText("InsertOrUpdateProduct()-->" & ex2.ToString)
                Throw ex2
            End Try

        End Try
    End Function

    ''' <summary>
    ''' 更新多个产品的产品信息到Products表中
    ''' </summary>
    ''' <param name="listProduct"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateListProduct(ByVal listProduct As List(Of Product), ByVal strCategory As String, ByVal siteId As Integer) As List(Of Integer)
        Dim nowTime As DateTime = Now
        Dim listProductId As New List(Of Integer)

        For Each li As Product In listProduct
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(p) p.Url = li.Url)
            If Not (String.IsNullOrEmpty(li.Prodouct)) Then
                updateProduct.Prodouct = li.Prodouct
            End If
            If Not (li.Price Is Nothing) Then
                updateProduct.Price = li.Price
            End If
            If Not (li.Discount Is Nothing) Then
                updateProduct.Discount = li.Discount
            End If
            If Not (li.Sales Is Nothing) Then
                updateProduct.Sales = li.Sales
            End If
            If Not (String.IsNullOrEmpty(li.PictureUrl)) Then
                updateProduct.PictureUrl = li.PictureUrl.Trim()
            End If
            updateProduct.LastUpdate = nowTime

            If Not (String.IsNullOrEmpty(li.Description)) Then
                updateProduct.Description = li.Description
            End If

            If Not (String.IsNullOrEmpty(li.Currency)) Then
                updateProduct.Currency = li.Currency
            End If

            If Not (String.IsNullOrEmpty(li.PictureAlt)) Then
                updateProduct.PictureAlt = li.PictureAlt
            End If

            If Not (li.SizeWidth Is Nothing) Then
                updateProduct.SizeWidth = li.SizeWidth
            End If

            If Not (li.SizeHeight Is Nothing) Then
                updateProduct.SizeHeight = li.SizeHeight
            End If
            If Not (li.TbScore Is Nothing) Then
                updateProduct.TbScore = li.TbScore
            End If
            If Not (li.TbComment Is Nothing) Then
                updateProduct.TbComment = li.TbComment
            End If

            '2013/06/08 added
            'FreeShipping picture and ships24 picture update,begin
            If Not (String.IsNullOrEmpty(li.FreeShippingImg)) Then
                updateProduct.FreeShippingImg = li.FreeShippingImg
            End If
            If Not (String.IsNullOrEmpty(li.ShipsImg)) Then
                updateProduct.ShipsImg = li.ShipsImg
            End If
            'FreeShipping picture and ships24 picture update,end

            'li.Categories.Add(category) '出错的地方，此处虽然没有添加product但是，它会自动添加，而我们只需要更新表数据即可，导致错误
            efContext.SaveChanges()
            listProductId.Add(updateProduct.ProdouctID)

            '2013/4/7新增-begin
            Dim queryProduct = From p In efContext.Products Where p.ProdouctID = updateProduct.ProdouctID Select p
            Dim cate = queryProduct.FirstOrDefault.Categories '关系表存在的Category
            Dim queryCategory = efContext.Categories.FirstOrDefault(Function(c) c.Category1 = strCategory AndAlso c.SiteID = siteId) '新增记录的Category
            Dim counter As Integer = 0
            For Each c In cate.ToList
                If (c.CategoryID = queryCategory.CategoryID) Then
                    Exit For
                Else
                    counter = counter + 1
                End If
                If (counter = cate.Count) Then
                    queryCategory.Products.Add(updateProduct)
                End If
            Next
            efContext.SaveChanges()
            '2013/4/7新增加-end

        Next
        Return listProductId
    End Function

    ''' <summary>
    ''' 更新一个产品的产品信息到Products表中；
    ''' 如果该产品不存咋，则添加一条新纪录
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="strCategory"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateSingleProduct(ByVal myProduct As Product, ByVal strCategory As String, ByVal siteId As Integer, ByVal nowTime As DateTime) As Integer
        Dim returnProductId As Integer = 0
        Dim updateProduct As New Product
        Try
            updateProduct = efContext.Products.Single(Function(p) p.Url = myProduct.Url)
            If Not (String.IsNullOrEmpty(myProduct.Prodouct)) Then
                updateProduct.Prodouct = myProduct.Prodouct
            End If
            If Not (myProduct.Price Is Nothing) Then
                updateProduct.Price = myProduct.Price
            End If
            If Not (myProduct.Discount Is Nothing) Then
                updateProduct.Discount = myProduct.Discount
            End If
            If Not (myProduct.Sales Is Nothing) Then
                updateProduct.Sales = myProduct.Sales
            End If
            If Not (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                updateProduct.PictureUrl = myProduct.PictureUrl.Trim()
            End If
            updateProduct.LastUpdate = nowTime

            If Not (String.IsNullOrEmpty(myProduct.Description)) Then
                updateProduct.Description = myProduct.Description
            End If

            If Not (String.IsNullOrEmpty(myProduct.Currency)) Then
                updateProduct.Currency = myProduct.Currency
            End If

            If Not (String.IsNullOrEmpty(myProduct.PictureAlt)) Then
                updateProduct.PictureAlt = myProduct.PictureAlt
            End If

            If Not (myProduct.SizeWidth Is Nothing) Then
                updateProduct.SizeWidth = myProduct.SizeWidth
            End If

            If Not (myProduct.SizeHeight Is Nothing) Then
                updateProduct.SizeHeight = myProduct.SizeHeight
            End If
            If Not (myProduct.TbScore Is Nothing) Then
                updateProduct.TbScore = myProduct.TbScore
            End If
            If Not (myProduct.TbComment Is Nothing) Then
                updateProduct.TbComment = myProduct.TbComment
            End If
            'li.Categories.Add(category) '出错的地方，此处虽然没有添加product但是，它会自动添加，而我们只需要更新表数据即可，导致错误
            efContext.SaveChanges()
            returnProductId = updateProduct.ProdouctID

            '2013/4/7新增-begin
            Dim queryProduct = From p In efContext.Products Where p.ProdouctID = updateProduct.ProdouctID Select p
            Dim cate = queryProduct.Single.Categories '关系表存在的Category
            Dim queryCategory = efContext.Categories.Single(Function(c) c.Category1 = strCategory AndAlso c.SiteID = siteId) '新增记录的Category
            Dim counter As Integer = 0

            '2013/4/26新增，如果以前数据已经存在处理
            If (cate.Count <= 0) Then
                queryCategory.Products.Add(updateProduct)
            End If

            For Each c In cate.ToList
                If (c.CategoryID = queryCategory.CategoryID) Then
                    Exit For
                Else
                    counter = counter + 1
                End If
                If (counter = cate.Count) Then
                    queryCategory.Products.Add(updateProduct)
                End If
            Next
            efContext.SaveChanges()
        Catch ex As Exception
            efContext.AddToProducts(myProduct)
            efContext.SaveChanges()
            returnProductId = myProduct.ProdouctID
        End Try
        Return returnProductId
    End Function

    ''' <summary>
    ''' 更新一个产品的产品信息到Products表中；
    ''' 如果该产品不存咋，则添加一条新纪录（改进UpdateSingleProduct()方法）
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="categoryUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateSingleProduct2(ByVal myProduct As Product, ByVal categoryUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime) As Integer
        Dim returnProductId As Integer = 0
        Dim updateProduct As New Product
        Try
            updateProduct = efContext.Products.Single(Function(p) p.Url = myProduct.Url)
            If Not (String.IsNullOrEmpty(myProduct.Prodouct)) Then
                updateProduct.Prodouct = myProduct.Prodouct
            End If
            If Not (myProduct.Price Is Nothing) Then
                updateProduct.Price = myProduct.Price
            End If
            If Not (myProduct.Discount Is Nothing) Then
                updateProduct.Discount = myProduct.Discount
            End If
            If Not (myProduct.Sales Is Nothing) Then
                updateProduct.Sales = myProduct.Sales
            End If
            If Not (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                updateProduct.PictureUrl = myProduct.PictureUrl.Trim()
            End If
            updateProduct.LastUpdate = nowTime

            If Not (String.IsNullOrEmpty(myProduct.Description)) Then
                updateProduct.Description = myProduct.Description
            End If

            If Not (String.IsNullOrEmpty(myProduct.Currency)) Then
                updateProduct.Currency = myProduct.Currency
            End If

            If Not (String.IsNullOrEmpty(myProduct.PictureAlt)) Then
                updateProduct.PictureAlt = myProduct.PictureAlt
            End If

            If Not (myProduct.SizeWidth Is Nothing) Then
                updateProduct.SizeWidth = myProduct.SizeWidth
            End If

            If Not (myProduct.SizeHeight Is Nothing) Then
                updateProduct.SizeHeight = myProduct.SizeHeight
            End If
            If Not (myProduct.TbScore Is Nothing) Then
                updateProduct.TbScore = myProduct.TbScore
            End If
            If Not (myProduct.TbComment Is Nothing) Then
                updateProduct.TbComment = myProduct.TbComment
            End If
            'li.Categories.Add(category) '出错的地方，此处虽然没有添加product但是，它会自动添加，而我们只需要更新表数据即可，导致错误
            efContext.SaveChanges()
            returnProductId = updateProduct.ProdouctID

            '2013/4/7新增-begin
            Dim queryProduct = From p In efContext.Products Where p.ProdouctID = updateProduct.ProdouctID Select p
            Dim cate = queryProduct.Single.Categories '关系表存在的Category
            Dim queryCategory = efContext.Categories.Single(Function(c) c.Category1 = categoryUrl AndAlso c.SiteID = siteId) '新增记录的Category
            Dim counter As Integer = 0

            '2013/4/26新增，如果以前数据已经存在处理
            If (cate.Count <= 0) Then
                queryCategory.Products.Add(updateProduct)
            End If

            For Each c In cate.ToList
                If (c.CategoryID = queryCategory.CategoryID) Then
                    Exit For
                Else
                    counter = counter + 1
                End If
                If (counter = cate.Count) Then
                    queryCategory.Products.Add(updateProduct)
                End If
            Next
            efContext.SaveChanges()
        Catch ex As Exception
            efContext.AddToProducts(myProduct)
            efContext.SaveChanges()
            returnProductId = myProduct.ProdouctID
        End Try
        Return returnProductId
    End Function

    ''' <summary>
    ''' 判断指定时间内产品是否已经在邮件中发送   如果是HO、HA则在HO、HA内进行比较，如果是HPI 则在HO、HA、HPI内进行比较
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="productUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsProductSent(ByVal siteID As String, ByVal productUrl As String, ByVal StartDate As String, ByVal EndDate As String, ByVal planType As String) As Boolean
        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        'where (new int?[] {1,2}).Contains(p.CategoryID)  sql语句中的‘in'在linq的实现

        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID AndAlso (New String() {"HO", "HA", planType}).Contains(i.PlanType)
                        Select i.IssueID).ToList()

        Dim lastproductUrls As List(Of String)
        For Each issues In lastIssueIDS
            lastproductUrls = (From p In efContext.Products
                               Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                               Where pi.IssueID = issues And pi.SiteId = siteID
                               Select p.Url.Trim().ToLower()).ToList()
            If (lastproductUrls.Contains(productUrl.Trim().ToLower())) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' 判断指定时间内产品是否已经在邮件中发送，在同一个PlanType内进行比较，缩小IsProductSent()的查询范围
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="productUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsProductSent2(ByVal siteID As String, ByVal productUrl As String, ByVal StartDate As String, ByVal EndDate As String, ByVal planType As String) As Boolean
        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        'where (new int?[] {1,2}).Contains(p.CategoryID)  sql语句中的‘in'在linq的实现

        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID AndAlso i.PlanType.Trim() = planType.Trim
                        Select i.IssueID).ToList()

        Dim lastproductUrls As List(Of String)
        For Each issues In lastIssueIDS
            lastproductUrls = (From p In efContext.Products
                               Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                               Where pi.IssueID = issues And pi.SiteId = siteID
                               Select p.Url.Trim().ToLower()).ToList()
            If (lastproductUrls.Contains(productUrl.Trim().ToLower())) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' 判断指定时间内产品是否已经在邮件中发送，同所有的plantype进行比较
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="productUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsProductSent(ByVal siteID As String, ByVal productUrl As String, ByVal StartDate As String, ByVal EndDate As String) As Boolean  'IsProductSent(Site,ProductUrl,StartDate,EndDate) as boolean
        Dim recentIssues As List(Of Long)
        recentIssues = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID
                        Select i.IssueID).ToList()

        Dim lastproductUrls As List(Of String)
        For Each issues In recentIssues
            lastproductUrls = (From p In efContext.Products
                               Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                               Where pi.IssueID = issues And pi.SiteId = siteID
                               Select p.Url.Trim().ToLower()).ToList()
            If (lastproductUrls.Contains(productUrl.Trim().ToLower())) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' 获取指定站点的products表中的所有产品，Dora，2014.02.11
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProductList(ByVal siteId As Integer) As List(Of Product)
        Dim query = From p In efContext.Products
                    Where p.SiteID = siteId
                    Select p
        Dim listProduct As New List(Of Product)
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency, .FreeShippingImg = q.FreeShippingImg})
        Next
        Return listProduct
    End Function

    ''' <summary>
    ''' 将一个product数据插入到Product表中，如果Product表中已存在此条记录，则update，同时不修改此product的CategoryID  Dora 2014-02-11
    ''' </summary>
    ''' <param name="aProductInList"></param>
    ''' <param name="now"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertProduct(ByVal aProductInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal list As List(Of Product)) As Integer

        Dim result As Integer = -1
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = aProductInList.Prodouct
        product.Url = aProductInList.Url
        product.Price = aProductInList.Price
        product.PictureUrl = aProductInList.PictureUrl
        product.LastUpdate = now
        product.Description = aProductInList.Description
        product.SiteID = aProductInList.SiteID
        product.Currency = aProductInList.Currency
        product.PictureAlt = aProductInList.PictureAlt
        product.Discount = aProductInList.Discount
        product.ShipsImg = aProductInList.ShipsImg
        product.FreeShippingImg = aProductInList.FreeShippingImg
        product.ExpiredDate = aProductInList.ExpiredDate
        If (isNewProduct(product, list)) Then
            Try
                product.Categories.Add(category)
                efContext.Products.AddObject(product)
                efContext.SaveChanges()
                result = product.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct error: " & ex.ToString)
            End Try

        Else

            Try
                Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = aProductInList.Url And m.Prodouct = aProductInList.Prodouct)
                updateProduct.Prodouct = product.Prodouct
                updateProduct.Price = product.Price
                updateProduct.Description = product.Description
                updateProduct.PictureUrl = product.PictureUrl
                updateProduct.PictureAlt = product.PictureAlt
                updateProduct.Currency = product.Currency
                updateProduct.LastUpdate = now
                updateProduct.Discount = product.Discount
                updateProduct.FreeShippingImg = product.FreeShippingImg
                updateProduct.ShipsImg = product.ShipsImg
                updateProduct.ExpiredDate = product.ExpiredDate

                efContext.SaveChanges()
                result = updateProduct.ProdouctID
            Catch ex As Exception
                LogText("UpdateProduct error: " & ex.ToString)
            End Try

        End If

        Return result
    End Function



    Public Function InsertUnduplicateProduct(ByVal aProductInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal list As List(Of Product)) As Integer

        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = aProductInList.Prodouct
        product.Url = aProductInList.Url
        product.Price = aProductInList.Price
        product.PictureUrl = aProductInList.PictureUrl
        product.LastUpdate = now
        product.Description = aProductInList.Description
        product.SiteID = aProductInList.SiteID
        product.Currency = aProductInList.Currency
        product.PictureAlt = aProductInList.PictureAlt
        product.Discount = aProductInList.Discount
        product.ShipsImg = aProductInList.ShipsImg
        product.FreeShippingImg = aProductInList.FreeShippingImg
        product.ExpiredDate = aProductInList.ExpiredDate
        If (isNewProduct(product, list)) Then
            Try
                product.Categories.Add(category)
                efContext.Products.AddObject(product)
                efContext.SaveChanges()
                Return product.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct ERROR: " & ex.ToString)
                Return -1
            End Try

        Else

            Try
                Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = aProductInList.Url And m.Prodouct = aProductInList.Prodouct)
                updateProduct.Prodouct = product.Prodouct
                updateProduct.Price = product.Price
                updateProduct.Description = product.Description
                updateProduct.PictureUrl = product.PictureUrl
                updateProduct.PictureAlt = product.PictureAlt
                updateProduct.Currency = product.Currency
                updateProduct.LastUpdate = now
                updateProduct.Discount = product.Discount
                updateProduct.FreeShippingImg = product.FreeShippingImg
                updateProduct.ShipsImg = product.ShipsImg
                updateProduct.ExpiredDate = product.ExpiredDate


                '2014/2/21新增，防止一个产品有多个productCategory关系
                Dim updateCategory = updateProduct.Categories
                'Dim queryCate = From p In efContext.Products
                '                Where p.ProdouctID = updateProduct.ProdouctID
                '                Select p
                'Dim cate = queryCate.Single.Categories
                'Dim counter As Integer = updateCategory.Count
                For i As Integer = 0 To updateCategory.Count - 1
                    If updateCategory(0).CategoryID <> category.CategoryID Then 'product belong more than 1 category ,next one
                        Return 0
                    End If
                    updateCategory(0).Products.Remove(updateProduct) '移除ProductCategory表 这个产品对应的productid和categoryid绑定记录
                Next
                efContext.SaveChanges()

                If Not updateCategory.Contains(category) Then
                    category.Products.Add(updateProduct)  'ProductCategory表 绑定新的productid-categoryid
                End If
                efContext.SaveChanges()
                Return updateProduct.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct ERROR2: " & ex.ToString)
                Return -1
            End Try

        End If
    End Function

    ''' <summary>
    ''' 将一个K11 facebook 的product数据插入到Product表中，
    ''' 如果Product表中已存在此条记录，则update
    ''' 上一个方法InsertProduct()不同，这个产品重复判断的方法不一样,使用isNewFbProduct()
    ''' </summary>
    ''' <param name="aProductInList"></param>
    ''' <param name="now"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    Public Function InsertK11Product(ByVal aProductInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal list As List(Of Product)) As Integer

        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()

        Common.LogText("开始插入产品表")

        product.Prodouct = aProductInList.Prodouct
        product.Url = aProductInList.Url
        product.Price = aProductInList.Price
        product.PictureUrl = aProductInList.PictureUrl
        product.LastUpdate = now
        product.Description = aProductInList.Description
        product.SiteID = aProductInList.SiteID
        product.Currency = aProductInList.Currency
        product.PictureAlt = aProductInList.PictureAlt
        product.Discount = aProductInList.Discount
        product.ShipsImg = aProductInList.ShipsImg
        product.FreeShippingImg = aProductInList.FreeShippingImg
        product.ExpiredDate = aProductInList.ExpiredDate
        If (isNewK11Product(product, list)) Then
            Try
                product.Categories.Add(category)
                efContext.Products.AddObject(product)
                efContext.SaveChanges()
                Return product.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct ERROR: " & ex.ToString)
                Return -1
            End Try
        Else
            Try
                Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) (m.Description = aProductInList.Description And m.PictureAlt = aProductInList.PictureAlt) Or m.FreeShippingImg = aProductInList.FreeShippingImg)
                updateProduct.Prodouct = product.Prodouct
                updateProduct.Price = product.Price
                updateProduct.Description = product.Description
                updateProduct.PictureUrl = product.PictureUrl
                updateProduct.PictureAlt = product.PictureAlt
                updateProduct.Currency = product.Currency
                updateProduct.LastUpdate = now
                updateProduct.Discount = product.Discount
                updateProduct.FreeShippingImg = product.FreeShippingImg
                updateProduct.ShipsImg = product.ShipsImg
                updateProduct.ExpiredDate = product.ExpiredDate


                '2014/2/21新增，防止一个产品有多个productCategory关系
                Dim updateCategory = updateProduct.Categories
                'Dim queryCate = From p In efContext.Products
                '                Where p.ProdouctID = updateProduct.ProdouctID
                '                Select p
                'Dim cate = queryCate.Single.Categories
                'Dim counter As Integer = updateCategory.Count
                For i As Integer = 0 To updateCategory.Count - 1
                    If updateCategory(0).CategoryID <> category.CategoryID Then 'product belong more than 1 category ,next one
                        Return 0
                    End If
                    updateCategory(0).Products.Remove(updateProduct) '移除ProductCategory表 这个产品对应的productid和categoryid绑定记录
                Next
                efContext.SaveChanges()

                If Not updateCategory.Contains(category) Then
                    category.Products.Add(updateProduct)  'ProductCategory表 绑定新的productid-categoryid
                End If
                efContext.SaveChanges()
                Return updateProduct.ProdouctID
            Catch ex As Exception
                Common.LogText("对不起，我又出错了2,InsertProduct ERROR2: " & ex.ToString)
                LogText("InsertProduct ERROR2: " & ex.ToString)
                Return -1
            End Try

        End If
    End Function

    ''' <summary>
    ''' 插入或更新product表，（即使有多个category包含此产品）
    ''' </summary>
    ''' <param name="aProductInList"></param>
    ''' <param name="now"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="existProducts"></param>
    ''' <returns></returns>
    Public Function InsertOrUpdateProduct(ByVal aProductInList As Product, ByVal now As DateTime, ByVal categoryId As Integer, ByVal existProducts As List(Of Product)) As Integer

        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = aProductInList.Prodouct
        product.Url = aProductInList.Url
        product.Price = aProductInList.Price
        product.PictureUrl = aProductInList.PictureUrl
        product.LastUpdate = now
        product.Description = aProductInList.Description
        product.SiteID = aProductInList.SiteID
        product.Currency = aProductInList.Currency
        product.PictureAlt = aProductInList.PictureAlt
        product.Discount = aProductInList.Discount
        product.ShipsImg = aProductInList.ShipsImg
        product.FreeShippingImg = aProductInList.FreeShippingImg
        product.ExpiredDate = aProductInList.ExpiredDate
        If (isNewProduct(product, existProducts)) Then
            Try
                product.Categories.Add(category) 'add to ProductCategory表?
                efContext.Products.AddObject(product)
                efContext.SaveChanges()
                Return product.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct ERROR: " & ex.ToString)
                Return -1
            End Try

        Else
            Try
                Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = aProductInList.Url And m.Prodouct = aProductInList.Prodouct)
                updateProduct.Prodouct = product.Prodouct
                updateProduct.Price = product.Price
                updateProduct.Description = product.Description
                updateProduct.PictureUrl = product.PictureUrl
                updateProduct.PictureAlt = product.PictureAlt
                updateProduct.Currency = product.Currency
                updateProduct.LastUpdate = now
                updateProduct.Discount = product.Discount
                updateProduct.FreeShippingImg = product.FreeShippingImg
                updateProduct.ShipsImg = product.ShipsImg
                updateProduct.ExpiredDate = product.ExpiredDate

                Dim updateCategory = updateProduct.Categories
                For i As Integer = 0 To updateCategory.Count - 1
                    If updateCategory(0).CategoryID <> category.CategoryID Then 'product belong more than 1 category ,next one
                        'Return 0
                    End If
                    updateCategory(0).Products.Remove(updateProduct) '移除ProductCategory表 这个产品对应的productid和categoryid绑定记录
                Next
                efContext.SaveChanges()

                If Not updateCategory.Contains(category) Then
                    category.Products.Add(updateProduct)  'ProductCategory表 绑定新的productid-categoryid
                End If
                efContext.SaveChanges()
                Return updateProduct.ProdouctID
            Catch ex As Exception
                LogText("InsertProduct ERROR2: " & ex.ToString)
                Return -1
            End Try

        End If
    End Function



    'Public Function InsertOneProduct(ByVal productName As String ,ByVal url As String ,ByVal price As Double , _
    '                                 ByVal discount As Double ,)
    ''' <summary>
    ''' 判断即将插入的Productd的URL是否在数据库中已经存在（siteid,days天内），如果存在，返回false
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function JudgeOneProduct(ByVal url As String, ByVal siteid As Integer, ByVal issueid As Long, ByVal days As Integer) As Boolean
        Dim issuesDate As DateTime = (From iss In efContext.Issues
                                      Where iss.IssueID = issueid
                                      Select iss.IssueDate).SingleOrDefault
        issuesDate = issuesDate.AddDays(-days)

        Dim query = (From p In efContext.Products
                     Where p.Url = url And p.LastUpdate > issuesDate And p.SiteID = siteid).Count
        If (query = 0) Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 判断即将插入的Productd的URL是否在数据库中已经存在（siteid），如果存在，返回Productid,不存在返回-1
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function JudgeOneProductOfSite(ByVal url As String, ByVal siteid As Integer) As Long

        Dim query = (From p In efContext.Products
                     Where p.Url = url And p.SiteID = siteid)
        If (query.Count = 0) Then
            Return -1
        Else
            Return query.SingleOrDefault.ProdouctID
        End If
    End Function

    ''' <summary>
    ''' 判断即将插入的Productd的URL是否在数据库中已经存在，如果存在，返回false Dora 2014-02-11
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isNewProduct(ByVal product As Product, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url.Trim() = product.Url.Trim()) And product.Prodouct = li.Prodouct Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 判断即将插入的Productd的url和description，对应Facebook的产品，不能判断Prodouct，因为明明相同的产品，每期Prodouct会自动根据日期命名，会变动，所以不能比较
    ''' </summary>
    ''' <param name="product"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    Private Function isNewK11Product(ByVal product As Product, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If li.FreeShippingImg = "10158611237485274" Then
                li.FreeShippingImg = li.FreeShippingImg
            End If
            If li.FreeShippingImg = product.FreeShippingImg Then
                Return False
            End If
            If (li.Url.Trim() = product.Url.Trim()) And product.Description = li.Description Then
                Return False
            End If

        Next
        Return True
    End Function


    Public Shared Function GetOneProductPath(ByVal siteid As Integer, ByVal plantype As String, ByVal cateid As Integer) As ProductPath
        Dim myprodPaht As New ProductPath()
        myprodPaht = (From pp In efContext.ProductPaths
                      Where pp.siteId = siteid AndAlso pp.planType = plantype AndAlso pp.prodcate = cateid
                      Select pp).FirstOrDefault()
        Return myprodPaht
    End Function
#End Region

#Region "Products_Issue表"
    ''' <summary>
    ''' 将产品添加到Products_Issue表中，并使用1个产品作为Promotion的内容
    ''' </summary>
    ''' <param name="listProductId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub InsertPOProductIssue(ByVal listProductId As List(Of Integer), ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal addString As String)
        'Dim queryPOPromotion = (From issue In efContext.Products_Issue
        '                        Join iss In efContext.Issues On issue.IssueID Equals iss.IssueID
        '                     Where issue.SectionID = "PO" AndAlso issue.SiteId = siteId AndAlso Not iss.Subject = ""
        '                     Order By issue.IssueID Descending
        '                     Select issue.ProductId).Take(8)
        Dim queryPOPromotion = (From issue In efContext.Products_Issue
                                Join iss In efContext.Issues On issue.IssueID Equals iss.IssueID
                                Where issue.SiteId = siteId AndAlso Not iss.Subject = "" AndAlso issue.SectionID = sectionId
                                Order By iss.IssueID Descending
                                Select issue.ProductId).Take(3)

        Dim listPOProductId As List(Of Long) = queryPOPromotion.ToList()
        Dim counter As Integer = 1
        For Each li In listProductId
            Dim proIssue As New Products_Issue()
            proIssue.ProductId = li
            proIssue.SiteId = siteId
            proIssue.IssueID = issueId
            'If (Not (listPOProductId.Contains(li)) AndAlso counter < 2) Then
            '    proIssue.SectionID = "CA"
            '    counter = counter + 1
            'Else
            '    proIssue.SectionID = "CA"
            'End If
            proIssue.SectionID = sectionId
            efContext.AddToProducts_Issue(proIssue)
        Next
        efContext.SaveChanges()
        For Each li In listProductId
            If (Not (listPOProductId.Contains(li))) Then
                Dim productName As String = efContext.Products.Single(Function(p) p.ProdouctID = li).Prodouct
                Dim myIssue = efContext.Issues.Single(Function(i) i.IssueID = issueId)
                myIssue.Subject = addString & productName & " ★ Weekly Deals"
                efContext.SaveChanges()
                Exit For
            End If
        Next


        '2013/4/5修改，不需要返回大图的产品信息
        ''返回需要更新大图的产品信息
        'Dim queryPOProduct = From issue In efContext.Products_Issue
        '                     Join p In efContext.Products On issue.ProductId Equals p.ProdouctID
        '                   Where issue.SectionID = "PO" AndAlso issue.IssueID = issueId
        '                   Select p
        'Dim bigPicProducts As List(Of Product) = queryPOProduct.ToList()
        'Return bigPicProducts
    End Sub

    ''' <summary>
    ''' 将产品添加到Products_Issue表中
    ''' </summary>
    ''' <param name="listProductId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub InsertProductIssue(ByVal listProductId As List(Of Integer), ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String)
        '2013/4/7添加
        Dim queryProIss = From pIssue In efContext.Products_Issue
                          Where pIssue.IssueID = issueId AndAlso pIssue.SiteId = siteId AndAlso pIssue.SectionID = sectionId
                          Select pIssue.ProductId
        'Dim listProIssue As List(Of Long) = queryProIss.ToList()'2013/4/10删除，务必要判断
        Dim listProIssue As New HashSet(Of Integer) '2013/4/15 添加，否则会插入Issue表中数据冲突
        For Each q In queryProIss
            listProIssue.Add(q)
        Next

        For Each li In listProductId
            Dim proIssue As New Products_Issue()
            proIssue.ProductId = li
            proIssue.SiteId = siteId
            proIssue.IssueID = issueId
            proIssue.SectionID = sectionId
            If Not (listProIssue.Contains(li)) Then
                efContext.AddToProducts_Issue(proIssue)
            End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 将N个产品每期轮番添加到Products_Issue表中
    ''' </summary>
    ''' <param name="listProductId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="iTakeCount">模板中产品的个数</param>
    ''' <remarks></remarks>
    Public Sub InsertProductIssue(ByVal listProductId As List(Of Integer), ByVal siteId As Integer, ByVal issueId As Integer,
                                  ByVal sectionId As String, ByVal iTakeCount As Integer, ByVal iTotalCount As Integer)
        Dim counter As Integer
        If (iTotalCount <= iTakeCount) Then
            counter = 0
        Else
            counter = iTotalCount - iTakeCount
        End If
        Dim queryLastProdIss = (From pIssue In efContext.Products_Issue
                                Where pIssue.SiteId = siteId AndAlso pIssue.SectionID = sectionId
                                Order By pIssue.IssueID Descending
                                Select pIssue.ProductId).Take(counter)
        Dim listLastProductId As New HashSet(Of Integer)
        For Each query In queryLastProdIss
            listLastProductId.Add(query)
        Next
        Dim listInsertProdId As New List(Of Integer)
        Dim productCounter As Integer = 0
        For Each li In listProductId
            If Not (listLastProductId.Contains(li)) Then
                If (productCounter = iTakeCount) Then 'change a bug of first
                    Exit For
                End If
                listInsertProdId.Add(li)
                productCounter = productCounter + 1
            End If
        Next
        InsertProductIssue(listInsertProdId, siteId, issueId, sectionId)
    End Sub

    ''' <summary>
    ''' 查询Products_Issue表中指定SectionID的产品
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SearchPromotionProduct(ByVal issueId As Integer, ByVal sectionId As String) As List(Of String)
        Dim listProductName As New List(Of String)
        Dim promotionProduct = From p In efContext.Products
                               Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                               Where pIssue.SectionID = sectionId AndAlso pIssue.IssueID = issueId
                               Select p.Prodouct
        listProductName = promotionProduct.ToList()
        Return listProductName
    End Function

    ''' <summary>
    ''' 获取Products_Issue表中SectionID为"PO"的产品名
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPromotionProduct(ByVal issueId As Integer) As String
        Dim promotionProduct = From p In efContext.Products
                               Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                               Where pIssue.SectionID = "PO" AndAlso pIssue.IssueID = issueId
                               Select p.Prodouct
        Dim strSubject As String = promotionProduct.Single()
        Return strSubject
    End Function

    ''' <summary>
    ''' 添加单个产品到Products_Issue表中
    ''' </summary>
    ''' <param name="productId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <remarks></remarks>
    Public Sub InsertSinglePIssue(ByVal productId As Integer, ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String)
        Try
            Dim pIssue As New Products_Issue
            pIssue.ProductId = productId
            pIssue.SiteId = siteId
            pIssue.IssueID = issueId
            pIssue.SectionID = sectionId
            efContext.AddToProducts_Issue(pIssue)
            efContext.SaveChanges()

        Catch ex As Exception
            LogText("InsertSinglePIssue()-->" & ex.ToString)
            Throw ex
        End Try


    End Sub

    ''' <summary>
    ''' 将获取的productIDList 中的productid对应IssueIdea插入到ProductIssue表中 Dora，2014-02-11
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="listProductId"></param>
    ''' <param name="iProIssueCount"></param>
    ''' <remarks></remarks>
    Public Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
                                    ByVal iProIssueCount As Integer)
        Try
            Dim queryProductId = (From pro In efContext.Products_Issue
                                  Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId
                                  Select pro.ProductId).ToList() '获取该issue中已经插入的产品id，同一个产品分属于不同的类别，避免同一个issue获取到相同的产品
            Dim i As Integer = 0
            For Each li In listProductId
                If i < iProIssueCount AndAlso Not (queryProductId.Contains(li)) Then
                    Try
                        Dim pIssue As New Products_Issue
                        pIssue.ProductId = li
                        pIssue.SiteId = siteId
                        pIssue.IssueID = issueId
                        pIssue.SectionID = sectionId
                        efContext.AddToProducts_Issue(pIssue)
                        efContext.SaveChanges()
                        i = i + 1
                    Catch ex As Exception
                        LogText("InsertProductsIssue()-->" & ex.ToString)
                        i = i - 1 'next one 
                    End Try

                End If
                If (queryProductId.Contains(li)) Then
                    LogText(String.Format(" duplicate product in the same issue ,jump ---where issueid={0},siteId={1},productid={2}", issueId, siteId, li))
                    i = i - 1
                End If
                If (i >= iProIssueCount) Then
                    Exit For
                End If
            Next

        Catch ex As Exception
            LogText("InsertProductsIssue()-->" & ex.ToString)
            Throw New Exception("error occured when insert into tabel:productIssue:" & ex.Message.ToString)
        End Try
    End Sub
#End Region


#Region "Products表和Products_Issue表联合查询"
    ''' <summary>
    ''' 获取第一个产品的ProductName
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFirstProductSubject(ByVal issueId As Integer) As String
        Dim product = From p In efContext.Products
                      Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                      Where pIssue.IssueID = issueId
                      Select p.Prodouct
        Dim subject As String = product.First()
        Return subject
    End Function

    ''' <summary>
    ''' 获取第一个产品的ProductName
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="issueId">产品section，CA，NE，DA等</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFirstProductSubject(ByVal issueId As Integer, ByVal section As String) As String
        Dim product = (From p In efContext.Products
                       Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                       Where (pIssue.IssueID = issueId And pIssue.SectionID = section)
                       Select p.Prodouct).FirstOrDefault()

        Dim subject As String = product
        Return subject
    End Function

    ''' <summary>
    ''' 获取某个分类下的前N个产品
    ''' </summary>
    ''' <param name="cateUrl"></param>
    ''' <param name="issueId"></param>
    ''' <param name="topN">前N期产品的个数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTopNCateProdUrl(ByVal cateUrl As String, ByVal issueId As Integer, ByVal topN As Integer, ByVal siteId As Integer)
        Dim listUrl As New List(Of String)
        Dim products = From p In efContext.Products
                       Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                       Where pIssue.IssueID < issueId AndAlso pIssue.SectionID = "CA" AndAlso p.SiteID = siteId
                       Order By pIssue.IssueID Descending
                       Select p
        Dim counter As Integer = 0  '查询上一期某个分类产品的计数器
        For Each p In products
            If (p.Categories.First.Url = cateUrl) Then
                listUrl.Add(p.Url)
            End If
            counter = counter + 1
            If (counter >= topN) Then
                Exit For
            End If
        Next
        Return listUrl
    End Function

    ''' <summary>
    ''' 获取某个Section下的前N期的产品URL
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="topN">页面的某一个块的产品个数-填充产品的个数</param>
    ''' <param name="sectionId">产品所属的SectionId，在表Sections中；如果为空，则获取所有</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTopNSectionProdUrl(ByVal issueId As Integer, ByVal topN As Integer, ByVal sectionId As String, ByVal siteId As Integer)
        Dim listUrl As New List(Of String)
        '如果sectionId为空，则获取前几期所有的已经发送过的产品
        If (String.IsNullOrEmpty(sectionId)) Then
            listUrl = (From p In efContext.Products
                       Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                       Join issue In efContext.Issues On pIssue.IssueID Equals issue.IssueID
                       Where pIssue.IssueID < issueId AndAlso p.SiteID = siteId _
                       AndAlso issue.SentStatus = "ES" AndAlso Not String.IsNullOrEmpty(issue.Subject)
                       Order By pIssue.IssueID Descending
                       Select p.Url).Take(topN).ToList()
        Else
            listUrl = (From p In efContext.Products
                       Join pIssue In efContext.Products_Issue On p.ProdouctID Equals pIssue.ProductId
                       Join issue In efContext.Issues On pIssue.IssueID Equals issue.IssueID
                       Where pIssue.IssueID < issueId AndAlso pIssue.SectionID = sectionId AndAlso p.SiteID = siteId _
                       AndAlso issue.SentStatus = "ES" AndAlso Not String.IsNullOrEmpty(issue.Subject)
                       Order By pIssue.IssueID Descending
                       Select p.Url).Take(topN).ToList()
        End If
        Return listUrl
    End Function
#End Region

#Region "Issues表"
    ''' <summary>
    ''' 把Subject内容插入到Issues表中
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub InsertIssueSubject(ByVal subject As String, ByVal issueId As Integer)
        Dim proIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        proIssue.Subject = subject
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 把Subject写入到Issue表中
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="subject"></param>
    ''' <remarks></remarks>
    Public Sub InsertIssueSubject(ByVal issueId As Integer, ByVal subject As String)
        Try
            Using EF As New EmailAlerterEntities
                Dim queryIssue As Issue = EF.Issues.Single(Function(i) i.IssueID = issueId)
                queryIssue.Subject = subject
                EF.SaveChanges()
            End Using
        Catch ex As Exception
            LogText("InsertIssueSubject()-->" & ex.ToString)
        End Try

        'Dim queryIssue = efContext.Issues.Single(Function(i) i.IssueID = issueId)
        'queryIssue.Subject = subject
        'efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 获取今年某个账户下面的某个PlanType发送的次数
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSentCount(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String)
        Dim count As Integer = 0
        count = efContext.Issues.Where(Function(issue) issue.SiteID = siteId AndAlso
                                           issue.PlanType = planType AndAlso issue.IssueID < issueId _
                                           AndAlso issue.SentStatus = "ES" AndAlso issue.IssueDate.Year = Now.Year).Count
        Return count
    End Function

    ''' <summary>
    ''' 获得邮件的subject, siteName Deals[New Arrivals] For MMM.yyyy.Vol.** ,Dora 2014-02-11
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal siteName As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal section As String)
        Dim preSubject As String
        If (section = "DA") Then
            preSubject = siteName & " Deals For " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (section = "NE") Then
            preSubject = siteName & " New Arrivals For " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")
        Dim query = From i In efContext.Issues
                    Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                    Select i
        Dim subject As String = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + 1).ToString.PadLeft(2, "0")
        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub
#End Region

#Region "Ads表"


    ''' <summary>
    ''' 使用正则表达式匹配所有的banner出来
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="adType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FetchBanner(ByVal url As String, ByVal adType As String)

    End Function


    ''' <summary>
    ''' 根据siteId返回Ads表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListAd(ByVal siteId As Integer) As List(Of Ad)
        Dim listAds As New List(Of Ad)
        listAds = efContext.Ads.Where(Function(ad) ad.SiteID = siteId).ToList()
        Return listAds
    End Function
    ''' <summary>
    ''' get exisit adid of site
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAdidForIssue(ByVal siteId As Integer, ByVal issueid As Long, ByVal count As Integer) As Long
        Dim listid As New List(Of Long)
        listid = (From a In efContext.Ads
                  Where a.SiteID = siteId
                  Select a.AdID).ToList
        Dim adids As New List(Of Long)
        adids = (From ai In efContext.Ads_Issue
                 Where (ai.SiteId = siteId And ai.IssueID < issueid)
                 Order By ai.IssueID Descending
                 Select ai.AdId).Take(count).ToList
        For Each id In listid
            If Not (adids.Contains(id)) Then
                Return id
            End If
        Next
    End Function

    Public Function GetAdidbyImgUrl(ByVal imgurl As String, ByVal siteid As Integer) As Long
        Dim query = (From a In efContext.Ads
                     Where a.SiteID = siteid And a.PictureUrl.Trim = imgurl
                     Select a.AdID)
        If (query.Count <> 0) Then
            Return query.Single
        Else
            Return -1
        End If
    End Function

    ''' <summary>
    ''' 根据siteId返回Ads表中的前N条数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="counter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTopNAds(ByVal siteId As Integer, ByVal counter As Integer) As List(Of Ad)
        Dim listAds As New List(Of Ad)
        listAds = efContext.Ads.Where(Function(ad) ad.SiteID = siteId).OrderByDescending(Function(o) o.Lastupdate).Take(counter).ToList()
        Return listAds
    End Function

    ''' <summary>
    ''' 根据siteId返回Ads表中的前N条数据的Ad Url
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="counter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTopNAdsUrl(ByVal siteId As Integer, ByVal counter As Integer) As List(Of String)
        Dim listAdsUrl As New List(Of String)
        listAdsUrl = (From ad In efContext.Ads
                      Where ad.SiteID = siteId AndAlso String.IsNullOrEmpty(ad.Type)
                      Order By ad.Lastupdate Descending
                      Select ad.Url).Take(counter).ToList()
        Return listAdsUrl
    End Function


    ''' <summary>
    ''' 插入数据到Ads表中
    ''' </summary>
    ''' <param name="insertListAds"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertAds(ByVal insertListAds As List(Of Ad), ByVal siteId As Integer) As List(Of Integer)
        Dim listAdId As New List(Of Integer)
        Try
            Dim queryBannerCategory As Category = efContext.Categories.Single(Function(c) c.SiteID = siteId AndAlso c.Category1 = "Banner")
            For Each li In insertListAds
                queryBannerCategory.Ads.Add(li) '添加一条记录到AdsCategory表中
                efContext.AddToAds(li)
                efContext.SaveChanges()
                listAdId.Add(li.AdID)
            Next
        Catch ex As Exception
            LogText("InsertAds error " & ex.ToString)
        End Try

        Return listAdId
    End Function
    ''' <summary>
    ''' 更新Ads表中的数据
    ''' </summary>
    ''' <param name="updateListAds"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateAds(ByVal updateListAds As List(Of Ad), ByVal siteId As Integer) As List(Of Integer)
        Dim listAdId As New List(Of Integer)
        Try
            For Each li In updateListAds
                Dim updateAd = efContext.Ads.Single(Function(a) a.Url = li.Url)
                If Not (String.IsNullOrEmpty(li.Ad1)) Then
                    updateAd.Ad1 = li.Ad1
                End If
                If Not (String.IsNullOrEmpty(li.PictureUrl)) Then
                    updateAd.PictureUrl = li.PictureUrl
                End If
                If Not (String.IsNullOrEmpty(li.Description)) Then
                    updateAd.Description = li.Description
                End If
                If Not (li.SizeHeight Is Nothing) Then
                    updateAd.SizeHeight = li.SizeHeight
                End If
                If Not (li.SizeWidth Is Nothing) Then
                    updateAd.SizeWidth = li.SizeWidth
                End If
                If Not (String.IsNullOrEmpty(li.Type)) Then
                    updateAd.Type = li.Type
                End If
                updateAd.Lastupdate = li.Lastupdate
                listAdId.Add(updateAd.AdID)
            Next
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("UpdateAds error " & ex.ToString)
        End Try

        Return listAdId
    End Function

    ''' <summary>
    ''' 插入或者更新一个产品到Ads表中
    ''' </summary>
    ''' <param name="ad"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrUpdateAd(ByVal ad As Ad) As Integer
        Try
            Dim updateAd As Ad = efContext.Ads.Where(Function(a) a.Url = ad.Url).Single()
            If Not (String.IsNullOrEmpty(ad.Ad1)) Then
                updateAd.Ad1 = ad.Ad1
            End If
            If Not (String.IsNullOrEmpty(ad.PictureUrl)) Then
                updateAd.PictureUrl = ad.PictureUrl
            End If
            If Not (String.IsNullOrEmpty(ad.Description)) Then
                updateAd.Description = ad.Description
            End If
            If Not (ad.SizeHeight Is Nothing) Then
                updateAd.SizeHeight = ad.SizeHeight
            End If
            If Not (ad.SizeWidth Is Nothing) Then
                updateAd.SizeWidth = ad.SizeWidth
            End If
            If Not (String.IsNullOrEmpty(ad.Type)) Then
                updateAd.Type = ad.Type
            End If
            updateAd.Lastupdate = ad.Lastupdate
            efContext.SaveChanges()
            Return updateAd.AdID
        Catch ex As Exception
            Dim queryBannerCategory As Category = efContext.Categories.Single(Function(c) c.SiteID = ad.SiteID AndAlso c.Category1 = "Banner")
            queryBannerCategory.Ads.Add(ad) '添加一条记录到AdsCategory表中
            efContext.AddToAds(ad)
            efContext.SaveChanges()
            Return ad.AdID
        End Try
    End Function

    'Public Function InsertOneAd(ByVal ad As Ad, ByVal categoryID As Integer) As Integer

    'End Function

    ''' <summary>
    ''' 将获得的Ad信息保存至Ads表中。,Droa,2014.02.08
    ''' </summary>
    ''' <param name="ad">ad即一个banner对象</param>
    ''' <param name="listAd"></param>
    ''' <param name="categoryID"></param>
    ''' <returns></returns>
    ''' <remarks>若不存在则new一个ad若存在则更新Hoyho2016/2/25</remarks>
    Public Function InsertAd(ByVal ad As Ad, ByVal listAd As List(Of Ad), ByVal categoryID As Integer) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryID Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim newAd As New Ad()
        If Not (String.IsNullOrEmpty(ad.Ad1)) Then
            newAd.Ad1 = ad.Ad1
        End If
        If Not (String.IsNullOrEmpty(ad.Description)) Then
            newAd.Description = ad.Description
        End If
        If Not (String.IsNullOrEmpty(ad.Type)) Then
            newAd.Type = ad.Type
        End If

        newAd.Url = ad.Url
        newAd.PictureUrl = ad.PictureUrl
        newAd.SizeHeight = ad.SizeHeight
        newAd.SizeWidth = ad.SizeWidth
        newAd.SiteID = ad.SiteID
        newAd.Lastupdate = ad.Lastupdate


        If (IsNewBanner(newAd, listAd)) Then
            Try
                newAd.Categories.Add(category)
                efContext.AddToAds(newAd)
                efContext.SaveChanges()
                Return newAd.AdID
            Catch ex As Exception
                LogText("InsertAd Error : " & ex.ToString)
                Return -1
            End Try

        Else
            Try
                Dim updateAd = efContext.Ads.FirstOrDefault(Function(m) m.Url = ad.Url AndAlso m.SiteID = ad.SiteID AndAlso m.PictureUrl = ad.PictureUrl)
                If Not (String.IsNullOrEmpty(ad.Ad1)) Then
                    updateAd.Ad1 = ad.Ad1
                End If
                If Not (String.IsNullOrEmpty(ad.Description)) Then
                    updateAd.Description = ad.Description
                End If
                If Not (String.IsNullOrEmpty(ad.Type)) Then
                    updateAd.Type = ad.Type
                End If
                updateAd.Url = ad.Url
                updateAd.PictureUrl = ad.PictureUrl
                updateAd.SizeHeight = ad.SizeHeight
                updateAd.SizeWidth = ad.SizeWidth
                updateAd.SiteID = ad.SiteID
                updateAd.Lastupdate = ad.Lastupdate

                '2014/2/24新增，防止一个Banner有多个adCategory关系
                Dim updateCategory = updateAd.Categories

                'Dim queryCate = From p In efContext.Products
                '                Where p.ProdouctID = updateProduct.ProdouctID
                '                Select p
                'Dim cate = queryCate.Single.Categories
                Dim counter As Integer = updateCategory.Count
                For i As Integer = 0 To counter - 1
                    updateCategory(0).Ads.Remove(updateAd)
                Next
                efContext.SaveChanges()

                If Not updateCategory.Contains(category) Then
                    category.Ads.Add(updateAd)
                End If
                efContext.SaveChanges()

                Return updateAd.AdID
            Catch ex As Exception
                LogText("InsertAd Error2 : " & ex.ToString)
                Return -1
            End Try

        End If
    End Function
    ''' <summary>
    ''' Insert one ad into Ads table,return its id
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="picurl"></param>
    ''' <param name="siteid"></param>
    ''' <param name="type"></param>
    ''' <param name="lastUpdate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function InsertSingleAd(ByVal newAd As Ad, ByVal categoryID As Integer) As Integer
        Try
            Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryID Select c
            Dim category As Category = queryCategory.FirstOrDefault()
            newAd.Categories.Add(category)
            efContext.AddToAds(newAd)
            efContext.SaveChanges()
            Return newAd.AdID
        Catch ex As Exception
            LogText("InsertSingleAd error ,categoryID=" & categoryID & ex.ToString)
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 判断即将插入的Ad的URL是否在数据库中已经存在，如果存在，返回false,Droa,2014.02.08
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsNewBanner(ByVal ad As Ad, ByVal list As List(Of Ad)) As Boolean
        For Each li In list
            If (li.Url.Trim() = ad.Url.Trim() And ad.PictureUrl = li.PictureUrl) Then
                Return False
            End If
        Next
        Return True
    End Function
#End Region

    ''' <summary>
    ''' 判断ad是否已存入数据库，存入返回adid，没存返回-1
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isADexist(ByVal PictureUrl As String, ByVal siteid As Integer) As Integer
        Dim query = (From a In efContext.Ads
                     Where a.PictureUrl = PictureUrl And a.SiteID = siteid)
        If query.Count = 0 Then
            Return -1
        Else
            Return query.Take(1).SingleOrDefault.AdID
        End If
    End Function



#Region "Ads_Issue表"
    ''' <summary>
    ''' 插入数据到Ads_Issue，每次轮番发送
    ''' </summary>
    ''' <param name="listAdsId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="adNumbers">在Promotion处插入产品的个数</param>
    ''' <param name="totalAds">在Ads表中Type="P"的产品的个数</param>
    ''' <remarks></remarks>
    Public Sub InsertAdsIssue(ByVal listAdsId As List(Of Integer), ByVal siteId As Integer, ByVal issueId As Integer, ByVal adNumbers As Integer, ByVal totalAds As Integer)
        Dim adId As New HashSet(Of Integer)
        Dim takeCounter As Integer
        If (totalAds <= adNumbers) Then
            takeCounter = 0
        Else
            takeCounter = totalAds - adNumbers
        End If
        Dim queryAdIssue = (From iss In efContext.Ads_Issue
                            Where iss.SiteId = siteId
                            Order By iss.IssueID Descending
                            Select iss).Take(takeCounter)
        For Each q In queryAdIssue
            adId.Add(q.AdId)
        Next
        Dim counter As Integer = 0
        For Each li In listAdsId
            Try
                If (Not (adId.Contains(li)) AndAlso counter < adNumbers) Then
                    Dim adIssue As New Ads_Issue()
                    adIssue.AdId = li
                    adIssue.SiteId = siteId
                    adIssue.IssueID = issueId
                    efContext.AddToAds_Issue(adIssue)
                    counter = counter + 1
                End If
                efContext.SaveChanges()
            Catch ex As Exception
                LogText("InsertAdsIssue()-->" & ex.ToString)
            End Try

        Next

    End Sub

    ''' <summary>
    ''' 将一条Banner的数据插入到Ads_Issue表中，轮番发送
    ''' </summary>
    ''' <param name="listAdsId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="totalNumber"></param>
    ''' <param name="insertNumber"></param>
    ''' <remarks></remarks>
    Public Sub InsertAdsIssue2(ByVal listAdsId As List(Of Integer), ByVal siteId As Integer, ByVal issueId As Integer,
                              ByVal totalNumber As Integer, ByVal insertNumber As Integer)
        'Dim listId As New List(Of Long)
        Dim listTopNAdsId As New List(Of Long)
        Dim takeNumber As Integer = totalNumber - insertNumber
        If (takeNumber <= 0) Then
            takeNumber = insertNumber
        End If
        LogText("takeNumber:" & takeNumber & ",totalNumber:" & totalNumber)
        listTopNAdsId = (From ad In efContext.Ads_Issue
                         Where ad.SiteId = siteId
                         Order By ad.IssueID Descending
                         Select ad.AdId).Take(takeNumber).ToList()
        If (listTopNAdsId.Count > 0) Then
            listTopNAdsId = listTopNAdsId.Take(takeNumber).ToList()
        End If
        For Each li In listAdsId
            If Not (listTopNAdsId.Contains(li)) Then '前N期如果没有发送，则这期发送
                Dim adIssue As New Ads_Issue
                adIssue.AdId = li
                adIssue.SiteId = siteId
                adIssue.IssueID = issueId
                efContext.AddToAds_Issue(adIssue)
                efContext.SaveChanges()
                Exit For
            End If
        Next
        'efContext.Ads_Issue.Where(Function(ad) ad.SiteId = siteId).OrderByDescending(Function(a) a.IssueID).Take(100)

    End Sub

    ''' <summary>
    ''' 添加一条记录到Ads_Issue表中
    ''' </summary>
    ''' <param name="adId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub InsertSingleAdsIssue(ByVal adId As Integer, ByVal siteId As Integer, ByVal issueId As Integer)
        Try
            Dim adIssue As New Ads_Issue
            adIssue.AdId = adId
            adIssue.SiteId = siteId
            adIssue.IssueID = issueId
            efContext.Ads_Issue.AddObject(adIssue)
            'efContext.AddToAds_Issue(adIssue)
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertSingleAdsIssue Error ，SiteID=" & siteId & " adId=" & adId & " issueId=" & issueId & " Detail:" & ex.ToString)
            Throw ex
        End Try

    End Sub


    ''' <summary>
    ''' 添加记录到ads_issue表，
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="listAdId"></param>
    ''' <param name="iProIssueCount"></param>
    ''' <remarks></remarks>
    Public Sub InsertAdsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal listAdId As List(Of Integer), ByVal iProIssueCount As Integer)
        Dim i As Integer = 0
        For Each li In listAdId
            If i < iProIssueCount Then
                Dim aIssue As New Ads_Issue
                aIssue.AdId = li
                aIssue.SiteId = siteId
                aIssue.IssueID = issueId
                efContext.AddToAds_Issue(aIssue)
                i = i + 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub
#End Region

#Region "Ads表和Ads_Issue表联合查询"
    ''' <summary>
    ''' 根据siteId 和IssueId查询Ads_Issue表中数据，获取Ads表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListAd(ByVal siteId As Integer, ByVal IssueId As Integer) As List(Of Ad)
        Dim listAds As New List(Of Ad)
        Dim queryAd = From iss In efContext.Ads_Issue
                      Join ad In efContext.Ads On iss.AdId Equals ad.AdID
                      Where iss.SiteId = siteId AndAlso iss.IssueID = IssueId
                      Select ad
        listAds = queryAd.ToList()
        Return listAds
    End Function

    ''' <summary>
    ''' 判断指定时间内Ad是否已经在邮件中发送,已发送，则返回True，否则false
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="adUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isAdSent(ByVal siteID As String, ByVal ad As Ad, ByVal StartDate As String, ByVal EndDate As String)
        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID
                        Select i.IssueID).ToList()

        Dim lastAdUrls As List(Of String)
        Dim lastAdPicUrls As List(Of String)
        For Each issues In lastIssueIDS
            lastAdUrls = (From p In efContext.Ads
                          Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                          Where pi.IssueID = issues And pi.SiteId = siteID
                          Select p.Url.Trim().ToLower()).ToList()

            lastAdPicUrls = (From p In efContext.Ads
                             Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                             Where pi.IssueID = issues And pi.SiteId = siteID
                             Select p.PictureUrl.Trim().ToLower()).ToList()

            If (lastAdUrls.Contains(ad.Url.Trim.ToLower) And lastAdPicUrls.Contains(ad.PictureUrl.Trim.ToLower)) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' 判断指定时间内Ad是否已经在邮件中发送,已发送，则返回True，否则false,判断重复的条件是跳转到地址的Url
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="adUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="planType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isAdSent(ByVal siteID As String, ByVal adUrl As String, ByVal StartDate As String, ByVal EndDate As String, ByVal planType As String)
        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID AndAlso (New String() {"HO", "HA", planType}).Contains(i.PlanType)
                        Select i.IssueID).ToList()

        Dim lastAdUrls As List(Of String)
        For Each issues In lastIssueIDS
            lastAdUrls = (From p In efContext.Ads
                          Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                          Where pi.IssueID = issues And pi.SiteId = siteID
                          Select p.Url.Trim().ToLower()).ToList()
            If (lastAdUrls.Contains(adUrl.Trim().ToLower())) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' 控制内获取的Ad在最近6次发送的issue不重复不重复,重复返回ad此次应发的adid，
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="adUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="planType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isAdSentInRecentSixIssue(ByVal siteID As String, ByVal adUrl As String, ByVal planType As String) As Integer
        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID AndAlso (New String() {"HO", "HA", planType}).Contains(i.PlanType)
                        Select i.IssueID
                        Order By IssueID Descending).Skip(1).Take(6).ToList()

        Dim lastAdUrls As List(Of String)
        For Each issues In lastIssueIDS
            lastAdUrls = (From p In efContext.Ads
                          Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                          Where pi.IssueID = issues And pi.SiteId = siteID
                          Select p.Url.Trim().ToLower()).ToList()
            If (lastAdUrls.Contains(adUrl.Trim().ToLower())) Then
                Return True
            End If
        Next
        Return False
    End Function


    ''' <summary>
    ''' 判断指定时间内Ad是否已经在邮件中发送,已发送，则返回True，否则false
    ''' </summary>
    ''' <param name="siteID"></param>
    ''' <param name="adUrl"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="planType"></param>
    ''' <param name="isImg">True：判断重复的条件是img的链接；false：判断重复的条件是跳转到地址的Url</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isAdSent(ByVal siteID As String, ByVal adUrl As String, ByVal StartDate As String, ByVal EndDate As String, ByVal planType As String, ByVal isImg As Boolean)
        Dim lastIssueIDS As List(Of Long)
        lastIssueIDS = (From i In efContext.Issues
                        Where i.IssueDate >= StartDate AndAlso i.IssueDate <= EndDate AndAlso i.SentStatus = "ES" AndAlso i.Subject <> "" AndAlso i.SiteID = siteID AndAlso (New String() {"HO", "HA", planType}).Contains(i.PlanType)
                        Select i.IssueID).ToList()

        Dim lastAdUrls As List(Of String)
        If (Not isImg) Then
            For Each issues In lastIssueIDS
                lastAdUrls = (From p In efContext.Ads
                              Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                              Where pi.IssueID = issues And pi.SiteId = siteID
                              Select p.Url.Trim().ToLower()).ToList()
                If (lastAdUrls.Contains(adUrl.Trim().ToLower())) Then
                    Return True
                End If
            Next
        Else
            For Each issues In lastIssueIDS
                lastAdUrls = (From p In efContext.Ads
                              Join pi In efContext.Ads_Issue On p.AdID Equals pi.AdId
                              Where pi.IssueID = issues And pi.SiteId = siteID
                              Select p.PictureUrl.Trim().ToLower()).ToList()
                If (lastAdUrls.Contains(adUrl.Trim().ToLower())) Then
                    Return True
                End If
            Next
        End If

        Return False
    End Function

    Public Sub updateAutomationUrl(ByVal siteid As Integer, ByVal url As String)
        Try
            Dim query = (From p In efContext.AutomationPlans
                         Where p.SiteID = siteid).SingleOrDefault
            query.URL = url
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("updateAutomationUrl()-->" & ex.ToString)
        End Try

    End Sub
#End Region

#Region "ContactLists_Issue"
    ''' <summary>
    ''' 把进箱率的添加到ContactLists_Issue表中
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub InsertInbox(ByVal issueId As Integer)
        Dim contactIssue As New ContactLists_Issue()
        contactIssue.IssueID = issueId
        contactIssue.ContactList = "inboxtest"
        efContext.AddToContactLists_Issue(contactIssue)
        efContext.SaveChanges()
    End Sub


    ''' <summary>
    ''' 将某个特定的List名加入到ContactLists_Issue表中
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="contactList"></param>
    ''' <remarks></remarks>
    Public Sub InsertContactList(ByVal IssueID As Integer, ByVal contactList As String, ByVal sendingStatus As String)
        Try
            Dim myContactList As New ContactLists_Issue()
            myContactList.IssueID = IssueID
            myContactList.ContactList = contactList
            myContactList.SendingStatus = sendingStatus
            efContext.AddToContactLists_Issue(myContactList)
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertContactList()-->" & ex.ToString())
        End Try

    End Sub

    ''' <summary>
    ''' 将一个或多个List名插入到发送表ContactLists_Issue表中
    ''' </summary>
    ''' <param name="IssueId"></param>
    ''' <param name="ContactListArr"></param>
    ''' <param name="sendingStatus"></param>
    ''' <remarks></remarks>
    Public Sub InsertContactList(ByVal IssueId As Integer, ByVal ContactListArr As String(), ByVal sendingStatus As String)
        Try
            For i As Integer = 0 To ContactListArr.Length - 1
                Dim myContactList As New ContactLists_Issue()
                myContactList.IssueID = IssueId
                myContactList.ContactList = ContactListArr(i)
                myContactList.SendingStatus = sendingStatus
                efContext.AddToContactLists_Issue(myContactList)
            Next
            efContext.SaveChanges()
        Catch ex As Exception
            LogText("InsertContactList error:" & "IssueId:" & IssueId & ex.ToString)
            Throw ex

        End Try


    End Sub

    ''' <summary>
    ''' 插入一条SplitContactList记录
    ''' </summary>
    ''' <param name="s"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertSplitContactList(ByVal s As SplitContactList)
        efContext.AddToSplitContactLists(s)
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 获得指定SiteId下的Split Contact Lists
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSplitContactLists(ByVal siteId As Integer) As List(Of SplitContactList)
        Dim splitContactLists As List(Of SplitContactList)
        Dim query = From s In efContext.SplitContactLists
                    Where s.SiteID = siteId
                    Select s
        splitContactLists = query.ToList()
        Return splitContactLists
    End Function
#End Region

#Region "RssSubscriptions表"
    ''' <summary>
    ''' 返回RssSubscription对象
    ''' </summary>
    ''' <param name="SiteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRssSubscription(ByVal SiteId As Integer) As RssSubscription
        'Dim efContext As New EmailAlerterEntities()
        Dim query = (From r In efContext.RssSubscriptions
                     Where r.SiteID = SiteId
                     Select r).Single()
        Return query
    End Function

    ''' <summary>
    ''' 返回最大的RssSubscriptions RssId
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaxRssId() As Integer
        Dim efContext As New EmailAlerterEntities()
        Dim maxRssId As Integer = efContext.RssSubscriptions.Max(Function(r) r.RssId)
        Return maxRssId
    End Function
#End Region

#Region "日志文件"
    ''' <summary>
    ''' 写错误日志到日志文件中
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <remarks></remarks>
    Public Shared Sub Log(ByVal Ex As Exception)
        Try
            LogText(Now & ", " & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
    Public Shared Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
#End Region

#Region "AutoPlan"
    Public Shared Function AddAutoPlan(ByVal autoSite As AutomationSite, ByVal autoPlan As AutomationPlan, ByVal templateInfo As TemplateInfo) As Integer
        efContext.AddToAutomationSites(autoSite)
        efContext.SaveChanges()
        autoPlan.SiteID = autoSite.siteid
        efContext.AddToAutomationPlans(autoPlan)
        templateInfo.SiteId = autoSite.siteid
        efContext.AddToTemplateInfoes(templateInfo)
        efContext.SaveChanges()
        Return autoPlan.SiteID
    End Function

    Public Shared Sub UpdateAutoPlan(ByVal autoPlan As AutomationPlan)
        Dim query = (From a In efContext.AutomationPlans
                     Where a.SiteID = autoPlan.SiteID
                     Select a).Single()
        query.IntervalDay = autoPlan.IntervalDay
        query.WeekDays = autoPlan.WeekDays
        query.SenderEmail = autoPlan.SenderEmail
        query.SenderName = autoPlan.SenderName
        efContext.SaveChanges()
    End Sub

    Public Shared Function GetAutoPlan(ByVal shopUrl As String) As List(Of AutomationPlan)
        Dim query = From a In efContext.AutomationPlans
                    Where a.URL = shopUrl
                    Select a
        Dim planlist As List(Of AutomationPlan) = query.ToList()
        If planlist.Count = 0 Then
            Return Nothing
        Else
            Return planlist
        End If
    End Function

    Public Shared Function GetAutoPlan(ByVal siteId As Integer) As List(Of AutomationPlan)
        Dim query = From a In efContext.AutomationPlans
                    Where a.SiteID = siteId
                    Select a
        Dim planlist As List(Of AutomationPlan) = query.ToList()
        If planlist.Count = 0 Then
            Return Nothing
        Else
            Return planlist
        End If
    End Function

    Public Shared Function GetAutoPlanById(ByVal siteId As Integer) As AutomationPlan
        Dim query = From a In efContext.AutomationPlans
                    Where a.SiteID = siteId
                    Select a
        Dim automationPlan As AutomationPlan = query.Single()
        Return automationPlan
    End Function


    ''' <summary>
    ''' 获取到一个目前没有使用的plantype，仅供HO,HA类型
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPlantype(ByVal siteid As Integer) As String
        Dim allPlantype As List(Of String) = (From p In efContext.PlanTypes
                                              Where p.PlanType1.Contains("HO") Or p.PlanType1.Contains("HA")
                                              Select p.PlanType1).ToList()
        Dim plantypes As List(Of String) = (From a In efContext.AutomationPlans
                                            Where a.SiteID = siteid
                                            Select a.PlanType).ToList()
        For Each item As String In allPlantype
            If Not (plantypes.Contains(item)) Then
                Return item
            End If
        Next
        Return ""
    End Function

    ''' <summary>
    ''' 获取到一个目前没有使用的plantype,根据是否是触发返回HP组别还是{"HO","HA"}组别
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPlantype(ByVal siteid As Integer, ByVal isTrigger As Boolean) As String
        Dim allPlantype As List(Of String)
        If (isTrigger) Then '触发类型
            allPlantype = (From p In efContext.PlanTypes
                           Where p.PlanType1.Contains("HP")
                           Select p.PlanType1).ToList()
        Else '基本类型
            allPlantype = (From p In efContext.PlanTypes
                           Where p.PlanType1.Contains("HO") Or p.PlanType1.Contains("HA")
                           Select p.PlanType1).ToList()
        End If

        Dim plantypes As List(Of String) = (From a In efContext.AutomationPlans
                                            Where a.SiteID = siteid
                                            Select a.PlanType).ToList()
        For Each item As String In allPlantype
            If Not (plantypes.Contains(item)) Then
                Return item
            End If
        Next
        Return ""
    End Function
#End Region

#Region "AutoSite"
    Public Shared Sub ReduceVolumn(ByVal sid As Integer, ByVal volumn As Long)
        Dim query = (From a In efContext.AutomationSites
                     Where a.siteid = sid
                     Select a).Single()
        query.volumn = query.volumn - volumn
        efContext.SaveChanges()
    End Sub

    Public Shared Sub IncreVolumn(ByVal sid As Integer, ByVal volumn As Long)
        Dim query = (From a In efContext.AutomationSites
                     Where a.siteid = sid
                     Select a).Single()
        query.volumn = query.volumn + volumn
        efContext.SaveChanges()
    End Sub

    Public Shared Function GetAutoSiteBySid(ByVal sid As Integer) As AutomationSite
        Dim query = (From a In efContext.AutomationSites
                     Where a.siteid = sid
                     Select a).Single()
        Return query
    End Function

    Public Function GetSiteIdBySpreadaccount(ByVal spreadAccount As String) As Integer
        spreadAccount = spreadAccount.Trim
        Dim autositeid As Integer = (From ausite In efContext.AutomationSites
                                     Where ausite.SpreadLogin = spreadAccount
                                     Select ausite.siteid).FirstOrDefault()
        Return autositeid
    End Function

    Public Function getAutoSites(ByVal spreadAccount As String) As List(Of AutomationSite)
        spreadAccount = spreadAccount.Trim
        Dim listAutoSites As List(Of AutomationSite) = (From autosite In efContext.AutomationSites
                                                        Where autosite.SpreadLogin = spreadAccount
                                                        Select autosite).ToList
        Return listAutoSites
    End Function

    Public Function getAutoSite(ByVal spreadAccount As String, ByVal siteid As Integer) As AutomationSite
        spreadAccount = spreadAccount.Trim
        Dim autoSite As AutomationSite = (From ausite In efContext.AutomationSites
                                          Where ausite.SpreadLogin = spreadAccount And ausite.siteid = siteid).FirstOrDefault()
        Return autoSite
    End Function
#End Region

#Region "Template"

    Public Function GetTemplate(ByVal templateID As Integer) As Template
        Dim temp As Template = (From t In efContext.Templates
                                Where t.Tid = templateID
                                Select t).FirstOrDefault
        Return temp
    End Function

    Public Function GetTemplate(ByVal templateID As Integer, ByVal spreadAccount As String) As Template
        Dim temp As Template = (From t In efContext.Templates
                                Join autosite In efContext.AutomationSites On t.SiteID Equals autosite.siteid
                                Where t.Tid = templateID And autosite.SpreadLogin = spreadAccount
                                Select t).FirstOrDefault
        Return temp
    End Function

    Public Function GetTemplates(ByVal spreadAccount As String) As List(Of Template)
        spreadAccount = spreadAccount.Trim
        Dim listTemplate As List(Of Template) = (From t In efContext.Templates
                                                 Join autosite In efContext.AutomationSites On t.SiteID Equals autosite.siteid
                                                 Where (autosite.SpreadLogin = spreadAccount)
                                                 Select t).ToList()
        Return listTemplate
    End Function


    Public Function GetCommonTemplates(ByVal tType As String) As List(Of Template)
        tType = tType.Trim
        Dim listTemplate As List(Of Template) = (From t In efContext.Templates
                                                 Where t.tType = tType
                                                 Select t).ToList()
        Return listTemplate
    End Function

    ''' <summary>
    ''' 同一个店铺下是否存在此模板名，存在返回true，否则返回false
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="tempName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isTempNameExist(ByVal siteid As Integer, ByVal tempName As String) As Boolean
        Dim temp As Template = (From t In efContext.Templates
                                Where t.SiteID = siteid AndAlso t.TemplateName = tempName
                                Select t).FirstOrDefault()
        If (temp Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function AddTemp(ByVal temp As Template)
        temp.LatestModifyTime = DateTime.Now
        efContext.Templates.AddObject(temp)
        efContext.SaveChanges()
    End Function

    Public Function UpdateTemp(ByVal temp As Template, ByVal tid As Integer)
        Dim update As Template = efContext.Templates.FirstOrDefault(Function(t) t.Tid = tid)
        update.UserName = temp.UserName
        update.TemplateName = temp.TemplateName
        update.Contents = temp.Contents
        update.SiteID = temp.SiteID
        update.LatestModifyTime = DateTime.Now
        efContext.SaveChanges()
    End Function

    Public Function DeleteTemp(ByVal temp As Template, ByVal tid As Integer) As Boolean
        Try
            Dim query = (From a In efContext.AutomationPlans  'autoplan所有引用此Tid的改为默认
                         Where a.TemplateID = tid
                         Select a).ToList()
            For Each q In query
                q.TemplateID = 115
                efContext.SaveChanges()
            Next
            Dim target As Template = efContext.Templates.FirstOrDefault(Function(t) t.Tid = tid)
            efContext.Templates.DeleteObject(target)
            efContext.SaveChanges()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region

#Region "TemplateInfo"
    Public Shared Sub AddTemplateInfo(ByVal templateInfo As TemplateInfo)
        efContext.AddToTemplateInfoes(templateInfo)
        efContext.SaveChanges()
    End Sub

    Public Shared Sub UpdateTemplateInfo(ByVal templateInfo As TemplateInfo)
        Dim query = (From a In efContext.TemplateInfoes
                     Where a.SiteId = templateInfo.SiteId
                     Select a).Single()
        query.PlanType = templateInfo.PlanType
        query.SortedCates = templateInfo.SortedCates
        query.DisplayCount = templateInfo.DisplayCount
        query.Modified = templateInfo.Modified
        efContext.SaveChanges()
    End Sub

    Public Shared Function GetTemplateInfo(ByVal siteId As Integer)
        Dim query = From a In efContext.TemplateInfoes
                    Where a.SiteId = siteId
                    Select a
        Dim templist As List(Of TemplateInfo) = query.ToList()
        If templist.Count = 0 Then
            Return Nothing
        Else
            Return templist
        End If
    End Function
#End Region

#Region "Tempalte library"
    Public Shared Sub AddTemplateLibrary(ByVal templibra As TemplateLibrary)
        efContext.AddToTemplateLibraries(templibra)
        efContext.SaveChanges()
    End Sub

    Public Shared Function GetTemplateLibrary(ByVal tid As Long)
        Dim query = From t In efContext.TemplateLibraries
                    Where t.Tid = tid
                    Select t
        Dim templist As List(Of TemplateLibrary) = query.ToList()
        If templist.Count = 0 Then
            Return Nothing
        Else
            Return templist
        End If
    End Function
#End Region


#Region "Subejct"

    ''' <summary>
    ''' 从数据库Promotion产品中获得并设置Subject信息如个性化，刊号，日期等
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks>可以新增自定义标签，并在此处实现</remarks>
    Public Sub SetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal categoryName As String, ByVal productListOfCate As List(Of String), ByVal planType As String, ByVal preSubjdect As String)
        Dim subject As String

        'Dim productList As List(Of String) = GetProductsByCateId(issueId, cateName, siteId)
        Dim allIssues = From i In efContext.Issues
                        Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                        Select i  '所有成功的issue记录

        If productListOfCate.Count > 0 AndAlso Not (String.IsNullOrEmpty(productListOfCate.Item(0))) Then
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", productListOfCate.Item(0))
        Else
            subject = preSubjdect.Replace("[FIRST_PRODUCT]", "")
        End If
        If productListOfCate.Count > 1 AndAlso Not (String.IsNullOrEmpty(productListOfCate.Item(1))) Then
            subject = subject.Replace("[SECOND_PRODUCT]", productListOfCate.Item(1))
        Else
            subject = subject.Replace("[SECOND_PRODUCT]", "")
        End If


        subject = subject.Replace("[VOL_NUMBER]", (allIssues.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", categoryName.Trim)
        subject = subject.Replace("[YYYY]", Now.Year.ToString)
        subject = subject.Replace("[MM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))
        subject = subject.Replace("[MMM]", Now.ToString("MMMM", New Globalization.CultureInfo("en-US")))

        InsertIssueSubject(issueId, subject)
    End Sub

#End Region


#Region "SeachContacts"
    ''' <summary>
    ''' 调用spreadApi SearchContacts（） 函数，创建一个未点击任何分类的联系人列表 
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="saveListName"></param>
    ''' <param name="daySpan"></param>
    ''' <param name="strategy"></param>
    ''' <param name="sendStatus"></param>
    ''' <param name="loginEmail"></param>
    ''' <param name="appId"></param>
    ''' <remarks></remarks>
    Public Function CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal saveListName As String, ByVal daySpan As Integer,
                                         ByVal strategy As ChooseStrategy, ByVal sendStatus As String, ByVal loginEmail As String, ByVal appId As String) As Integer
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(-daySpan).ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.Service = New SpreadWebReference.Service()
        mySpread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
        mySpread.Timeout = 2400000
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
            Return count
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
        'Dim arrContactList As String() = New String() {saveListName}
        'Dim helper As New EFHelper
        'helper.InsertContactList(issueId, saveListName, sendStatus)
    End Function

    ''' <summary>
    ''' 调用spreadApi SearchContacts（） 函数，创建一个点击了指定分类的联系人列表 
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="saveListName"></param>
    ''' <param name="daySpan"></param>
    ''' <param name="strategy"></param>
    ''' <param name="sendStatus"></param>
    ''' <param name="loginEmail"></param>
    ''' <param name="appId"></param>
    ''' <remarks></remarks>
    Public Function CreateContactList(ByVal siteId As Integer, ByVal issueId As Integer, ByVal planType As String, ByVal categoryName As String, ByVal saveListName As String,
                                         ByVal daySpan As Integer, ByVal strategy As ChooseStrategy, ByVal sendStatus As String, ByVal loginEmail As String, ByVal appId As String) As Integer
        Dim categoryId As String = EFHelper.GetCategoryId(siteId, categoryName.Trim())
        Dim QuerySubscriber As New QuerySubscriber
        QuerySubscriber.Favorite = categoryId
        QuerySubscriber.Strategy = strategy
        QuerySubscriber.CountryList = New String() {}
        QuerySubscriber.StartDate = Date.Now.AddDays(-daySpan).ToString("yyyy-MM-dd")
        QuerySubscriber.EndDate = Date.Now.ToString("yyyy-MM-dd")
        Dim CriteriaString As String = QuerySubscriber.ToJsonString
        Dim mySpread As SpreadWebReference.Service = New SpreadWebReference.Service()
        mySpread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
        mySpread.Timeout = 2400000
        Dim count As Integer = -1
        Try
            count = mySpread.SearchContacts(loginEmail, appId, CriteriaString, Integer.MaxValue, saveListName, True)
            Return count
        Catch ex As Exception
            'Throw New Exception(ex.ToString())
            Common.LogText(ex.ToString() & "请检查收件人名单是否已存在")
            If ex.ToString().Contains("Contact List Name already exist") Then
                Return 1
            Else
                Return count
            End If
        End Try
        'Dim arrContactList As String() = New String() {saveListName}
    End Function

#End Region

#Region "SentLogs表"
    Public Shared Sub InsertSentLog(ByVal siteId As Integer, ByVal planType As String, ByVal logText As String, ByVal lastSentTime As DateTime)
        Try
            Dim sentLog As New SentLog()
            sentLog.SiteID = siteId
            sentLog.PlanType = planType
            sentLog.Subject = logText
            sentLog.LastSentAt = lastSentTime
            efContext.AddToSentLogs(sentLog)
            efContext.SaveChanges()
        Catch ex As Exception
            EFHelper.LogText("InsertContactList()-->" & ex.ToString())
        End Try

    End Sub
#End Region

#Region "BannerRegex以及ProductPaths"

    ''' <summary>
    ''' 获取第一条bannerRegex记录
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="plantype"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBannerRegex(ByVal siteid As Integer, ByVal plantype As String) As BannerRegex
        Dim myBannerRegex As BannerRegex = (From b In efContext.BannerRegexes
                                            Where b.siteId = siteid AndAlso b.planType = plantype
                                            Select b).FirstOrDefault()
        Return myBannerRegex
    End Function

    Public Function UpdateBannerRegex(ByVal myBannerRegex As BannerRegex)
        Dim siteid As Integer = myBannerRegex.siteId
        Dim planType As String = myBannerRegex.planType.Trim
        Dim oldBannerRegex As BannerRegex = GetBannerRegex(siteid, planType)
        If (oldBannerRegex Is Nothing) Then
            efContext.BannerRegexes.AddObject(myBannerRegex)
        Else
            Dim toupdate As BannerRegex = (From b In efContext.BannerRegexes
                                           Where b.ID = oldBannerRegex.ID
                                           Select b).FirstOrDefault()
            toupdate.srcRegText = myBannerRegex.srcRegText
            toupdate.pageEncoding = myBannerRegex.pageEncoding
            toupdate.noRepeatSentDays = myBannerRegex.noRepeatSentDays
            toupdate.hrefRegText = myBannerRegex.hrefRegText
            toupdate.cookie = myBannerRegex.cookie
            toupdate.bannerStartIndex = myBannerRegex.bannerStartIndex
            toupdate.bannerRegText = myBannerRegex.bannerRegText
            toupdate.bannerFormUrl = myBannerRegex.bannerFormUrl
            toupdate.bannerEndIndex = myBannerRegex.bannerEndIndex
        End If
        efContext.SaveChanges()

    End Function

    Public Function GetProductPathGroup(ByVal siteid As Integer, ByVal plantype As String) As List(Of ProductPath)
        plantype = plantype.Trim()
        Dim listProductPath As List(Of ProductPath) = (From p In efContext.ProductPaths
                                                       Where p.siteId = siteid AndAlso p.planType = plantype
                                                       Select p).ToList()
        Return listProductPath
    End Function



    ''' <summary>
    ''' 获取cates相关联的productPath记录
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="plantype"></param>
    ''' <param name="cates">格式如：'Cate1','Cate2','Cate2'</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProductPaths(ByVal siteid As Integer, ByVal plantype As String, ByVal cates As String) As List(Of CateProdpath)
        Dim myListProductPath As List(Of CateProdpath) = efContext.GetProdpathbyCates(cates, siteid, plantype).ToList()
        Return myListProductPath
    End Function


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="listproductPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateProductPathGroup(ByVal listproductPath As List(Of ProductPath))
        If (listproductPath.Count > 0) Then
            Dim siteid As Integer = listproductPath(0).siteId
            Dim planType As String = listproductPath(0).planType.Trim

            Dim addProdpaths As New List(Of ProductPath)
            For Each item In listproductPath
                Dim updateProdpath As ProductPath = efContext.ProductPaths.FirstOrDefault(Function(p) p.siteId = siteid And p.planType = planType And p.prodcate = item.prodcate)
                If (updateProdpath Is Nothing) Then
                    efContext.ProductPaths.AddObject(item)
                Else
                    updateProdpath.fileType = item.fileType
                    updateProdpath.cateParam = item.cateParam
                    updateProdpath.prodPath = item.prodPath
                    updateProdpath.productPath1 = item.productPath1
                    updateProdpath.productAttri = item.productAttri
                    updateProdpath.urlPath = item.urlPath
                    updateProdpath.urlAttri = item.urlAttri
                    updateProdpath.pricePath = item.pricePath
                    updateProdpath.priceAttri = item.priceAttri
                    updateProdpath.discountPath = item.discountPath
                    updateProdpath.discountAttri = item.discountAttri
                    updateProdpath.salesPath = item.salesPath
                    updateProdpath.salesAttri = item.salesAttri
                    updateProdpath.pictureUrlPath = item.pictureUrlPath
                    updateProdpath.pictureUrlAttri = item.pictureUrlAttri
                    updateProdpath.pictureAltPath = item.pictureAltPath
                    updateProdpath.pictureAltAttri = item.pictureAltAttri
                    updateProdpath.descriptionPath = item.descriptionPath
                    updateProdpath.descriptionAttri = item.descriptionAttri
                    updateProdpath.currencyChar = item.currencyChar
                    updateProdpath.prodDisplayCount = item.prodDisplayCount
                    updateProdpath.cookie = item.cookie
                    updateProdpath.pageEncoding = item.pageEncoding
                    updateProdpath.noRepeatSentDays = item.noRepeatSentDays
                    updateProdpath.issueDate = item.issueDate
                    updateProdpath.issueDateAttri = item.issueDateAttri
                    updateProdpath.validityPeriod = item.validityPeriod
                End If
            Next

            Dim allProdpaths As List(Of ProductPath) = GetProductPathGroup(siteid, planType)
            Dim flag As Boolean = False
            Dim toDeleteids As New List(Of Int64)
            If (allProdpaths.Count > listproductPath.Count) Then
                For Each item In allProdpaths
                    flag = False
                    For Each newitem In listproductPath
                        If (item.prodcate = newitem.prodcate) Then
                            flag = True
                        End If
                    Next
                    If Not (flag) Then
                        toDeleteids.Add(item.ID)
                    End If
                Next

                For Each itemid In toDeleteids
                    efContext.ProductPaths.DeleteObject(efContext.ProductPaths.FirstOrDefault(Function(p) p.ID = itemid))
                Next
            End If

            efContext.SaveChanges()
        End If
    End Function


#End Region

#Region "Taobao&Tmall"
    ''' <summary>
    ''' 获取banner图片及对应url(仅适用tmall/taobao)
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="Issueid"></param>
    ''' <param name="categoryName">banner所属分类名</param>
    ''' <param name="planType"></param>
    ''' <param name="pageUrl">banner图所在链接</param>
    ''' <param name="bannerImgIndex">banner正则匹配结果的索引</param>
    ''' <remarks></remarks>
    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String,
                     ByVal pageUrl As String, ByVal bannerImgIndex As Integer)

        Dim ad As New Ad()
        Dim imgRegex As String = "<img.*?(?:src="").*?(?:>|\/>)"
        Try
            Dim txtDocHtml As String = EFHelper.GetHtmlStringByUrlAli(pageUrl)
            Dim mCollection As MatchCollection = Regex.Matches(txtDocHtml, imgRegex)
            Dim matched As Match = mCollection(bannerImgIndex - 1)
            If (matched.Success) Then
                Dim bannerImg As String = mCollection(bannerImgIndex - 1).Value
                Dim srcRegex As String = "data-ks-lazyload=[\'\""]?([^\'\""]*)[\'\""]?"

                Dim srcMatched As Match = Regex.Match(bannerImg, srcRegex)
                If (srcMatched.Success) Then
                    ad.PictureUrl = srcMatched.Groups(1).Value.Trim

                    Dim ahrefReg As String = "<a.*?href=""(.*?)"".*?(?:" & ad.PictureUrl & ").*?(?:>|\/>)"
                    Dim ahrefMatched As Match = Regex.Match(txtDocHtml, ahrefReg)
                    If (ahrefMatched.Success) Then
                        ad.Url = ahrefMatched.Groups(1).Value.Trim
                    Else
                        ad.Url = pageUrl
                    End If
                End If
            End If
        Catch ex As Exception
            LogText("获取banner失败" & "siteId=" & siteId & "imgRegex=" & imgRegex)
        End Try


        If Not String.IsNullOrEmpty(ad.PictureUrl) Then '对banner不重复时长未做限定
            ad.Url = AddHttpForAli(ad.Url)
            ad.PictureUrl = AddHttpForAli(ad.PictureUrl)
            ad.SiteID = siteId
            ad.Lastupdate = Now

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = GetListAd(siteId)
            Dim categoryId As Integer = EFHelper.GetCategoryId(siteId, categoryName)
            Dim adid As Integer = InsertAd(ad, listAd, categoryId)
            If adid > 0 Then
                InsertSingleAdsIssue(adid, siteId, Issueid)
            End If

        End If
    End Sub

    ''' <summary>
    ''' 获取banner图片及对应url(仅适用tmall/taobao)
    ''' 根据[AutomationPlan]表的bannerImgex，bannerImgIndex和pageurl获取banner图片
    ''' </summary>
    ''' 
    ''' <param name="siteId"></param>
    ''' <param name="Issueid"></param>
    ''' <param name="categoryName">banner所属分类名</param>
    ''' <param name="planType"></param>
    ''' <param name="pageUrl">banner图所在链接</param>
    ''' <param name="bannerImgRegex">匹配banner的正则表达式</param>
    ''' <param name="bannerImgIndex">banner正则匹配结果的索引</param>
    ''' <remarks>2016.2.25Hoyho注</remarks>
    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String,
                     ByVal pageUrl As String, ByVal bannerImgRegex As String, ByVal bannerImgIndex As Integer)

        Dim ad As New Ad()

        Try
            Dim txtDocHtml As String = GetHtmlStringByUrlAli(pageUrl)
            Dim mCollection As MatchCollection = Regex.Matches(txtDocHtml, bannerImgRegex)
            If bannerImgIndex >= mCollection.Count() Then
                bannerImgIndex = 0
            End If
            If Not (IsGoodBanner(bannerImgIndex, mCollection, siteId, planType)) Then '该banner不适合
                For i As Integer = 0 To mCollection.Count - 1
                    If IsGoodBanner(i, mCollection, siteId, planType) Then
                        bannerImgIndex = i
                        Exit For
                    End If
                    'no suitable banner
                Next
            End If
            Dim matched As Match = mCollection(bannerImgIndex)
            If (matched.Success) Then
                'Dim bannerImg As String = matched.Groups(1).Value
                'Dim srcRegex As String = "data-ks-lazyload=[\'\""]?([^\'\""]*)[\'\""]?"

                'Dim srcMatched As Match = Regex.Match(bannerImg, srcRegex)
                'If (srcMatched.Success) Then
                ad.PictureUrl = matched.Groups(1).Value

                'comment followed 6 lines debug#
                'Dim ahrefReg As String = "<a.*?href=""(.*?)"".*?(?:" & ad.PictureUrl & ").*?(?:>|\/>)"  '此正则存在巨大回溯，严重影响性能，耗时1.5h++，故换一条并且观察
                Dim ahrefReg As String = "<a.*?href=""(?<1>[^""]+)"".*?(?:" & ad.PictureUrl & ").*?(?:>|\/>)"
                Dim ahrefMatched As Match = Regex.Match(txtDocHtml, ahrefReg)
                If (ahrefMatched.Success) Then
                    ad.Url = ahrefMatched.Groups(1).Value.Trim
                Else
                    ad.Url = pageUrl
                End If

                'End If
            End If
        Catch ex As Exception
            LogText("获取banner失败" & "siteId=" & siteId & "imgRegex=" & bannerImgRegex)
        End Try

        If Not String.IsNullOrEmpty(ad.PictureUrl) Then '对baner不重复时长未做限定
            ad.Url = AddHttpForAli(ad.Url)
            ad.PictureUrl = AddHttpForAli(ad.PictureUrl)
            ad.SiteID = siteId
            ad.Lastupdate = Now

            Dim listAd As New List(Of Ad)
            listAd = GetListAd(siteId) '获取该站的的Ads表中的所有ad，根据siteid返回Ads表列表
            Dim categoryId As Integer = EFHelper.GetCategoryId(siteId, categoryName) '根据siteId, categoryName从Categories表获取分类ID，categoryId是初始化时自分配的
            Dim adid As Integer = InsertAd(ad, listAd, categoryId) 'ad信息保存到Ads表
            If adid > 0 Then
                InsertSingleAdsIssue(adid, siteId, Issueid) '添加记录到Ads_Issue表
            End If

        End If
    End Sub

    ''' <summary>
    '''  callback of generate  banner thumbnail to scan every pixel of a picture
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ThumbnailCallback() As Boolean
        LogText("GetThumbnailImageAbort,exit and try next banner")
        Return False
    End Function


    ''' <summary>
    ''' 判断当前index是否匹配最佳尺寸banner，是则返回true，并更新数据库该BannerIndex，否则只返回false
    ''' </summary>
    ''' <param name="index">数据库bannerImgIndex</param>
    ''' <param name="mCollection">当前正则匹配结果集</param>
    ''' <param name="siteID">用于更新AutomationPlan</param>
    ''' <param name="planType">用于更新AutomationPlan</param>
    ''' <returns>匹配则返回则返回true</returns>
    ''' <remarks></remarks>
    Public Function IsGoodBanner(ByVal index As Integer, ByVal mCollection As MatchCollection, ByVal siteID As Integer, ByVal planType As String) As Boolean
        Dim height As Double = 0
        Dim width As Double = 0
        Dim matched As Match = mCollection(index)
        If matched.Success Then
            Dim imgurl = addDominForUrl("https:", matched.Groups(1).Value.Trim)
            Dim abortCallback As New System.Drawing.Bitmap.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
            Dim imgRequst As WebRequest = System.Net.WebRequest.Create(imgurl)
            Dim originalImage As System.Drawing.Image = System.Drawing.Image.FromStream(imgRequst.GetResponse().GetResponseStream())
            Dim thumbnails As System.Drawing.Bitmap = originalImage.GetThumbnailImage(100, 80, abortCallback, IntPtr.Zero)
            Dim color As System.Drawing.Color
            Dim colorDictionary As Dictionary(Of System.Drawing.Color, Integer) = New Dictionary(Of System.Drawing.Color, Integer)
            originalImage.Save(System.AppDomain.CurrentDomain.BaseDirectory.ToString + "banneroriginal")
            thumbnails.Save(System.AppDomain.CurrentDomain.BaseDirectory.ToString + "bannerthumnail")
            Dim watch As Stopwatch = Stopwatch.StartNew()
            For y As Integer = 0 To thumbnails.Height - 1   'Count all pixs with dictionary
                For x As Integer = 0 To thumbnails.Width - 1
                    color = thumbnails.GetPixel(x, y)
                    If Not colorDictionary.ContainsKey(color) Then
                        colorDictionary.Add(color, 1)
                    Else
                        colorDictionary.Item(color) = colorDictionary.Item(color) + 1
                    End If
                Next
            Next
            watch.Stop()
            Dim processTime As Long = watch.ElapsedMilliseconds
            'sort colorDictionary by value to get max pixel 
            colorDictionary = (From entry In colorDictionary Order By entry.Value Descending Select entry).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            Dim maxPixelRate As Double = colorDictionary.Values(0) / (thumbnails.Width * thumbnails.Height)  '单一像素最大占比，通常为白色像素最大占比
            height = originalImage.Size.Height
            width = originalImage.Size.Width

            If (height > 300 And width > 700 And height / width > 0.26 And height / width < 0.74) Then  'Level1 size filter
                If (colorDictionary.Count > 10 And maxPixelRate < 0.85) Then  'level2 pixel filter, simple color or wide range pure color are not good banner
                    Try 'update bannner index
                        Dim autoPlan As AutomationPlan = efContext.AutomationPlans.Where(Function(c) c.PlanType = planType AndAlso c.SiteID = siteID).Single()
                        autoPlan.bannerIndex = index
                    Catch ex As Exception
                        'efContext.AddToAutomationPlans(autoPlan)
                    End Try
                    efContext.SaveChanges()
                    Return True  '找到合适bannerindex，返回结果
                Else
                    Return False
                End If
            Else
                Return False
            End If

        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' 获取X月X日类型的新品链接及文字(subject)
    ''' </summary>
    ''' <param name="shopUrl"></param>
    ''' <returns>list.iten(0)是新品链接，list.item(1)是主题</returns>
    ''' <remarks></remarks>
    Public Function GetXinpinUrlandSubject(ByVal shopUrl As String) As List(Of String)
        Dim urlandSubject As New List(Of String)
        Dim categoryDoc As HtmlDocument
        Try
            categoryDoc = EFHelper.GetHtmlDocByUrlTmall(shopUrl)
        Catch ex As Exception
            categoryDoc = EFHelper.GetHtmlDocByUrlTmall(shopUrl)
        End Try
        'Dim productLines As HtmlNodeCollection = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")
        Dim xinpinLi As HtmlNode = categoryDoc.DocumentNode.SelectNodes("//li[@class='cat snd-cat  ']")(0)
        Dim xinpinUrl As String = xinpinLi.SelectSingleNode("h4/a").GetAttributeValue("href", "").Trim()
        xinpinUrl = AddHttpForAli(xinpinUrl)
        urlandSubject.Add(xinpinUrl)
        Dim subject As String = xinpinLi.SelectSingleNode("h4/a").InnerText().Trim()
        urlandSubject.Add(subject)
        Return urlandSubject
    End Function


    ''' <summary>
    ''' 获取指定链接(tmall/taobao)上的产品，并将对应数据插入到product及productsIssue表，无返回值
    ''' </summary>
    ''' <param name="url">产（新）品来源url</param>
    ''' <param name="categoryName">产品类别名称</param>
    ''' <param name="section">CA、NE、DA等</param>
    ''' <param name="planType"></param>
    ''' <param name="productCount">需要获取的产品个数</param>
    ''' <param name="startTime">产品不重复的开始时间</param>
    ''' <param name="endTime">产品不重复的结束时间</param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <returns>获取到的产品均插入到product及productsIssue表，无返回值</returns>
    ''' <remarks></remarks>
    Public Function FetchProduct(ByVal url As String, ByVal categoryName As String, ByVal section As String, ByVal planType As String,
                                 ByVal productCount As Integer, ByVal startTime As DateTime, ByVal endTime As DateTime, ByVal siteId As Integer,
                                 ByVal issueId As Integer)
        Dim listProducts As New List(Of Product)
        Dim maxpageRange As Integer = 10
        Do
            LogText("请求产品的url：" & url)
            listProducts = GetAsynProductList(url, siteId) '获取指定url的product对象列表
            If (listProducts.Count > 0) Then
                Exit Do
            End If
            '翻页
            Dim pageNostr As String
            Dim constPageNo As String = "pageNo="
            If (Regex.Match(url, constPageNo & "(\d*)").Success) Then
                pageNostr = Regex.Match(url, constPageNo & "(\d*)").Groups(1).Value
                Dim pageNo As Integer
                If (Integer.TryParse(pageNostr, pageNo)) Then
                    pageNo = pageNo + 1
                Else
                    pageNo = 2
                End If
                url = url.Replace(constPageNo & pageNostr, constPageNo & pageNo)
            Else
                Dim conjStr As String
                If (url.Contains("?")) Then
                    conjStr = "&"
                Else
                    conjStr = "?"
                End If
                url = url & conjStr & constPageNo & "2"
            End If
            maxpageRange = maxpageRange - 1
        Loop While (maxpageRange > 0)



        Dim IsDuplicated As Boolean = False
        Dim DedupeList As New List(Of Product)
        Dim NoneDedupeList As List(Of Product) = listProducts

        For i As Integer = 0 To listProducts.Count - 1

            For Each p As Product In DedupeList
                If p.Prodouct = listProducts(i).Prodouct Then
                    IsDuplicated = True
                    Exit For
                End If
            Next

            If Not IsDuplicated Then
                DedupeList.Add(listProducts(i))
            End If
        Next
        listProducts = DedupeList
        Dim addedProduct As List(Of Product) = New List(Of Product)


        Dim existProducts As New List(Of Product)
        existProducts = GetProductList(siteId) '获取数据库（Products表）中该siteID下的所有product(返回列表)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

        For Each li In listProducts
            If Not (IsProductSent(siteId, li.Url, startTime, endTime, planType)) AndAlso Not addedProduct.Contains(li) Then
                Dim productId As Integer = InsertProduct(li, Now, categoryId, existProducts)
                If productId > 0 Then
                    listProductId.Add(productId)
                    addedProduct.Add(li)
                End If
                If (listProductId.Count = productCount) Then
                    Exit For
                End If
            End If
        Next

        If listProductId.Count < productCount Then '说明产品不够了，则把条件放宽到允许15天前发送的
            LogText("prodouct insufficient set startTime = " & Now.AddDays(-15).ToString)
            For Each li In listProducts
                startTime = Now.AddDays(-15)
                If Not (IsProductSent(siteId, li.Url, startTime, endTime, planType)) AndAlso Not addedProduct.Contains(li) Then
                    Dim productId As Integer = InsertProduct(li, Now, categoryId, existProducts)
                    If productId > 0 And (Not listProductId.Contains(productId)) Then
                        listProductId.Add(productId)
                    End If
                    If (listProductId.Count = productCount) Then
                        Exit For
                    End If
                End If
            Next
        End If

        If listProductId.Count < productCount Then '说明产品不够了，则把条件放宽到允许8天前发送的
            LogText("prodouct insufficient set startTime = " & Now.AddDays(-8).ToString)
            For Each li In listProducts
                startTime = Now.AddDays(-8)
                If Not (IsProductSent(siteId, li.Url, startTime, endTime, planType)) AndAlso Not addedProduct.Contains(li) Then
                    Dim productId As Integer = InsertProduct(li, Now, categoryId, existProducts)
                    If productId > 0 And (Not listProductId.Contains(productId)) Then
                        listProductId.Add(productId)
                    End If
                    If (listProductId.Count = productCount) Then
                        Exit For
                    End If
                End If
            Next
        End If

        Dim executeResult As Integer = ProductIssueDAL.InsertProductsIssue(siteId, issueId, section, listProductId, productCount, categoryId)
        ' InsertProductsIssue(siteId, issueId, section, listProductId, productCount)
    End Function


    ''' <summary>
    ''' 获取tmall、taobao产品页的产品,taobao和tmall店铺链接均适用
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="siteid"></param>
    ''' <returns>装product类型的列表listProducts</returns>
    ''' <remarks>指定URL上获得prodoct的Url和productName,SiteID，Discount等？hoyho2016.2.25</remarks>
    Public Function GetAsynProductList(ByVal url As String, ByVal siteid As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim index As Integer
        ' Dim categoryDoc As HtmlDocument = GetAsynHtmlDocByUrlTmall(url)
        Dim cookie As String = ConfigurationManager.AppSettings("aliTmallCookie").ToString.Trim
        Dim categoryDoc As HtmlDocument = GetHtmlDocument(url, cookie, "refer", "gb2312")
        Dim productLines As HtmlNodeCollection
        Dim loopcount As Integer = 2
        Do
            If (url.ToLower.Contains("taobao.com")) Then
                productLines = categoryDoc.DocumentNode.SelectNodes("//div[@class='shop-hesper-bd grid']/div")
            Else
                productLines = categoryDoc.DocumentNode.SelectNodes("//div[@class='J_TItems']/div")
            End If
            If (Not productLines Is Nothing OrElse loopcount < 0) Then
                Exit Do
            End If
            'Threading.Thread.Sleep(5 * 60 * 1000)
            categoryDoc = GetAsynHtmlDocByUrlTmall(url)
            loopcount = loopcount - 1
        Loop
        If Not (productLines Is Nothing) Then
            For Each productLine In productLines
                If (productLine.GetAttributeValue("class", "").ToString.ToLower() = "pagination") Then
                    LogText("遇到分隔符了，退出")
                    Exit For
                End If
                Dim productDls As HtmlNodeCollection = productLine.SelectNodes("dl")
                If Not (productDls Is Nothing) Then
                    For Each dl In productDls
                        Dim myProduct As New Product
                        Dim productName As String = dl.SelectSingleNode("dd[@class='detail']/a").InnerText
                        If productName.Contains("运费") Or productName.Contains("差额") Or productName.Contains("专拍") Or productName.Contains("补拍") Or productName.Contains("差价") Or productName.Contains("補拍") Or productName.Contains("鏈接") Or productName.Contains("專拍") Or productName.Contains("運費") Or productName.Contains("链接") Then
                            Continue For
                        End If
                        myProduct.Prodouct = productName
                        myProduct.Url = dl.SelectSingleNode("dd[@class='detail']/a").GetAttributeValue("href", "").Trim()
                        If (myProduct.Url.Contains("&rn")) Then
                            index = myProduct.Url.IndexOf("&rn")
                            myProduct.Url = myProduct.Url.Remove(index)
                        End If
                        myProduct.Url = AddHttpForAli(myProduct.Url)

                        myProduct.Discount = dl.SelectSingleNode("dd[@class='detail']/div[1]/div[1]/span[2]").InnerText.Trim()
                        myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("data-ks-lazyload", "")
                        If (String.IsNullOrEmpty(myProduct.PictureUrl)) Then
                            myProduct.PictureUrl = dl.SelectSingleNode("dt/a/img").GetAttributeValue("src", "")
                        End If
                        myProduct.PictureUrl = AddHttpForAli(myProduct.PictureUrl)
                        myProduct.Description = productName
                        myProduct.LastUpdate = DateTime.Now
                        myProduct.SiteID = siteid

                        Dim IsDuplicated As Boolean = False
                        For Each catched As Product In listProducts
                            If catched.Prodouct = myProduct.Prodouct AndAlso catched.Url = myProduct.Url Then
                                IsDuplicated = True
                                Exit For
                            End If
                        Next

                        If Not IsDuplicated Then
                            listProducts.Add(myProduct)
                        End If

                    Next
                End If
            Next
        End If
        Return listProducts
    End Function
#End Region


#Region "独立站"
    ''' <summary>
    ''' 读取本站点的本类型计划的BannerRegex表中的第一条规则，根据其中的数据获取本期邮件的banner，并填充至ads及ads_issue表
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="issueId"></param>
    ''' <param name="cateName">banner所属分类名</param>
    ''' <param name="planType"></param>
    ''' <param name="domain">站点域名（当站点的src及url使用的是相对路径时，使用此参数补充成绝对路径）</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBanner(ByVal siteid As Integer, ByVal issueId As Integer, ByVal cateName As String, ByVal planType As String, ByVal domain As String)
        Dim bannerRegex As BannerRegex = (From br In efContext.BannerRegexes
                                          Where br.siteId = siteid And br.planType = planType
                                          Select br).FirstOrDefault()
        If Not (bannerRegex Is Nothing) Then
            '这个cookie是fashion71站点专用且必须的cookie
            Dim cookie As String = "visit_id=zvwm9furfyxagxx52ung; visit_id_user=0; ECS[history]=87403%2C88138; __cfduid=d18948503b84261b12e857250993d77f41434333139; ECS_ID=86cd5acda129a1d940c3e592d9f5bd1a22b6985e; a4316_pages=1; a4316_times=8; CNZZDATA4486438=cnzz_eid%3D560226944-1432116528-%26ntime%3D1434327995; _cnzz_CV4486438=toJSONString%7C%7C; comm100standby_session144462=-43345; comm100_guid_144462=13c7aa2b260a47c9974ba5049cb88032"
            Try
                Dim htmlStr As String = GetHtmlStringByUrl(bannerRegex.bannerFormUrl, bannerRegex.cookie, "")
                htmlStr = System.Net.WebUtility.HtmlDecode(htmlStr)
                Dim mCollection As MatchCollection = Regex.Matches(htmlStr, bannerRegex.bannerRegText)
                Dim index As Integer = bannerRegex.bannerStartIndex
                Dim bMatch As Match
                For i As Integer = index To bannerRegex.bannerEndIndex
                    bMatch = mCollection(i)
                    Dim myAd As New Ad()

                    myAd.PictureUrl = bMatch.Groups(1).Value.Trim ' src及href的顺序不确定，用是否含有图片后缀判断是src还是href
                    If (myAd.PictureUrl.ToLower.Contains(".bmp") OrElse myAd.PictureUrl.ToLower.Contains(".jpg") OrElse
                        myAd.PictureUrl.ToLower.Contains(".jpeg") OrElse myAd.PictureUrl.ToLower.Contains(".png") OrElse
                        myAd.PictureUrl.ToLower.Contains(".gif")) Then
                        myAd.PictureUrl = addDominForUrl(domain, bMatch.Groups(1).Value.Trim)
                        myAd.Url = bMatch.Groups(2).Value.Trim
                        myAd.Url = addDominForUrl(domain, IIf(String.IsNullOrEmpty(myAd.Url), bannerRegex.bannerFormUrl, myAd.Url))
                    Else
                        myAd.Url = addDominForUrl(domain, bMatch.Groups(1).Value.Trim)
                        myAd.PictureUrl = bMatch.Groups(2).Value.Trim
                        If (String.IsNullOrEmpty(myAd.PictureUrl)) Then
                            Continue For
                        End If
                        myAd.PictureUrl = addDominForUrl(domain, bMatch.Groups(2).Value.Trim)
                    End If

                    If Not (isAdSent(siteid, myAd, DateTime.Now.AddDays(0 - bannerRegex.noRepeatSentDays), DateTime.Now)) Then
                        myAd.SiteID = siteid
                        myAd.Lastupdate = Now

                        Dim listAd As List(Of Ad) = GetListAd(siteid) '获取本站的的Ads表中的所有ad
                        Dim categoryId As Integer = GetCategoryId(siteid, cateName)
                        Dim adid As Integer = InsertAd(myAd, listAd, categoryId)
                        InsertSingleAdsIssue(adid, siteid, issueId)
                        '可以配多个连续的banner index，一旦找到合适则exit
                        Exit For
                    End If
                Next
            Catch ex As Exception
                LogText("获取banner失败" & "siteId=" & siteid.ToString() & "bannerRegex=" & bannerRegex.ToString())
            End Try

        End If
    End Function


    ''' <summary>
    ''' 遍历本站点的本类型计划的productPaths表中的规则，根据其中的设定获取产品，并填充至products及products_issue表
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="planType"></param>
    ''' <param name="domain"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProducts(ByVal siteid As Integer, ByVal planType As String, ByVal domain As String, ByVal issueId As Integer)
        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath).ToList()
        For Each item As ProductPath In listProPath
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()

            If cate.Category1.ToLower = "hotclick" Then
                '往期热门，后面从数据库中提取，不需要爬站点
                Continue For
            End If


            If (item.fileType.Trim.ToLower() = "tmall") Then
                '天猫走独立站流程，获取产品时
                Dim requestUrl As String = addParamForUrl(cate.Url, IIf(item.cateParam Is Nothing, "", item.cateParam))
                FetchProduct(requestUrl, cate.Category1, "CA", planType, item.prodDisplayCount, DateTime.Now.AddDays(0 - item.noRepeatSentDays), DateTime.Now, siteid, issueId)
            Else
                Dim mylistProduct As List(Of Product) = GetListProducts(cate, item, domain)
                Dim listProductId As List(Of Integer) = insertProducts(mylistProduct, cate.Category1, "CA", planType, item.prodDisplayCount, siteid, issueId)
                'InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
                ProductIssueDAL.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount, cate.CategoryID)
            End If
        Next
    End Function







    ''' <summary>
    ''' 将listproduct中前productCount个产品插入到产品表，并返回list of those productids
    ''' </summary>
    ''' <param name="listProducts"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="section">CA，DA，NE等</param>
    ''' <param name="planType"></param>
    ''' <param name="productCount"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <returns>返回list of those productids</returns>
    ''' <remarks></remarks>
    Public Function insertProducts(ByVal listProducts As List(Of Product), ByVal categoryName As String, ByVal section As String, ByVal planType As String,
                         ByVal productCount As Integer, ByVal siteId As Integer,
                         ByVal issueId As Integer) As List(Of Integer)
        Dim listProduct As List(Of Product) = GetProductList(siteId)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = GetCategoryId(siteId, categoryName)

        For Each li In listProducts
            Dim returnId As Integer = InsertProduct(li, Now, categoryId, listProduct)
            If returnId > 0 Then
                listProductId.Add(returnId)
            End If
            If (listProductId.Count = productCount) Then
                Exit For
            End If
        Next
        Return listProductId
    End Function


    ''' <summary>
    ''' 将listproduct中前productCount个产品插入到产品表，并返回list of those productids
    ''' 此方法针对K11的站点使用，因为product比较特殊
    ''' </summary>
    ''' <param name="listProducts"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="section"></param>
    ''' <param name="planType"></param>
    ''' <param name="productCount"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <returns></returns>
    Public Function insertK11Products(ByVal listProducts As List(Of Product), ByVal categoryName As String, ByVal section As String, ByVal planType As String,
                     ByVal productCount As Integer, ByVal siteId As Integer,
                     ByVal issueId As Integer) As List(Of Integer)
        Dim listProduct As List(Of Product) = GetProductList(siteId)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
        For Each li In listProducts

            Dim returnId As Integer = InsertK11Product(li, Now, categoryId, listProduct)
            Common.LogText("返回插入的id" + returnId.ToString())
            If returnId > 0 Then
                listProductId.Add(returnId)
            End If
            If (listProductId.Count = productCount) Then
                Exit For
            End If
        Next
        Return listProductId
    End Function
    Public Function GetListProducts(ByVal siteid As Integer, ByVal plantype As String, ByVal cate As Category, ByVal item As ProductPath, Optional ByVal domain As String = "") As List(Of Product)
        Dim site As AutomationSite = GetAutoSiteBySid(siteid)
        If site.DllType = "rss" Then
            Dim plan = efContext.AutomationPlans.First(Function(p) p.SiteID = siteid AndAlso p.PlanType = plantype)
            Dim doc = LoadXmlDoc(plan.URL)
            Return GetListProducts(doc, cate, item)
        ElseIf site.DllType = "fb" Then 'fb
            'Dim myHtmlDom As HtmlDocument = New HtmlDocument
            'Dim myProdDom As HtmlNodeCollection = Nothing
            'myHtmlDom = GetHtmlDocument(cate.Url, item.cookie, "refer", item.pageEncoding)
            'Dim jsonNode As HtmlNode = myHtmlDom.DocumentNode.SelectSingleNode(item.cateParam)
            'Dim jobj = JObject.Parse(jsonNode.InnerHtml)
            'Return GetJsonProducts(jobj, cate, item)
        Else
            If String.IsNullOrEmpty(domain) Then
                Try
                    Dim uri = New Uri(cate.Url)
                    domain = uri.Scheme & "://" & uri.Host
                Catch ex As Exception

                End Try
            End If
            Return GetListProducts(cate, item, domain)
        End If
    End Function

    ''' <summary>
    ''' 根据一条productpath记录获取产品
    ''' </summary>
    ''' <param name="cate"></param>
    ''' <param name="item"></param>
    ''' <param name="domain"></param>
    ''' <returns>返回Product类型的list</returns>
    ''' <remarks></remarks>
    Public Function GetListProducts(ByVal cate As Category, ByVal item As ProductPath, ByVal domain As String) As List(Of Product)

        Dim requestUrl As String = addParamForUrl(cate.Url, IIf(item.cateParam Is Nothing, "", item.cateParam))

        If Not String.IsNullOrEmpty(domain) Then
            Try
                Dim uri = New Uri(domain)
                domain = uri.Scheme & "://" & uri.Host
            Catch ex As Exception

            End Try
        End If
        Dim listProds As New List(Of Product)

        Dim issueDate As Date
        If (item.fileType.ToLower.Trim = "html") Then
            'debug
            Dim myHtmlDom As HtmlDocument = New HtmlDocument
            Dim myProdDom As HtmlNodeCollection = Nothing
            myHtmlDom = GetHtmlDocument(requestUrl, item.cookie, "refer", item.pageEncoding)
            myProdDom = myHtmlDom.DocumentNode.SelectNodes(item.prodPath)

            Dim userBrowser As Boolean = False

            'use webbrowser to load again
            If myProdDom Is Nothing Then
fetchagain:
                userBrowser = True
                LogText("webbrowser load:" & requestUrl & "  Category:" & cate.CategoryID)
                Dim path = GetBrowserConsolePath()
                Dim value = ApplicationHelper.Run(path, requestUrl, item.cateParam, item.pageEncoding, 10)
                Dim message = JsonHelper.ParseFromJson(Of ApplicationMessage)(value)
                If message.Code = 1 Then
                    myHtmlDom.LoadHtml(message.Value)
                    myProdDom = myHtmlDom.DocumentNode.SelectNodes(item.prodPath)
                End If
            End If

            If myProdDom Is Nothing Then
                LogText("myHtmlDom.DocumentNode.SelectNodes(item.prodPath) is null where  domain:" & domain)
                NotificationEmail.SentNotificationEmail("autoedm@reasonable.cn", "Please update crawler rule", String.Format("The website structure may have changed, you have to update thr rule,Url={0},siteid={1}", cate.Url, cate.SiteID))
                Return listProds
            End If

            For Each inode As HtmlNode In myProdDom
                Try
                    Dim newProd As New Product()
                    Dim decodeStr As String
                    '因为某些网站含有奇怪的转义字符比如<a href="https&#58;&#47;&#47;www.vipme.com&#47;clothing_c900027"
                    '因此需要htmldecode()   (-hoyho, 2016 - 3 - 18)
                    If (String.IsNullOrEmpty(item.urlAttri)) Then
                        decodeStr = inode.SelectSingleNode(item.urlPath).InnerText.Trim
                        newProd.Url = Web.HttpUtility.HtmlDecode(decodeStr).Trim
                    Else
                        If (String.IsNullOrEmpty(item.urlPath)) Then
                            decodeStr = inode.GetAttributeValue(item.urlAttri, "").Trim()
                            newProd.Url = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                        Else
                            Dim hnode As HtmlNode = inode.SelectSingleNode(item.urlPath)
                            decodeStr = inode.SelectSingleNode(item.urlPath).GetAttributeValue(item.urlAttri, "").Trim
                            newProd.Url = Web.HttpUtility.HtmlDecode(decodeStr).Trim()

                        End If
                    End If
                    newProd.Url = addDominForUrl(domain, newProd.Url)
                    If Not (IsProductSent(item.siteId, newProd.Url, DateTime.Now.AddDays(0 - item.noRepeatSentDays), DateTime.Now)) Then
                        If Not (String.IsNullOrEmpty(item.productPath1)) Then
                            If (String.IsNullOrEmpty(item.productAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.productPath1).InnerText.Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            Else
                                decodeStr = inode.SelectSingleNode(item.productPath1).GetAttributeValue(item.productAttri, "").Trim
                                newProd.Prodouct = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(item.pricePath)) Then
                            If (String.IsNullOrEmpty(item.priceAttri)) Then
                                'priceStr = "US&#36; 51.99" getPriceNum(priceStr)于是price和discount全都变成只提取到36，因此解码
                                Dim priceStr As String = inode.SelectSingleNode(item.pricePath).InnerText.Trim
                                priceStr = Web.HttpUtility.HtmlDecode(priceStr).Trim()
                                newProd.Price = getPriceNum(priceStr)
                            Else
                                decodeStr = getPriceNum(inode.SelectSingleNode(item.pricePath).GetAttributeValue(item.priceAttri, "").Trim)
                                newProd.Price = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(item.discountPath)) Then
                            If (String.IsNullOrEmpty(item.discountAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.discountPath).InnerText.Trim
                                decodeStr = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                                newProd.Discount = getPriceNum(decodeStr)
                            Else
                                decodeStr = getPriceNum(inode.SelectSingleNode(item.discountPath).GetAttributeValue(item.discountAttri, "").Trim)
                                newProd.Discount = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(item.salesPath)) Then
                            If (String.IsNullOrEmpty(item.salesAttri)) Then
                                newProd.Sales = inode.SelectSingleNode(item.salesPath).InnerText.Trim
                            Else
                                decodeStr = inode.SelectSingleNode(item.salesPath).GetAttributeValue(item.salesAttri, "").Trim
                                newProd.Sales = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If
                        If Not (String.IsNullOrEmpty(item.pictureUrlPath)) Then
                            If (String.IsNullOrEmpty(item.pictureUrlAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.pictureUrlPath).InnerText.Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            Else
                                decodeStr = inode.SelectSingleNode(item.pictureUrlPath).GetAttributeValue(item.pictureUrlAttri, "").Trim
                                newProd.PictureUrl = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                            newProd.PictureUrl = addDominForUrl(domain, newProd.PictureUrl)
                        End If

                        If Not (String.IsNullOrEmpty(item.descriptionPath)) Then
                            If (String.IsNullOrEmpty(item.descriptionAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.descriptionPath).InnerText.Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            Else
                                decodeStr = inode.SelectSingleNode(item.descriptionPath).GetAttributeValue(item.descriptionAttri, "").Trim
                                newProd.Description = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If
                        newProd.Currency = item.currencyChar
                        If Not (String.IsNullOrEmpty(item.pictureAltPath)) Then
                            If (String.IsNullOrEmpty(item.pictureAltAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.pictureAltPath).InnerText.Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            Else
                                decodeStr = inode.SelectSingleNode(item.pictureAltPath).GetAttributeValue(item.pictureAltAttri, "").Trim
                                newProd.PictureAlt = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                            End If
                        End If

                        newProd.SiteID = item.siteId
                        newProd.LastUpdate = DateTime.Now
                        If Not (String.IsNullOrEmpty(item.issueDate.Trim)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                            If (String.IsNullOrEmpty(item.issueDateAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.issueDate).InnerText.Trim
                                issueDate = Date.Parse(decodeStr)
                                newProd.PublishDate = issueDate
                            Else
                                decodeStr = inode.SelectSingleNode(item.issueDate).GetAttributeValue(item.issueDateAttri, "").Trim
                                issueDate = Date.Parse(decodeStr)
                                newProd.PublishDate = issueDate
                            End If
                            If (Now.Date - issueDate > TimeSpan.FromDays(item.validityPeriod)) Then  '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        Else '抓取数固定
                            If (listProds.Count >= (item.prodDisplayCount * 2)) Then
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    LogText("Error occurs, skip this product\n")
                    LogText(ex.Message) 'skip this product
                End Try

            Next
            'If Not userBrowser AndAlso listProds.Count = 0 Then
            '    GoTo fetchagain
            'End If

        ElseIf (item.fileType.ToLower.Trim = "xml") Then
            Dim xmlDoc As XmlDocument = LoadXmlDoc(requestUrl, item.pageEncoding)
            Dim prodXmlNodeList As XmlNodeList = xmlDoc.SelectNodes(item.prodPath)
            Dim decodeStr As String = ""
            For Each inode As XmlNode In prodXmlNodeList
                Try
                    Dim newProd As New Product()
                    newProd.Url = inode.SelectSingleNode(item.urlPath).InnerText.Trim
                    newProd.Url = addDominForUrl(domain, newProd.Url)
                    If Not (IsProductSent(item.siteId, newProd.Url, DateTime.Now.AddDays(0 - item.noRepeatSentDays), DateTime.Now)) Then
                        If Not (String.IsNullOrEmpty(item.productPath1)) Then
                            newProd.Prodouct = inode.SelectSingleNode(item.productPath1).InnerText.Trim
                        End If
                        If Not (String.IsNullOrEmpty(item.pricePath)) Then
                            Try
                                newProd.Price = getPriceNum(inode.SelectSingleNode(item.pricePath).InnerText.Trim)
                            Catch ex As Exception

                            End Try
                        End If
                        If Not (String.IsNullOrEmpty(item.discountPath)) Then
                            Try
                                newProd.Discount = getPriceNum(inode.SelectSingleNode(item.discountPath).InnerText.Trim)
                            Catch ex As Exception

                            End Try
                        End If
                        If Not (String.IsNullOrEmpty(item.salesPath)) Then
                            Try
                                newProd.Sales = inode.SelectSingleNode(item.salesPath).InnerText.Trim
                            Catch ex As Exception

                            End Try
                        End If
                        If Not (String.IsNullOrEmpty(item.pictureUrlPath)) Then
                            Try
                                If (String.IsNullOrEmpty(item.pictureUrlAttri)) Then
                                    newProd.PictureUrl = inode.SelectSingleNode(item.pictureUrlPath).InnerText.Trim
                                Else
                                    newProd.PictureUrl = inode.SelectSingleNode(item.pictureUrlPath).Attributes(item.pictureUrlAttri, "").InnerXml.ToString()
                                End If
                                newProd.PictureUrl = addDominForUrl(domain, newProd.PictureUrl)
                            Catch ex As Exception

                            End Try
                        End If
                        If Not (String.IsNullOrEmpty(item.descriptionPath)) Then
                            Try
                                newProd.Description = inode.SelectSingleNode(item.descriptionPath).InnerText.Trim
                            Catch ex As Exception

                            End Try
                        End If

                        newProd.Currency = item.currencyChar
                        If Not (String.IsNullOrEmpty(item.pictureAltPath)) Then
                            Try
                                newProd.PictureAlt = inode.SelectSingleNode(item.pictureAltPath).InnerText.Trim
                            Catch ex As Exception

                            End Try
                        End If
                        If Not (String.IsNullOrEmpty(item.issueDate)) Then

                        End If

                        newProd.SiteID = item.siteId
                        newProd.LastUpdate = DateTime.Now

                        If Not (String.IsNullOrEmpty(item.issueDate)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                            If (String.IsNullOrEmpty(item.issueDateAttri)) Then
                                decodeStr = inode.SelectSingleNode(item.issueDate).InnerText.Trim
                                issueDate = Date.Parse(decodeStr)
                                newProd.PublishDate = issueDate
                            Else
                                decodeStr = inode.SelectSingleNode(item.issueDate).Attributes(item.issueDateAttri, "").InnerXml.ToString()
                                issueDate = Date.Parse(decodeStr)
                                newProd.PublishDate = issueDate
                            End If
                            If (Now.Date - issueDate > TimeSpan.FromDays(item.validityPeriod)) Then '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        Else '抓取数固定
                            If (listProds.Count >= (item.prodDisplayCount * 2)) Then
                                Exit For
                            ElseIf Not isProductDuplicate(listProds, newProd) Then
                                listProds.Add(newProd)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    LogText("Error occurs,skip this product\n")
                    LogText(ex.Message) 'skip this product
                End Try
            Next
        End If
        Return listProds
    End Function
    Public Function GetListProducts(xmlDoc As XmlDocument, cate As Category, item As ProductPath) As List(Of Product)
        Dim prodXmlNodeList As XmlNodeList = xmlDoc.SelectNodes(item.prodPath)
        Dim decodeStr As String = ""
        Dim uri = New Uri(cate.Url)
        Dim domain As String = uri.Scheme & "://" & uri.Host
        Dim listProds As New List(Of Product)
        Dim issueDate As Date

        For Each inode As XmlNode In prodXmlNodeList
            Try
                Dim cateName As String = inode.SelectSingleNode("category").InnerText.Trim
                If cate.Category1 <> cateName Then
                    Continue For
                End If
                Dim newProd As New Product()

                If Not (String.IsNullOrEmpty(item.urlPath)) Then
                    If String.IsNullOrEmpty(item.urlAttri) Then
                        newProd.Url = inode.SelectSingleNode(item.urlPath).InnerText.Trim
                    Else
                        newProd.Url = inode.SelectSingleNode(item.urlPath).Attributes(item.urlAttri).InnerText.Trim
                    End If
                    newProd.Url = addDominForUrl(domain, newProd.Url)
                End If
                If Not (IsProductSent(item.siteId, newProd.Url, DateTime.Now.AddDays(0 - item.noRepeatSentDays), DateTime.Now)) Then

                    If Not (String.IsNullOrEmpty(item.productPath1)) Then
                        If String.IsNullOrEmpty(item.productAttri) Then
                            newProd.Prodouct = inode.SelectSingleNode(item.productPath1).InnerText.Trim
                        Else
                            newProd.Prodouct = inode.SelectSingleNode(item.productPath1).Attributes(item.productAttri).InnerText.Trim
                        End If
                    End If

                    If Not (String.IsNullOrEmpty(item.pricePath)) Then
                        Try
                            If String.IsNullOrEmpty(item.priceAttri) Then
                                newProd.Price = getPriceNum(inode.SelectSingleNode(item.pricePath).InnerText.Trim)
                            Else
                                newProd.Price = getPriceNum(inode.SelectSingleNode(item.pricePath).Attributes(item.priceAttri).InnerText.Trim)
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(item.discountPath)) Then
                        Try
                            If String.IsNullOrEmpty(item.discountAttri) Then
                                newProd.Discount = getPriceNum(inode.SelectSingleNode(item.discountPath).InnerText.Trim)
                            Else
                                newProd.Discount = getPriceNum(inode.SelectSingleNode(item.discountPath).Attributes(item.discountAttri).InnerText.Trim)
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(item.salesPath)) Then
                        Try
                            If String.IsNullOrEmpty(item.salesAttri) Then
                                newProd.Sales = inode.SelectSingleNode(item.salesPath).InnerText.Trim
                            Else
                                newProd.Sales = inode.SelectSingleNode(item.salesPath).Attributes(item.salesAttri).InnerText.Trim
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(item.pictureUrlPath)) Then
                        Try
                            If (String.IsNullOrEmpty(item.pictureUrlAttri)) Then
                                newProd.PictureUrl = inode.SelectSingleNode(item.pictureUrlPath).InnerText.Trim
                            Else
                                newProd.PictureUrl = inode.SelectSingleNode(item.pictureUrlPath).Attributes(item.pictureUrlAttri, "").InnerXml.ToString()
                            End If
                            newProd.PictureUrl = addDominForUrl(domain, newProd.PictureUrl)
                        Catch ex As Exception

                        End Try
                    End If
                    If Not (String.IsNullOrEmpty(item.descriptionPath)) Then
                        Try
                            If String.IsNullOrEmpty(item.descriptionAttri) Then
                                newProd.Description = inode.SelectSingleNode(item.descriptionPath).InnerText.Trim
                            Else
                                newProd.Description = inode.SelectSingleNode(item.descriptionPath).Attributes(item.descriptionAttri).InnerText.Trim
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    newProd.Currency = item.currencyChar
                    If Not (String.IsNullOrEmpty(item.pictureAltPath)) Then
                        Try
                            If String.IsNullOrEmpty(item.pictureAltAttri) Then
                                newProd.PictureAlt = inode.SelectSingleNode(item.pictureAltPath).InnerText.Trim
                            Else
                                newProd.PictureAlt = inode.SelectSingleNode(item.pictureAltPath).Attributes(item.pictureAltAttri).InnerText.Trim
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    newProd.SiteID = item.siteId
                    newProd.LastUpdate = DateTime.Now

                    If Not (String.IsNullOrEmpty(item.issueDate)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数
                        If (String.IsNullOrEmpty(item.issueDateAttri)) Then
                            decodeStr = inode.SelectSingleNode(item.issueDate).InnerText.Trim
                            issueDate = Date.Parse(decodeStr)
                            newProd.PublishDate = issueDate
                        Else
                            decodeStr = inode.SelectSingleNode(item.issueDate).Attributes(item.issueDateAttri, "").InnerXml.ToString()
                            issueDate = Date.Parse(decodeStr)
                            newProd.PublishDate = issueDate
                        End If
                        If (Now.Date - issueDate > TimeSpan.FromDays(item.validityPeriod)) Then '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                            Exit For
                        ElseIf Not isProductDuplicate(listProds, newProd) Then
                            listProds.Add(newProd)
                        End If
                    Else '抓取数固定
                        If (listProds.Count >= (item.prodDisplayCount * 2)) Then
                            Exit For
                        ElseIf Not isProductDuplicate(listProds, newProd) Then
                            listProds.Add(newProd)
                        End If
                    End If
                End If
            Catch ex As Exception
                EFHelper.LogText("Error occurs,skip this product\n")
                EFHelper.LogText(ex.Message) 'skip this product
            End Try
        Next
        Return listProds
    End Function
    Public Function GetJsonProducts(jsonObj As JObject, cate As Category, item As ProductPath) As List(Of Product)


        Dim uri = New Uri(cate.Url)
        Dim domain As String = uri.Scheme & "://" & uri.Host
        Dim listProduct As New List(Of Product)

        Dim postJarr As JArray = jsonObj.SelectToken(item.prodPath)
        Dim issueDate As Date
        For i As Integer = 0 To postJarr.Count - 1
            Dim inode As JObject = postJarr(i)

            Try
                Dim newProd As New Product()
                Dim decodeStr As String
                '因为某些网站含有奇怪的转义字符比如<a href="https&#58;&#47;&#47;www.vipme.com&#47;clothing_c900027"
                '因此需要htmldecode()   (-hoyho, 2016 - 3 - 18)
                If Not (String.IsNullOrEmpty(item.urlPath)) Then
                    decodeStr = inode.SelectToken(item.urlPath).ToString.Trim
                    newProd.Url = Web.HttpUtility.HtmlDecode(decodeStr).Trim
                End If

                newProd.Url = addDominForUrl(domain, newProd.Url)
                If Not (IsProductSent(item.siteId, newProd.Url, DateTime.Now.AddDays(0 - item.noRepeatSentDays), DateTime.Now)) Then

                    If Not (String.IsNullOrEmpty(item.productPath1)) Then
                        decodeStr = inode.SelectToken(item.productPath1).ToString.Trim
                        newProd.Prodouct = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                    End If

                    If Not (String.IsNullOrEmpty(item.pricePath)) Then
                        'priceStr = "US&#36; 51.99" getPriceNum(priceStr)于是price和discount全都变成只提取到36，因此解码
                        Dim priceStr As String = inode.SelectToken(item.pricePath).ToString.Trim
                        priceStr = Web.HttpUtility.HtmlDecode(priceStr).Trim()
                        newProd.Price = getPriceNum(priceStr)

                    End If
                    If Not (String.IsNullOrEmpty(item.discountPath)) Then
                        decodeStr = inode.SelectToken(item.discountPath).ToString.Trim
                        decodeStr = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                        newProd.Discount = getPriceNum(decodeStr)
                    End If

                    If Not (String.IsNullOrEmpty(item.salesPath)) Then
                        newProd.Sales = inode.SelectToken(item.salesPath).ToString.Trim
                    End If
                    If Not (String.IsNullOrEmpty(item.pictureUrlPath)) Then
                        decodeStr = inode.SelectToken(item.pictureUrlPath).ToString.Trim
                        newProd.PictureUrl = Web.HttpUtility.HtmlDecode(decodeStr).Trim()

                        newProd.PictureUrl = addDominForUrl(domain, newProd.PictureUrl)
                    End If

                    If Not (String.IsNullOrEmpty(item.descriptionPath)) Then
                        decodeStr = inode.SelectToken(item.descriptionPath).ToString.Trim
                        newProd.Description = Web.HttpUtility.HtmlDecode(decodeStr).Trim()
                    End If
                    newProd.Currency = item.currencyChar
                    If Not (String.IsNullOrEmpty(item.pictureAltPath)) Then
                        decodeStr = inode.SelectToken(item.pictureAltPath).ToString.Trim
                        newProd.PictureAlt = Web.HttpUtility.HtmlDecode(decodeStr).Trim()

                    End If

                    newProd.SiteID = item.siteId
                    newProd.LastUpdate = DateTime.Now
                    If Not (String.IsNullOrEmpty(item.issueDate.Trim)) Then '抓取项目数不定，由抓取日期决定，prodDisplayCount字段无法在此更新，但至少应该设置成大于当前抓到的产品数

                        decodeStr = inode.SelectToken(item.issueDate).ToString.Trim
                        issueDate = Date.Parse(decodeStr)
                        newProd.PublishDate = issueDate
                        If (Now.Date - issueDate > TimeSpan.FromDays(item.validityPeriod)) Then  '仅抓取validityPeriod内的产品,超出时间范围的则退出（与当前时间比较）
                            Exit For
                        ElseIf Not isProductDuplicate(listProduct, newProd) Then
                            listProduct.Add(newProd)
                        End If
                    Else '抓取数固定
                        If (listProduct.Count >= (item.prodDisplayCount * 2)) Then
                            Exit For
                        ElseIf Not isProductDuplicate(listProduct, newProd) Then
                            listProduct.Add(newProd)
                        End If
                    End If
                End If
            Catch ex As Exception
                LogText("Error occurs, skip this product\n")
                LogText(ex.Message) 'skip this product
            End Try
        Next
        Return listProduct
    End Function
    Public Function AddSubject(ByVal siteID As Integer, ByVal issueID As Integer, ByVal preSubject As String)
        Dim querySubject = (From p In efContext.Products
                            Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                            Where pi.IssueID = issueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteID
                            Select p.Prodouct).FirstOrDefault()
        Dim subject As String

        If Not (querySubject Is Nothing) Then
            subject = preSubject
        End If

        InsertIssueSubject(issueID, preSubject & querySubject.ToString)
    End Function

    Private Function addParamForUrl(ByVal url As String, ByVal param As String) As String
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

    Private Function getPriceNum(ByVal priceStr As String) As Double
        Dim reg As String = "([0-9.]+)"
        Dim price As Double = Double.Parse(Regex.Matches(priceStr, reg)(0).Value())
        Return price
    End Function

    Private Function isProductDuplicate(ByVal productList As List(Of Product), ByVal prodcut As Product) As Boolean
        Dim isDuplicate As Boolean = False
        For Each p As Product In productList
            If p.Prodouct = prodcut.Prodouct AndAlso p.Price = prodcut.Price AndAlso
                IIf(p.Discount Is Nothing, -1, p.Discount) = IIf(prodcut.Discount Is Nothing, -1, prodcut.Discount) Then
                isDuplicate = True
                Exit For
            End If
        Next
        Return isDuplicate
    End Function

    Private Function GetBrowserConsolePath() As String
        If System.Web.HttpContext.Current IsNot Nothing Then
            Return System.Web.HttpContext.Current.Server.MapPath("/") & "/bin/EmailFetchConsole.exe"
        End If
        Return System.IO.Path.GetDirectoryName(GetType(EFHelper).Assembly.Location) + "/EmailFetchConsole.exe"
    End Function

#End Region

#Region "facebook"
    ''' <summary>
    ''' 获取fb的一个page上发布的post，并返回List(Of Product)
    ''' </summary>
    ''' <param name="fbPageName">fb的page名</param>
    ''' <param name="limit">指定获取前几个post（按从新到旧）</param>
    ''' <param name="accessToken"></param>
    ''' <param name="siteID"></param>
    ''' <param name="planType"></param>
    ''' <param name="issueID"></param>
    ''' <param name="categoryName">产品所属分类名</param>
    ''' <param name="section">"CA","DA","NE"等</param>
    ''' <returns>List(Of Product)</returns>
    ''' <remarks></remarks>
    Public Function FetchfbPosts(ByVal fbPageName As String, ByVal limit As Integer, ByVal accessToken As String, ByVal siteID As Integer) As List(Of Product)
        Dim listProduct As New List(Of Product)
        Dim postJarr As JArray = GetfbPosts(fbPageName, limit, accessToken)
        For i As Integer = 0 To postJarr.Count - 1
            Dim item As JObject = postJarr(i)

            Dim myPro As New Product()
            If Not (item.TryGetValue("message", myPro.Description)) Then
                Continue For
            End If

            Dim photoid As String = ""
            'If Not (item.TryGetValue("object_id", photoid)) Then
            '    Continue For
            'End If

            If Not (item.TryGetValue("id", photoid)) Then
                Continue For
            End If

            Dim photoJson As JObject = GetfbPhotobyid(photoid, accessToken)
            If Not (photoJson.TryGetValue("full_picture", myPro.PictureUrl)) Then
                Continue For
            End If

            item.TryGetValue("created_time", myPro.ExpiredDate)
            'item.TryGetValue("link", myPro.Url)
            photoJson.TryGetValue("permalink_url", myPro.Url)


            Dim rmatch As Match = Regex.Match(myPro.Description, "http(?:|s):\S+\s?")
            If (rmatch.Success) Then
                myPro.Prodouct = rmatch.Value.Trim 'post文中的超链接（“立即购买”按钮）
            Else
                myPro.Prodouct = myPro.Url
            End If

            myPro.Description = AddfbPosttextHyperlink(myPro.Description)
            If (myPro.Description.Contains(vbLf)) Then
                myPro.PictureAlt = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                myPro.Description = myPro.Description.Remove(0, myPro.Description.IndexOf(vbLf))
            End If
            myPro.Description = myPro.Description.Replace(vbLf, "</br>")
            myPro.Description = myPro.Description
            myPro.SiteID = siteID
            myPro.LastUpdate = DateTime.Now
            If Not (IsProductSent(siteID, myPro.Url, Now.AddDays(-100), Now)) Then
                listProduct.Add(myPro)
            End If
        Next
        Return listProduct
    End Function

    ''' <summary>
    ''' 获取一个fb页面的post的Json数据
    ''' </summary>
    ''' <param name="fbPageName">fb页面名字</param>
    ''' <param name="limit">int，获取几个最新的post</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以Jarry返回post数据</returns>
    ''' <remarks></remarks>
    Public Function GetfbPosts(ByVal fbPageName As String, ByVal limit As Integer, ByVal accessToken As String) As JArray

        Dim requestUrl As String = "https://graph.facebook.com/" & fbPageName & "?access_token=" & accessToken & "&fields=posts.limit(" & limit & ")&format=json"
        'Dim requestUrl As String = "https://graph.facebook.com/" & fbPageName & "?access_token=" & accessToken & "&fields=&format=json"
        Common.LogText(requestUrl)
        Dim postsStr As String = GetHtmlStringByUrl(requestUrl, "", "", "", 3)
        Dim postsJson As JObject = JObject.Parse(postsStr)
        Dim postsJArr As JArray = postsJson("posts")("data")
        Return postsJArr
    End Function
    ''' <summary>
    ''' 根据id获取一条帖子的信息，如type，link
    ''' </summary>
    ''' <param name="帖子的id">相片的id</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以Jarry返回post数据</returns>
    ''' <remarks></remarks>
    ''' <summary>
    Public Function GetfbTypeAndLink(ByVal id As String, ByVal accessToken As String) As JObject
        Dim requestUrl As String = "https://graph.facebook.com/v2.10/" & id & "?access_token=" & accessToken & "&fields=type,link&format=json"

        Dim postsStr As String = GetHtmlStringByUrl(requestUrl, "", "", "", 3)
        Dim postsJson As JObject = JObject.Parse(postsStr)
        Return postsJson
    End Function
    ''' 根据id获取一条photo的数据
    ''' </summary>
    ''' <param name="id">帖子的id</param> --by Roy
    ''' <param name="accessToken">token</param>
    ''' <returns>以jObject返回相片数据</returns>
    ''' <remarks></remarks>
    Public Function GetfbPhotobyid(ByVal id As String, ByVal accessToken As String) As JObject
        Dim requestUrl As String = "https://graph.facebook.com/v2.10/" & id & "?access_token=" & accessToken & "&fields=full_picture,permalink_url,picture&format=json"
        Dim photoStr As String = GetHtmlStringByUrl(requestUrl, "", "", "")
        Common.LogText("图片的url" + requestUrl)
        Dim photoJson As JObject = JObject.Parse(photoStr)
        Return photoJson
    End Function
    ''' <summary>
    ''' 给post的文字描述中的http纯文本添加超链接
    ''' </summary>
    ''' <param name="postText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddfbPosttextHyperlink(ByVal postText As String) As String
        Dim hrefRegexCollection As MatchCollection = Regex.Matches(postText, "http:\S+\s ?")
        Dim hrefUrl As String
        For i = 0 To hrefRegexCollection.Count - 1
            hrefUrl = hrefRegexCollection.Item(i).Value
            postText = postText.Replace(hrefUrl, "<a href=""" & hrefUrl.Trim & """ target=""_blank"" style=""text-decoration:none;color: #3b5998"">" & hrefUrl.Trim & "</a></br>")
        Next
        Return postText
    End Function

    ''' <summary>
    ''' 获取指定id的long-time-token,tokenId=1-->facebook的token信息，tokenId=2-->新浪微博的token信息
    ''' </summary>
    ''' <param name="tokenId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLongTimeToken(ByVal tokenId As Integer) As String
        Dim longtimeToken = (From t In efContext.Tokens
                             Where t.id = tokenId
                             Select t).FirstOrDefault()
        Return longtimeToken.longTimeToken.Trim
    End Function


    ''' <summary>
    ''' 到期时间在30天之内则自动更新facebook的token
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function UpdateFbToken() As Boolean
        'Dim oldLongTimeToken As String = GetLongTimeToken(1) 'old
        Dim token = (From t In efContext.Tokens
                     Where t.id = 1
                     Select t).FirstOrDefault()
        'If token Is Nothing Then
        '    LogText(efContext.Connection.ConnectionString)
        'End If
        Dim oldLongTimeToken As String = token.longTimeToken
        Dim expireDate As Date = token.longTokenExpireDate

        If (expireDate - Now.Date < TimeSpan.FromDays(30)) Then
            Dim requestUrl As String = "https://graph.facebook.com/v2.3/oauth/access_token?grant_type=fb_exchange_token&client_id=874014232612712&client_secret=01e0ec489cb4a17546e2f0674242db69&fb_exchange_token=" & oldLongTimeToken.Trim
            Dim resultStr As String = ""
            Try
                'resultStr = efHelper.GetHtmlStringByUrl(requestUrl, "", "", Encoding.UTF8)
                resultStr = EFHelper.GetHtmlStringByUrl(requestUrl, 5)
                Dim tokenJobject As JObject = JObject.Parse(resultStr)
                Dim newlongToken As String = tokenJobject("access_token").ToString.Trim
                Dim myToken As Token = (From t In efContext.Tokens
                                        Where t.id = 1
                                        Select t).FirstOrDefault()
                myToken.longTokenExpireDate = Date.Now.AddDays(30)
                myToken.longTimeToken = newlongToken
                efContext.SaveChanges()
                Return True
            Catch ex As Exception
                EFHelper.LogText("fail to update Token " & ex.InnerException.ToString)
                Dim subject As String = "Please check FaceBook Token"
                Dim body As String
                body = "program fail to update token <br/>"
                body = body & "https://developers.facebook.com/tools/explorer/874014232612712/?method=GET&path=me%3Ffields%3Did%2Cname" & "<br/>"
                body = body & "From Auto Reminder"
                NotificationEmail.SentStartEmail(subject, body)
            End Try
        End If

    End Function
    ''' <summary>
    ''' 匹配一串文本中的“#** ” 或“#***#”类型的文字，这些文字组提取出来作为k11Seo关键字用
    ''' </summary>
    ''' <param name="postText"></param>
    ''' <returns>返回文字组list</returns>
    ''' <remarks></remarks>
    Public Function FilterKeyword(ByVal postText As String) As List(Of String)
        Dim listKeyWrod As New List(Of String)
        Dim regexp As String = "#([\w]+?)(?:#|\s)"
        Dim mCollection As MatchCollection = Regex.Matches(postText, regexp)
        For i = 0 To mCollection.Count - 1
            listKeyWrod.Add(mCollection.Item(i).Value.Trim)
        Next
        Return listKeyWrod
    End Function
#End Region


#Region "标题Subject"
    ''' <summary>
    ''' 根据preSubject标题字段替换标题标签，并保存至数据库
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="siteName"></param>
    ''' <param name="planType"></param>
    ''' <param name="preSubject"></param>
    ''' <param name="cateName">替换成[CATE_NAME]</param>
    ''' <remarks></remarks>
    Public Sub HandleSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal siteName As String, ByVal planType As String,
                                  ByVal preSubject As String, ByVal cateName As String, ByVal section As String)
        Dim subject As String
        Dim querySubject As String = GetFirstProductSubject(issueId, section)
        Dim query = From i In efContext.Issues
                    Where i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                    Select i
        If (String.IsNullOrEmpty(preSubject)) Then
            If (planType.Contains("HO") Or planType.Contains("HA")) Then
                If Not (String.IsNullOrEmpty(querySubject)) Then
                    subject = "[FIRSTNAME] 你好," & siteName & "为你带来： " & querySubject & "(AD)"
                Else
                    subject = "[FIRSTNAME] 你好,更多惊喜尽在" & siteName
                End If
            ElseIf (planType.Contains("HP")) Then

                subject = "[FIRSTNAME] 您的" & siteName & "专属资讯（第" & (query.Count + 1).ToString.PadLeft(2, "0") & "）期"
            End If
        Else
            If Not (String.IsNullOrEmpty(querySubject)) Then
                subject = preSubject.Replace("[FIRST_PRODUCT]", querySubject.Trim)
            Else
                subject = preSubject.Replace("[FIRST_PRODUCT]", "")
            End If
            subject = subject.Replace("[VOL_NUMBER]", (query.Count + 1).ToString.PadLeft(2, "0")).Replace("[CATE_NAME]", cateName.Trim)
        End If

        InsertIssueSubject(issueId, subject)
    End Sub
#End Region

#Region "Catagory"
    ''' <summary>
    ''' 判断用户是否需要通过网易邮件判断
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function IsNtes(siteid As Integer, planType As String) As Boolean
        Dim s = efContext.AutomationPlans.Where(Function(cs) cs.SiteID = siteid AndAlso cs.PlanType = planType).FirstOrDefault()
        Return s IsNot Nothing AndAlso s.ContactType.HasValue AndAlso s.ContactType.Value = 1
    End Function
    ''' <summary>
    ''' 获取名单里面的订阅人数
    ''' </summary>
    ''' <param name="account"></param>
    ''' <param name="cate"></param>
    ''' <returns></returns>
    Public Shared Function GetSubscriberCount(account As String, cate As String) As Integer
        Return SpreadHelper.GetSubscriptionCount(account, SEmailEfHelper.getApiKey(account), cate)
    End Function
    ''' <summary>
    ''' 获取发送名单
    ''' </summary>
    ''' <param name="cate"></param>
    ''' <returns></returns>
    Public Shared Function GetFilterCategory(account As String, cates As String(), ByVal isNetEase As Boolean) As String
        Return SpreadHelper.GetFilterCategory(account, SEmailEfHelper.getApiKey(account), cates, isNetEase)

    End Function

    Public Shared Function GetNtesLevelCount(account As String) As Integer
        Return SpreadHelper.GetNtesLevelCount(account, SEmailEfHelper.getApiKey(account))
    End Function


#End Region

End Class

