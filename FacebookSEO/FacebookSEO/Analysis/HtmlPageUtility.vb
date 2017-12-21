Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack

Public Class HtmlPageUtility


    Public Function GetHtmlDocument(ByVal pageUrl As String, ByVal cookie As String, ByVal refer As String, ByVal pageEncoding As String, Optional ByVal dg As Integer = 3) As HtmlDocument
        Try
            cookie = cookie.Trim
            pageEncoding = pageEncoding.Trim
            refer = refer.Trim
            Dim myhtmlDoc As New HtmlDocument()
            Dim resultSting As String
            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
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
            Dim response As WebResponse = request.GetResponse()
            Dim resStream As Stream = response.GetResponseStream()
            Dim myencoding As Encoding
            If (String.IsNullOrEmpty(pageEncoding)) Then
                myencoding = Encoding.UTF8
            Else
                myencoding = Encoding.GetEncoding(pageEncoding)
            End If
            Dim resStreamReader As StreamReader = New StreamReader(resStream, myencoding)
            resultSting = resStreamReader.ReadToEnd()
            myhtmlDoc.LoadHtml(resultSting)
            Common.LogText("递归次数dg=" & dg)
            Return myhtmlDoc
        Catch ex As Exception
            Common.LogText("GetHtmlDocument()->RequestURL:" & pageUrl)
            If (dg > 0) Then
                dg = dg - 1
                Return GetHtmlDocument(pageUrl, cookie, refer, pageEncoding, dg)
            Else
                Common.LogText("dg error occured.")
                Throw ex
            End If
        End Try
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
            Common.LogText("LoadXmlDoc()->RequestURL:" & url)
            Common.LogText(ex.ToString)
        End Try

        xmlDoc.LoadXml(result)
        Return xmlDoc
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

        Dim resultSting As String = ""
        Try
            cookie = cookie.Trim
            pageEncoding = pageEncoding.Trim

            Dim request As HttpWebRequest = HttpWebRequest.Create(pageUrl)
#If DEBUG Then
            '+debug# proxy
            'Dim myProxy As New WebProxy("127.0.0.1:1080", True)
            'request.Proxy = myProxy
#End If
            '模拟https请求，这个方法对tmall失败
            request.Timeout = 120000
            request.Headers.Add("Accept-Language", "zh-CN")
            request.Referer = refer
            request.Headers.Add("Cookie", cookie)
            request.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11"
            request.Method = "GET"
            request.AllowAutoRedirect = True
            Dim response As WebResponse = request.GetResponse()
            Dim resStream As Stream = response.GetResponseStream()
            Dim myencoding As Encoding
            If (pageEncoding Is Nothing Or String.IsNullOrEmpty(pageEncoding)) Then
                myencoding = Encoding.UTF8
            Else
                myencoding = Encoding.GetEncoding(pageEncoding)
            End If
            Dim resStreamReader As StreamReader = New StreamReader(resStream, myencoding)
            resultSting = resStreamReader.ReadToEnd()
            Common.LogText("递归次数dg=" & dg)

        Catch ex As Exception
            Common.LogText("GetHtmlStringByUrl()->RequestURL:" & pageUrl)
            If dg > 0 Then
                dg = dg - 1
                Return GetHtmlStringByUrl(pageUrl, cookie, refer, pageEncoding, dg)
            Else
                Common.LogText("dg error occured.")

                Throw ex
            End If
        End Try

        Return resultSting

    End Function






End Class
