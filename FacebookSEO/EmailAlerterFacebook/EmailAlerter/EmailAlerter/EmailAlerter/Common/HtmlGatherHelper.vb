Imports HtmlAgilityPack
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions

Public Class HtmlGatherHelper
    Public Shared Function GetPageContentWithAgent(url As String, pageEncode As Encoding) As HtmlDocument
        Dim htmlDoc As HtmlDocument = Nothing
        Try

            htmlDoc = New HtmlDocument()
            Dim httpWebRequest As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            httpWebRequest.UserAgent = "Mozilla/4.0(compatibleMSIE5.01WindowsNT5.0)"
            httpWebRequest.Method = "GET"
            httpWebRequest.AllowAutoRedirect = False
            httpWebRequest.Timeout = 15000

            Dim response As HttpWebResponse = TryCast(httpWebRequest.GetResponse(), HttpWebResponse)
            htmlDoc.Load(response.GetResponseStream(), pageEncode)
            response.Close()
        Catch ex As Exception

        End Try

        Return htmlDoc
    End Function

    Public Shared Function GetMatchValue(regexPattern As String, strValue As String) As String
        If String.IsNullOrEmpty(regexPattern) OrElse String.IsNullOrEmpty(strValue) Then Return strValue
        Dim reg As Regex = New Regex(regexPattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline)
        Dim match As Match = reg.Match(strValue)
        Return CStr(IIf(match.Success, match.Value, String.Empty))
    End Function

    Public Shared Function GetSingleValueFromHtml(htmlDoc As HtmlDocument, xpath As String, regexPattern As String) _
        As String
        Return GetSingleValue(htmlDoc, xpath, regexPattern, True)
    End Function

    Public Shared Function GetSingleValueFromText(htmlDoc As HtmlDocument, xpath As String, regexPattern As String) _
        As String
        Return GetSingleValue(htmlDoc, xpath, regexPattern, False)
    End Function

    Private Shared Function GetSingleValue(htmlDoc As HtmlDocument, xpath As String, regexPattern As String,
                                           htmlOrText As Boolean) As String
        If htmlDoc Is Nothing Then Return String.Empty
        If Not String.IsNullOrEmpty(xpath) Then
            Dim node As HtmlNode = htmlDoc.DocumentNode.SelectSingleNode(xpath)
            If node IsNot Nothing Then
                Return GetMatchValue(regexPattern, CStr(IIf(htmlOrText, node.InnerHtml, node.InnerText)))
            End If
        End If

        Return String.Empty
    End Function

    Private Shared Function GetValueList(htmlDoc As HtmlDocument, xpath As String, regexPattern As String,
                                         htmlOrText As Boolean) As IList(Of String)
        Dim result As IList(Of String) = New List(Of String)()

        If htmlDoc Is Nothing Or String.IsNullOrEmpty(xpath) Then Return result

        Dim nodelist As HtmlNodeCollection = htmlDoc.DocumentNode.SelectNodes(xpath)
        Dim singlevalue As String
        If nodelist IsNot Nothing AndAlso nodelist.Count > 0 Then
            For i As Integer = 0 To nodelist.Count - 1
                singlevalue = GetMatchValue(regexPattern,
                                            CStr(IIf(htmlOrText, nodelist(i).InnerHtml, nodelist(i).InnerText)))
                If Not String.IsNullOrEmpty(singlevalue) Then
                    result.Add(singlevalue)
                End If
            Next
        End If
        Return result
    End Function

    Public Shared Function GetTextList(htmlDoc As HtmlDocument, xpath As String, regexPattern As String) _
        As IList(Of String)
        Return GetValueList(htmlDoc, xpath, regexPattern, False)
    End Function

    Public Shared Function GetHtmlList(htmlDoc As HtmlDocument, xpath As String, regexPattern As String) _
        As IList(Of String)
        Return GetValueList(htmlDoc, xpath, regexPattern, True)
    End Function

    Public Shared Function GetLinksFromHtml(htmlDoc As HtmlDocument, xpath As String) As IList(Of String)
        Dim result As IList(Of String) = New List(Of String)()

        If htmlDoc Is Nothing Or String.IsNullOrEmpty(xpath) Then Return result

        Dim nodelist As HtmlNodeCollection = htmlDoc.DocumentNode.SelectNodes(xpath)
        Dim singlevalue As String
        If nodelist IsNot Nothing AndAlso nodelist.Count > 0 Then
            For i As Integer = 0 To nodelist.Count - 1
                singlevalue = GetLinkFromNode(nodelist(i))
                If Not String.IsNullOrEmpty(singlevalue) Then
                    result.Add(singlevalue)
                End If
            Next

        End If

        Return result
    End Function

    Public Shared Function GetLinkFromHtml(htmlDoc As HtmlDocument, xpath As String) As String
        If htmlDoc Is Nothing Or String.IsNullOrEmpty(xpath) Then Return String.Empty
        Dim node As HtmlNode = htmlDoc.DocumentNode.SelectSingleNode(xpath)
        Return GetLinkFromNode(node)
    End Function

    Private Shared Function GetLinkFromNode(node As HtmlNode) As String
        Dim result As String = String.Empty
        If node IsNot Nothing Then
            If Not node.Attributes("href") Is Nothing Then
                result = node.Attributes("href").Value
            End If
        End If
        Return result
    End Function
End Class
