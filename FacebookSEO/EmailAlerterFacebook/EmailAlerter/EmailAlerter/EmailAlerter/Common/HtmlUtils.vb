Imports System.Text.RegularExpressions


Public Class HtmlUtils
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clean removes any HTML Tags, Entities (and optionally any punctuation) from
    ''' a string
    ''' </summary>
    ''' <remarks>
    ''' Encoded Tags are getting decoded, as they are part of the content!
    ''' </remarks>
    ''' <param name="HTML">The Html to clean</param>
    ''' <param name="RemovePunctuation">A flag indicating whether to remove punctuation</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	11/16/2004	created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Clean(ByVal HTML As String, ByVal RemovePunctuation As Boolean) As String

        'First remove any HTML Tags ("<....>")
        HTML = StripTags(HTML, True)

        'Second replace any HTML entities (&nbsp; &lt; etc) through their char symbol
        HTML = System.Web.HttpUtility.HtmlDecode(HTML)

        'Finally remove any punctuation
        If RemovePunctuation Then
            HTML = StripPunctuation(HTML, True)
        End If

        Return HTML
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' FormatText replaces <br/> tags by LineFeed characters
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="HTML">The HTML content to clean up</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	12/13/2004	created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FormatText(ByVal HTML As String, ByVal RetainSpace As Boolean) As String

        'Match all variants of <br> tag (<br>, <BR>, <br/>, including embedded space
        Dim brMatch As String = "\s*<\s*[bB][rR]\s*/\s*>\s*"

        'Set up Replacement String
        Dim RepString As String
        If RetainSpace Then
            RepString = " "
        Else
            RepString = ""
        End If

        'Replace Tags by replacement String and return mofified string
        Return System.Text.RegularExpressions.Regex.Replace(HTML, brMatch, ControlChars.Lf)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format a domain name including link
    ''' </summary>
    ''' <param name="Website">The domain name to format</param>
    ''' <returns>The formatted domain name</returns>
    ''' <history>
    '''		[cnurse]	09/29/2005	moved from Globals
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FormatWebsite(ByVal Website As Object) As String

        FormatWebsite = ""

        If Not IsDBNull(Website) Then
            If Trim(Website.ToString()) <> "" Then
                If Convert.ToBoolean(InStr(1, Website.ToString(), ".")) Then
                    FormatWebsite = "<a href=""" &
                                    IIf(Convert.ToBoolean(InStr(1, Website.ToString(), "://")), "", "http://").ToString &
                                    Website.ToString() & """>" & Website.ToString() & "</a>"
                Else
                    FormatWebsite = Website.ToString()
                End If
            End If
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shorten returns the first (x) characters of a string
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="txt">The text to reduces</param>
    ''' <param name="length">The max number of characters to return</param>
    ''' <param name="suffix">An optional suffic to append to the shortened string</param>
    ''' <returns>The shortened string</returns>
    ''' <history>
    '''		[cnurse]	11/16/2004	created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Shorten(ByVal txt As String, ByVal length As Integer, ByVal suffix As String) As String
        Dim results As String
        If txt.Length > length Then
            results = txt.Substring(0, length) & suffix
        Else
            results = txt
        End If
        Return results
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' StripTags removes the HTML Tags from the content
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="HTML">The HTML content to clean up</param>
    ''' <param name="RetainSpace">Indicates whether to replace the Tag by a space (true) or nothing (false)</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	11/16/2004	documented
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripTags(ByVal HTML As String, ByVal RetainSpace As Boolean) As String

        'Set up Replacement String
        Dim RepString As String
        If RetainSpace Then
            RepString = " "
        Else
            RepString = ""
        End If

        'Replace Tags by replacement String and return mofified string
        Return System.Text.RegularExpressions.Regex.Replace(HTML, "<[^>]*>", RepString)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' StripPunctuation removes the Punctuation from the content
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="HTML">The HTML content to clean up</param>
    ''' <param name="RetainSpace">Indicates whether to replace the Punctuation by a space (true) or nothing (false)</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	11/16/2004	documented
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripPunctuation(ByVal HTML As String, ByVal RetainSpace As Boolean) As String

        'Create Regular Expression objects
        Dim punctuationMatch As String = "[~!#\$%\^&*\(\)-+=\{\[\}\]\|;:\x22'<,>\.\?\\\t\r\v\f\n]"
        Dim afterRegEx As New System.Text.RegularExpressions.Regex(punctuationMatch & "\s")
        Dim beforeRegEx As New System.Text.RegularExpressions.Regex("\s" & punctuationMatch)

        'Define return string
        Dim retHTML As String = HTML & " "
        'Make sure any punctuation at the end of the String is removed

        'Set up Replacement String
        Dim RepString As String
        If RetainSpace Then
            RepString = " "
        Else
            RepString = ""
        End If

        While beforeRegEx.IsMatch(retHTML)
            'Strip punctuation from beginning of word
            retHTML = beforeRegEx.Replace(retHTML, RepString)
        End While

        While afterRegEx.IsMatch(retHTML)
            'Strip punctuation from end of word
            retHTML = afterRegEx.Replace(retHTML, RepString)
        End While

        ' Return modified string
        Return retHTML
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' StripWhiteSpace removes the WhiteSpace from the content
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="HTML">The HTML content to clean up</param>
    ''' <param name="RetainSpace">Indicates whether to replace the WhiteSpace by a space (true) or nothing (false)</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	12/13/2004	documented
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripWhiteSpace(ByVal HTML As String, ByVal RetainSpace As Boolean) As String

        'Set up Replacement String
        Dim RepString As String
        If RetainSpace Then
            RepString = " "
        Else
            RepString = ""
        End If

        'Replace Tags by replacement String and return mofified string
        If HTML = Null.NullString Then
            Return Null.NullString
        Else
            Return System.Text.RegularExpressions.Regex.Replace(HTML, "\s+", RepString)
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' StripNonWord removes any Non-Word Character from the content
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <param name="HTML">The HTML content to clean up</param>
    ''' <param name="RetainSpace">Indicates whether to replace the Non-Word Character by a space (true) or nothing (false)</param>
    ''' <returns>The cleaned up string</returns>
    ''' <history>
    '''		[cnurse]	1/28/2005	created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripNonWord(ByVal HTML As String, ByVal RetainSpace As Boolean) As String

        'Set up Replacement String
        Dim RepString As String
        If RetainSpace Then
            RepString = " "
        Else
            RepString = ""
        End If

        'Replace Tags by replacement String and return mofified string
        If HTML Is Nothing Then
            Return HTML
        Else
            Return System.Text.RegularExpressions.Regex.Replace(HTML, "\W*", RepString)
        End If
    End Function

    ''' <summary>
    ''' 过滤指定HTML中的Script标记及其内容
    ''' </summary>
    ''' <param name="HTML"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StripScript(ByVal HTML As String) As String

        If HTML Is Nothing Then
            Return HTML
        End If
        ' TODO 过滤HTML中的Script标签
        Dim ScriptPattern As String = "<script.*?</script>"
        Return Regex.Replace(HTML, ScriptPattern, "", RegexOptions.IgnoreCase)
    End Function

    ''' <summary>
    ''' 净化HTML代码,如移除Script、link等标签
    ''' </summary>
    ''' <param name="HTML"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SanitizeHtml(ByVal HTML As String) As String
        ' TODO 过滤HTML中有威胁的标签
        Return HTML
    End Function
End Class
