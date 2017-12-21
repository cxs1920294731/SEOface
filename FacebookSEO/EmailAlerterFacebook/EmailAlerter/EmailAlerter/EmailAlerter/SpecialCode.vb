Imports System.Text
Imports HtmlAgilityPack

Public Class SpecialCode
    ''' <summary>
    ''' 给FocalPrice加上对应的追踪代码
    ''' </summary>
    ''' <param name="template">模板</param>
    ''' <param name="siteName">站点名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddSpecialCodeForFocalPrice(ByVal template As String, ByVal siteName As String)
        Dim utmsource As String = "RSpread"
        Dim utmMedia As String = "EM"
        Dim byteArray As Byte() = Encoding.GetEncoding("gb2312").GetBytes(template) 'System.Text.Encoding.UTF8.GetBytes(SpreadTemplate)
        Dim stream As New System.IO.MemoryStream(byteArray)
        Dim doc As New HtmlDocument
        doc.Load(stream)
        Dim hrefNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a")
        Dim ci As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim cal As System.Globalization.Calendar = ci.Calendar
        Dim cwr As System.Globalization.CalendarWeekRule = ci.DateTimeFormat.CalendarWeekRule
        Dim dow As DayOfWeek = ci.DateTimeFormat.FirstDayOfWeek
        Dim nowWeek As String = cal.GetWeekOfYear(DateTime.Now, cwr, dow).ToString() '当前时间是第几周
        For Each node As HtmlNode In hrefNodes
            Dim href As String = node.GetAttributeValue("href", "")
            If Not (href.Contains("?")) Then '该链接没有添加追踪代码的链接
                Dim firstIndex As Integer = 0 ' href.LastIndexOf("/", href.LastIndexOf("/") - 1) + 1
                Dim lastIndex As Integer = 0 ' href.LastIndexOf("/")
                Dim skuName As String = "" ' href.Substring(firstIndex, lastIndex - firstIndex)
                Try
                    firstIndex = href.LastIndexOf("/", href.LastIndexOf("/") - 1) + 1
                    lastIndex = href.LastIndexOf("/")
                    skuName = href.Substring(firstIndex, lastIndex - firstIndex)
                Catch ex As Exception
                    '待确定处理
                End Try
                
                Dim utmCampaign As String = "DM_" & nowWeek & siteName & "_" & skuName
                href = href & "?utm_source=" & utmsource & "&utm_media=" & utmMedia & "&utm_campaign=" & utmCampaign & "&source=rspread"
                node.SetAttributeValue("href", href)
            ElseIf (href.Contains("?uri=")) Then  '该链接已经添加追踪代码参数
                'href特殊类型：http://special.focalprice.com/other-advinfo-id-130.html?uri=aHR0cDovL3NwZWNpYWwuZm9jYWxwcmljZS5jb20vMjMxLUxpcXVpZGFjaSVDMyVCM24uaHRtbD91dG1fc291cmNlPUZQJnV0bV9tZWRpdW09MTMxMkVDUyZ1dG1fY2FtcGFpZ249RlBfMTMxMkVDU19IU0JO
                Dim index1 As Integer = href.IndexOf("uri=")
                href = href.Substring(index1 + 4, href.Length - index1 - 4)
                Dim bytes() As Byte = Convert.FromBase64String(href)
                href = Encoding.GetEncoding("utf-8").GetString(bytes)
                'href = href.Substring(href.IndexOf("?"), href.Length - href.IndexOf("?"))
                href = href.Substring(0, href.IndexOf("?"))

                Dim efHelper As New Analysis.EFHelper
                Dim hd As HtmlDocument   '= efHelper.GetHtmlDocument2(href, 120000)
                Dim bannerText As String = "Banner" ' hd.DocumentNode.SelectSingleNode("//div[@class='topmenu_l']").InnerText
                Dim index As Integer = 0  'bannerText.LastIndexOf("»")
                'bannerText = bannerText.Substring(index + 1, bannerText.Length - index - 1)
                Try
                    hd = efHelper.GetHtmlDocument2(href, 120000)
                    bannerText = hd.DocumentNode.SelectSingleNode("//div[@class='topmenu_l']").InnerText
                    index = bannerText.LastIndexOf("»")
                    bannerText = bannerText.Substring(index + 1, bannerText.Length - index - 1)
                Catch ex As Exception
                    'Ignore
                End Try
                Dim utmCampaign As String = "DM_" & nowWeek & siteName & "_" & bannerText
                href = href & "?utm_source=" & utmsource & "&utm_media=" & utmMedia & "&utm_campaign=" & utmCampaign & "&source=rspread"
                node.SetAttributeValue("href", href)
            End If
        Next
        template = doc.DocumentNode.OuterHtml.ToString()
        Return template
    End Function

    ''' <summary>
    ''' 添加固定格式的追踪代码
    ''' </summary>
    ''' <param name="codeStyle">追踪代码的格式</param>
    ''' <param name="template">填充好产品的模板</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddSpecialCode(ByVal codeStyle As String, ByVal template As String) As String
        If (codeStyle.Contains("[yyyyMMdd]")) Then '增加时间标签：[yyyyMMdd]
            codeStyle = codeStyle.Replace("[yyyyMMdd]", Now.ToString("yyyyMMdd"))
        End If
        '使用load（）读取template经常出现乱码现象。遂修改为loadhtml（）方法读取
        'Dim byteArray As Byte() = Encoding.GetEncoding("gb2312").GetBytes(template) 'System.Text.Encoding.UTF8.GetBytes(SpreadTemplate)
        'Dim stream As New System.IO.MemoryStream(byteArray)
        'doc.Load(stream)
        Dim doc As New HtmlDocument
        doc.LoadHtml(template)
        Dim hrefNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a")
        For Each node As HtmlNode In hrefNodes
            Dim href As String = node.GetAttributeValue("href", "")
            If (href.Contains("?")) Then
                href = href & codeStyle.Replace("?", "&")
            Else
                href = href & codeStyle
            End If
            node.SetAttributeValue("href", href)
        Next
        template = doc.DocumentNode.OuterHtml.ToString()
        Return template
    End Function
End Class
