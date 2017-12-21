Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying
Imports System.Net
Imports HtmlAgilityPack

Public Class Blog

    Structure EmailCampaignElement
        Dim Subject As String
        Dim description As String
        Dim Link As String
        Dim postDate As DateTime
        Dim pictureSrc As String
        Dim categoryName As String
        Dim categoryLink As String
        Dim facebookLike As String
    End Structure

    Function ReadRSS(ByVal URL As String) As EmailCampaignElement()
        Dim A As New alerter
        Try
            System.Net.ServicePointManager.Expect100Continue = False
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
            '2013/07/18 added
            'xmlDoc.Load(URL)
            xmlDoc = GetXmlDocument(URL)

            'xmlDoc.Load("http://hk.myblog.yahoo.com/master-dragon/rss")
            'xmlDoc.Load("http://rss.sina.com.cn/news/marquee/ddt.xml")
            'xmlDoc.Load("C:\Users\SamTest\Desktop\EmailAlerter 2.0\EmailAlerter 1.0\XMLFile1.xml")

            Dim n As Integer = 5 '填放产品的个数，此处为6个
            Dim itemNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes("//item") 'xmlDoc.SelectNodes(String.Format("//item[position()<{0}]", n + 2))
            Dim EmailCampaignElements(n) As EmailCampaignElement
            Dim i As Integer = 0
            For Each node As Xml.XmlNode In itemNodeList
                'Dim titleNodeText As String = node.SelectSingleNode("title").InnerText.Trim()
                'Dim desNodeText As String = node.SelectSingleNode("description").InnerText.Replace("appeared first on %%http://lungchuntin.com%%.", "").Trim()
                'Dim postDateNodeText As String = node.SelectSingleNode("pubDate").InnerText
                'Dim linkNodeText As String = node.SelectSingleNode("link").InnerText
                'EmailCampaignElements(i).Subject = titleNodeText
                'EmailCampaignElements(i).Content = desNodeText
                'EmailCampaignElements(i).Link = linkNodeText
                Dim subject As String = node.SelectSingleNode("title").InnerText.Trim()
                Dim description As String = node.SelectSingleNode("description").InnerText.Trim()
                'Dim reg As String = "<p><p>(.*)</p>"   '2013/11/27 revised, former: <p>(.*)</p><p>
                'description = Regex.Matches(description, reg)(0).Groups(1).Value
                Dim reg As String = "<p><p>(.*\s*.*)</p>"
                If (description.Contains("<p>")) Then
                    Try
                        description = Regex.Matches(description, reg)(0).Groups(1).Value
                    Catch ex As Exception
                        reg = "<p>(.*\s*.*)</p><p>"
                        description = Regex.Matches(description, reg)(0).Groups(1).Value
                    End Try
                End If
                ''description存放的是图片信息，begin
                If (description.Contains(".jpg") OrElse description.Contains(".JPG")) Then
                    description = ""
                End If
                '--------------------------------end
                Dim postDate As DateTime = DateTime.Parse(DateTime.Parse(node.SelectSingleNode("pubDate").InnerText.Trim()))
                Dim link As String = node.SelectSingleNode("link").InnerText
                Dim doc As HtmlDocument
                Try
                    doc = GetHtmlDoc(link)
                Catch ex As Exception
                    Continue For
                End Try
                Dim picSrc As String = ""
                'Dim postContent As String = ""
                Try
                    picSrc = doc.DocumentNode.SelectSingleNode("//div[@class='entry-content clearfix']/p/a/img").GetAttributeValue("src", "")

                Catch ex As Exception
                    Continue For
                End Try

                Dim categoryName As String = ""
                Dim categoryLink As String = ""
                Dim linkNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='post-meta clearfix']/p/a")
                For Each nodeItem As HtmlNode In linkNodes
                    If Not (nodeItem.InnerText.Contains("所有文章")) Then
                        categoryName = nodeItem.InnerText
                        categoryLink = nodeItem.GetAttributeValue("href", "")
                        Exit For
                    End If
                Next
                Dim facebookLike As String = "http://www.facebook.com/sharer.php?p[title]=" & subject & "&p[url]=" & link & "&p[summary]=" & description & "&p[images][0]=" & picSrc
                
                EmailCampaignElements(i).Subject = subject
                EmailCampaignElements(i).description = description
                If (String.IsNullOrEmpty(picSrc)) Then
                    EmailCampaignElements(i).description = description
                End If
                EmailCampaignElements(i).postDate = postDate
                EmailCampaignElements(i).Link = link
                EmailCampaignElements(i).pictureSrc = picSrc
                EmailCampaignElements(i).categoryName = categoryName
                EmailCampaignElements(i).categoryLink = categoryLink
                EmailCampaignElements(i).facebookLike = facebookLike
                i = i + 1
                If (i >= 6) Then
                    Exit For
                End If
            Next
            Return EmailCampaignElements
        Catch ex As Exception
            'A.LogText(ex.ToString)
            Throw New Exception(ex.ToString()) '外部处理异常，把异常以邮件形式发出
        End Try
    End Function


    Function CustomizeTempalate(ByVal Template As String, ByVal EmailTemplateElements As EmailCampaignElement()) As String
        Dim A As New alerter
        Try
            'update send template from xml data
            For i As Integer = 0 To EmailTemplateElements.Length - 1
                'Template = System.Text.RegularExpressions.Regex.Replace(Template, "<a style=""font-weight: bold;text-decoration:none;color:white;"" href=""[^""]*"" sign=" & i + 1 & "t>([^<]*)</a>", "<a style=""font-weight: bold;text-decoration:none;color:white;"" href=" & """" & _
                '                                                        EmailTemplateElements(i).Link & """" & " sign=" & i + 1 & "t>" & EmailTemplateElements(i).Subject & "</a>")
                ''Template = System.Text.RegularExpressions.Regex.Replace(Template, "<td\s*style=""color: #494949;padding-top:8px""\s*sign=" & i + 1 & "c>([^<]*)</td>", "<td style=""color: #494949;padding-top:8px"" sign=" & i + 1 & "c>" & EmailTemplateElements(i).Content.Remove(100) & "......" & "</td>")
                'Dim content As String = System.Text.RegularExpressions.Regex.Replace(EmailTemplateElements(i).description, "<(.|\n)*?>", "")
                'Template = System.Text.RegularExpressions.Regex.Replace(Template, "<td\s*style=""color: #494949;padding-top:8px""\s*sign=" & i + 1 & "c>([^<]*)</td>", "<td style=""color: #494949;padding-top:8px"" sign=" & i + 1 & "c>" & content.Remove(IIf(content.Length > 100, 100, content.Length - 1)) & "......" & "</td>")

                'Template = System.Text.RegularExpressions.Regex.Replace(Template, "<a href=""[^""]*"" style=""text-decoration:none;color:#a21818"" sign=" & i + 1 & "l>", "<a href=" & """" & EmailTemplateElements(i).Link & """" & " style=""text-decoration:none;color:#a21818"" sign=" & i + 1 & "l>")


                Dim beginIndex As Int32 = Template.IndexOf("[BEGIN_PRODUCT]")
                Dim endIndex As Int32 = Template.IndexOf("[END_PRODUCT]")
                'Dim beginLength As Int32 = "[BEGIN_PRODUCT]".Length

                If (String.IsNullOrEmpty(EmailTemplateElements(i).Subject)) Then
                    Template = Template.Remove(beginIndex, endIndex + "[END_PRODUCT]".Length - beginIndex)
                Else
                    Dim newTemplate As String = Template.Substring(beginIndex + "[BEGIN_PRODUCT]".Length, endIndex - beginIndex - "[BEGIN_PRODUCT]".Length)
                    newTemplate = newTemplate.Replace("[LINK]", EmailTemplateElements(i).Link)
                    newTemplate = newTemplate.Replace("[PICTURESRC]", EmailTemplateElements(i).pictureSrc)
                    newTemplate = newTemplate.Replace("[CATEGORYNAME]", EmailTemplateElements(i).categoryName)
                    newTemplate = newTemplate.Replace("[CATEGORYLINK]", EmailTemplateElements(i).categoryLink)
                    newTemplate = newTemplate.Replace("[SUBJECT]", EmailTemplateElements(i).Subject)
                    newTemplate = newTemplate.Replace("[DESCRIPTION]", EmailTemplateElements(i).description)
                    If (Now.DayOfWeek.ToString() = "6") Then
                        newTemplate = newTemplate.Replace("[FACEBOOKLIKE]", "")
                    Else
                        newTemplate = newTemplate.Replace("[FACEBOOKLIKE]", EmailTemplateElements(i).facebookLike)
                    End If
                    Template = Template.Remove(beginIndex, endIndex + "[END_PRODUCT]".Length - beginIndex)
                    Template = Template.Insert(beginIndex, newTemplate)
                End If
                'newTemplate = newTemplate.Format("", )
            Next
            Return Template
        Catch ex As Exception
            'A.LogText(ex.ToString)
            Throw New Exception(ex.ToString()) '外部处理异常，把异常以邮件形式发出
        End Try
        'System.Web.HttpUtility.HtmlEncode(">")

    End Function

    Function CreateCampaignBlog(ByVal LoginName As String, ByVal AppID As String, ByVal senderName As String, ByVal SenderEmail As String, ByVal EmailBody As String, ByVal EmailCampaignElements As EmailCampaignElement()) As String
        Dim A As New alerter
        Try

            'get data from local database
            'Dim emailAlerterData As EmailTableAdapters.RssSubscriptionsTableAdapter = New EmailTableAdapters.RssSubscriptionsTableAdapter()
            'Dim dt As System.Data.DataTable = emailAlerterData.GetDataById(1)
            'Dim rssId As String = dt.Rows(0)("rssid").ToString()
            'Dim url As String = dt.Rows(0)("url").ToString()

            'get contact lists from spread service
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            Dim subscriptionData As System.Data.DataSet = mySpread.getAllSubscription(LoginName, AppID)

            'Dim temp As String = ""
            Dim ContactLists(subscriptionData.Tables(0).Select("status='Invisible'").Length - 1) As String
            Dim i As Integer
            For Each row As System.Data.DataRow In subscriptionData.Tables(0).Rows
                If row("status").ToString() = "Invisible" Then
                    ContactLists(i) = row("Subscription Name") 'temp = temp & "," & row(0).ToString()
                    i += 1
                End If
            Next

            'create campaign
            'Dim targetSubscription As String() = temp.Split(",")
            Dim intervalTime As Integer = -1
            Dim campaignId As Integer
            Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()
            'For j As Integer = 0 To Math.Min(2, EmailCampaignElements.Length)
            '    myCampaign.campaignName &= EmailCampaignElements(j).Subject & IIf(j = 0, "/", "/")
            'Next
            ' myCampaign.campaignName = EmailCampaignElements(2).Subject



            For j As Integer = 0 To EmailCampaignElements.Length - 1
                If Not EmailCampaignElements(j).Subject.Contains("每日生肖運程") Then
                    '2013/08/29 added，"把龙震天的发送的Subject调整为第一篇博客的title"
                    'myCampaign.campaignName = EmailCampaignElements(j + 1).Subject
                    myCampaign.campaignName = EmailCampaignElements(j).Subject
                    Exit For
                End If

            Next

            If String.IsNullOrEmpty(myCampaign.campaignName) Then
                myCampaign.campaignName = "龍震天-每日生肖運程"
            End If


            myCampaign.from = senderName
            myCampaign.fromEmail = SenderEmail
            myCampaign.subject = myCampaign.campaignName
            myCampaign.content = EmailBody
            Dim de As Date = Now
            myCampaign.schedule = de
            'send campaign
            campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
            mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            'add by gary start 20130801
            Try
                Dim subject2 = senderName & " 开始发送邮件"
                Dim SpreadTemplate2 = subject2 & "<br />状态：正在发送<br />" & " 标题:" & myCampaign.campaignName & "<br />内容：<br />" & EmailBody
                mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "自动化发送", "emailalerter@reasonables.com", subject2, SpreadTemplate2)
            Catch ex As Exception
                'A.LogText(ex.ToString)
                Throw New Exception(ex.ToString()) '外部处理异常，把异常以邮件形式发出去
            End Try
            'add by gary end 20130801
            Return myCampaign.subject
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' 根据站点的Rss Url获取
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetXmlDocument(ByVal url As String)
        Dim request As System.Net.HttpWebRequest = System.Net.WebRequest.Create(url)
        Dim response As System.Net.HttpWebResponse
        Try
            request.Timeout = 60000
            response = request.GetResponse()
        Catch ex As Exception
            request.Timeout = 120000
            response = request.GetResponse()
        End Try
        Dim stream As System.IO.Stream = response.GetResponseStream()
        Dim streamReader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8) 'System.Text.Encoding.UTF8

        Dim result As String = streamReader.ReadToEnd()

        ''2013/09/09,xml出错排除特殊字符串，暂时不用
        'Dim builder As New System.Text.StringBuilder()
        'For i As Integer = 0 To result.Length - 1
        '    If (Regex.IsMatch(result(i), "[\x01-\x7f]|[\xc0-\xdf][\x80-\xbf]|[\xe0-\xef][\x80-\xbf]{2}|[\xf0-\xff][\x80-\xbf]{3}")) OrElse _
        '        (Regex.IsMatch(result(i), "[\u4e00-\u9fbb]+$")) OrElse (Regex.IsMatch(result(i), "\xEA84BF-\xEAA980")) OrElse _
        '         (Regex.IsMatch(result(i), "/pP$/u")) OrElse (result(i) = "「" OrElse result(i) = "」") Then
        '        builder.Append(result(i))
        '    End If
        'Next
        'result = builder.ToString()
        ''-----------------------------------------------------

        Dim xmldoc As New System.Xml.XmlDocument
        xmldoc.LoadXml(result)
        Return xmldoc
    End Function

    Private Function GetHtmlDoc(ByVal pageUrl As String)
        Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(pageUrl), HttpWebRequest)
        request.Timeout = 120000
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim myStream As System.IO.Stream = response.GetResponseStream()
        Dim doc As HtmlDocument = New HtmlDocument()
        doc.Load(myStream, System.Text.Encoding.GetEncoding("utf-8"))
        Return doc
    End Function
End Class
