Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying

Public Class Ladies
    Structure EmailCampaignElementMagazine
        Dim Subject As String
        Dim Content As String
        Dim Link As String
        Dim Img As String
    End Structure


    Sub SendEmailMagazine(ByVal r As RssSubscription)
        Dim A As New alerter
        Try
            Dim EmailCampaignElementsM(9) As EmailCampaignElementMagazine
            Dim rss As String() = r.URL.Split("^")
            For i As Integer = 0 To rss.Length - 1
                Dim RssLink As String = rss(i)
                Dim EmailElement As EmailCampaignElementMagazine = ReadRSSmagazine(RssLink)
                EmailCampaignElementsM(i).Subject = EmailElement.Subject
                EmailCampaignElementsM(i).Link = EmailElement.Link
                EmailCampaignElementsM(i).Content = EmailElement.Content
                EmailCampaignElementsM(i).Img = EmailElement.Img
            Next
            Dim subject As String = ""
            If Now.DayOfWeek = 1 Or Now.DayOfWeek = 3 Then
                subject = EmailCampaignElementsM(0).Subject
            ElseIf Now.DayOfWeek = 5 Then
                subject = EmailCampaignElementsM(1).Subject
            End If
            Dim template As String() = r.Template.Split("^")
            Dim html As String = CustomizeTempalateMagazine(template(0), template(1), template(2),
                                                            EmailCampaignElementsM)
            Dim scheduleTime As DateTime = Date.Parse("12:30:00")
            A.CreateCampaign(subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, html, scheduleTime)
            A.UpdateSentLog(r.RssId, subject)
            Common.LogText(subject)

        Catch ex As Exception
            Common.LogException(ex)
        End Try
    End Sub

    Function ReadRSSmagazine(ByVal URL As String) As EmailCampaignElementMagazine
        Dim A As New alerter
        Dim EmailCampaignElements As New EmailCampaignElementMagazine
        Try
            System.Net.ServicePointManager.Expect100Continue = False
            ' Dim ElementAndTitle As ElementAndTitle
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
            xmlDoc.Load(URL)


            Dim nsmgr As Xml.XmlNamespaceManager = New XmlNamespaceManager(xmlDoc.NameTable)
            nsmgr.AddNamespace("content", "http://purl.org/rss/1.0/modules/content/")

            Dim itemNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes("//item")

            'For Each node As Xml.XmlNode In itemNodeList
            Dim node As Xml.XmlNode = itemNodeList(0)
            Dim titleNodeText As String = node.SelectSingleNode("title").InnerText.Trim()
            Dim desNodeText As String = node.SelectSingleNode("description").InnerText.Trim()
            'Dim postDateNodeText As String = node.SelectSingleNode("pubDate").InnerText
            Dim linkNodeText As String = node.SelectSingleNode("link").InnerText
            Dim contentText As String = node.SelectSingleNode("content:encoded", nsmgr).InnerText
            ' New System.Xml.XmlNamespaceManager) '    GetPrefixOfNamespace("content")
            ' Dim ImgLinkCollection As MatchCollection = Regex.Matches(contentText, "<img[\s\S]*?src=""([^""]*)"" a", RegexOptions.IgnoreCase)
            Dim ImgLinkCollection As MatchCollection = Regex.Matches(contentText, "<img[\s\S]*?src=""([^""]*)""",
                                                                     RegexOptions.IgnoreCase)
            Dim ImgLink As String = ""
            If ImgLinkCollection.Count > 0 Then
                ImgLink = ImgLinkCollection.Item(0).Groups(1).Value
            End If
            If String.IsNullOrEmpty(ImgLink) Or ImgLink.Contains("sina_weibo") Then
                ImgLink = "https://app.reasonablespread.com//SpreaderFiles/3602/files/upload/Ladies_150x150.png"
            End If
            EmailCampaignElements.Subject = titleNodeText
            '2014/05/04 added,邮件模板排版出错，begin
            'If desNodeText.Length >= 111 Then
            '    EmailCampaignElements.Content = desNodeText.Remove(110)
            'Else
            '    EmailCampaignElements.Content = desNodeText
            'End If
            '2014/05/04 end

            EmailCampaignElements.Link = linkNodeText
            EmailCampaignElements.Img = ImgLink
        Catch ex As Exception
            Common.LogException(ex)
        End Try

        '  Next
        Return EmailCampaignElements
    End Function


    Function CustomizeTempalateMagazine(ByVal Template As String, ByVal TopicTemplate As String,
                                        ByVal ArticleTemplate As String,
                                        ByVal Elements() As EmailCampaignElementMagazine) As String
        Dim A As New alerter
        Try
            TopicTemplate = String.Format(TopicTemplate, Elements(0).Link, Elements(0).Subject, Elements(0).Link,
                                          Elements(0).Img, Elements(0).Content, Elements(0).Link,
                                          Elements(1).Link, Elements(1).Subject, Elements(1).Link, Elements(1).Img,
                                          Elements(1).Content, Elements(1).Link,
                                          Elements(2).Link, Elements(2).Subject, Elements(2).Link, Elements(2).Img,
                                          Elements(2).Content, Elements(2).Link)
            ArticleTemplate = String.Format(ArticleTemplate, Elements(3).Link, Elements(3).Img, Elements(3).Link,
                                            Elements(3).Subject,
                                            Elements(4).Link, Elements(4).Img, Elements(4).Link, Elements(4).Subject,
                                            Elements(5).Link, Elements(5).Img, Elements(5).Link, Elements(5).Subject,
                                            Elements(6).Link, Elements(6).Img, Elements(6).Link, Elements(6).Subject,
                                            Elements(7).Link, Elements(7).Img, Elements(7).Link, Elements(7).Subject,
                                            Elements(8).Link, Elements(8).Img, Elements(8).Link, Elements(8).Subject)

            Template = String.Format(Template, TopicTemplate, ArticleTemplate)
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
        Return Template
    End Function
End Class
