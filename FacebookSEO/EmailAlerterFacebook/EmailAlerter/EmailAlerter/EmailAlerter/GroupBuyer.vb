Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying
Public Class GroupBuyer

    Structure Campaign
        Dim Subject As String
        Dim Content As String
    End Structure

    Structure EmailCampaignElementDeal
        Dim Subject As String
        'Dim Content As String
        Dim Link As String
        Dim Price As String
        Dim Discount As String
        Dim DiscountPrice As String
        Dim Saved As String
        Dim LargeBuyButton As String
        Dim BuyButton As String
        Dim ProductImage As String
        Dim PubDate As String
        Dim DealCategory As String
        Dim Index As Integer

    End Structure

    Structure ElementAndTitle
        Dim ChannelTitle As String
        Dim PubDate As DateTime
        Dim EmailCampaignElements() As EmailCampaignElementDeal
        Dim AllCategories As String
    End Structure

    Sub SendEmailDeal(ByVal r As RssSubscription)
        Dim A As New alerter
        Try
            '  r.URL = "C:\Users\Suwen\Desktop\fff.xml"
            Dim ElementAndTitle = ReadRSSDeal(r.URL)

            Dim PubDateTime As DateTime = ElementAndTitle.pubdate
            Dim pubdate As String = PubDateTime.ToString("yyyy年MM月dd日")
            Dim subject As String = ""

            Dim Templates As String() = r.Template.Split("^")
            Dim FTemplate As String = Templates(0)
            Dim s As String = FTemplate.Replace(vbCrLf, "")
            Dim SubTemplate1 As String = Templates(1)
            Dim SubTemplate2 As String = Templates(2)
            Dim SubTemplate3 As String = Templates(3)

            Dim categories As String() = r.Categories.Split(",")
            For Each category As String In categories
                Try
                    Dim campaign As Campaign = CustomizeTempalateDeal(r.SpreadLogin, r.AppID, category, FTemplate, SubTemplate1, SubTemplate2, SubTemplate3, ElementAndTitle.EmailCampaignElements, pubdate)
                    If Not String.IsNullOrEmpty(campaign.Subject) Then
                        subject = campaign.Subject
                    Else
                        subject = ElementAndTitle.channeltitle
                    End If
                    Dim CampaignName As String = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & category & ")"
                    '(Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year).ToString & "(favorite:" & category & ")"
                    Dim SentSuccess As Boolean = CreateCampaignDeal(category, CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, campaign.Content)
                    If SentSuccess = True Then
                        A.UpdateSentLog(r.RssId, subject)
                    End If
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Sub

    Function ReadRSSDeal(ByVal URL As String) As ElementAndTitle
        Dim A As New alerter
        Try

            Dim AllCategories As String = ""

            System.Net.ServicePointManager.Expect100Continue = False
            Dim ElementAndTitle As ElementAndTitle
            Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument()
            Dim result As String = ""
            Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(URL)
            Dim response As System.Net.HttpWebResponse = request.GetResponse()
            Dim receiveStream As IO.Stream = response.GetResponseStream()
            Dim readStream As New IO.StreamReader(receiveStream, System.Text.Encoding.UTF8)
            result = readStream.ReadToEnd()


            Try
                Dim Collection As MatchCollection = Regex.Matches(result, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", RegexOptions.IgnoreCase)

                If Collection.Count > 0 Then
                    Dim myTitle As String = Collection.Item(0).Groups(1).Value
                    Dim myTitleNode As String = "<title><![CDATA[" & myTitle & "]]></title>"
                    result = System.Text.RegularExpressions.Regex.Replace(result, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", myTitleNode)
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

            xmlDoc.LoadXml(result)
            'Try
            '    xmlDoc.Load(URL)
            'Catch ex As Exception
            '    LogText(ex.ToString)
            'End Try


            'xmlDoc.Load("http://hk.myblog.yahoo.com/master-dragon/rss")
            'xmlDoc.Load("http://rss.sina.com.cn/news/marquee/ddt.xml")
            'xmlDoc.Load("C:\Users\SamTest\Desktop\EmailAlerter 2.0\EmailAlerter 1.0\XMLFile1.xml")
            Dim n As Integer = xmlDoc.SelectNodes("//item").Count
            Dim itemNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes("//item")
            'Dim channelNode As Xml.XmlNodeList = xmlDoc.SelectNodes("//channel")
            Dim ChannelTitle As String = System.Web.HttpUtility.HtmlDecode(xmlDoc.SelectSingleNode("//channel").SelectSingleNode("title").InnerText.Trim())
            Dim PubDateString As String = xmlDoc.SelectSingleNode("//channel").SelectSingleNode("pubDate").InnerText.Trim()

            Dim PubDate As DateTime = PubDateString.Substring(5, 20)
            'Dim s As DateTime = "29 Jan 2012 10:42:25"


            'Dim n As Integer = 6
            'Dim itemNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes(String.Format("//item[position()<{0}]", n + 2))
            Dim EmailCampaignElements(n - 1) As EmailCampaignElementDeal
            Dim i As Integer
            For Each node As Xml.XmlNode In itemNodeList
                Try

                    Dim titleNodeText As String = System.Web.HttpUtility.HtmlEncode(node.SelectSingleNode("title").InnerText.Trim())
                    Dim link_category As String = node.SelectSingleNode("category").InnerText.Trim()
                    Dim priceText As String = node.SelectSingleNode("description").SelectSingleNode("price").InnerText.Trim()
                    Dim discountText As String = node.SelectSingleNode("description").SelectSingleNode("discount").InnerText.Trim()
                    Dim savedText As String = node.SelectSingleNode("description").SelectSingleNode("saved").InnerText.Trim()
                    Dim discountPriceText As String = node.SelectSingleNode("description").SelectSingleNode("discountPrice").InnerText.Trim()
                    Dim productImageText As String = node.SelectSingleNode("description").SelectSingleNode("productImage").InnerText.Trim()
                    Dim largeBuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("largeBuyButton").InnerText.Trim()
                    Dim BuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("BuyButton").InnerText.Trim()
                    Dim postDateNodeText As String = node.SelectSingleNode("pubDate").InnerText

                    Dim linkNodeText As String = node.SelectSingleNode("link").InnerText
                    linkNodeText = linkNodeText & """ link_category=""" & link_category

                    EmailCampaignElements(i).Subject = titleNodeText
                    EmailCampaignElements(i).Link = linkNodeText
                    EmailCampaignElements(i).Price = "HK$" & priceText
                    EmailCampaignElements(i).Discount = discountText
                    EmailCampaignElements(i).Saved = "HK$" & savedText
                    EmailCampaignElements(i).DiscountPrice = "HK$" & discountPriceText
                    EmailCampaignElements(i).LargeBuyButton = largeBuyButtonText
                    EmailCampaignElements(i).BuyButton = BuyButtonText
                    EmailCampaignElements(i).ProductImage = productImageText
                    EmailCampaignElements(i).DealCategory = link_category

                    EmailCampaignElements(i).PubDate = postDateNodeText
                    EmailCampaignElements(i).Index = i
                    i += 1
                    AllCategories = AllCategories & link_category
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next
            ElementAndTitle.EmailCampaignElements = EmailCampaignElements
            If String.IsNullOrEmpty(ChannelTitle) Then
                ElementAndTitle.ChannelTitle = "Group Buyer 今日團購資訊"
            Else
                ElementAndTitle.ChannelTitle = ChannelTitle
            End If

            ElementAndTitle.PubDate = PubDate
            ElementAndTitle.AllCategories = AllCategories
            Return ElementAndTitle
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Function

    Function CreateCampaignDeal(ByVal Favorite As String, ByVal CampaignName As String, ByVal Subject As String, ByVal LoginName As String, ByVal AppID As String, ByVal senderName As String, ByVal SenderEmail As String, ByVal EmailBody As String) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try

            'get data from local database
            'Dim emailAlerterData As EmailTableAdapters.RssSubscriptionsTableAdapter = New EmailTableAdapters.RssSubscriptionsTableAdapter()
            'Dim dt As System.Data.DataTable = emailAlerterData.GetDataById(1)
            'Dim rssId As String = dt.Rows(0)("rssid").ToString()
            'Dim url As String = dt.Rows(0)("url").ToString()

            Try
                If Subject.Length >= 121 Then
                    Dim Sub1 As String = Subject.Substring(0, 120)
                    Dim Index As Integer = Sub1.LastIndexOf("/")
                    Subject = Sub1.Substring(0, Index)
                    'Subject.
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName
            Dim QuerySubscriber As New QuerySubscriber
            QuerySubscriber.Favorite = Favorite

            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0
            Try
                Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            If Count = 0 And Favorite = "0" Then
                Try
                    Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
                Catch ex As Exception
                    Common.LogText("retry to create f-0 failed" & ex.ToString)

                End Try
            End If
            If Count = 0 And Favorite = "0" Then
                Try
                    Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613760312261" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
                    Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
                    request.GetResponse()
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            End If


            If Count > 0 Then
                Dim ContactLists() As String = New String() {FavoriteContactsList}




                Dim intervalTime As Integer = -1
                Dim campaignId As Integer
                Dim myCampaign As SpreadService.Campaign = New SpreadService.Campaign()

                myCampaign.campaignName = CampaignName

                myCampaign.from = senderName
                myCampaign.fromEmail = SenderEmail
                myCampaign.subject = Subject
                myCampaign.content = EmailBody
                Dim de As Date = Now
                myCampaign.schedule = de
                'send campaign
                Try
                    campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
                Common.LogText(Subject)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Else
                Common.LogText("No contact in this favorite:" & Favorite)
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
        Return SentResult
    End Function

    Function CustomizeTempalateDeal(ByVal LoginEmail As String, ByVal AppID As String, ByVal category As String, ByVal Template As String, ByVal subTemplate1 As String, ByVal subTemplate2 As String, ByVal subTemplate3 As String, ByVal EmailTemplateElements As EmailCampaignElementDeal(), ByVal PubDate As String) As Campaign
        Dim A As New alerter
        'begin sort,the favorite deals will up to top
        Dim subject As String = ""
        Dim Comparer As New MyComparer
        Comparer.Favorite = category
        Try
            Array.Sort(EmailTemplateElements, Comparer)
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try

        'insert popular deals to position 2, 3
        Try
            EmailTemplateElements = InsertPopularDeals(LoginEmail, AppID, EmailTemplateElements)
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
        'generate the subject
        For Each EmailTemplateElement As EmailCampaignElementDeal In EmailTemplateElements

            Dim DealSubject As String = EmailTemplateElement.Subject
            Dim s As String = EmailTemplateElement.PubDate.Replace("HKT", "")
            Dim DealPubDate As Date = Date.Parse(s)
            If DealPubDate.Date = Now.Date Then
                If DealSubject.IndexOf("【") > 0 Then
                    Try
                        DealSubject = System.Web.HttpUtility.HtmlDecode(DealSubject.Substring(DealSubject.IndexOf("【"), DealSubject.IndexOf("】") - DealSubject.IndexOf("【") + 1))
                    Catch ex As Exception
                        Common.LogText(ex.ToString)
                    End Try
                    subject = subject & "/" & DealSubject
                End If
            End If
        Next
        'remove the first "/"
        If Not String.IsNullOrEmpty(subject) Then
            subject = subject.Remove(0, 1)
        End If

        Try
            'update publish date in the content
            Template = System.Text.RegularExpressions.Regex.Replace(Template, "<td style=""text-align: right""><font color=""#ffffff"" size=""2""[\s\S]*?serif"">2012年01月17日[\s\S]*?</font></td>", "<td style=""text-align: right""><font color=""#ffffff"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & PubDate & "</font></td>")
            'update send template from xml data
            ' For i As Integer = 0 To EmailTemplateElements.Length - 1


            Dim MainDealsHtml As String = ""
            Dim OtherDealsHtml As String = ""
            'sign for other deal position
            Dim sign As Integer = 0
            For i As Integer = 0 To EmailTemplateElements.Length - 1
                Try


                    ' Dim DealPubDate As Date = EmailTemplateElements(i).PubDate.Substring(5, 11)
                    Dim s As String = EmailTemplateElements(i).PubDate.Replace("HKT", "")
                    Dim DealPubDate As Date = Date.Parse(s)
                    If DealPubDate.Date = Now.Date Then
                        'put today's deal into html
                        Dim MDealHtml As String = String.Format(subTemplate1, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                        MainDealsHtml = MainDealsHtml & MDealHtml
                    ElseIf DealPubDate < Now.Date Then
                        'not today's deal and put is into certain position
                        If sign = 0 Then
                            Dim ODealHtml As String = String.Format(subTemplate2, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                            OtherDealsHtml = OtherDealsHtml & ODealHtml
                            sign = 1
                        ElseIf sign = 1 Then

                            Dim ODealHtml As String = String.Format(subTemplate3, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                            OtherDealsHtml = OtherDealsHtml & ODealHtml
                            sign = 0
                        End If
                    End If
                Catch ex As Exception
                    System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
                    Common.LogText(ex.ToString)
                End Try

            Next
            Try
                If String.IsNullOrEmpty(OtherDealsHtml) Then
                    OtherDealsHtml = "<td></td>"
                Else
                    Template = System.Text.RegularExpressions.Regex.Replace(Template, " <tr bgcolor=""#ffffff"">", "<tr bgcolor=""#1d4474""><td>&nbsp;</td><td style=""text-align: center"" height=""35""><strong><font color=""#ffffff"" size=""4"" face=""Arial,Helvetica,sans-serif"">其他團購 </font></strong></td><td>&nbsp;</td></tr><tr bgcolor=""#ffffff"">")
                End If
            Catch ex As Exception
                System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
                Common.LogText(ex.ToString)
            End Try

            Dim html As String = String.Format(Template, MainDealsHtml, OtherDealsHtml)

            Dim campaign As New Campaign
            campaign.Subject = subject
            campaign.Content = html

            Return campaign
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    Function UpdateCategories(ByVal OldCategories As String, ByVal newCategories As String)
        Dim OldC As String() = OldCategories.Split(",")
        Dim NewC As String() = newCategories.Split(New Char() {"-"}, System.StringSplitOptions.RemoveEmptyEntries)
        Dim OldL As New List(Of String)(OldC)
        Dim NewL As New List(Of String)(NewC)
        For i As Integer = 0 To NewL.Count - 1
            If Not OldL.Contains(NewL(i).ToString) Then
                OldL.Add(NewL(i).ToString)
            End If
        Next
        Dim UpdatedCategories As String = ""
        For j As Integer = 0 To OldL.Count - 1
            UpdatedCategories = UpdatedCategories & OldL(j).ToString
        Next


        'For i As Integer = 0 To NewC.Length - 1
        '    Dim c As Integer = Array.IndexOf(OldC, NewC(i).ToString)
        'Next
    End Function


    'update EmailCampaignElements and put the popular deals into today deals
    Function InsertPopularDeals(ByVal LoginEmail As String, ByVal AppID As String, ByVal EmailTemplateElements As EmailCampaignElementDeal()) As EmailCampaignElementDeal()
        Dim A As New alerter
        Dim N As Integer = ConfigurationManager.AppSettings("NumOfPopularDeals")
        Dim DealLists As List(Of EmailCampaignElementDeal) = New List(Of EmailCampaignElementDeal)(EmailTemplateElements)

        Dim PopularDeals As New List(Of EmailCampaignElementDeal)


        Dim CampaignName As String = Format(Now.Date.AddDays(-1), "dd/MM/yyyy")
        Dim spreadiws As New EmailAlerter.IntrawWebService.Service
        Dim ds As DataSet = spreadiws.dsGetURLTrackings(LoginEmail, AppID, CampaignName)
        '24/12/2012
        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To N - 1
                Try
                    Dim URL As String = dt.Rows(i).Item(0).ToString

                    For j As Integer = DealLists.Count - 1 To 0 Step -1
                        Dim DealLink As String = Regex.Match(DealLists(j).Link.ToString, "http[^""]*").ToString
                        If String.Compare(URL, DealLink, True) = 0 Then
                            Dim Deal As EmailCampaignElementDeal = New EmailCampaignElementDeal
                            Deal.DealCategory = DealLists(j).DealCategory
                            Deal.Subject = DealLists(j).Subject
                            Deal.Link = DealLists(j).Link
                            Deal.Price = DealLists(j).Price
                            Deal.Discount = DealLists(j).Discount
                            Deal.DiscountPrice = DealLists(j).DiscountPrice
                            Deal.Saved = DealLists(j).Saved
                            Deal.LargeBuyButton = DealLists(j).LargeBuyButton
                            Deal.BuyButton = DealLists(j).BuyButton
                            Deal.ProductImage = DealLists(j).ProductImage
                            Deal.PubDate = Format(Now, "yyyy-MM-dd 00:00:00") ' DealLists(j).PubDate
                            Deal.Index = j
                            PopularDeals.Add(Deal)
                            DealLists.Remove(DealLists(j))
                        End If

                    Next

                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next

            Try
                For i As Integer = 0 To PopularDeals.Count - 1
                    DealLists.Insert(i + 1, PopularDeals(i))
                Next
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

        End If

        Return DealLists.ToArray()
    End Function
End Class
