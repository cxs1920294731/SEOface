Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Configuration
Imports EDM4Everbuying
Imports HtmlAgilityPack
Imports System.Linq

Public Class GroupBuyer2


    Public Structure Campaign
        Dim Subject As String
        Dim Content As String
    End Structure

    ''' <summary>
    ''' 单个产品信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure EmailCampaignElementDeal
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
        Dim ExpiredDate As String
        Dim DealCategory As String
        Dim Index As Integer '产品的索引值，从0开始

        '添加时间：2013/05/07
        '描述：Mobile模板增加一个Content内容
        Dim Content As String

        '添加时间：2014/02/08
        '描述：增加新需求产品购买数量的排序
        Dim noofPurchased As Integer
    End Structure

    'Added in 2014-07-19,增加一个Banner
    Public Structure Banner
        Dim ahrefURL As String 'banner 的跳转链接
        Dim imgSRC As String 'banner的图片链接
        Dim linkCategory As String
    End Structure

    ''' <summary>
    ''' Rss的产品信息和频道信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure ElementAndTitle
        Dim ChannelTitle As String
        Dim PubDate As DateTime
        Dim EmailCampaignElements() As EmailCampaignElementDeal
        Dim AllCategories As String
        Dim Banner() As Banner
    End Structure

    Sub SendEmailDeal(ByVal r As RssSubscription)
        Dim A As New alerter

        Dim siteId As Integer = 59 '真实环境：59；本地测试：60（真实环境siteId=59，测试环境siteId=60）

        Try
            '  r.URL = "C:\Users\Suwen\Desktop\fff.xml"
            Dim ElementAndTitle = ReadRSSDeal(r.Url, siteId)

            Dim PubDateTime As DateTime = ElementAndTitle.pubdate
            Dim pubdate As String = PubDateTime.ToString("yyyy年MM月dd日")
            Dim subject As String = ""

            '2013/06/21，不发送desktop版本，begin
            Dim Templates As String() = r.Template.Split("^")
            Dim FTemplate As String = ""
            Dim s As String = ""
            Dim NLeftTemplate As String = ""
            Dim NRightTemplate As String = ""
            Dim OLeftTemplate As String = ""
            Dim OMiddleTemplate As String = ""
            Dim ORightTemplate As String = ""

            '修改时间：2013/05/14
            '修改人：彭震
            '修改原因：兼容手机版模板
            If Templates.Length >= 6 Then '电脑版模板分成6个模块处理
                FTemplate = Templates(0)
                s = FTemplate.Replace(vbCrLf, "")
                NLeftTemplate = Templates(1)
                NRightTemplate = Templates(2)
                OLeftTemplate = Templates(3)
                OMiddleTemplate = Templates(4)
                ORightTemplate = Templates(5)
            Else '手机版模板分成4个模块处理
                FTemplate = Templates(0)
                s = FTemplate.Replace(vbCrLf, "")
                NLeftTemplate = Templates(1)
                NRightTemplate = Templates(2)
                OLeftTemplate = Templates(3)
            End If
            Dim categories As String() = r.Categories.Split(",") 'Categories：0,1,2,3,4,5,6,7,8,9,10,13,14,16,17,18,19
            Dim SentCat As String = DateTime.Now.ToString("yyyy-MM-dd") & " Groupbuyer sent favorites:"
            Dim listNames As String = ""
            For Each category As String In categories
                Try
                    'Campaign元素：campaignName,from,fromEmail,subject,content,schedule
                    'r.ExcludeSubject,这个字段的内容其实是Google Analytics code,借用此字段放置追踪代码
                    Dim campaign As Campaign = CustomizeTempalateDeal(r.SpreadLogin, r.AppID, category, FTemplate, NLeftTemplate, NRightTemplate, OLeftTemplate, OMiddleTemplate, ORightTemplate, ElementAndTitle, pubdate, r.ExcludeSubject)
                    If Not String.IsNullOrEmpty(campaign.Subject) Then
                        subject = campaign.Subject
                    Else
                        subject = ElementAndTitle.channeltitle
                    End If
                    'subject = "Hi [FIRSTNAME]," & subject '2014/03/24 added,

                    '2014/02/08 added,增加新需求：请将Groupbuyer的CompaignName里添加明确的分类标识，如CampaignName: 07/02/2014(favorite:5) -> 07/02/2014(favorite:food-5)
                    'Dim CampaignName As String = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & category & ")"
                    Dim CampaignName As String = ""
                    Select Case category
                        Case "0"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "default-" & category & ")"
                        Case "1"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "wine-" & category & ")"
                        Case "2"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & " beauty-" & category & ")"
                        Case "3"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & " movie-" & category & ")"
                        Case "4"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & " deposit-" & category & ")"
                        Case "5"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & " food-" & category & ")"
                        Case "6"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & " product-" & category & ")"
                        Case "7"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "course/family-" & category & ")"
                        Case "8"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "travel-" & category & ")"
                        Case "9"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "baby-" & category & ")"
                        Case "10"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "computer-" & category & ")"
                        Case "13"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "yuenlong-" & category & ")"
                        Case "14"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "hitech-" & category & ")"
                        Case "16"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "fashion-" & category & ")"
                        Case "17"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "entertainment-" & category & ")"
                        Case "18"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "special-" & category & ")"
                        Case "19"
                            CampaignName = Format(Now.Date, "dd/MM/yyyy").ToString & "(favorite:" & "Beauty product-" & category & ")"
                    End Select

                    '(Date.Now.Day & "/" & Date.Now.Month & "/" & Date.Now.Year).ToString & "(favorite:" & category & ")"

                    Try
                        Dim sentResult As Boolean =CreateCampaignDeal(category, CampaignName, subject, r.SpreadLogin, r.AppID, r.SenderName, r.SenderEmail, campaign.Content, listNames, siteId)
                        If (sentResult = True) Then
                            SentCat = SentCat & category & ","
                            '2014/04/25 added，請嘗試把【xxx】改成 | | 分隔subject，如：产品1|产品2，begin
                            If (subject.Contains("【") OrElse subject.Contains("】")) Then
                                subject = subject.Replace("【", "").Replace("】", " | ")
                                subject = subject.Substring(0, subject.LastIndexOf("|"))
                            End If
                            'end
                            A.UpdateSentLog(r.RssID, subject)
                        End If
                    Catch ex As Exception
                        '2013/09/29 added
                        Try
                            Dim exStr As String = SentCat & "sent failed"
                            Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613424160512" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & exStr & "&type=1&returnMode=1&sentDirect=1"
                            Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
                            request.GetResponse()
                        Catch ex1 As Exception
                            Common.LogText(ex1.ToString())
                        End Try
                        '---------------------------------
                    End Try
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next
            SentCat = SentCat & "end."
            Try

                Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613424160512" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & SentCat & "&type=1&returnMode=1&sentDirect=1"
                Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
                request.GetResponse()
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            '2013/06/21，不发送desktop版本，end

        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' 根据网站的Rss URL获取网站的产品信息和所有的CategoryId、channel title、channel pubDate等产品信息
    ''' </summary>
    ''' <param name="URL">Rss URL</param>
    ''' <returns>多个产品信息和频道信息</returns>
    ''' <remarks></remarks>
    Public Function ReadRSSDeal(ByVal URL As String, ByVal siteId As Integer) As ElementAndTitle
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
                    '获取channel title，并为channel title加上![CDATA，
                    '防止获取到的channel title没加上![CDATA,使用.net方法读取XML文件时出错
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
                    Dim titleNodeText As String
                    Dim title As String = node.SelectSingleNode("title").InnerText.Trim()

                    '2013/05/14新增
                    Dim content As String = node.SelectSingleNode("title").InnerText.Replace("<![CDATA[", "").Replace("]]>", "")

                    If title.IndexOf("【") <> -1 And title.IndexOf("】") <> -1 Then
                        titleNodeText = title.Substring(title.IndexOf("【"), title.IndexOf("】") - title.IndexOf("【") + 1)
                    Else
                        If (title.Length >= 50) Then
                            titleNodeText = title.Substring(0, 49) & "..."
                        Else
                            titleNodeText = title
                        End If
                    End If
                    Dim link_category As String = node.SelectSingleNode("category").InnerText.Trim()
                    Dim priceText As String = node.SelectSingleNode("description").SelectSingleNode("price").InnerText.Trim()
                    Dim discountText As String = node.SelectSingleNode("description").SelectSingleNode("discount").InnerText.Trim()
                    Dim savedText As String = node.SelectSingleNode("description").SelectSingleNode("saved").InnerText.Trim()
                    Dim discountPriceText As String = node.SelectSingleNode("description").SelectSingleNode("discountPrice").InnerText.Trim()
                    Dim productImageText As String = node.SelectSingleNode("description").SelectSingleNode("productImage").InnerText.Trim()
                    Dim largeBuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("largeBuyButton").InnerText.Trim()
                    Dim BuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("BuyButton").InnerText.Trim()
                    Dim postDateNodeText As String = node.SelectSingleNode("pubDate").InnerText.Trim()
                    Dim endDateNodeText As String = node.SelectSingleNode("endDate").InnerText.Trim()

                    '2014/02/08 added,增加新需求产品购买数量的排序
                    Dim iNoofPurchased As Integer = Int32.Parse(node.SelectSingleNode("noofPurchased").InnerText.Trim())

                    Dim linkNodeText As String = node.SelectSingleNode("link").InnerText
                    linkNodeText = linkNodeText & """ link_category=""" & link_category

                    'Start by Gary, 2014/01/28 
                    '处理RSS description里没有价格时，获取外层的价格'
                    Try
                        If Not IsNumeric(priceText) Then
                            priceText = node.SelectSingleNode("orginalPrice").InnerText.Trim
                        End If
                        If Not IsNumeric(discountPriceText) Then
                            discountPriceText = node.SelectSingleNode("actualPrice").InnerText.Trim
                        End If
                    Catch ex As Exception

                    End Try
                    'End by Gary

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
                    EmailCampaignElements(i).noofPurchased = iNoofPurchased

                    '2013/05/14新增，Mobile版模板增加Content
                    EmailCampaignElements(i).Content = content

                    EmailCampaignElements(i).PubDate = postDateNodeText
                    EmailCampaignElements(i).ExpiredDate = endDateNodeText
                    EmailCampaignElements(i).Index = i
                    i += 1
                    AllCategories = AllCategories & link_category
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next

            '20140719,get banner from rssFile
            Dim bannerNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes("//banner")
            Dim mybanner(bannerNodeList.Count - 1) As Banner
            Dim ibanner As Integer = 0
            For Each node As Xml.XmlNode In bannerNodeList
                Try
                    mybanner(ibanner).ahrefURL = node.SelectSingleNode("link").InnerText.Trim
                    mybanner(ibanner).imgSRC = node.SelectSingleNode("bannerImage").InnerText.Trim
                    mybanner(ibanner).imgSRC = mybanner(ibanner).imgSRC.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com")
                    mybanner(ibanner).linkCategory = node.SelectSingleNode("category").InnerText.Trim
                    ibanner += 1
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next
            '20140719,get banner from rssFile

            '2014/03/20,Insert Data to Datebase for notification,begin
            Try
                UpdateGroupBuyerData.Start(siteId, EmailCampaignElements)
            Catch ex As Exception
                'Ignore
            End Try
            '2014/03/20,Insert Data to Datebase for notification,end

            ElementAndTitle.EmailCampaignElements = EmailCampaignElements
            If String.IsNullOrEmpty(ChannelTitle) Then
                ElementAndTitle.ChannelTitle = "Group Buyer 今日團購資訊"
            Else
                ElementAndTitle.ChannelTitle = ChannelTitle
            End If

            ElementAndTitle.PubDate = PubDate
            ElementAndTitle.AllCategories = AllCategories
            ElementAndTitle.Banner = mybanner
            Return ElementAndTitle
        Catch ex As Exception
            Common.LogText(ex.ToString)
        End Try
    End Function

    Function CreateCampaignDeal(ByVal Favorite As String, ByVal CampaignName As String, ByVal Subject As String, _
                                ByVal LoginName As String, ByVal AppID As String, ByVal senderName As String, _
                                ByVal SenderEmail As String, ByVal EmailBody As String, ByRef listNames As String, _
                                ByVal siteId As Integer) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            Try
                '修改时间：2013/05/15修改
                '修改人：彭震 
                '修改原因：Subject长度>120,截取字符串会出错,如：...【12次24小時無間斷物理豐胸】/【西藏薰蒸足部
                If Subject.Length >= 121 Then
                    Dim Sub1 As String = Subject.Substring(0, 120)
                    Dim Index As Integer = Sub1.LastIndexOf("【")
                    Subject = Sub1.Substring(0, Index)
                End If

                '2014/04/25 added，請嘗試把【xxx】改成 | | 分隔subject，如：产品1|产品2，begin
                If (Subject.Contains("【") OrElse Subject.Contains("】")) Then
                    Subject = Subject.Replace("【", "").Replace("】", " | ")
                    Subject = Subject.Substring(0, Subject.LastIndexOf("|"))
                End If
                'end
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

            'Generate favorite contact list from spread service,return the count of contacts
            Dim mySpread As SpreadService.Service = New SpreadService.Service()
            mySpread.Timeout = 1200000
            Dim FavoriteContactsList As String = CampaignName

            Dim QuerySubscriber As New QuerySubscriber

            'QuerySubscriber.Strategy = ChooseStrategy.Favorite
            'add by Gary Start
            If (Favorite = "0") Then
                'Open but no click,发送最进4个月的有开启没点击的
                QuerySubscriber.Strategy = ChooseStrategy.OpenExcludeCategory
                QuerySubscriber.StartDate = Date.Now.AddMonths(-4).ToString("yyyy-MM-dd")
            Else
                QuerySubscriber.Strategy = ChooseStrategy.Favorite
                QuerySubscriber.Favorite = Favorite
            End If

            '2013/06/06注释,by Gary
            'QuerySubscriber.Favorite = Favorite
            'add by Gary End

            QuerySubscriber.CountryList = New String() {}
            Dim CriteriaString As String = QuerySubscriber.ToJsonString
            Dim Count As Integer = 0

            Try
                Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
                If Not (Favorite = "0") Then
                    listNames = listNames & FavoriteContactsList & ";"
                End If
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
            'End If
            '2013/06/06注释,by Gary End

            'Favourite="0",发送特定的list，list名为：Opened last 3-4 months (as of 17May13)，开始
            If Count = 0 And Favorite = "0" Then
                Try
                    Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
                Catch ex As Exception
                    Common.LogText("retry to create f-0 failed" & ex.ToString)
                End Try
            End If
            '2013/05/16测试注释,2013/06/06解注释By Gary
            If Count = 0 And Favorite = "0" Then
                Try
                    FavoriteContactsList = "Opens2014042516" '创建失败，则使用手动创建的List，"Opened last 3-4 months (as of 17May13)"
                    Count = 1
                    Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613424160512" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
                    Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
                    request.GetResponse()
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            End If
            If Count > 0 Then
                'Dim ContactLists() As String = New String() {FavoriteContactsList} '2013/05/28修改
                Dim ContactLists() As String
                If (Favorite = "0") Then
                    ContactLists = New String() {FavoriteContactsList, "Subscribers", "fb", "Members", "Puzzle", "newsletter", "InboxTest", "Group Dollar_20150107"}
                Else
                    ContactLists = New String() {FavoriteContactsList}
                End If

                Dim campaignId As Integer
                Try
                    If (Subject.Contains("【") OrElse Subject.Contains("】")) Then
                        Subject = Subject.Replace("【", "").Replace("】", " | ")
                        Subject = Subject.Substring(0, Subject.LastIndexOf("|"))
                    End If
                    '2014/01/16 added，favorite(0)设定为draft状态,begin
                    Dim campaignCreative(0) As EmailAlerter.SpreadService.CampaignCreatives
                    Dim c As EmailAlerter.SpreadService.CampaignCreatives = New EmailAlerter.SpreadService.CampaignCreatives
                    With c
                        .creativeContent = EmailBody
                        .displayName = senderName
                        .fromAddress = SenderEmail
                        .isCampaignDefault = True
                        .replyTo = ""
                        .subject = Subject
                        .target = "D"
                    End With
                    campaignCreative(0) = c
                    Dim interval As Integer = -1
                    If (Favorite = "0") Then
                        campaignId = mySpread.createCampaign2(LoginName, AppID, CampaignName, campaignCreative, ContactLists, interval, Now, "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
                        '因default0会发送一些固定名单，这些固定名单会包含有点击的，所以需要将其排除, comment by dora
                        listNames = listNames.Substring(0, listNames.LastIndexOf(";"))
                        mySpread.ExcludeContactList(LoginName, AppID, CampaignName, listNames)
                    Else
                        campaignId = mySpread.createCampaign2(LoginName, AppID, CampaignName, campaignCreative, ContactLists, interval, Now.AddHours(1.3), "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
                    End If
                    mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
                    '2014/01/16 added，favorite(0)设定为draft状态,end

                    '2014/04/04 added，添加一条数据到Issues表中,begin
                    Try
                        Dim efContext As New FaceBookForSEOEntities
                        Dim myIssue As New AutoIssue
                        myIssue.IssueDate = Now
                        myIssue.Subject = Subject
                        myIssue.SiteID = siteId
                        myIssue.PlanType = "HO"
                        myIssue.SentStatus = "ES"
                        myIssue.SpreadCampId = campaignId
                        efContext.AddToAutoIssues(myIssue)
                        efContext.SaveChanges()
                    Catch ex As Exception
                        'Ignore
                    End Try
                    '2014/04/04 added,添加一条数据到Issues表中，end


                    'campaignId = mySpread.createCampaign(LoginName, AppID, myCampaign, ContactLists, intervalTime)
                    'add by gary start 20130802
                    Try
                        Dim subject2 = senderName & " 开始发送邮件"
                        Dim myList As String = ""
                        For Each list As String In ContactLists
                            myList = myList & list & ","
                        Next

                        Dim SpreadTemplate2 As String
                        If (Favorite = 0) Then
                            SpreadTemplate2 = subject2 & "<br />邮件状态：正在发送<br />" & " 标题:" & Subject _
                                              & "<br />发送对象：" & myList & "<br />发件人Email：" & SenderEmail & "<br />发件人：" & senderName _
                                              & "<br />Campaign名：" & CampaignName & "<br />是否Publish to Newsletter Archive：" & "是" _
                                              & "<br />内容(*部分Spread系统按钮在提醒邮件里会失效，但不影响正常的发送)：<br />" & EmailBody
                        Else
                            SpreadTemplate2 = subject2 & "<br />邮件状态：待发<br />" & "预定发出时间:" & Now.AddHours(1.3).ToString("yyyy-MM-dd HH:mm:ss") & "<br />" & " 标题:" & Subject _
                                                  & "<br />发送对象：" & myList & "<br />发件人Email：" & SenderEmail & "<br />发件人：" & senderName _
                                                  & "<br />Campaign名：" & CampaignName & "<br />是否Publish to Newsletter Archive：" & "是" _
                                                  & "<br />内容(*部分Spread系统按钮在提醒邮件里会失效，但不影响正常的发送)：<br />" & EmailBody
                        End If
                       
                        NotificationEmail.SentStartEmail(subject2, SpreadTemplate2)
                    Catch ex As Exception
                        Common.LogText(ex.ToString)
                        NotificationEmail.SentErrorEmail(Favorite & "发送失败")
                    End Try
                    'add by gary end 20130802
                Catch ex As Exception
                    Common.LogText(ex.ToString)
                    Throw ex
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
            Throw ex
        End Try
        Return SentResult
    End Function

    ''' <summary>
    ''' 兼容手机版本和电脑版本的创建Campaign
    ''' </summary>
    ''' <param name="Favorite"></param>
    ''' <param name="CampaignName"></param>
    ''' <param name="Subject"></param>
    ''' <param name="LoginName"></param>
    ''' <param name="AppID"></param>
    ''' <param name="senderName"></param>
    ''' <param name="SenderEmail"></param>
    ''' <param name="EmailBody"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CreateCampaignDeal(ByVal Favorite As String, ByVal CampaignName As String, ByVal Subject As String, ByVal LoginName As String, ByVal AppID As String, ByVal senderName As String, ByVal SenderEmail As String, ByVal EmailBody() As String) As Boolean
        Dim A As New alerter
        Dim SentResult As Boolean = False
        Try
            If Subject.Length >= 121 Then 'Subject的字符数>120，就截取最后一个【...】
                Dim Sub1 As String = Subject.Substring(0, 120)
                Dim Index As Integer = Sub1.LastIndexOf("【")
                Subject = Sub1.Substring(0, Index)
            End If
        Catch ex As Exception
            Common.LogText(ex.ToString())
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

        'modify by gary start 2013/06/22
        If (Favorite = "0") Then
            'Open but no click,发送最进4个月的有开启没点击的
            QuerySubscriber.Strategy = ChooseStrategy.OpenExcludeCategory
            QuerySubscriber.StartDate = Date.Now.AddMonths(-4).ToString("yyyy-MM-dd")
        Else
            QuerySubscriber.Strategy = ChooseStrategy.Favorite
            QuerySubscriber.Favorite = Favorite
        End If

        CriteriaString = QuerySubscriber.ToJsonString

        Try
            Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
        Catch ex As Exception
            If ex.ToString().Contains("Contact List Name already exist") Then
                Count = 1
            End If
            Common.LogText(ex.ToString())
            NotificationEmail.SentErrorEmail(Favorite & "SearchContacts error:" & ex.ToString())
            Throw New Exception
        End Try

        'If (Favorite = "0") Then
        '    FavoriteContactsList = "Opened last 3-4 months (as of 17May13)"
        '    Count = 1
        'Else
        '    Try
        '        Count = mySpread.SearchContacts(LoginName, AppID, CriteriaString, Integer.MaxValue, FavoriteContactsList, True)
        '    Catch ex As Exception
        '        A.LogText(ex.ToString())
        '    End Try
        'End If
        If Count = 0 And Favorite = "0" Then
            Try
                FavoriteContactsList = "Opened last 3-4 months (as of 17May13)" '创建失败，则使用手动创建的List
                Count = 1
                Dim url As String = "http://mdtechcorp.com:20000/openapi/?destinatingAddress=" & "8613424160512" & "&username=spread&password=sms190854&originatingAddress=spread&SMS=" & "favorite 0 sent failed" & "&type=1&returnMode=1&sentDirect=1"
                Dim request As System.Net.HttpWebRequest = Net.WebRequest.Create(url)
                request.GetResponse()
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try
        End If
        'modify by gary end 2013/06/22


        If Count > 0 Then
            'modify by gary start 2013/06/22
            'Dim ContactLists() As String = New String() {FavoriteContactsList} '联系人list
            Dim ContactLists() As String
            If (Favorite = "0") Then
                Dim Favorite0ContactsList As String = "Subcribers"
                ContactLists = New String() {FavoriteContactsList, Favorite0ContactsList}
            Else
                ContactLists = New String() {FavoriteContactsList}
            End If
            'modify by gary end 2013/06/22

            Dim interval As Integer = -1

            '2013/05/10添加gb code
            'EmailBody(0) = AddSpecialCode(EmailBody(0), CampaignName) '电脑版本，添加追踪代码
            'EmailBody(1) = AddSpecialCode(EmailBody(1), CampaignName) '手机版本，添加追踪代码

            Dim campaigCreative(1) As SpreadService.CampaignCreatives
            Dim c1 As SpreadService.CampaignCreatives = New SpreadService.CampaignCreatives()
            With c1
                .creativeContent = EmailBody(0)
                .displayName = "GroupBuyer"
                .fromAddress = LoginName
                .isCampaignDefault = True
                .replyTo = ""
                .subject = Subject
                .target = "D"
            End With
            Dim c2 As SpreadService.CampaignCreatives = New SpreadService.CampaignCreatives()
            With c2
                .creativeContent = EmailBody(1)
                .displayName = "GroupBuyer"
                .fromAddress = LoginName
                .isCampaignDefault = True
                .replyTo = ""
                .subject = Subject
                .target = "M"
            End With
            campaigCreative(0) = c1
            campaigCreative(1) = c2
            Try
                Dim campaignId As Integer = mySpread.createCampaign2(LoginName, AppID, CampaignName, campaigCreative, ContactLists, interval, Now, "Spread", EmailAlerter.SpreadService.CampaignStatus.Waiting)
                SentResult = True
                mySpread.ChangePublishStatus(LoginName, AppID, campaignId, True)
            Catch ex As Exception
                Common.LogText(ex.ToString())
                NotificationEmail.SentErrorEmail(Favorite & "createCampaign2 error:")
                Throw New Exception '2013/09/29 added
            End Try
            'Dim MySpread As New SpreadService.SpreadWebService()
            'MySpread.createCampaign2(

        Else
            Common.LogText("No contact in this favorite:" & Favorite)
        End If
        Return SentResult
    End Function

    ''' <summary>
    ''' 填充Desktop模板和获取Subject，并为Campaign对象属性Content、Subject赋值；
    ''' 模板被分成6部分；
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="category"></param>
    ''' <param name="FTemplate"></param>
    ''' <param name="NLeftTemplate"></param>
    ''' <param name="NRightTemplate"></param>
    ''' <param name="OLeftTemplate"></param>
    ''' <param name="OMiddleTemplate"></param>
    ''' <param name="ORightTemplate"></param>
    ''' <param name="EmailTemplateElements"></param>
    ''' <param name="PubDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CustomizeTempalateDeal(ByVal LoginEmail As String, ByVal AppID As String, ByVal category As String, _
                                    ByVal FTemplate As String, ByVal NLeftTemplate As String, ByVal NRightTemplate As String, _
                                    ByVal OLeftTemplate As String, ByVal OMiddleTemplate As String, ByVal ORightTemplate As String, _
                                    ByVal ElementAndTitle As ElementAndTitle, ByVal PubDate As String, ByVal urlSpecialCode As String) As Campaign
        Dim A As New alerter
        Dim EmailTemplateElements() As EmailCampaignElementDeal = ElementAndTitle.EmailCampaignElements
        'begin sort,the favorite deals will up to top
        Dim subject As String = ""
        Dim Comparer As New MyComparer
        Comparer.Favorite = category
        Try
            Array.Sort(EmailTemplateElements, Comparer)
        Catch ex As Exception
            'A.LogText(ex.ToString) '2013/05/15注释，不需要写Log
        End Try


        ''2014/02/08 added,增加新需求产品购买数量的排序,begin
        ''增加新需求1：将当天的Foods分类（分类Num: 5）的Best selling的一个Deals放置在所有的分类邮件里的第一位，Best Selling的判断根据RSS里的<noofPurchased>标签判断
        ''增加新需求2：以前根据Spread点击获取最受欢迎的某一个Deal，改为根据<noofPurchased>获取
        'Try
        'EmailTemplateElements = InsertPopularDeals(LoginEmail, AppID, EmailTemplateElements)
        'Catch ex As Exception
        '    A.LogText(ex.ToString)
        'End Try
        Try
            EmailTemplateElements = InsertSpecialCateAndSort.InsertSpecialCateAndSort(LoginEmail, AppID, EmailTemplateElements, 5)  '主推分类5：Food
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
        ''2014/02/08 added,增加新需求产品购买数量的排序,end

        '2014/03/20 added,add top6 to a new block,begin
        Dim elementsCounter As Integer = EmailTemplateElements.Count
        Dim listElements As List(Of EmailCampaignElementDeal) = EmailTemplateElements.Cast(Of EmailCampaignElementDeal).ToList()
        Dim Top6TemplateElements() As EmailCampaignElementDeal
        Top6TemplateElements = listElements.OrderByDescending(Function(e) e.noofPurchased).Skip(1).Take(6).ToList().ToArray()
        Dim listOtherElements As List(Of EmailCampaignElementDeal) = listElements.OrderByDescending(Function(e) e.noofPurchased).Skip(7).Take(elementsCounter - 7).ToList()
        listOtherElements.Insert(0, listElements(0))
        listOtherElements = InsertSpecialCateAndSort.InsertCurrentCateProduct(category, listOtherElements)
        listOtherElements = listOtherElements.Distinct().ToList()
        EmailTemplateElements = listOtherElements.ToArray()
        '2014/03/20 added,add top6 to a new block,end

        'generate the subject
        Dim NumNew As Integer = 0
        Dim NumOld As Integer = 0
        For Each EmailTemplateElement As EmailCampaignElementDeal In EmailTemplateElements

            Dim DealSubject As String = EmailTemplateElement.Subject
            Dim s As String = EmailTemplateElement.PubDate.Replace("HKT", "")
            Dim DealPubDate As Date = Date.Parse(s).AddDays(1)
            If DealPubDate.Date = Now.Date Then
                If DealSubject.IndexOf("【") > -1 Then
                    Try
                        DealSubject = System.Web.HttpUtility.HtmlDecode(DealSubject.Substring(DealSubject.IndexOf("【"), DealSubject.IndexOf("】") - DealSubject.IndexOf("【") + 1))
                        If (DealSubject.Length > 80) Then
                            DealSubject = DealSubject.Substring(0, 79)
                            If Not (DealSubject.Contains("【")) Then
                                DealSubject = "【" & DealSubject
                            End If
                            If Not (DealSubject.Contains("】")) Then
                                DealSubject = DealSubject & "】"
                            End If
                        End If
                    Catch ex As Exception
                        Common.LogText(ex.ToString)
                    End Try
                    subject = subject & DealSubject
                End If
            End If
            Dim d1 As String = Format(DateTime.Parse(EmailTemplateElement.PubDate).AddDays(1), "yyyy-MM-dd") 'Dim d1 As String = Format(DateTime.Parse(EmailTemplateElement.PubDate), "yyyy-MM-dd")
            Dim d2 As String = Format(Now, "yyyy-MM-dd")
            If d1 = d2 Then
                NumNew = NumNew + 1
            Else
                NumOld = NumOld + 1
            End If
        Next

        'remove the first "/"
        'If Not String.IsNullOrEmpty(subject) Then
        '    subject = subject.Remove(0, 1)
        'End If
        'subject = subject.Replace("/", "●")

        '填充模板
        Try
            Dim MainDealsHtml As String = ""
            Dim OtherDealsHtml As String = ""
            Dim sign As Integer = 0
            Dim sign2 As Integer = 0
            For i As Integer = 0 To EmailTemplateElements.Length - 1
                '  For i As Integer = 0 To 8
                Try
                    Dim s As String = EmailTemplateElements(i).PubDate.Replace("HKT", "")
                    Dim DealPubDate As Date = Date.Parse(s).AddDays(1)  '2014-05-27 added，今晚刊出的DEAL會改成10:00出街，Dim DealPubDate As Date = Date.Parse(s)

                    '2014/02/07 added,增加需求：分享该产品到facebook
                    Dim link As String = EmailTemplateElements(i).Link.Substring(0, EmailTemplateElements(i).Link.IndexOf(""""))
                    Dim shareOnFacebook As String = "https://www.facebook.com/sharer/sharer.php?u=" & System.Web.HttpUtility.UrlEncode(link)

                    If DealPubDate.Date = Now.Date Then
                        'put today's deal into html
                        Dim MDealHtml As String = ""  ' String.Format(subTemplate1, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                        If sign = 0 Then
                            '2014/02/07 added,增加需求：分享该产品到facebook;2014/03/17 cut掉这个功能;
                            Dim NLdeal As String = String.Format(NLeftTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link)
                            'Dim NLdeal As String = String.Format(NLeftTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                            '2013/07/13 added,add GroupBuyer mobile friendly Template
                            'NLdeal = "<tr><td width=""700"" height=""16""></td></tr>" & NLdeal
                            'Dim deal As String = "<tr><td><table width=""700"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td width=""25""></td><td width=""315"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top"">" & Today & " </td><td width=""16""></td><td width=""315"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top"">"
                            'NLdeal为两个产品之间的空格table
                            'NLdeal = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>" & NLdeal
                            MainDealsHtml = MainDealsHtml & NLdeal

                            sign = 1
                        ElseIf sign = 1 Then
                            '2014/02/07 added,增加需求：分享该产品到facebook；2014/03/17 cut掉这个功能；
                            Dim NRdeal As String = String.Format(NRightTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link)
                            'Dim NRdeal As String = String.Format(NRightTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                            MainDealsHtml = MainDealsHtml & NRdeal
                            sign = 0
                        End If

                    Else  'ElseIf DealPubDate < Now.Date Then '2014/05/28 revised

                        'not today's deal and put it into certain position
                        If sign2 = 0 Then
                            '2014/02/07 added,增加需求：分享该产品到facebook；2014/03/17 cut掉这个功能；
                            Dim OLdeal1 As String = String.Format(OLeftTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link)
                            'Dim OLdeal1 As String = String.Format(OLeftTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                            '2013/07/13 added,add GroupBuyer mobile friendly Template
                            'OLdeal1 = "<tr><td width=""700"" height=""16""></td></tr>" & OLdeal1
                            'OLdeal1为两个产品之间的空格table
                            'OLdeal1 = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>" & OLdeal1
                            OtherDealsHtml = OtherDealsHtml & OLdeal1
                            sign2 = sign2 + 1
                        ElseIf sign2 = 1 Then
                            '2014/02/07 added,增加需求：分享该产品到facebook；2014/03/17 cut掉这个功能；
                            Dim OMdeal1 As String = String.Format(OMiddleTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link)
                            'Dim OMdeal1 As String = String.Format(OMiddleTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                            OtherDealsHtml = OtherDealsHtml & OMdeal1
                            sign2 = sign2 + 1
                        ElseIf sign2 = 2 Then
                            '2014/02/07 added,增加需求：分享该产品到facebook；2014/03/17 cut掉这个功能；
                            Dim ORdeal1 As String = String.Format(ORightTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link)
                            'Dim ORdeal1 As String = String.Format(ORightTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                            'Dim ODealHtml As String = String.Format(subTemplate3, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                            OtherDealsHtml = OtherDealsHtml & ORdeal1
                            sign2 = 0
                        End If
                    End If


                Catch ex As Exception
                    System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())

                End Try

            Next
            If NumNew Mod 2 = 1 Then
                'MainDealsHtml = MainDealsHtml & "<td width=""16""></td><td width=""315"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""></td><td width=""25""></td></tr></table></td></tr>"
                MainDealsHtml = MainDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" class=""mobile_hidden"" width=""13""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table>"
                MainDealsHtml = MainDealsHtml & "<table width=""315"" align=""left"" border=""0"" cellspacing=""0"" valign=""top"" class=""mobile_hidden""><tbody><tr><td><table align=""left"" width=""315"" height=""333"" border=""0"" cellspacing=""0"" cellpadding=""5"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><tbody><tr><td style=""width:317px;height:333px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""305"" height=""367px"" border=""0"" alt=""""></a></td></tr></tbody></table></td></tr></tbody></table>"
                MainDealsHtml = MainDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
                MainDealsHtml = MainDealsHtml & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>"
                MainDealsHtml = MainDealsHtml & "</td></tr></table></td></tr>"
            End If
            If NumOld Mod 3 = 1 Then
                'OtherDealsHtml = OtherDealsHtml & "<td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""16""></td><td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/special"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_2.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""25""></td></tr>"
                OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
                OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
                OtherDealsHtml = OtherDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
                OtherDealsHtml = OtherDealsHtml & "</td></tr>"
            ElseIf NumOld Mod 3 = 2 Then

                OtherDealsHtml = OtherDealsHtml & "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289"" border=""0"" alt="""" /></a></td></tr></table>"
                OtherDealsHtml = OtherDealsHtml & "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
                OtherDealsHtml = OtherDealsHtml & "</td></tr>"
            End If

            '2014/03/20 added,add top6 to a new block,begin
            Dim middleDealsHtml As String = ""
            Dim sign3 As Integer = 0
            For Each element As EmailCampaignElementDeal In Top6TemplateElements
                If (sign3 = 0) Then
                    Dim OLdeal1 As String = String.Format(OLeftTemplate, element.Link, element.ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), element.Subject, element.DiscountPrice, element.Price, element.Link)
                    middleDealsHtml = middleDealsHtml & OLdeal1
                    sign3 = sign3 + 1
                ElseIf (sign3 = 1) Then
                    Dim OMdeal1 As String = String.Format(OMiddleTemplate, element.Link, element.ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), element.Subject, element.DiscountPrice, element.Price, element.Link)
                    'Dim OMdeal1 As String = String.Format(OMiddleTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Price, EmailTemplateElements(i).Link, shareOnFacebook)

                    middleDealsHtml = middleDealsHtml & OMdeal1
                    sign3 = sign3 + 1
                ElseIf (sign3 = 2) Then
                    Dim ORdeal1 As String = String.Format(ORightTemplate, element.Link, element.ProductImage.Replace("http://www.groupbuyer.com.hk", "http://c.groupbuyermail.com"), element.Subject, element.DiscountPrice, element.Price, element.Link)
                    middleDealsHtml = middleDealsHtml & ORdeal1
                    sign3 = 0
                End If
            Next
            '2014/03/20 added,add top6 to a new block,end

            '填充两个Banner，header和footer
            Dim bannerHtmlString As String = "<tr><td width=""700"" ><a href=""{0}"" link_category=""{2}"" target=""_blank""><img src=""{1}"" border=""0""/></a>  </td></tr>"
            Dim headerBanner As String = ""
            Dim footerBanner As String = ""
            Try
                headerBanner = String.Format(bannerHtmlString, ElementAndTitle.Banner(0).ahrefURL, ElementAndTitle.Banner(0).imgSRC, ElementAndTitle.Banner(0).linkCategory)
                footerBanner = String.Format(bannerHtmlString, ElementAndTitle.Banner(1).ahrefURL, ElementAndTitle.Banner(1).imgSRC, ElementAndTitle.Banner(1).linkCategory)
                ''为降低出错率，强制header相同footer
                'footerBanner = headerBanner
            Catch ex As Exception

            End Try

            '填充两个Banner，header和footer
            Dim Deals As String = FTemplate.Replace("[ALL_NEW_DEALS]", MainDealsHtml)
            Dim html As String = Deals.Replace("[ALL_OLD_DEALS]", OtherDealsHtml).Replace("[ALL_TOP_DEALS]", middleDealsHtml)
            html = html.Replace("[HEADER_BANNER]", headerBanner)
            html = html.Replace("[FOOTER_BANNER]", footerBanner)
            html = SpecialCode.AddSpecialCode(urlSpecialCode, html)
            Dim campaign As New Campaign
            campaign.Subject = subject
            campaign.Content = html

            Return campaign
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' 添加时间：2013/05/14；
    ''' 添加人：彭震；
    ''' 实现功能：填充mobile模板和获取Subject，并为Campaign对象属性Content、Subject赋值；
    ''' 模板被分成4个部分；
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="category"></param>
    ''' <param name="FirstTemplate"></param>
    ''' <param name="SecondTemplate"></param>
    ''' <param name="ThirdTemplate"></param>
    ''' <param name="ForthTemplate"></param>
    ''' <param name="EmailTemplateElements"></param>
    ''' <param name="PubDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CustomizeTempalateDeal(ByVal LoginEmail As String, ByVal AppID As String, ByVal category As String, ByVal FirstTemplate As String, ByVal SecondTemplate As String, ByVal ThirdTemplate As String, ByVal ForthTemplate As String, ByVal EmailTemplateElements As EmailCampaignElementDeal(), ByVal PubDate As String) As Campaign
        Dim A As New alerter
        'begin sort "the most favorite deals" will up to to
        Dim subject As String = ""
        Dim comparer As New MyComparer

        comparer.Favorite = category
        Try
            Array.Sort(EmailTemplateElements, comparer)
        Catch ex As Exception
            'A.LogText(ex.ToString)'2013/05/15注释，不需要写Log
        End Try
        Try
            EmailTemplateElements = InsertPopularDeals2(LoginEmail, AppID, EmailTemplateElements)
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try
        'generate the subject
        Dim NumNew As Integer = 0
        Dim NumOld As Integer = 0
        For Each EmailTemplateElement As EmailCampaignElementDeal In EmailTemplateElements
            Dim DealSubject As String = EmailTemplateElement.Subject
            Dim strChannelPubDate As String = EmailTemplateElement.PubDate.Replace("HKT", "")
            Dim DealPubDate As Date = Date.Parse(strChannelPubDate)
            If DealPubDate.Date = Now.Date Then
                If DealSubject.IndexOf("【") > -1 Then
                    Try
                        DealSubject = System.Web.HttpUtility.HtmlDecode(DealSubject.Substring(DealSubject.IndexOf("【"), DealSubject.IndexOf("】") - DealSubject.IndexOf("【") + 1))
                    Catch ex As Exception
                        Common.LogText(ex.ToString())
                    End Try

                    '2013/05/15修改，去掉【】【】之间的"/"，更改成：【兜有●电影2D单人电影票】【银通之旅宾馆住宿】
                    'subject = subject & "/" & DealSubject
                    subject = subject & DealSubject

                End If
            End If
            Dim strPubDate As String = Format(DateTime.Parse(EmailTemplateElement.PubDate), "yyyy-MM-dd")
            Dim strNowDate As String = Format(Now, "yyyy-MM-dd")
            If strPubDate = strNowDate Then
                NumNew = NumNew + 1
            Else
                NumOld = NumOld + 1
            End If
        Next

        '修改时间：2013/05/15
        '修改原因：不用添加"/"，并将所有的"/"修改成"●"
        'remove the first "/"
        'If Not String.IsNullOrEmpty(subject) Then
        '    subject = subject.Remove(0, 1)
        'End If
        'subject = subject.Replace("/", "●")

        '填充模板
        Dim SpreadTemplate As String = FirstTemplate
        Dim todayProductTemplate As String = ""
        Dim otherProductTemplate As String = ""
        Try
            For i As Integer = 0 To EmailTemplateElements.Length - 1
                Dim strPubDate As Date = EmailTemplateElements(i).PubDate.Replace("HKT", "")
                Dim DealPubDate As Date = Date.Parse(strPubDate)
                If DealPubDate.Date = Now.Date Then
                    todayProductTemplate = todayProductTemplate & String.Format(SecondTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Price, EmailTemplateElements(i).DiscountPrice)
                Else
                    otherProductTemplate = otherProductTemplate & String.Format(SecondTemplate, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Price, EmailTemplateElements(i).DiscountPrice)
                End If
            Next
            SpreadTemplate = SpreadTemplate & todayProductTemplate & ThirdTemplate & otherProductTemplate & ForthTemplate
        Catch ex As Exception
            Common.LogText(ex.ToString())
        End Try

        Dim campaign As New Campaign
        campaign.Subject = subject
        campaign.Content = SpreadTemplate
        Return campaign
    End Function


    Sub UpdateCategories(ByVal OldCategories As String, ByVal newCategories As String)
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
    End Sub


    ''' <summary>
    ''' windows Template,update EmailCampaignElements and put the popular deals into today deals
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="EmailTemplateElements"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function InsertPopularDeals(ByVal LoginEmail As String, ByVal AppID As String, ByVal EmailTemplateElements As EmailCampaignElementDeal()) As EmailCampaignElementDeal()
        Dim A As New alerter
        Dim NumNew As Integer = 0 'Rss中今日团购的产品个数，根据时间判断
        Dim NumOld As Integer = 0 'Rss中往日团购的产品个数，根据时间判断
        For i As Integer = 0 To EmailTemplateElements.Length - 1
            Dim d1 As String = Format(DateTime.Parse(EmailTemplateElements(i).PubDate), "yyyy-MM-dd")
            Dim d2 As String = Format(Now, "yyyy-MM-dd")
            If d1 = d2 Then
                NumNew = NumNew + 1
            Else
                NumOld = NumOld + 1
            End If
        Next

        '今日团购为偶数，则添加2个点击量最大的产品；
        '若今日团购个数为奇数，则添加1个点击量最大的产品。
        Dim NumNeed As Integer = 2 - NumNew Mod 2

        Dim DealLists As List(Of EmailCampaignElementDeal) = New List(Of EmailCampaignElementDeal)(EmailTemplateElements)

        Dim PopularDeals As New List(Of EmailCampaignElementDeal)


        Dim CampaignName As String = Format(Now.Date.AddDays(-1), "dd/MM/yyyy")
        Dim spreadiws As New EmailAlerter.IntrawWebService.Service
        Dim ds As DataSet = spreadiws.dsGetURLTrackings(LoginEmail, AppID, CampaignName)
        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then

            Dim t As Integer = 0
            If NumNeed = 0 Then NumNeed = 2
            For m As Integer = 0 To NumNeed - 1
                Try
                    If t < dt.Rows.Count Then
                        Dim URL As String = dt.Rows(t).Item(0).ToString
                        Dim FoundTag As Boolean = False
                        For j As Integer = 0 To DealLists.Count - 1
                            Dim DealLink As String = Regex.Match(DealLists(j).Link.ToString, "http[^""]*").ToString
                            '判断点击量最高的产品URL和rss中的产品URL是否相同，忽略字符大小写匹配
                            '2013/05/15 ,the URL add ga code
                            If (URL.LastIndexOf("?") > -1) Then
                                URL = URL.Substring(0, URL.LastIndexOf("?"))
                            End If

                            If String.Compare(URL, DealLink, True) = 0 Then
                                Dim d1 As String = Format(DateTime.Parse(DealLists(j).PubDate), "yyyy-MM-dd")
                                Dim d2 As String = Format(Now, "yyyy-MM-dd")
                                If d1 <> d2 Then
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
                                    Deal.Content = DealLists(j).Content '2013/5/16，EmailCampaignElementDeal类新增Content字段
                                    PopularDeals.Add(Deal) '将产品信息添加到最喜欢的List中
                                    DealLists.Remove(DealLists(j)) '在所有产品List中,在相应位置移除最喜欢产品
                                    FoundTag = True
                                    Exit For
                                End If
                            End If
                        Next

                        If FoundTag = False Then
                            m = m - 1
                        End If
                        t = t + 1

                    End If

                Catch ex As Exception
                    Common.LogText(ex.ToString)
                End Try
            Next



            Try
                '在第二、第三个位子添加上一期点击量最大的产品
                For i As Integer = 0 To PopularDeals.Count - 1
                    DealLists.Insert(i + 1, PopularDeals(i))
                Next
            Catch ex As Exception
                Common.LogText(ex.ToString)
            End Try

        End If
        '24/12/2012
        Return DealLists.ToArray()
    End Function

    ''' <summary>
    ''' 添加时间：2013/05/14,
    ''' 添加人：彭震,
    ''' 功能实现：Mobile Template,update EmailCampaignElements and put the popular deals into today deals
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="EmailTemplateElements"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function InsertPopularDeals2(ByVal LoginEmail As String, ByVal AppID As String, ByVal EmailTemplateElements As EmailCampaignElementDeal()) As EmailCampaignElementDeal()
        Dim A As New alerter
        Dim NumNeed As Integer = 2 '往期点击量靠前的2个产品
        Dim DealLists As List(Of EmailCampaignElementDeal) = New List(Of EmailCampaignElementDeal)(EmailTemplateElements)
        Dim PopularDeals As New List(Of EmailCampaignElementDeal)
        Dim CampaignName As String = Format(Now.Date.AddDays(-1), "dd/MM/yyyy")
        Dim spreadiws As New EmailAlerter.IntrawWebService.Service

        '获取昨日点击量最高的产品URL
        Dim ds As DataSet = spreadiws.dsGetURLTrackings(LoginEmail, AppID, CampaignName)
        Dim dt As DataTable = ds.Tables(0)
        Dim counter As Integer = 0
        If dt.Rows.Count > 0 Then
            For t As Integer = 0 To dt.Rows.Count - 1
                Try
                    If counter < NumNeed Then
                        Dim URL As String = dt.Rows(t).Item(0).ToString()
                        Dim FoundFlag As Boolean = False
                        For j As Integer = 0 To DealLists.Count - 1
                            Dim DealLink As String = Regex.Match(DealLists(j).Link.ToString, "http[^""]*").ToString
                            '判断点击量最高的产品URL和rss中的产品URL是否相同，忽略字符大小写匹配
                            If (URL.LastIndexOf("?") > -1) Then
                                URL = URL.Substring(0, URL.LastIndexOf("?"))
                            End If
                            If String.Compare(URL, DealLink, True) = 0 Then
                                Dim strPubDate As String = Format(DateTime.Parse(DealLists(j).PubDate), "yyyy-MM-dd")
                                Dim strNowDate As String = Format(Now, "yyyy-MM-dd")
                                If strPubDate <> strNowDate Then
                                    Dim Deal As EmailCampaignElementDeal = New EmailCampaignElementDeal
                                    Deal.DealCategory = DealLists(j).DealCategory
                                    Deal.Subject = DealLists(j).Subject
                                    Deal.Link = DealLists(j).Link
                                    Deal.Price = DealLists(j).Price
                                    Deal.Discount = DealLists(j).Discount
                                    Deal.DiscountPrice = DealLists(j).DiscountPrice
                                    Deal.Saved = DealLists(j).Saved
                                    Deal.LargeBuyButton = DealLists(j).LargeBuyButton
                                    Deal.ProductImage = DealLists(j).ProductImage
                                    Deal.PubDate = Format(Now, "yyyy-MM-dd 00:00:00")
                                    Deal.Index = j
                                    Deal.BuyButton = DealLists(j).BuyButton
                                    Deal.Content = DealLists(j).Content
                                    '把昨日团购点击量最大的产品信息放到新对象中，并修改新对象的PubDate
                                    PopularDeals.Add(Deal)
                                    '在相应的位置移除点击量最大的昨日产品
                                    DealLists.Remove(DealLists(j))
                                    counter = counter + 1
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        Exit For
                    End If
                Catch ex As Exception
                    Common.LogText(ex.ToString())
                End Try
            Next
            Try
                '把点击量最高的产品信息添加到第二、三个位置
                For i As Integer = 0 To PopularDeals.Count - 1
                    DealLists.Insert(i + 1, PopularDeals(i))
                Next
            Catch ex As Exception
                Common.LogText(ex.ToString())
            End Try
        End If
        Return DealLists.ToArray()
    End Function



    Function CustomizeTempalateDeal2(ByVal LoginEmail As String, ByVal AppID As String, ByVal category As String, ByVal Template As String, ByVal subTemplate1 As String, ByVal subTemplate2 As String, ByVal subTemplate3 As String, ByVal EmailTemplateElements As EmailCampaignElementDeal(), ByVal PubDate As String) As Campaign
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
            Dim Ldeal As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\todayleft.txt")
            Dim Rdeal As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\todayright.txt")
            Dim OLdeal As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\Oleft.txt")
            Dim OMiddle As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\Omiddle.txt")
            Dim ORdeal As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\Oright.txt")
            Dim FTemplate As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\FTemplate.txt")
            ' Dim today As String = IO.File.ReadAllText("C:\Users\Suwen\Desktop\groupbuyer2\todayright.txt")
            Dim MainDealsHtml As String = ""
            Dim OtherDealsHtml As String = ""
            Dim sign1 As Integer = 0
            Dim sign2 As Integer = 0
            For i As Integer = 0 To EmailTemplateElements.Length - 1
                'For i As Integer = 0 To 5
                Try
                    Dim s As String = EmailTemplateElements(i).PubDate.Replace("HKT", "")
                    Dim DealPubDate As Date = Date.Parse(s)
                    If DealPubDate.Date = Now.Date Then
                        'put today's deal into html
                        If sign1 = 0 Then
                            Dim Ldeal1 As String = "<tr><td width=""700"" height=""16""></td></tr>" & Ldeal
                            MainDealsHtml = MainDealsHtml & Ldeal1
                            sign1 = 1
                        ElseIf sign1 = 1 Then
                            MainDealsHtml = MainDealsHtml & Rdeal
                            sign1 = 0
                        End If
                    ElseIf DealPubDate < Now.Date Then
                        'not today's deal and put it into certain position
                        If sign2 = 0 Then
                            Dim OLdeal1 As String = "<tr><td width=""700"" height=""16""></td></tr>" & OLdeal
                            OtherDealsHtml = OtherDealsHtml & OLdeal1
                            sign2 = sign2 + 1
                        ElseIf sign2 = 1 Then

                            'Dim ODealHtml As String = String.Format(subTemplate3, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                            OtherDealsHtml = OtherDealsHtml & OMiddle
                            sign2 = sign2 + 1
                        ElseIf sign2 = 2 Then

                            'Dim ODealHtml As String = String.Format(subTemplate3, EmailTemplateElements(i).Link, EmailTemplateElements(i).Subject, EmailTemplateElements(i).Link, EmailTemplateElements(i).ProductImage, EmailTemplateElements(i).DiscountPrice, EmailTemplateElements(i).Link, EmailTemplateElements(i).Price, EmailTemplateElements(i).Discount, EmailTemplateElements(i).Saved)
                            OtherDealsHtml = OtherDealsHtml & ORdeal
                            sign2 = 0
                        End If
                    End If

                Catch ex As Exception
                    System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())

                End Try

            Next


            Dim NewDeals As String = FTemplate.Replace("<tr><td>[ALL_NEW_DEALS]</td></tr>", MainDealsHtml)
            Dim html As String = NewDeals.Replace("<tr><td>[ALL_OLD_DEALS]</td></tr>", OtherDealsHtml)

            Dim campaign As New Campaign
            campaign.Subject = subject
            campaign.Content = html


            Return campaign
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
            Common.LogText(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' 添加日期：2013/05/10，
    ''' 功能实现：给SpreadTemplate添加追踪代码
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="compaignName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddSpecialCode(ByVal SpreadTemplate As String, ByVal compaignName As String) As String
        Dim byteArray As Byte() = System.Text.Encoding.GetEncoding("gb2312").GetBytes(SpreadTemplate) 'GB2312为简体中文
        Dim stream As New System.IO.MemoryStream(byteArray)
        '给模板最后加工
        Dim doc As New HtmlDocument()
        doc.Load(stream)
        Dim linkNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a")
        For Each link As HtmlNode In linkNodes
            Dim hrefStr As String = link.GetAttributeValue("href", "")
            Dim newUrl As String = hrefStr & "?utm_source=spread&utm_medium=email&utm_term=" & "GROUPBUYER" & "&utm_campaign=" & compaignName
            link.SetAttributeValue("href", newUrl)
        Next
        SpreadTemplate = doc.DocumentNode.OuterHtml.ToString()
        Return SpreadTemplate
    End Function


End Class
