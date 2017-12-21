Imports System.Text.RegularExpressions
Imports System.Xml
Imports HtmlAgilityPack

Public Class GroupBuyer

    Private Shared efContext As New EmailAlerterEntities()
    'Private Shared beautyWeekdayNum As String = "4" '新增的beauty专刊的活跃星期，在真实服务器上此值应为4。

    Public Shared Sub Start(ByVal issueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, _
                     ByVal appId As String, ByVal Url As String, ByVal nowTime As DateTime)
        Dim helper As New EFHelper
        GetCategory(siteId, nowTime)
        Dim categoryName As String = "-5"
  
        If (planType.Trim = "HO1") Then
            categoryName = "-2"
        End If
        GetProducts(siteId, issueID, Url, categoryName)
        If (categoryName = "-5") Then
            helper.InsertContactList(issueID, "Members", "draft")
            helper.InsertContactList(issueID, "FB", "draft")
        ElseIf (categoryName = "-2") Then
            helper.InsertContactList(issueID, "Members", "draft")
            helper.InsertContactList(issueID, "Subscribers", "draft")
            helper.InsertContactList(issueID, "FB", "draft")
            helper.InsertContactList(issueID, "Newsletter", "draft")
            helper.InsertContactList(issueID, "Puzzle", "draft")
            helper.InsertContactList(issueID, "OSO", "draft")
        End If
        GetSubject(issueID, siteId, categoryName)
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer, ByVal lastUpdate As DateTime)
        Dim myCategory As New Category
        myCategory.Category1 = "-5"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        Try
            Dim queryCategory As Category = efContext.Categories.Where(Function(c) c.Category1 = myCategory.Category1 AndAlso myCategory.SiteID = siteId).Single()
        Catch ex As Exception
            efContext.AddToCategories(myCategory)
            efContext.SaveChanges()
        End Try
    End Sub

    Public Shared Sub GetProducts(ByVal siteId As Integer, ByVal issueId As Integer, ByVal rssUrl As String, ByVal categoryName As String)
        Dim helper As New EFHelper
        Dim doc As HtmlDocument = helper.GetHtmlDocument(rssUrl, 120000)
        'Dim productLis As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='fL deal_list_block']/a")
        Dim productLis As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//*[@id='deal_list']/div[@class='fL w100 deal_list_block']/a")
        Dim reg As String = "([0-9]+)"
        Dim ef As EFHelper = New EFHelper()
        For Each li As HtmlNode In productLis
            Dim productName As String = li.SelectSingleNode("div/div[@class='deal_list_right']/div[@class='deal-short_topic_box']").InnerText
            Dim productUrl As String = ef.addDominForUrl("http://www.groupbuyer.com.hk/", li.GetAttributeValue("href", ""))
            'Dim price As String = Regex.Matches(li.SelectSingleNode("a/div[@class='price-list clearfix']/span[1]").InnerText, reg)(0).Value
            'Dim discount As String = Regex.Matches(li.SelectSingleNode("a/div[@class='price-list clearfix']/span[2]").InnerText, reg)(0).Value
            Dim sales As String = "0" '销量
            Try
                sales = Regex.Matches(li.SelectSingleNode("div/div[@class='deal_list_right']/div[@class='deal_price']/div[@class='crGrey4A deal_price_list_right']").InnerText, reg)(0).Value
            Catch ex As Exception
                'Ignore
            End Try

            Dim pictureUrl As String = li.SelectSingleNode("div/div[@class='fL deal_list_left']/img").GetAttributeValue("src", "")
            If (pictureUrl.Contains("live_redeem_normal")) Then
                pictureUrl = li.SelectSingleNode("div/div[@class='fL deal_list_left']/img[2]").GetAttributeValue("src", "")
            End If
            Dim myProduct As New Product
            myProduct.Prodouct = productName
            myProduct.Url = productUrl
            Try
                Dim price1 As Double = Double.Parse(Regex.Matches(li.SelectSingleNode("div/div[@class='deal_list_right']/div[@class='deal_price']/div[@class='fL deal_price_list_left']/span[1]").InnerText, reg)(0).Value)
                Dim price2 As Double = Double.Parse(Regex.Matches(li.SelectSingleNode("div/div[@class='deal_list_right']/div[@class='deal_price']/div[@class='fL deal_price_list_left']/span[2]").InnerText, reg)(0).Value)
                If (price1 < price2) Then
                    myProduct.Discount = price1
                    myProduct.Price = price2
                Else
                    myProduct.Discount = price2
                    myProduct.Price = price1
                End If
            Catch ex As Exception
                'Ignore
            End Try
            myProduct.Sales = Int32.Parse(sales)
            myProduct.PictureUrl = pictureUrl
            myProduct.Description = productName
            myProduct.SiteID = siteId
            myProduct.Currency = "HK$"
            Dim productId As Integer = helper.InsertOrUpdateProduct(myProduct, "", categoryName, siteId)
            helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
        Next

        '2014/03/27 使用food页面获取所有的food产品
        'Dim helper As New EFHelper
        'Dim pageStr As String = EFHelper.GetRssPageString(rssUrl)
        'Try
        '    Dim Collection As MatchCollection = Regex.Matches(pageStr, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", RegexOptions.IgnoreCase)
        '    If Collection.Count > 0 Then
        '        '获取channel title，并为channel title加上![CDATA，
        '        '防止获取到的channel title没加上![CDATA,使用.net方法读取XML文件时出错
        '        Dim myTitle As String = Collection.Item(0).Groups(1).Value
        '        Dim myTitleNode As String = "<title><![CDATA[" & myTitle & "]]></title>"
        '        pageStr = System.Text.RegularExpressions.Regex.Replace(pageStr, "<title>((?!<\!\[CDATA|\<title)[\s\S])*</title>", myTitleNode)
        '    End If
        'Catch ex As Exception
        '    Throw New Exception(ex.ToString())
        'End Try
        'Dim xmlDoc As New XmlDocument
        'xmlDoc.LoadXml(pageStr)
        'Dim n As Integer = xmlDoc.SelectNodes("//item").Count
        'Dim itemNodeList As Xml.XmlNodeList = xmlDoc.SelectNodes("//item")
        ''Dim channelNode As Xml.XmlNodeList = xmlDoc.SelectNodes("//channel")
        'Dim ChannelTitle As String = System.Net.WebUtility.HtmlDecode(xmlDoc.SelectSingleNode("//channel").SelectSingleNode("title").InnerText.Trim())
        'Dim PubDateString As String = xmlDoc.SelectSingleNode("//channel").SelectSingleNode("pubDate").InnerText.Trim()
        'Dim PubDate As DateTime = PubDateString.Substring(5, 20)
        'For Each node As Xml.XmlNode In itemNodeList
        '    Try
        '        Dim titleNodeText As String
        '        Dim title As String = System.Net.WebUtility.HtmlEncode(node.SelectSingleNode("title").InnerText.Trim())

        '        '2013/05/14新增
        '        Dim content As String = node.SelectSingleNode("title").InnerText.Replace("<![CDATA[", "").Replace("]]>", "")

        '        If title.IndexOf("【") <> -1 And title.IndexOf("】") <> -1 Then
        '            titleNodeText = title.Substring(title.IndexOf("【"), title.IndexOf("】") - title.IndexOf("【") + 1)
        '        Else
        '            If (title.Length >= 50) Then
        '                titleNodeText = title.Substring(0, 49) & "..."
        '            Else
        '                titleNodeText = title
        '            End If
        '        End If
        '        Dim link_category As String = node.SelectSingleNode("category").InnerText.Trim()
        '        If Not (link_category.Contains("-5")) Then
        '            Continue For
        '        End If
        '        Dim priceText As String = node.SelectSingleNode("description").SelectSingleNode("price").InnerText.Trim()
        '        Dim discountText As String = node.SelectSingleNode("description").SelectSingleNode("discount").InnerText.Trim()
        '        Dim savedText As String = node.SelectSingleNode("description").SelectSingleNode("saved").InnerText.Trim()
        '        Dim discountPriceText As String = node.SelectSingleNode("description").SelectSingleNode("discountPrice").InnerText.Trim()
        '        Dim productImageText As String = node.SelectSingleNode("description").SelectSingleNode("productImage").InnerText.Trim()
        '        Dim largeBuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("largeBuyButton").InnerText.Trim()
        '        Dim BuyButtonText As String = node.SelectSingleNode("description").SelectSingleNode("BuyButton").InnerText.Trim()
        '        Dim postDateNodeText As String = node.SelectSingleNode("pubDate").InnerText
        '        Dim expireDateText As String = node.SelectSingleNode("endDate").InnerText
        '        '2014/02/08 added,增加新需求产品购买数量的排序
        '        Dim iNoofPurchased As Integer = Int32.Parse(node.SelectSingleNode("noofPurchased").InnerText.Trim())

        '        Dim linkNodeText As String = node.SelectSingleNode("link").InnerText
        '        linkNodeText = linkNodeText & """ link_category=""" & link_category

        '        'Start by Gary, 2014/01/28 
        '        '处理RSS description里没有价格时，获取外层的价格'
        '        Try
        '            If Not IsNumeric(priceText) Then
        '                priceText = node.SelectSingleNode("orginalPrice").InnerText.Trim  '原价
        '            End If
        '            If Not IsNumeric(discountPriceText) Then
        '                discountPriceText = node.SelectSingleNode("actualPrice").InnerText.Trim  '折扣价
        '            End If
        '        Catch ex As Exception
        '            'Ignore
        '        End Try
        '        'End by Gary
        '        Dim myProduct As New Product
        '        myProduct.Prodouct = titleNodeText
        '        myProduct.Description = titleNodeText
        '        myProduct.Url = linkNodeText
        '        myProduct.Price = Double.Parse(priceText)
        '        myProduct.Discount = Double.Parse(discountPriceText)
        '        myProduct.Currency = "HK$"
        '        myProduct.Sales = iNoofPurchased
        '        myProduct.PictureUrl = productImageText
        '        myProduct.LastUpdate = Format(DateTime.Parse(postDateNodeText), "yyyy-MM-dd hh:mm:ss")
        '        myProduct.ExpiredDate = Format(DateTime.Parse(expireDateText), "yyyy-MM-dd hh:mm:ss") 'DateTime.Parse(postDateNodeText)
        '        myProduct.SiteID = siteId
        '        Dim productId As Integer = helper.InsertOrUpdateProduct(myProduct, "", "-5", siteId)
        '        helper.InsertSinglePIssue(productId, siteId, issueId, "CA")
        '    Catch ex As Exception
        '        Throw New Exception(ex.ToString())
        '    End Try
        'Next
    End Sub

    Private Shared Sub GetSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal categoryName As String)
        'Dim subject As String = ""
        'Dim listSubject As List(Of String) = (From p In efContext.Products
        '                      Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
        '                      Where pi.SiteId = siteId AndAlso pi.IssueID = issueId
        '                      Select p.Prodouct).ToList()
        'For Each li In listSubject
        '    subject = subject & li
        'Next
        'subject = "Hi [FIRSTNAME],【食品專享】" & subject

        '2014/03/19修改
        Dim subject1 As String = "【號外】 本週至Hit飲食精選, 限時搶購!" & "5" '最后附加的这个CID是用于在Emailalerter中创建campaign时指定对应的campaignName
        '20140926 edited by Dora
        If (categoryName = "-2") Then
            subject1 = "美容最強！過百款美容團購優惠大匯集" & "2" '最后附加的这个CID是用于在Emailalerter中创建campaign时指定对应的campaignName
        End If
        '20140926 edited by Dora
        Dim helper As New EFHelper
        helper.InsertIssueSubject(issueId, subject1)
    End Sub
End Class
