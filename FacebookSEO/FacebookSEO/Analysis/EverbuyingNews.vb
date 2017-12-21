Imports System.Net.Cache
Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Security.Policy

Public Class EverbuyingNews
    Private Shared helper As New EFHelper()
    Private Shared htmlUrl As String = "http://www.everbuying.com/index1.html"
    Private Shared hd As HtmlDocument = GetHtmlDocByUrl(htmlUrl)
    Private Shared efContext As New EmailAlerterEntities()
    Private Shared listProduct As New List(Of Product)
    Private Shared listCategoryName As New HashSet(Of String)

    ''' <summary>
    ''' 程序开始入口
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="Url"></param>
    ''' <remarks></remarks>
    Public Shared Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        listProduct = GetProductList(siteId)

        '将Deals的四个分类插入数据库
        'GetCategories(siteId)
        GetNewArrival(siteId, IssueID, 20)
        GetSubject(IssueID)
        Dim contactListArr() As String = {"Open201307171", "Open201307172", "Open201307173", "Open201307174"}
        helper.InsertContactList(IssueID, contactListArr, "draft") 'draft 'waiting

        'CreateContacts(siteId, planType, spreadLogin, appId, splitContactCount, IssueID)

        'Dim contactLists As List(Of SplitContactList) = helper.GetSplitContactLists(siteId)
        'If contactLists.Count > 0 Then
        '    For Each contactList As SplitContactList In contactLists
        '        helper.InsertContactList(IssueID, contactList.ContactListName, contactList.SendingStatus)
        '    Next
        'End If
    End Sub

    Public Shared Sub GetNewArrival(ByVal siteId As Integer, ByVal issueId As Integer, ByVal count As Integer)
        Dim listProductId As New List(Of Integer)
        Dim lastUpdate As DateTime = DateTime.Now
        Dim currency As String = "USD"

        'Deals(Promotion)
        Dim newArrival As String = "http://www.everbuying.com/new-products.html"
        hd = GetHtmlDocByUrl(newArrival)

        'Dim nodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div/ul")
        'For Each node As HtmlNode In nodes.ToList()
        For i As Integer = 1 To count
            Dim node As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@id='g_pro']/div[" + i.ToString() + "]/ul")
            Dim imgNode As HtmlNode = node.SelectSingleNode("li[1]")
            Dim url As String = PathLink(imgNode.SelectSingleNode("a").GetAttributeValue("href", ""))
            Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("data-src", "")
            If String.IsNullOrEmpty(imgUrl) Then
                imgUrl = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
            End If
            Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
            Dim desc As String = node.SelectSingleNode("li[2]/a").InnerText
            Dim priceNode As HtmlNode = node.SelectSingleNode("li[3]")
            Dim price As New Decimal
            Try
                price = Decimal.Parse(priceNode.SelectSingleNode("em/span[2]").InnerText)
            Catch ex As Exception
                'price not exist, ignore
            End Try
            Dim discount As Decimal = Decimal.Parse(priceNode.SelectSingleNode("span[2]").InnerText)

            Dim freeship As String = ""
            Dim ship As String = ""
            Try
                freeship = node.SelectSingleNode("li[@class='fp24'][1]").SelectSingleNode("img").GetAttributeValue("src", "")
            Catch ex As Exception
                'Ignore
            End Try
            Try
                ship = node.SelectSingleNode("li[@class='fp24'][2]").SelectSingleNode("img").GetAttributeValue("src", "")
            Catch ex As Exception
                'Ignore
            End Try

            Dim categoryId As Integer
            Try
                categoryId = GetLinkCategoryId(url, siteId, issueId)
            Catch ex As Exception
                Continue For
            End Try

            Dim returnId As Integer = InsertProduct(desc, url, price, imgUrl, lastUpdate, desc, siteId, currency, imgAlt, categoryId, discount, freeship, ship)
            listProductId.Add(returnId)
        Next
        InsertProductIssue(siteId, issueId, "NE", listProductId)
    End Sub

    ''' <summary>
    ''' 添加数据到SplitContactLists表和ContactLists_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="spreaLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="splitContactCount"></param>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Shared Sub CreateContacts(ByVal siteId As Integer, ByVal planType As String, ByVal spreaLogin As String, _
                              ByVal appId As String, ByVal splitContactCount As Integer, _
                              ByVal issueId As Integer)
        Dim listName As String
        If (planType = "HA") Then
            listName = "NewArrivalOpen"
        ElseIf (planType = "HN") Then
            listName = "NewArrivalNotOpen"  '表示没开启邮件的用户
        End If
        If (splitContactCount >= 2) Then
            Dim queryContactList = From c In efContext.SplitContactLists _
                                 Where c.SiteID = siteId AndAlso c.PlanType = planType _
                                 Select c
            '如果SplitContackLists表中有数据，则遍历其中的数据
            If Not (queryContactList.Count = 0) Then
                Dim listContact As List(Of SplitContactList) = queryContactList.ToList()
                Dim i As Integer = 0
                For Each li In listContact
                    If (li.Flag = False) Then
                        Dim contactList As New ContactLists_Issue()
                        contactList.IssueID = issueId
                        contactList.ContactList = li.ContactListName

                        contactList.SendTime = li.SendTime '2013/4/11新增，增加新需求发送时间

                        efContext.AddToContactLists_Issue(contactList)
                        efContext.SaveChanges()
                        Dim myContact = efContext.SplitContactLists.Single(Function(m) m.SplitID = li.SplitID)
                        myContact.Flag = True
                        efContext.SaveChanges()
                        Exit For
                    End If
                    i = i + 1
                Next
                '如果根据siteId和planType查询到的SplitContactList表数据Flag=True
                If (i = listContact.Count) Then
                    Dim queryDelContact = From c In efContext.SplitContactLists
                                        Where c.SiteID = siteId AndAlso c.PlanType = planType
                                        Select c
                    '删除SplitContactLists表中数据
                    For Each del In queryDelContact
                        efContext.SplitContactLists.DeleteObject(del)
                    Next
                    efContext.SaveChanges()

                    InsertContactList(spreaLogin, appId, siteId, planType, issueId, listName, splitContactCount)
                End If
            Else '初始化时，SplitContactLists表中无数据，向表中添加数据
                InsertContactList(spreaLogin, appId, siteId, planType, issueId, listName, splitContactCount)
            End If
        Else 'splitContactCount=0,1
            InsertContactList(spreaLogin, appId, siteId, planType, issueId, listName, splitContactCount)

            'Dim myContactList As New ContactLists_Issue()
            'myContactList.IssueID = issueId
            'myContactList.ContactList = listName
            'efContext.AddToContactLists_Issue(myContactList)
            'efContext.SaveChanges()

        End If
    End Sub

    ''' <summary>
    ''' 将某一个联系人分组分割成多个联系人分组，并在ContactLists_Issue表中添加一条记录
    ''' </summary>
    ''' <param name="spreaLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="issueId"></param>
    ''' <param name="listName"></param>
    ''' <param name="splitContactCount"></param>
    ''' <remarks></remarks>
    Private Shared Sub InsertContactList(ByVal spreaLogin As String, ByVal appId As String, ByVal siteId As Integer, _
                                  ByVal planType As String, ByVal issueId As Integer, ByVal listName As String, _
                                  ByVal splitContactCount As Integer)
        Dim querySubscriber As New QuerySubscriber()
        querySubscriber.StartDate = Date.Now.AddMonths(-6).ToString("yyyy-MM-dd")
        If (planType = "HA") Then
            querySubscriber.Strategy = ChooseStrategy.Open '表示开启邮件的用户
        ElseIf (planType = "HN") Then
            querySubscriber.Strategy = ChooseStrategy.NotOpen '表示没开启邮件的用户
        End If
        querySubscriber.CountryList = New String() {}
        Dim criteria As String = querySubscriber.ToJsonString()
        Dim topN As Integer = Integer.MaxValue
        Dim saveAsList As String = ""

        If (splitContactCount >= 2) Then 'splitContactCount>=2
            For j As Integer = 1 To splitContactCount
                If Not (j = splitContactCount) Then
                    saveAsList = saveAsList & listName & Now.ToString("yyyyMMdd") & j & ";"
                Else
                    saveAsList = saveAsList & listName & Now.ToString("yyyyMMdd") & j
                End If
            Next
            Dim forceCreate As Boolean = True
            Dim mySpread As SpreadWebReference.Service = New SpreadWebReference.Service()
            mySpread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
            mySpread.Timeout = 1200000
            mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)

            Dim contactName() As String = saveAsList.Split(";")

            For j As Integer = 0 To splitContactCount - 1
                Dim splitList As New SplitContactList()
                splitList.ShopName = "EverbuyingNewArrival"
                splitList.ContactListName = contactName(j)
                splitList.SiteID = siteId
                splitList.PlanType = planType

                If (j = 0) Then
                    splitList.Flag = 1
                    efContext.AddToSplitContactLists(splitList)
                Else
                    splitList.Flag = 0
                    efContext.AddToSplitContactLists(splitList)
                End If
            Next
            Dim contactList As New ContactLists_Issue()
            contactList.IssueID = issueId
            contactList.ContactList = contactName(0)
            contactList.SendingStatus = "draft"
            efContext.AddToContactLists_Issue(contactList)
            efContext.SaveChanges()
        Else 'splitContactCount=0或1
            If (planType = "HA") Then
                saveAsList = "NewArrivalOpens" & planType & Now.ToString("yyyyMMdd") 'Now.ToString("yyyyMMdd") '
            ElseIf (planType = "HN") Then
                saveAsList = "NewArrivalNotOpen" & planType & Now.ToString("yyyyMMdd") 'Now.ToString("yyyyMMdd") '表示没开启邮件的用户
            End If
            Dim forceCreate As Boolean = True
            Dim mySpread As SpreadWebReference.Service = New SpreadWebReference.Service()
            mySpread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
            '2013/4/11新增返回结果判断，如果没有创建联系人列表失败，则不创建campaign
            Dim returnResult As Integer = 0
            Try
                mySpread.Timeout = 600000
                returnResult = mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
            Catch ex As Exception
                mySpread.Timeout = 1200000
                mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)
                LogText(ex.ToString())
            End Try
            If (returnResult > 0) Then
                Dim contactList As New ContactLists_Issue()
                contactList.IssueID = issueId
                contactList.ContactList = saveAsList
                contactList.SendingStatus = "draft"
                efContext.AddToContactLists_Issue(contactList)
            End If

            Dim contact1 As New ContactLists_Issue()
            contact1.IssueID = issueId
            contact1.ContactList = "Opens"
            contact1.SendingStatus = "draft"
            efContext.AddToContactLists_Issue(contact1)

            efContext.SaveChanges()
        End If
    End Sub


    ''' <summary>
    ''' 从Promotion产品中获得Subject信息
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Shared Sub GetSubject(ByVal issueId As Integer)
        Dim subject As String = "NEW ARRIVALS:" + helper.GetFirstProductSubject(issueId)
        helper.InsertIssueSubject(issueId, subject)
    End Sub


    Public Shared Sub GetCategories(siteId As Integer)
        Dim queryCategory = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
            listCategoryName.Add(q.Category1) '2013/4/27新增
        Next

        '更新所有类目
        For Each nds As HtmlNode In hd.DocumentNode.SelectNodes("//div[@class='sidebar_menu']/div/span").ToList()
            Dim cateUrl As String = PathLink(nds.SelectSingleNode("a").GetAttributeValue("href", ""))
            Dim catename As String = RemoveComma(nds.SelectSingleNode("a").InnerText)
            Dim cate As New Category()
            cate.Category1 = catename
            cate.SiteID = siteId
            cate.LastUpdate = DateTime.Now
            cate.Description = catename
            cate.Url = cateUrl
            cate.ParentID = -1
            If JudgeCategory(cate.Category1, cate.SiteID, cate.Url) Then
                efContext.AddToCategories(cate)
            End If
        Next
        efContext.SaveChanges()

        'Dim gdcate As New Category()
        'gdcate.Category1 = "Group Deals"
        'gdcate.SiteID = siteId
        'gdcate.LastUpdate = DateTime.Now
        'gdcate.Description = "Group Deals"
        'gdcate.Url = "http://www.everbuying.com/GroupDeals.html"
        'gdcate.ParentID = -1
        'If JudgeCategory(gdcate.Category1, gdcate.SiteID, gdcate.Url) Then
        '    efContext.AddToCategories(gdcate)
        'End If

        'Dim dccate As New Category()
        'dccate.Category1 = "Discount center"
        'dccate.SiteID = siteId
        'dccate.LastUpdate = DateTime.Now
        'dccate.Description = "Discount center"
        'dccate.Url = "http://www.everbuying.com/discount.html"
        'dccate.ParentID = -1
        'If JudgeCategory(dccate.Category1, dccate.SiteID, dccate.Url) Then
        '    efContext.AddToCategories(dccate)
        'End If

        'Dim pcate As New Category()
        'pcate.Category1 = "Promotions"
        'pcate.SiteID = siteId
        'pcate.LastUpdate = DateTime.Now
        'pcate.Description = "Promotions"
        'pcate.Url = "http://www.everbuying.com/special-c-10000.html"
        'pcate.ParentID = -1
        'If JudgeCategory(pcate.Category1, pcate.SiteID, pcate.Url) Then
        '    efContext.AddToCategories(pcate)
        'End If

        'Dim cate1 As New Category()
        'cate1.Category1 = "1.99DEALS"
        'cate1.SiteID = siteId
        'cate1.LastUpdate = DateTime.Now
        'cate1.Description = "1.99DEALS"
        'cate1.Url = "http://www.everbuying.com/promotional-products-under-1.99.html"
        'cate1.ParentID = -1
        'If JudgeCategory(cate1.Category1, cate1.SiteID, cate1.Url) Then
        '    efContext.AddToCategories(cate1)
        'End If

        'Dim cate2 As New Category()
        'cate2.Category1 = "2.99DEALS"
        'cate2.SiteID = siteId
        'cate2.LastUpdate = DateTime.Now
        'cate2.Description = "2.99DEALS"
        'cate2.Url = "http://www.everbuying.com/promotional-products-under-2.99.html"
        'cate2.ParentID = -1
        'If JudgeCategory(cate2.Category1, cate2.SiteID, cate2.Url) Then
        '    efContext.AddToCategories(cate2)
        'End If

        'Dim cate3 As New Category()
        'cate3.Category1 = "3.99DEALS"
        'cate3.SiteID = siteId
        'cate3.LastUpdate = DateTime.Now
        'cate3.Description = "3.99DEALS"
        'cate3.Url = "http://www.everbuying.com/promotional-products-under-3.99.html"
        'cate3.ParentID = -1
        'If JudgeCategory(cate3.Category1, cate3.SiteID, cate3.Url) Then
        '    efContext.AddToCategories(cate3)
        'End If

        'Dim cate4 As New Category()
        'cate4.Category1 = "4.99DEALS"
        'cate4.SiteID = siteId
        'cate4.LastUpdate = DateTime.Now
        'cate4.Description = "4.99DEALS"
        'cate4.Url = "http://www.everbuying.com/promotional-products-under-4.99.html"
        'cate4.ParentID = -1
        'If JudgeCategory(cate4.Category1, cate4.SiteID, cate4.Url) Then
        '    efContext.AddToCategories(cate4)
        'End If


    End Sub

    ''' <summary>
    ''' 当前Spread只支持最顶层的分类追踪,每期更新LastUpdate也会不相同，故加一个限制条件
    ''' </summary>
    ''' <param name="productUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetLinkCategoryId(ByVal productUrl As String, ByVal siteId As Integer, ByVal IssueId As Integer) As Integer
        Dim category As String = GetCategoryByLink(productUrl)
        Dim query = From c In efContext.Categories
                    Where c.Category1 = category AndAlso c.SiteID = siteId
                    Select c.CategoryID
        Dim returnId As Integer = query.Single()
        Return returnId
    End Function


    Public Shared Sub TestData(ByVal siteId As Integer)
        Dim listProductId As New List(Of Integer)
        Dim lastUpdate As DateTime = DateTime.Now
        Dim currency As String = "USD"
        'GetAds&Promotion Scroll Item
        'Dim seq As Integer = 1
        'For Each nds As HtmlNode In hd.GetElementbyId("ul-auto_scroll").SelectNodes("li[" & seq & "]/div[1]").ToList()
        '    Dim url As String = nds.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim fullurl As String = PathLink(url)
        'Next
        'For Each nds As HtmlNode In hd.GetElementbyId("ul-auto_scroll").SelectNodes("li[" & seq & "]/div[2]").ToList()
        '    For Each ndss As HtmlNode In nds.SelectNodes("ul/li").ToList()
        '        Dim imgnd As HtmlNode = ndss.SelectSingleNode("div[1]")
        '        Dim descnd As HtmlNode = ndss.SelectSingleNode("div[2]")

        '        Dim imgUrl As String = PathLink(imgnd.SelectSingleNode("a").GetAttributeValue("href", ""))
        '        Dim descUrl As String = PathLink(descnd.SelectSingleNode("p[1]/a").GetAttributeValue("href", ""))
        '        Dim desc As String = descnd.SelectSingleNode("p[1]").InnerHtml
        '        Dim price1 As String = descnd.SelectSingleNode("p[2]").InnerText
        '        Dim price2 As String = descnd.SelectSingleNode("p[3]").InnerText
        '    Next
        'Next

        'Group Deals
        'Dim groupdealurl As String = "http://www.everbuying.com/GroupDeals.html"
        'hd = GetHtmlDocByUrl(groupdealurl)
        'For Each gnode As HtmlNode In hd.DocumentNode.SelectNodes("//div[@class='tgzw']").ToList()
        '    Dim imgUrl As String = gnode.SelectSingleNode("//div[@class='tgzwnrr1']/a/img").GetAttributeValue("data-src", "")
        '    Dim imgAlt As String = gnode.SelectSingleNode("//div[@class='tgzwnrr1']/a/img").GetAttributeValue("alt", "")
        '    Dim url As String = gnode.SelectSingleNode("div[@class='tgzwbt']/a").GetAttributeValue("href", "")
        '    Dim desc As String = gnode.SelectSingleNode("div[@class='tgzwbt']/a").InnerText
        '    Dim price As String = gnode.SelectSingleNode("//div[@class='tgzwnrl2a']/span").InnerText
        '    Dim discount As String = gnode.SelectSingleNode("//div[@class='buypri']").InnerText

        '    Dim returnId As Integer = InsertProduct(desc, url, price, imgUrl, lastUpdate, desc, siteId, currency, imgAlt, discount, listProduct)
        '    listProductId.Add(returnId)

        'Next

        'Deals Promotion
        'Dim promotionsUrl As String = "http://www.everbuying.com/special-c-10000.html"
        'hd = GetHtmlDocByUrl(promotionsUrl)
        'For Each pnode As HtmlNode In hd.DocumentNode.SelectNodes("//div[@class='hotp31']").ToList()
        '    Dim imgNode As HtmlNode = pnode.SelectSingleNode("p[@class='hp11']")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = pnode.SelectSingleNode("p[@class='hp22']").InnerText
        '    Dim shipUrl As String = pnode.SelectSingleNode("p[@class='hp33']/img").GetAttributeValue("src", "")
        '    Dim price As String = pnode.SelectSingleNode("p[@class='hp44']/span[2]").InnerText
        '    Dim discount As String = pnode.SelectSingleNode("p[@class='hp55']/span[2]").InnerText
        'Next

        'Dim nodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div/ul")
        'For Each node As HtmlNode In nodes.ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("li[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imageUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim desc As String = node.SelectSingleNode("li[2]/a").InnerText
        '    Dim priceNode As HtmlNode = node.SelectSingleNode("li[3]")
        '    Dim price As String = priceNode.SelectSingleNode("em/span[2]").InnerText
        '    Dim discount As String = priceNode.SelectSingleNode("span[2]").InnerText
        '    Dim freeshipNode As HtmlNode = node.SelectSingleNode("li[@class='fp24'][1]")
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("li[@class='fp24'][2]")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If freeshipNode IsNot Nothing Then
        '        freeship = freeshipNode.SelectSingleNode("img").GetAttributeValue("src", "")
        '    End If
        '    If shipNode IsNot Nothing Then
        '        ship = shipNode.SelectSingleNode("img").GetAttributeValue("src", "")
        '    End If
        'Next

        'Deals discount center
        'Dim discountCenterUrl As String = "http://www.everbuying.com/discount.html"
        'hd = GetHtmlDocByUrl(discountCenterUrl)
        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//div[@class='w_975 prod_main'][1]/ul/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[2]").InnerText
        '    Dim price As String = node.SelectSingleNode("p[4]/span[1]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[5]")
        '    Dim freeship As String = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    Dim ship As String = shipNode.SelectSingleNode("img[2]").GetAttributeValue("src", "")
        'Next

        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//div[@class='w_975 prod_main'][2]/ul/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim price As String = node.SelectSingleNode("p[3]/span[1]/span[2]").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[3]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[@class='freeshipping']")
        '    Dim ship24Node As HtmlNode = node.SelectSingleNode("p[@class='ships24']")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If shipNode IsNot Nothing Then
        '        freeship = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        '    If ship24Node IsNot Nothing Then
        '        ship = ship24Node.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        'Next

        'Deals under price
        'Dim discountCenterUrl As String = "http://www.everbuying.com/promotional-products-under-4.99.html"
        'hd = GetHtmlDocByUrl(discountCenterUrl)
        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//ul[@id='switch-under1-trigger']/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim price As String = node.SelectSingleNode("p[3]/span[1]/span[2]").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[3]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[@class='freeshipping']")
        '    Dim ship24Node As HtmlNode = node.SelectSingleNode("p[@class='ships24']")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If shipNode IsNot Nothing Then
        '        freeship = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        '    If ship24Node IsNot Nothing Then
        '        ship = ship24Node.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        'Next

        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//ul[@id='switch-under2-trigger']/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim price As String = node.SelectSingleNode("p[3]/span[1]/span[2]").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[3]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[@class='freeshipping']")
        '    Dim ship24Node As HtmlNode = node.SelectSingleNode("p[@class='ships24']")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If shipNode IsNot Nothing Then
        '        freeship = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        '    If ship24Node IsNot Nothing Then
        '        ship = ship24Node.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        'Next

        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//ul[@id='switch-under3-trigger']/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim price As String = node.SelectSingleNode("p[3]/span[1]/span[2]").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[3]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[@class='freeshipping']")
        '    Dim ship24Node As HtmlNode = node.SelectSingleNode("p[@class='ships24']")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If shipNode IsNot Nothing Then
        '        freeship = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        '    If ship24Node IsNot Nothing Then
        '        ship = ship24Node.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        'Next

        'For Each node As HtmlNode In hd.DocumentNode.SelectNodes("//ul[@id='switch-under4-trigger']/li").ToList()
        '    Dim imgNode As HtmlNode = node.SelectSingleNode("p[1]")
        '    Dim url As String = imgNode.SelectSingleNode("a").GetAttributeValue("href", "")
        '    Dim imgUrl As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("src", "")
        '    Dim imgAlt As String = imgNode.SelectSingleNode("a/img").GetAttributeValue("alt", "")
        '    Dim desc As String = node.SelectSingleNode("p[2]/a").InnerText
        '    Dim price As String = node.SelectSingleNode("p[3]/span[1]/span[2]").InnerText
        '    Dim discount As String = node.SelectSingleNode("p[3]/span[3]").InnerText
        '    Dim shipNode As HtmlNode = node.SelectSingleNode("p[@class='freeshipping']")
        '    Dim ship24Node As HtmlNode = node.SelectSingleNode("p[@class='ships24']")
        '    Dim freeship As String = ""
        '    Dim ship As String = ""
        '    If shipNode IsNot Nothing Then
        '        freeship = shipNode.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        '    If ship24Node IsNot Nothing Then
        '        ship = ship24Node.SelectSingleNode("img[1]").GetAttributeValue("src", "")
        '    End If
        'Next

    End Sub

    ''' <summary>
    ''' 添加数据到Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="section"></param>
    ''' <remarks></remarks>
    Private Shared Sub InsertProductIssue(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal section As String, ByVal listProductId As List(Of Integer))
        'Dim listProductIssue As New List(Of Products_Issue)
        'Dim queryIssue = efContext.Products_Issue
        'For Each q In queryIssue
        '    listProductIssue.Add(New Products_Issue With {.ProductId = q.ProductId, .SiteId = q.SiteId, .IssueDate = q.IssueDate, .SectionID = q.SectionID})
        'Next
        'Dim queryProduct = efContext.Products
        For Each li In listProductId
            Dim pIssue As New Products_Issue()
            pIssue.ProductId = li
            pIssue.SiteId = siteId
            pIssue.IssueID = IssueId
            pIssue.SectionID = section
            'If (JudgeProductIssue(pIssue.ProductId, pIssue.SiteId, pIssue.IssueDate, listProductIssue)) Then
            efContext.AddToProducts_Issue(pIssue)
            'End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 判断将要插入Categories表中URL是否重复，若数据重复，则返回false
    ''' </summary>
    ''' <param name="category"></param>
    ''' <param name="siteId"></param>
    ''' <param name="Url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeCategory(ByVal category As String, ByVal siteId As Integer, ByVal Url As String) As Boolean
        '从数据库读取分类填充到List里面，判断记录是否重复
        Dim queryCategory = From c In efContext.Categories
                          Where c.SiteID = siteId
                          Select c
        Dim list As New List(Of Category)
        For Each q In queryCategory
            list.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next
        For Each li In list
            If (li.Url = Url) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 通过Url返回指定的HtmlDocument对象
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHtmlDocByUrl(ByVal pageUrl As String) As HtmlDocument
        Dim uri As Uri = New Uri(pageUrl)
        Dim url As String = CreateMagicUriString(uri).AbsoluteUri
        'Dim url As String = CreateMagicUriString(pageUrl)
        Dim request As WebRequest = WebRequest.Create(url)
        request.Timeout = 300000
        request.Headers.Add("Accept-Language", "zh-CN")
        Dim noCachePolicy As HttpRequestCachePolicy = New HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore)
        request.CachePolicy = noCachePolicy
        'WebRequest.Create方法，返回WebRequest的子类HttpWebRequest
        Dim response As WebResponse = request.GetResponse()
        'WebRequest.GetResponse方法，返回对 Internet 请求的响应
        Dim resStream As Stream = response.GetResponseStream()
        'WebResponse.GetResponseStream 方法，从 Internet 资源返回数据流。
        Dim pageEncoding As Encoding = Encoding.GetEncoding("gb2312")
        Dim document As New HtmlDocument()
        document.Load(resStream, pageEncoding)
        Return document
    End Function

    Private Shared Function CreateMagicUriString(uri As Uri) As Uri
        Dim ub = New UriBuilder(uri)
        Dim queryToAppend = Guid.NewGuid().ToString()
        If ub.Query.Length > 1 Then
            ub.Query = ub.Query.Substring(1) & "&" & Convert.ToString(queryToAppend)
        Else
            ub.Query = queryToAppend
        End If
        Return ub.Uri
    End Function

    ''' <summary>
    ''' 更新或插入数据到Products表中,并返回添加产品的Id
    ''' </summary>
    ''' <param name="myProduct"></param>
    ''' <param name="url"></param>
    ''' <param name="price"></param>
    ''' <param name="pictureUrl"></param>
    ''' <param name="lastUpdate"></param>
    ''' <param name="description"></param>
    ''' <param name="siteId"></param>
    ''' <param name="currency"></param>
    ''' <param name="pictureAlt"></param>
    ''' <param name="sizeWidth"></param>
    ''' <param name="sizeHeight"></param>
    ''' <param name="categoryId"></param>
    ''' <param name="list"></param>
    ''' <returns>产品的Id</returns>
    ''' <remarks></remarks>
    Public Shared Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                              ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal categoryId As Integer, ByVal discount As Decimal, ByVal freeship As String, ByVal ship24 As String) As Integer
        '获得类别
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.Single()

        '插入产品
        Dim product As New Product()
        product.Prodouct = myProduct
        product.Url = url
        product.Price = price
        product.PictureUrl = pictureUrl
        product.LastUpdate = lastUpdate
        product.Description = description
        product.SiteID = siteId
        product.Currency = currency
        product.PictureAlt = pictureAlt
        product.Discount = discount
        product.FreeShippingImg = freeship
        product.ShipsImg = ship24

        Dim list As List(Of Product) = GetProductList(siteId)
        If (JudgeProduct(product.Prodouct, product.Url, product.Price, product.PictureUrl, product.SiteID, product.Currency, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = (From m In efContext.Products Where m.Url = url And m.SiteID = siteId).Single()
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            updateProduct.Discount = discount
            updateProduct.FreeShippingImg = freeship
            updateProduct.ShipsImg = ship24

            Dim queryCate = From p In efContext.Products
                                Where p.ProdouctID = updateProduct.ProdouctID
                                Select p
            Dim cate = queryCate.Single.Categories
            If Not cate.Contains(category) Then
                category.Products.Add(updateProduct)
            End If
            efContext.SaveChanges()
            Return updateProduct.ProdouctID
        End If
    End Function


    ''' <summary>
    ''' 判断即将插入的数据URL是否在数据库中已经存在，如果存在，返回false
    ''' </summary>
    ''' <param name="product"></param>
    ''' <param name="url"></param>
    ''' <param name="price"></param>
    ''' <param name="pictureUrl"></param>
    ''' <param name="siteId"></param>
    ''' <param name="currency"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeProduct(ByVal product As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal siteId As Integer, ByVal currency As String, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url = url) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 将Products表中需要匹配列的信息添加到list中
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetProductList(ByVal siteId As Integer) As List(Of Product)
        '将Ads表中需要匹配的字段写入list中
        Dim query = From p In efContext.Products
                   Where p.SiteID = siteId
                   Select p
        Dim listProduct As New List(Of Product)
        For Each q In query
            listProduct.Add(New Product With {.Prodouct = q.Prodouct, .Url = q.Url, .Price = q.Price, .PictureUrl = q.PictureUrl, .SiteID = q.SiteID, .Currency = q.Currency})
        Next
        Return listProduct
    End Function

    '处理返回完整的超链接：1.返回图片来源的完整路径
    Private Shared Function PathLink(alink As String) As String
        Return alink.Replace(alink, "http://www.everbuying.com" + alink)
    End Function

    ''' <summary>
    ''' 获取随机数
    ''' </summary>
    ''' <param name="min"></param>
    ''' <param name="max"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetRnd(min As Long, max As Long) As Integer
        Randomize() '没有这个 产生的数会一样
        GetRnd = Rnd() * (max - min + 1) + min
    End Function

    ''' <summary>
    ''' 判断将要插入Ads表中的数据URL是否重复，若数据重复，则返回false
    ''' </summary>
    ''' <param name="ad"></param>
    ''' <param name="pictureUrl"></param>
    ''' <param name="url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function JudgeAds(ByVal ad As String, ByVal pictureUrl As String, ByVal url As String, ByVal list As List(Of Ad)) As Boolean
        For Each li In list
            'If (li.Ad = ad And li.PictureUrl = pictureUrl And li.Url = url) Then
            '2013/3/19，根据URL不同，判断是需要更新还是添加一条新纪录
            If (li.Url = url) Then
                Return False
            End If
        Next
        Return True
    End Function

    '获取产品分类，通过产品链接查找产品所属类别
    Private Shared Function GetCategoryByLink(url As String) As String
        Dim document As HtmlDocument = GetHtmlDocByUrl(url)
        Dim rootNode As HtmlNode = document.DocumentNode
        Return RemoveComma(rootNode.SelectSingleNode("//div[@class='classmenu']/a[2]").InnerText)
    End Function

    '格式化产品分类描述 1.去掉多余空格 2.转化特殊字符&
    Private Shared Function RemoveComma(catename As String) As String
        Return catename.Replace("&amp;", "&").Replace(",", "").Replace("  ", " ")
    End Function

    ''' <summary>
    ''' 写日志
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <remarks></remarks>
    Sub Log(ByVal Ex As Exception)
        Try
            LogText(Now & ", " & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

    Shared Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub

End Class

