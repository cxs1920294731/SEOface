Imports HtmlAgilityPack
Imports System.Math
Imports System.Text.RegularExpressions
Public Class jiacheng
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）
    Private domin As String = "http://www.e-lining.com/"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        GetCategory(siteId)
        Dim nexturl As String = ""
        Dim infourl As String = ""
        Dim bannerimgurl As String = ""
        ''GetBanner，固定用Banner Image:http://www.h5deal.com/gache2010/edm/A.jpg, 对应的URL：http://stores.ebay.com/astromarket99
        ''GetProduct 获取8个产品，如不够8个则只获取所展示的。无需判断重复，需要分析的URL为：http://stores.ebay.com/astromarket99/NEW-ARRIVAL-/_i.html?rt=nc&_dmd=2&_fsub=9850103012&_sid=1199926972&_trksid=p4634.c0.m14&_vc=1
        ''GetSubject，规则：Hi [FIRSTNAME],New Arrival for You：[第一个产品名称]
        ''更新下一次要发送的URL='http://stores.ebay.com/h5deal' 回AutomationPlan表里的URL字段

        Select Case url
            Case "http://stores.ebay.com/astromarket99"
                nexturl = "http://stores.ebay.com/h5deal"
                infourl = "http://stores.ebay.com/astromarket99/NEW-ARRIVAL-/_i.html?rt=nc&_dmd=2&_fsub=9850103012&_sid=1199926972&_trksid=p4634.c0.m14&_vc=1"
                bannerimgurl = "http://www.h5deal.com/gache2010/edm/A.jpg"

            Case "http://stores.ebay.com/h5deal"
                nexturl = "http://stores.ebay.com/gachemart"
                infourl = "http://stores.ebay.com/h5deal/New-Arrival-/_i.html?_fsub=6743972013&_sid=1019325113&_trksid=p4634.c0.m322"
                bannerimgurl = "http://www.h5deal.com/gache2010/edm/H.jpg"

            Case "http://stores.ebay.com/gachemart"
                nexturl = "http://stores.ebay.com/stage6store"
                infourl = "http://stores.ebay.com/gachemart/NEW-ARRIVAL-/_i.html?_fsub=4705185016&_sid=1098731856&_trksid=p4634.c0.m322"
                bannerimgurl = "http://www.h5deal.com/gache2010/edm/G.jpg"

            Case "http://stores.ebay.com/stage6store"
                nexturl = "http://stores.ebay.com/astromarket99"
                infourl = "http://stores.ebay.com/stage6store/NEW-ARRIVAL-/_i.html?_fsub=6487089015&_sid=1011196355&_trksid=p4634.c0.m322"
                bannerimgurl = "http://www.h5deal.com/gache2010/edm/S.jpg"
        End Select
        '"New Arrivals作为一个产品，以便能根据不同的店铺同步模板中的New Arrivals链接"
        Dim product As New Product()
        product.Url = infourl
        product.Prodouct = "New Arrivals"
        product.SiteID = siteId
        product.LastUpdate = Now
        product.Description = "New Arrivals"
        Dim listProduct As New List(Of Product)
        listProduct = efHelper.GetProductList(siteId)
        Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, "title")
        Dim returnId As Integer = efHelper.InsertProduct(product, Now, categoryId, listProduct)
        listProductId.Add(returnId)
        efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProductId, 1)

        GetBanner(siteId, IssueID, "xinpin", planType, bannerimgurl, infourl)
        GetProudcts(siteId, IssueID, "xinpin", planType, infourl)

        efHelper.InsertIssueSubject(IssueID, "Hi [FIRSTNAME],New Arrival for You：" & efHelper.GetFirstProductSubject(IssueID, "NE") & " ...")
        efHelper.updateAutomationUrl(siteId, nexturl)
    End Sub

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String, ByVal bannerimgurl As String, ByVal storeUrl As String)
        Dim Ad As Ad = New Ad()
        Ad.Url = storeUrl
        Ad.PictureUrl = bannerimgurl

        Ad.SiteID = siteId
        Ad.Lastupdate = DateTime.Now

        Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
        listAd = efHelper.GetListAd(siteId)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
        Dim adid As Integer = efHelper.InsertAd(Ad, listAd, categoryId)
        efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)
    End Sub

    Public Sub GetProudcts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = "NE"
       
        Dim liitemid As New List(Of Integer)
        Dim i As Integer = 0
        'If (siteurl = "http://stores.ebay.com/ultrapire/NEW-ARRIVAL-/_i.html?_fsub=6605456010") Then
        '    Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='gallery-grid clr']/div")
        '    liitemid.Clear()
        '    i = 0
        '    For Each node In itemNodes
        '        Dim pro As New Product
        '        pro.Prodouct = node.SelectSingleNode("div[@class='desc']/a").GetAttributeValue("title", "")
        '        pro.Url = node.SelectSingleNode("div[@class='desc']/a").GetAttributeValue("href", "")
        '        Dim p As String = Double.Parse(node.SelectSingleNode("div[@class='price']/span[@class='curr']").InnerText.Trim)
        '        If p.Contains("HKD") Then
        '            pro.Price = p.Replace("HKD", "").Trim
        '            pro.Currency = "HKD"
        '        ElseIf p.Contains("RMB") Then
        '            pro.Price = p.Replace("RMB", "").Trim
        '            pro.Currency = "RMB"
        '        ElseIf p.Contains("$") Then
        '            pro.Price = p.Replace("$", "").Trim
        '            pro.Currency = "$"
        '        End If

        '        pro.PictureUrl = node.SelectSingleNode("div[@class='img-cell']/a/img").GetAttributeValue("src", "")
        '        pro.LastUpdate = DateTime.Now
        '        pro.Description = pro.Prodouct
        '        pro.SiteID = siteid
        '        pro.PictureAlt = pro.Prodouct
        '        Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
        '        liitemid.Add(id)
        '        i = i + 1
        '        If (i >= 8) Then
        '            Exit For
        '        End If
        '    Next
        'Else
        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='pview rs-pview']/table/tr/td/table")
        If Not (itemNodes Is Nothing) Then
            For Each nds In itemNodes

                Dim pro As New Product
                Dim proinfos As HtmlNodeCollection = nds.SelectNodes("tr")
                pro.Prodouct = proinfos(2).SelectNodes("td/div")(0).InnerText.Trim
                pro.Url = proinfos(2).SelectNodes("td/div")(0).SelectSingleNode("a").GetAttributeValue("href", "")
                Dim p As String = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tr/td/span").InnerText.Trim
                'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
                pro.Description = pro.Prodouct
                pro.LastUpdate = DateTime.Now
                pro.SiteID = siteid
                pro.PictureAlt = pro.Prodouct
                If p.Contains("HKD") Then
                    pro.Price = p.Replace("HKD", "").Trim
                    pro.Currency = "HKD"
                ElseIf p.Contains("RMB") Then
                    pro.Price = p.Replace("RMB", "").Trim
                    pro.Currency = "RMB"
                ElseIf p.Contains("$") Then
                    pro.Price = p.Replace("$", "").Trim
                    pro.Currency = "$"
                End If
                Try
                    Dim imgDocument As HtmlDocument = efHelper.GetHtmlDocument(pro.Url, 120000)
                    Dim imgNode As HtmlNode = imgDocument.DocumentNode.SelectSingleNode("//img[@id='icImg']")
                    pro.PictureUrl = imgNode.GetAttributeValue("src", "")
                Catch ex As Exception
                    pro.PictureUrl = proinfos(0).SelectSingleNode("td/div/a/img").GetAttributeValue("src", "")
                End Try

                Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
                liitemid.Add(id)
                i = i + 1
                If (i >= 8) Then
                    Exit For
                End If
            Next

        End If
        'End If
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub


    Public Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "xinpin"
        myCategory.Description = "xinpin"
        myCategory.Url = ""
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim titleCategory As New Category
        titleCategory.Category1 = "title"
        titleCategory.Description = "title"
        titleCategory.Url = "title"
        titleCategory.SiteID = siteId
        titleCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(titleCategory, siteId)
    End Sub
End Class
