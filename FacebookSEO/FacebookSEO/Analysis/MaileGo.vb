
Imports HtmlAgilityPack
Imports System.Math
Imports System.Text.RegularExpressions

Public Class MaileGo
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）
    Private domin As String = "http://www.m6go.com/"
    Dim section As String

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        GetCategory(siteId)
       
        If (planType.Contains("HO3") OrElse planType.Contains("HO4")) Then
            section = "DA"
            GetDaZheProudcts(siteId, IssueID, "mailego", planType, url)
            Dim querySubject = From p In efContext.Products
                    Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                    Where pi.IssueID = IssueID AndAlso pi.SectionID = Section AndAlso pi.SiteId = siteId
                    Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "麦乐购限时抢购！" & querySubject.FirstOrDefault.ToString() & " ...")
        ElseIf (planType.Contains("HO2")) Then
            section = "CA"
            GetBanner(siteId, IssueID, "mailego", planType, url)
            GetOnsaleProudcts(siteId, IssueID, planType, url)
            GetBaokuanProducts(siteId, IssueID, "", planType, url)
            Dim querySubject = From p In efContext.Products
                    Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                    Where pi.IssueID = IssueID AndAlso pi.SectionID = Section AndAlso pi.SiteId = siteId
                    Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "麦乐购最新特卖开抢提醒：" & querySubject.FirstOrDefault.ToString() & " ...")
        ElseIf (planType.Contains("HO1")) Then
            section = "DA"
            GetProducts(siteId, IssueID, "热卖", planType, "http://www.m6go.com/top/")
            GetProducts(siteId, IssueID, "新品", planType, "http://www.m6go.com/newgoods/")
            GetProducts(siteId, IssueID, "收藏", planType, "http://www.m6go.com/favor/")
            GetRepingProducts(siteId, IssueID, "热评", planType, "http://www.m6go.com/top/")
        End If


    End Sub

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String, ByVal url As String)
        Dim htmlstring As String = efHelper.GetHtmlStringByUrl(Url)
        Dim matchRegex As String = "<a.*?href=""(.*?)"".*?background:url\('(http://.*?)'\) "
        Dim mCollection As MatchCollection = Regex.Matches(htmlstring, matchRegex)
        For i As Integer = 0 To mCollection.Count - 1
            If (efHelper.isAdSent(siteId, mCollection(i).Groups(1).Value.ToString.Trim, DateTime.Now.AddDays(-7), DateTime.Now)) Then

                Dim Ad As Ad = New Ad()
                Ad.Url = mCollection(i).Groups(1).Value.ToString.Trim
                Ad.PictureUrl = mCollection(i).Groups(2).Value.ToString.Trim

                Ad.SiteID = siteId
                Ad.Lastupdate = DateTime.Now

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteId)
                Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
                Dim adid As Integer = efHelper.InsertAd(Ad, listAd, categoryId)
                efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)
                Exit For
            End If
        Next

    End Sub

    Public Sub GetDaZheProudcts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = section

        Dim liitemid As New List(Of Integer)
        Dim i As Integer = 0

        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='sale_cont']/ul/li")
        For Each nds In itemNodes

            Dim pro As New Product
            Dim proinfos As HtmlNodeCollection = nds.SelectNodes("tr")
            pro.Prodouct = nds.SelectSingleNode("p[@class='psale03']").InnerText.Trim
            pro.Url = nds.SelectSingleNode("p[@class='psale03']/a").GetAttributeValue("href", "")
            pro.Url = efHelper.addDominForUrl(domin, pro.Url)

            'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
            pro.Description = pro.Prodouct
            pro.LastUpdate = DateTime.Now
            pro.SiteID = siteid
            pro.PictureAlt = pro.Prodouct
            pro.Currency = "¥"

            pro.Discount = Decimal.Parse(nds.SelectSingleNode("div[@class='saleBuy']/div[@class='saleBuy1']/p[@class='p1']").InnerText.Replace("¥", ""))
            pro.Price = Decimal.Parse(nds.SelectSingleNode("div[@class='saleBuy']/div[@class='saleBuy1']/p[@class='p2']/span[@class='span2']").InnerText.Replace("¥", ""))

            pro.PictureUrl = nds.SelectSingleNode("p[@class='psale02']/a/img").GetAttributeValue("src", "")


            Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
            liitemid.Add(id)
            i = i + 1
            If (i >= 12) Then
                Exit For
            End If
        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub

    Public Sub GetOnsaleProudcts(ByVal siteid As Integer, ByVal issueid As Long, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = section

        Dim liitemid As New List(Of Integer)
        Dim i As Integer
        Dim categoryName As String
        Dim tommorrowForecastIndex As Integer
        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='beginLeft']/*")
        For i = itemNodes.Count - 1 To 0 Step -1
            Dim nds As HtmlNode = itemNodes(i)
            Dim pro As New Product
            If (nds.GetAttributeValue("class", "") = "lastTit" AndAlso i <> itemNodes.Count - 1) Then
                categoryName = "明日预告"
                pro = fectchOnsaleProdct(itemNodes(i + 1), siteid)
                Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
                liitemid.Add(id)
            End If
            If (i <= 4 AndAlso nds.GetAttributeValue("class", "") = "beginTit") Then
                categoryName = "最新特卖"
                pro = fectchOnsaleProdct(itemNodes(i), siteid)
                Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
                liitemid.Add(id)
            End If
        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub

    Public Function fectchOnsaleProdct(ByVal aherfNode As HtmlNode, ByVal siteid As Integer) As Product
        Dim pro As New Product
        Dim salesInfoLeft As HtmlNode = aherfNode.SelectSingleNode("div[@class='salesInfo clear']/div[@class='salesInfoLeft']")
        Dim salesInfoRight As HtmlNode = aherfNode.SelectSingleNode("div[@class='salesInfo clear']/div[@class='salesInfoRight']")
        pro.Prodouct = salesInfoLeft.SelectSingleNode("p[@class='bandOnSale']").InnerText.Trim
        pro.Url = aherfNode.GetAttributeValue("href", "")
        pro.Url = efHelper.addDominForUrl(domin, pro.Url)

        'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
        pro.Description = salesInfoLeft.SelectSingleNode("p[@class='countForSale']").InnerText.Trim
        pro.LastUpdate = DateTime.Now
        pro.SiteID = siteid
        pro.PictureAlt = salesInfoLeft.SelectSingleNode("div[@class='LogoforSale']/img").GetAttributeValue("src", "")
        pro.Currency = "¥"

        'pro.Discount = Decimal.Parse(nds.SelectSingleNode("div[@class='saleBuy']/div[@class='saleBuy1']/p[@class='p1']").InnerText.Replace("¥", ""))
        'pro.Price = Decimal.Parse(nds.SelectSingleNode("div[@class='saleBuy']/div[@class='saleBuy1']/p[@class='p2']/span[@class='span2']").InnerText.Replace("¥", ""))

        pro.PictureUrl = salesInfoRight.SelectSingleNode("img").GetAttributeValue("src", "")
        Return pro
    End Function

    Public Sub GetBaokuanProducts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = section

        Dim liitemid As New List(Of Integer)
        Dim i As Integer = 0

        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='rushBuy']/a")
        For Each nds In itemNodes

            Dim pro As New Product
            'Dim proinfos As HtmlNodeCollection = nds.SelectNodes("tr")
            pro.Prodouct = nds.SelectSingleNode("p[@class='rushH']").InnerText.Trim
            pro.Url = nds.GetAttributeValue("href", "")
            pro.Url = efHelper.addDominForUrl(domin, pro.Url)

            'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
            pro.Description = pro.Prodouct
            pro.LastUpdate = DateTime.Now
            pro.SiteID = siteid
            pro.PictureAlt = pro.Prodouct
            pro.Currency = "¥"

            pro.Discount = nds.SelectSingleNode("dl[@class='rushInfo']/dd/p[@class='rushPrice']").InnerText.ToString.Trim.Split("¥")(1)
            'pro.Price = Decimal.Parse(nds.SelectSingleNode("div[@class='saleBuy']/div[@class='saleBuy1']/p[@class='p2']/span[@class='span2']").InnerText.Replace("¥", ""))

            pro.PictureUrl = nds.SelectSingleNode("dl[@class='rushInfo']/dt/img").GetAttributeValue("src", "")


            Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
            liitemid.Add(id)
            i = i + 1
            If (i >= 3) Then
                Exit For
            End If
        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub

    Public Sub GetProducts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = section

        Dim liitemid As New List(Of Integer)
        Dim i As Integer = 0

        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='result hotclicre']/ul/li")
        For Each nds In itemNodes

            Dim pro As New Product
            'Dim proinfos As HtmlNodeCollection = nds.SelectNodes("tr")
            pro.Prodouct = nds.SelectSingleNode("p").InnerText.Trim
            pro.Url = nds.SelectSingleNode("a").GetAttributeValue("href", "")
            pro.Url = efHelper.addDominForUrl(domin, pro.Url)

            'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
            pro.Description = pro.Prodouct
            pro.LastUpdate = DateTime.Now
            pro.SiteID = siteid
            pro.PictureAlt = pro.Prodouct
            pro.Currency = "¥"

            pro.Discount = nds.SelectSingleNode("strong").InnerText.ToString.Trim.Replace("¥", "")
            pro.Price = nds.SelectSingleNode("span").InnerText.ToString.Trim.Replace("¥", "")

            pro.PictureUrl = nds.SelectSingleNode("a/img").GetAttributeValue("src", "")


            Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
            liitemid.Add(id)
            i = i + 1
            If (i >= 3) Then
                Exit For
            End If
        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub

    Public Sub GetRepingProducts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String, ByVal siteurl As String)
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(siteurl, 120000)
        Dim sectionid As String = section

        Dim liitemid As New List(Of Integer)
        Dim i As Integer = 0

        liitemid.Clear()
        Dim itemNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='infoBox hotInfoBox'][2]/div")
        For Each nds In itemNodes

            Dim pro As New Product
            Dim hotL As HtmlNode = nds.SelectSingleNode("div[@class='hotL']")
            Dim hotR As HtmlNode = nds.SelectSingleNode("div[@class='hotR']")
            pro.Prodouct = hotR.SelectSingleNode("p").InnerText.Trim
            pro.Url = hotL.SelectSingleNode("a").GetAttributeValue("href", "")
            pro.Url = efHelper.addDominForUrl(domin, pro.Url)

            'pro.Price = proinfos(2).SelectNodes("td/div")(1).SelectSingleNode("table/tbody/tr/td/span").InnerText.Replace("$", "").Trim
            pro.Description = pro.Prodouct
            pro.LastUpdate = DateTime.Now
            pro.SiteID = siteid
            pro.PictureAlt = pro.Prodouct
            pro.Currency = "¥"

            pro.Discount = hotR.SelectSingleNode("strong").InnerText.ToString.Trim.Split("（")(0).Replace("¥", "")


            pro.PictureUrl = hotL.SelectSingleNode("a/img").GetAttributeValue("src", "")


            Dim id = efHelper.InsertSingleProduct(pro, categoryName, siteid)
            liitemid.Add(id)
            i = i + 1
            If (i >= 3) Then
                Exit For
            End If
        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liitemid, liitemid.Count)
    End Sub

    Public Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "mailego"
        myCategory.Description = "mailego"
        myCategory.Url = ""
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)

        Dim myCategory1 As New Category
        myCategory1.Category1 = "最新特卖"
        myCategory1.Description = "最新特卖"
        myCategory1.Url = "http://www.m6go.com/onsale"
        myCategory1.SiteID = siteId
        myCategory1.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory1, siteId)

        Dim myCategory2 As New Category
        myCategory2.Category1 = "爆款限时抢"
        myCategory2.Description = "爆款限时抢"
        myCategory2.Url = "http://www.m6go.com/dazhe/"
        myCategory2.SiteID = siteId
        myCategory2.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory2, siteId)

        Dim myCategory3 As New Category
        myCategory3.Category1 = "明日预告"
        myCategory3.Description = "明日预告"
        myCategory3.Url = "http://www.m6go.com/onsale/?date=tommorrow"
        myCategory3.SiteID = siteId
        myCategory3.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory3, siteId)

        Dim myCategory5 As New Category
        myCategory5.Category1 = "热卖"
        myCategory5.Description = "热卖"
        myCategory5.Url = "http://www.m6go.com/top/"
        myCategory5.SiteID = siteId
        myCategory5.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory5, siteId)

        Dim myCategory6 As New Category
        myCategory6.Category1 = "新品"
        myCategory6.Description = "新品"
        myCategory6.Url = "http://www.m6go.com/newgoods/"
        myCategory6.SiteID = siteId
        myCategory6.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory6, siteId)

        Dim myCategory4 As New Category
        myCategory4.Category1 = "收藏"
        myCategory4.Description = "收藏"
        myCategory4.Url = "http://www.m6go.com/favor/"
        myCategory4.SiteID = siteId
        myCategory4.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory4, siteId)
        'Dim titleCategory As New Category
        'titleCategory.Category1 = "title"
        'titleCategory.Description = "title"
        'titleCategory.Url = "title"
        'titleCategory.SiteID = siteId
        'titleCategory.LastUpdate = lastUpdate
        'helper.InsertOrUpdateCate(titleCategory, siteId)
    End Sub


End Class
