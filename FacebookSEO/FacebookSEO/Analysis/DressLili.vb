Imports HtmlAgilityPack
Imports System.Net
Imports System.IO
Imports System.Text
Imports Enumerable = System.Linq.Enumerable

Public Class DressLily
    Private htmlUrl As String = "http://www.dresslily.com"
    Private menShoesUrl As String = "http://www.dresslily.com/men-s-shoes-c-82.html?page_size=24"
    Private hd As HtmlDocument = GetHtmlDocByUrl(htmlUrl, 60000)
    Private efContext As New EmailAlerterEntities()
    Private listProduct As New List(Of Product)
    Private listCategoryName As New HashSet(Of String)
    Private helper As New EFHelper
    Private listProductUrl As New List(Of String)

    ''' <summary>
    ''' 程序入口函数
    ''' </summary>
    ''' <param name="IssueID"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <param name="splitContactCount"></param>
    ''' <param name="spreadLogin"></param>
    ''' <param name="appId"></param>
    ''' <param name="url"></param>
    ''' <remarks></remarks>
    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                     ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String)
        'Dim listName As String = "General"
        '2013/4/27新增，为增加Men和Women两个分类判断使用
        'Dim helper As New EFHelper()

        Try
            listProductUrl = helper.GetListProduct(siteId)

            listProduct = GetProductList(siteId)
            GetCategory(siteId, IssueID) 'by  Dora 注释掉2013-11-22 因为Category获取时出错，先注释掉
            'GetAds(siteId, IssueID) 'Promotion处图片 
            Dim updateTime As DateTime = Now

            'GetDailyDeal(siteId, IssueID) '2013/4/26修改，DailyDeal改版

            '2013/05/21使用特定Url，开始
            'GetNewArrivals(siteId, IssueID)
            'GetHotSales(siteId, IssueID)
            ''2013/4/11新增，获取页面Sexy Lingerie的产品并添加到Products表和Products_Issue表中
            'Dim sexyUrl As String = "http://www.dresslily.com/sexy-lingerie-b-97.html"
            ''Dim mySexy As New DressLilyWomen()
            ''mySexy.GetProductAndProductIssue(sexyUrl, 8, "Sexy Lingerie", siteId, IssueID, Now)
            'GetProductAndProductIssue(sexyUrl, 8, "Sexy Lingerie", siteId, IssueID, Now)
            'GetMenShoes(siteId, IssueID)
            '2013/05/21使用特定Url，结束

            '2013/05/21开始获取产品，开始
            Try 'New Arrival
                Dim newArrivalUrl As String = "http://www.dresslily.com/" 'Dora2013-11-22 dressLili 
                Dim maxProductCount As Integer = 16
                Dim iTakeCount As Integer = 8
                Dim iProIssueCount As Integer = 4
                GetNewArrivals(siteId, IssueID) 'Dora overload,so go to the other one
            Catch ex As Exception
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
                'LogText(ex.ToString())
            End Try
            Try 'Women's Clothing
                Dim womenClothingUrl As String = "http://www.dresslily.com/women-s-clothing-b-1.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, womenClothingUrl, "Women's Clothing", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            Try 'Women's Bags
                Dim menClothingUrl As String = "http://www.dresslily.com/bags-b-3.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, menClothingUrl, "Bags", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            Try 'Women's Shoes '2013/06/04新增
                Dim shoeUrl As String = "http://www.dresslily.com/women-s-shoes-c-81.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, shoeUrl, "Women's Shoes", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            Try 'Men's Watches
                Dim menShoeUrl As String = "http://www.dresslily.com/men-s-watches-c-125.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, menShoeUrl, "Men's Watches", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            Try 'Jewelry
                Dim jewelryUrl As String = "http://www.dresslily.com/jewelry-b-56.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, jewelryUrl, "Jewelry", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            Try 'Sexy Lingerie
                Dim sexUrl As String = "http://www.dresslily.com/sexy-lingerie-b-97.html?odr=new"
                Dim maxProductCount As Integer = 30
                Dim iTakeCount As Integer = 16
                Dim iProIssueCount As Integer = 8
                GetProducts(siteId, IssueID, sexUrl, "Sexy Lingerie", maxProductCount, iTakeCount, iProIssueCount, updateTime)
            Catch ex As Exception
                'LogText(ex.ToString())
                '2013/07/19 added, throw exception to outer layer
                Throw New Exception(ex.ToString())
            End Try
            '2013/05/21开始获取产品，结束

            AddIssueSubject(IssueID, siteId, planType)
            CreateContacts(siteId, planType, spreadLogin, appId, splitContactCount, IssueID)
        Catch ex As Exception
            '2013/07/19 added, throw exception to outer layer
            Throw New Exception(ex.ToString())
        End Try

        'Dim helper As New EFHelper() '2013/4/15新增
        'helper.InsertInbox(IssueID)
    End Sub

    ''' <summary>
    ''' 获取Category内容，并写入到Categories表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <remarks></remarks>
    Public Sub GetCategory(ByVal siteId As Integer, ByVal IssueId As Integer)
        '将Ads表中需要匹配的字段写入list中()
        Dim queryCategory = From c In efContext.Categories
                            Where c.SiteID = siteId
                            Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
            listCategoryName.Add(q.Category1) '2013/4/27新增
        Next
        Dim liNodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='n_nav']/ul/li") '''Dora daleilist is not exists.
        Dim nowTime As DateTime = Now
        For Each li In liNodes
            Dim cate As New Category()
            Dim hrefNode As HtmlNode = li.SelectSingleNode("a")
            cate.Category1 = hrefNode.InnerText

            If (cate.Category1.Trim().ToUpper().Contains("LINGERIE")) Then
                cate.Category1 = "Sexy Lingerie"
            End If
            'Dora 2013-11-29 改版，商品类别更新
            'If (cate.Category1 = "Women's Clothing" OrElse cate.Category1 = "Women's Clothing" OrElse _
            '    cate.Category1 = "Lingerie" OrElse cate.Category1 = "Clearance" OrElse _
            '    cate.Category1 = "Fashion Accessories" OrElse cate.Category1 = "Jewelry") Then
            '    cate.Gender = "W"  'Women
            'ElseIf (cate.Category1 = "Men's Clothing") Then
            '    cate.Gender = "M"  'Men
            'Else
            '    cate.Gender = "N" '性别不清
            'End If
            'Dora 2013-11-29 改版，商品类别更新
            cate.SiteID = siteId
            cate.LastUpdate = nowTime
            cate.Description = hrefNode.InnerText
            cate.Url = ParseLink(hrefNode.GetAttributeValue("href", "")).Trim()
            cate.ParentID = -1
            If (JudgeCategory(cate.Category1, cate.SiteID, cate.Url, listCategory)) Then
                efContext.AddToCategories(cate)
            Else
                Dim updateCate = efContext.Categories.Single(Function(m) m.Url = cate.Url)
                updateCate.Category1 = cate.Category1
                updateCate.LastUpdate = nowTime
                updateCate.Description = cate.Description
            End If
        Next

        '添加Women's Shoes,Men's Shoes,Women's Dresses作为大分类
        Dim cateWomen As New Category()
        Dim cateMen As New Category()
        Dim cateWomenDress As New Category()
        cateWomen.Category1 = "Women's Shoes"
        cateMen.Category1 = "Men's Shoes"
        cateWomenDress.Category1 = "Women's Dresses"
        cateWomen.SiteID = siteId
        cateMen.SiteID = siteId
        cateWomenDress.SiteID = siteId
        cateWomen.Url = "http://www.dresslily.com/women-s-shoes-c-81.html"
        cateMen.Url = "http://www.dresslily.com/men-s-shoes-c-82.html"
        cateWomenDress.Url = "http://www.dresslily.com/Dresses-c-6.html"
        cateWomen.LastUpdate = nowTime
        cateMen.LastUpdate = nowTime
        cateWomenDress.LastUpdate = nowTime
        cateWomen.Description = "Women's Shoes"
        cateMen.Description = "Men's Shoes"
        cateWomenDress.Description = "Women's Dresses"
        cateWomen.ParentID = -1
        cateMen.ParentID = -1
        cateWomenDress.ParentID = -1
        cateWomen.Gender = "W"   'Women
        cateMen.Gender = "M"   'Men
        cateWomenDress.Gender = "W"  'Women
        If (JudgeCategory(cateWomen.Category1, cateWomen.SiteID, cateWomen.Url, listCategory)) Then
            efContext.AddToCategories(cateWomen)
        End If
        If (JudgeCategory(cateMen.Category1, cateMen.SiteID, cateMen.Url, listCategory)) Then
            efContext.AddToCategories(cateMen)
        End If
        If (JudgeCategory(cateWomenDress.Category1, cateWomenDress.SiteID, cateWomenDress.Url, listCategory)) Then
            efContext.AddToCategories(cateWomenDress)
        End If

        '2013/4/27新增，增加Promotion的分类
        'If (listCategoryName.Contains("Men")) Then
        '    Dim cate1 As New Category()
        '    cate1.Category1 = "Men"
        '    cate1.LastUpdate = nowTime
        '    cate1.Description = "Men"
        '    cate1.Gender = "M"
        '    efContext.AddToCategories(cate1)
        'End If
        'If (listCategoryName.Contains("Women")) Then
        '    Dim cate2 As New Category()
        '    cate2.Category1 = "Women"
        '    cate2.LastUpdate = nowTime
        '    cate2.Description = "Women"
        '    cate2.Gender = "W"
        '    efContext.AddToCategories(cate2)
        'End If

        '增加Men大分类
        Dim counter1 As Integer = 0
        For Each strs In listCategoryName
            If (strs = "Men") Then
                Exit For
            End If
            counter1 = counter1 + 1
        Next
        If (counter1 = listCategoryName.Count) Then
            Dim cate1 As New Category()
            cate1.Category1 = "Men"
            cate1.LastUpdate = nowTime
            cate1.Description = "Men"
            cate1.Gender = "M"
            cate1.SiteID = siteId
            efContext.AddToCategories(cate1)
        End If
        '增加Women大分类
        Dim counter2 As Integer = 0
        For Each strs In listCategoryName
            If (strs = "Women") Then
                Exit For
            End If
            counter2 = counter2 + 1
        Next
        If (counter2 = listCategoryName.Count) Then
            Dim cate2 As New Category()
            cate2.Category1 = "Women"
            cate2.LastUpdate = nowTime
            cate2.Description = "Women"
            cate2.Gender = "W"
            cate2.SiteID = siteId
            efContext.AddToCategories(cate2)
        End If

        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 获取Promotion的信息，并写入Ads表中,每期只添加一条记录
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <remarks></remarks>
    Public Sub GetAds(ByVal siteId As Integer, ByVal IssueId As Integer)
        Try
            '将Ads表中需要匹配的字段写入list中
            Dim query = From ad1 In efContext.Ads
                       Where ad1.SiteID = siteId
                       Select ad1
            Dim listAd As New List(Of Ad)
            For Each q In query
                listAd.Add(New Ad With {.Ad1 = q.Ad1, .PictureUrl = q.PictureUrl, .Url = q.Url})
            Next
            Dim listAdId As New List(Of Integer) ''2013/3/19添加

            '2013/4/26新增
            Dim listMenAdId As New HashSet(Of Integer)
            Dim listWomenAdId As New HashSet(Of Integer)
            Dim queryMenCategory = efContext.Categories.Single(Function(c) c.Category1 = "Men" AndAlso c.SiteID = siteId)
            Dim queryWomenCategory = efContext.Categories.Single(Function(c) c.Category1 = "Women" AndAlso c.SiteID = siteId)
            For Each m In queryMenCategory.Ads
                listMenAdId.Add(m.AdID)
            Next
            For Each m In queryWomenCategory.Ads
                listWomenAdId.Add(m.AdID)
            Next

            ''遍历每个节点信息()
            Dim liNodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//ul[@id='idSlider2']/li")
            Dim nowTime As DateTime = Now
            For Each li In liNodes
                Dim ad As New Ad()
                Dim linkNode As HtmlNode = li.SelectSingleNode("a")
                Dim imgNode As HtmlNode = li.SelectSingleNode("a/img")
                ad.Ad1 = imgNode.GetAttributeValue("alt", "")
                ad.Url = ParseLink(linkNode.GetAttributeValue("href", "")).Trim()

                '2013/4/8添加，防止出现异常图片
                ad.PictureUrl = imgNode.GetAttributeValue("src", "")
                If (imgNode.GetAttributeValue("src", "").Contains("loading.gif")) Then
                    ad.PictureUrl = imgNode.GetAttributeValue("data-src", "")
                End If

                ad.Description = imgNode.GetAttributeValue("alt", "")
                ad.SiteID = siteId
                ad.Type = "P"
                ad.Lastupdate = nowTime
                If (JudgeAds(ad.Ad1, ad.PictureUrl, ad.Url, listAd)) Then
                    '2013/4/26新增，增加一个关系表AdsCategory
                    If (ad.Ad1.Contains("Men")) Then
                        ad.Categories.Add(queryMenCategory)
                    Else
                        ad.Categories.Add(queryWomenCategory)
                    End If

                    efContext.AddToAds(ad)
                    efContext.SaveChanges()
                    listAdId.Add(ad.AdID) ''2013/3/19添加

                Else
                    Dim updateCate = efContext.Ads.Single(Function(m) m.Url = ad.Url)
                    updateCate.Ad1 = ad.Ad1
                    updateCate.PictureUrl = ad.PictureUrl
                    updateCate.Description = ad.Description
                    updateCate.Lastupdate = nowTime
                    efContext.SaveChanges()
                    listAdId.Add(updateCate.AdID) ''2013/3/19添加

                    '2013/4/26添加，增加AdsCategory表
                    If (ad.Ad1.Contains("Men")) Then
                        If Not (listMenAdId.Contains(updateCate.AdID)) Then
                            queryMenCategory.Ads.Add(updateCate)
                        End If
                    Else
                        If Not (listWomenAdId.Contains(updateCate.AdID)) Then
                            queryWomenCategory.Ads.Add(updateCate)
                        End If
                    End If

                End If
            Next
            '2013/4/3添加,防止每次发送的Subject都相同处理
            Dim adId As Integer
            Dim queryAdsIssue = From adIss In efContext.Ads_Issue
                                Join iss In efContext.Issues On adIss.IssueID Equals iss.IssueID
                                Where adIss.SiteId = siteId AndAlso Not iss.Subject = ""
                                Order By iss.IssueID Descending
                                Select adIss
            Dim counter As Integer = 1
            Dim takeCounter As Integer = liNodes.Count - 1
            Dim listIssueId As New List(Of Integer)
            For Each q In queryAdsIssue
                If (counter <= takeCounter) Then
                    listIssueId.Add(q.AdId)
                    counter = counter + 1
                End If
            Next
            For Each li In listAdId
                If (listIssueId.Count = 0) OrElse (Not listIssueId.Contains(li)) Then
                    adId = li
                    Exit For
                End If
            Next

            'Dim Generator As System.Random = New System.Random()
            'Dim Min As Integer = listAdId.Min()
            'Dim Max As Integer = listAdId.Max()
            'Dim adId As Integer = Generator.Next(Min, Max)
            'For Each li In listAdId
            '    If Not (adId = li) Then
            '        adId = Generator.Next(Min, Max)
            '    End If
            'Next

            Dim adIssue As New Ads_Issue()
            adIssue.AdId = adId
            adIssue.IssueID = IssueId
            adIssue.SiteId = siteId
            efContext.AddToAds_Issue(adIssue)
            efContext.SaveChanges()
        Catch ex As Exception
            LogText(ex.ToString())
        End Try


    End Sub

    ''' <summary>
    ''' 获取Daily Deals的信息，并写入Products表中，同时生成Products_Issue表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <remarks></remarks>
    Private Sub GetDailyDeal(ByVal siteId As Integer, ByVal IssueId As Integer)
        Dim divNodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='dailyc']/div")
        Dim listProductId As New List(Of Integer)
        Dim nowTime As DateTime = DateTime.Now
        For i As Integer = 1 To divNodes.Count - 1
            Dim hrefNode As HtmlNode = divNodes(i).SelectNodes("div")(0).SelectSingleNode("a")
            Dim picNode As HtmlNode = hrefNode.SelectSingleNode("img")
            Dim rightDiv As HtmlNode = divNodes(i).SelectNodes("div")(1)
            Dim product As String = rightDiv.SelectSingleNode("h2/a").InnerHtml
            Dim url As String = ParseLink(hrefNode.GetAttributeValue("href", "")).Trim()
            Dim price As Decimal = Convert.ToDecimal(rightDiv.SelectNodes("p/span")(1).InnerText)

            '2013/4/8添加，防止出现异常图片
            Dim pictureUrl As String = picNode.GetAttributeValue("src", "")
            If (picNode.GetAttributeValue("src", "").Contains("loading.gif")) Then
                pictureUrl = picNode.GetAttributeValue("data-src", "").Trim()
            End If

            Dim lastUpdate As DateTime = nowTime
            Dim description As String = divNodes(i).SelectNodes("div")(1).SelectNodes("p/span")(2).InnerText.Trim()
            Dim currency As String = rightDiv.SelectNodes("p/span")(0).InnerText.Trim()
            Dim pictureAlt As String = picNode.GetAttributeValue("alt", "")
            Dim sizeWidth As Integer = Convert.ToInt32(picNode.GetAttributeValue("width", "").Trim())
            Dim sizeHeight As Integer = Convert.ToInt32(picNode.GetAttributeValue("height", "").Trim())
            Dim categoryId As Integer = InsertCategory(url, siteId, IssueId)
            'Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, sizeWidth, sizeHeight, categoryId, listProduct)
            'listProductId.Add(returnId)
        Next
        InsertProductIssue(siteId, IssueId, "DA", listProductId)
    End Sub

    ''' <summary>
    ''' 获取New Arrivals信息，并添加到Products表，同时生成Products_Issue表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <remarks></remarks>
    Private Sub GetNewArrivals(ByVal siteId As Integer, ByVal IssueId As Integer)
        Dim listProductId As New List(Of Integer)
        Dim liNodes As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='n_proListWrap']/ul[2]/li")
        Dim nowTime As DateTime = Now
        For Each li In liNodes
            Dim divNodes As HtmlNodeCollection = li.ChildNodes 'li 下面所有的子节点，包括<a/>，<p/>等等
            Dim imgNode As HtmlNode = li.SelectSingleNode("a")
            Dim titleNode As HtmlNode = li.SelectSingleNode("p[@class='n_proName']/a")
            Dim moneyNode As HtmlNodeCollection = li.SelectNodes("p[@class='pt5 n_price']/span")
            Dim freeShipImgNode As HtmlNode
            Dim shipsImgNode As HtmlNode
            Dim freeShipImg As String = ""
            Dim shipsImg As String = ""

            '2013/05/18添加，网站改版
            'If Not (li.SelectSingleNode("div[3]").InnerText.Contains("USD")) Then
            '    moneyNode = li.SelectNodes("div[4]/span/span/span")
            'End If
            If (divNodes.Count = 9) Then
                Dim myImgNode As HtmlNode = li.SelectSingleNode("p[2]/img")
                If (myImgNode.GetAttributeValue("src", "").Contains("icon_ships24.gif")) Then
                    shipsImgNode = myImgNode
                    shipsImg = shipsImgNode.GetAttributeValue("src", "")
                ElseIf (myImgNode.GetAttributeValue("src", "").Contains("freepic.gif")) Then
                    freeShipImgNode = myImgNode
                    freeShipImg = freeShipImgNode.GetAttributeValue("src", "")
                End If
                moneyNode = li.SelectNodes("p[3]/span")
            ElseIf (divNodes.Count = 11) Then
                Dim myImgNode As HtmlNode = li.SelectSingleNode("p[2]/img")

                If (myImgNode.GetAttributeValue("src", "").Contains("freepic.gif")) Then
                    freeShipImgNode = myImgNode
                    shipsImgNode = li.SelectSingleNode("p[3]/img")
                    freeShipImg = freeShipImgNode.GetAttributeValue("src", "")
                    shipsImg = shipsImgNode.GetAttributeValue("src", "")
                ElseIf (myImgNode.GetAttributeValue("src", "").Contains("icon_ships24.gif")) Then
                    shipsImgNode = myImgNode
                    freeShipImgNode = li.SelectSingleNode("p[3]/img")
                    freeShipImg = freeShipImgNode.GetAttributeValue("src", "")
                    shipsImg = shipsImgNode.GetAttributeValue("src", "")
                End If
                moneyNode = li.SelectNodes("p[4]/span")
            End If

            Dim product As String = titleNode.InnerText
            Dim url As String = ParseLink(titleNode.GetAttributeValue("href", "")).Trim()

            '2013/4/16修改，版面修改
            Dim price As Decimal = Decimal.Parse(moneyNode(0).InnerText.Trim("$").Trim("D"))

            '2013/4/8添加
            Dim pictureUrl As String = imgNode.SelectSingleNode("img").GetAttributeValue("src", "")
            If (imgNode.SelectSingleNode("img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                pictureUrl = imgNode.SelectSingleNode("img").GetAttributeValue("data-src", "")
            End If

            Dim lastUpdate As DateTime = nowTime
            Dim description As String = titleNode.InnerText

            '2013/4/16修改，版面修改
            Dim currency As String = "USD" 'moneyNode(0).InnerText.Trim() 'Mid(moneyNode.InnerText.Trim(), 1, 3)'Dora 因为获取很复杂，所有直接赋值

            Dim pictureAlt As String = imgNode.SelectSingleNode("img").GetAttributeValue("alt", "")
            Dim sizeWidth As Integer = 150
            Dim sizeHeight As Integer = 150
            Dim categoryId As Integer = InsertCategory(url, siteId, IssueId)
            Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, sizeWidth, sizeHeight, categoryId, listProduct, freeShipImg, shipsImg)
            listProductId.Add(returnId)
        Next
        efContext.SaveChanges()
        InsertProductIssue(siteId, IssueId, "NE", listProductId)
    End Sub

    ''' <summary>
    ''' 获取制定NewArrival页面的信息，并添加到Products表，同时生成Products_Issue表中的数据，
    ''' 每次轮番更替产品
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="newArrivalURL">http://www.dresslily.com/</param>'Dora2013-11-22 dressLili 
    ''' <param name="maxCount">获取页面产品的个数</param>
    ''' <remarks></remarks>
    Private Sub GetNewArrivals(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal newArrivalURL As String, ByVal maxCount As Integer, _
                               ByVal updateTime As DateTime, ByVal iTakeCount As Integer, ByVal iProIssueCount As Integer)
        'Dim pageUrl As String = "http://www.dresslily.com/"'Dora2013-11-22 dressLili 
        'Dim productCount As Integer = 30
        Dim listProducts As List(Of Product) = GetListProducts(newArrivalURL, maxCount, siteId, updateTime)
        Dim listProductId As New List(Of Integer)
        For Each li In listProducts
            Dim categoryId As Integer = InsertCategory(li.Url, siteId, IssueId)
            Dim returnId As Integer = InsertProduct(li.Prodouct, li.Url.Trim(), li.Price, li.Discount, li.PictureUrl, Now, li.Description, li.SiteID, li.Currency, li.PictureAlt, li.SizeWidth, li.SizeHeight, categoryId, listProduct, li.FreeShippingImg, li.ShipsImg)
            listProductId.Add(returnId)
        Next
        InsertProductsIssue(siteId, IssueId, "NE", listProductId, iTakeCount, iProIssueCount)
    End Sub

    ''' <summary>
    ''' 获取某个特定URL分类的产品信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="categoryUrl"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="iMaxCount">从页面中获取产品的个数</param>
    ''' <param name="iTakeCount">获取前几期Products_Issue表产品个数</param>
    ''' <param name="iProIssueCount">本期写入Products_Issue表中产品个数</param>
    ''' <param name="updateTime"></param>
    ''' <remarks></remarks>
    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, _
                            ByVal categoryName As String, ByVal iMaxCount As Integer, ByVal iTakeCount As Integer, _
                            ByVal iProIssueCount As Integer, ByVal updateTime As DateTime)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, iMaxCount, siteId, updateTime)
            Dim listProductId As New List(Of Integer)
            Dim categoryId As Integer = GetCategoryId(siteId, categoryName)
            For Each li In listProducts
                Dim returnId As Integer = InsertProduct(li.Prodouct, li.Url, li.Price, li.Discount, li.PictureUrl, Now, li.Description, li.SiteID, li.Currency, li.PictureAlt, li.SizeWidth, li.SizeHeight, categoryId, listProduct, li.FreeShippingImg, li.ShipsImg)
                listProductId.Add(returnId)
            Next
            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iTakeCount, iProIssueCount, categoryName)
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 获取某个店铺的一个CategoryId
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCategoryId(ByVal siteId As Integer, ByVal categoryName As String) As Integer
        Dim queryCategory = From c In efContext.Categories
                                  Where c.SiteID = siteId AndAlso c.Category1.Trim.ToUpper = categoryName.ToUpper
                                  Select c
        Dim categoryId As Integer = 0
        For Each q In queryCategory
            categoryId = q.CategoryID
        Next
        Return categoryId
    End Function

    ''' <summary>
    ''' 轮番发布特定URL页面的产品信息
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="listProductId"></param>
    ''' <param name="iTakeCount"></param>
    ''' <param name="iProIssueCount"></param>
    ''' <param name="categoryName"></param>
    ''' <remarks></remarks>
    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer), _
                                    ByVal iTakeCount As Integer, ByVal iProIssueCount As Integer, ByVal categoryName As String)
        Dim queryLastProductId = From p In efContext.Products
                                Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                               Where pi.IssueID < issueId AndAlso pi.SiteId = siteId AndAlso pi.SectionID = sectionId
                               Order By pi.IssueID Descending
                               Select p
        Dim lastProductsId As New List(Of Integer)
        Dim j As Integer = 0

        'get the ProductId of the same CategoryName in past time
        Dim queryCategory = efContext.Categories.Where(Function(c) c.Category1 = categoryName AndAlso c.SiteID = siteId).FirstOrDefault
        For Each query In queryLastProductId
            If (query.Categories.Contains(queryCategory) AndAlso j < iTakeCount) Then
                lastProductsId.Add(query.ProdouctID)
                j = j + 1
            End If
            If (j >= iTakeCount) Then
                Exit For
            End If
        Next

        Dim queryProductId = (From pro In efContext.Products_Issue
                           Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId
                           Select pro.ProductId).ToList()
        Dim i As Integer = 0
        For Each li In listProductId
            If Not (lastProductsId.Contains(li)) AndAlso i < iProIssueCount AndAlso Not (queryProductId.Contains(li)) Then
                Dim pIssue As New Products_Issue
                pIssue.ProductId = li
                pIssue.SiteId = siteId
                pIssue.IssueID = issueId
                pIssue.SectionID = sectionId
                efContext.AddToProducts_Issue(pIssue)
                i = i + 1
            End If
            If (queryProductId.Contains(li)) Then
                i = i - 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' SectionId != 'CA'的产品添加到Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="listProductId"></param>
    ''' <param name="iTakeCount"></param>
    ''' <param name="iProIssueCount"></param>
    ''' <remarks></remarks>
    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer), _
                                    ByVal iTakeCount As Integer, ByVal iProIssueCount As Integer)
        Dim queryLastProductId = (From pi In efContext.Products_Issue
                               Where pi.IssueID < issueId AndAlso pi.SiteId = siteId AndAlso pi.SectionID = sectionId
                               Order By pi.IssueID Descending
                               Select pi.ProductId).Take(iTakeCount)
        Dim queryProductId = (From pro In efContext.Products_Issue
                           Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId
                           Select pro.ProductId).ToList()
        Dim i As Integer = 0
        For Each li In listProductId
            If Not (queryLastProductId.Contains(li)) AndAlso i < iProIssueCount AndAlso Not (queryProductId.Contains(li)) Then
                Dim pIssue As New Products_Issue
                pIssue.ProductId = li
                pIssue.SiteId = siteId
                pIssue.IssueID = issueId
                pIssue.SectionID = sectionId
                efContext.AddToProducts_Issue(pIssue)
                i = i + 1
            End If
            If (queryProductId.Contains(li)) Then
                i = i - 1
            End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 获取某个分类URL的产品信息，如：http://www.dresslily.com/dresses-c-6.html
    ''' </summary>
    ''' <param name="pageUrl">特定分类的URL</param>
    ''' <param name="productCount"></param>
    ''' <param name="siteId"></param>
    ''' <param name="nowTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListProducts(ByVal pageUrl As String, ByVal productCount As Integer, _
                                    ByVal siteId As Integer, ByVal nowTime As DateTime) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            For counter As Integer = 0 To productCount - 1
                Dim liNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("ul/li")
                Dim liNode1 As HtmlNode = liNodes(0)
                Dim liNode2 As HtmlNode = liNodes(1)
                Dim liNode3 As HtmlNode = liNodes(2) '原价
                ' Dim liNode4 As HtmlNode = liNodes(3) '折扣价
                Dim product As New Product()
                product.Prodouct = liNode2.SelectSingleNode("h4/a").InnerText
                product.Url = ParseLink(liNode2.SelectSingleNode("h4/a").GetAttributeValue("href", "")).Trim()
                '2013/07/19 added,price and discount price layer changed
                'product.Discount = Double.Parse(liNode3.SelectNodes("span")(2).InnerText.Trim())
                'product.Price = Double.Parse(liNode3.SelectNodes("span")(0).SelectNodes("span")(1).InnerText.Trim())
                product.Discount = Double.Parse(liNode3.SelectNodes("span")(0).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Price = Double.Parse(liNode3.SelectNodes("span")(1).GetAttributeValue("orgp", "").Trim("$").Trim())
                product.Currency = "USD" 'Dora 2013-11-25 不好获取，且经常改版故直接赋值

                '特殊图片处理
                product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                If (liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                    product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("data-src", "")
                End If

                product.LastUpdate = nowTime
                product.Description = liNode2.SelectSingleNode("h4/a").InnerText
                product.SiteID = siteId
                product.PictureAlt = liNode1.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                product.SizeHeight = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("width", ""))
                product.SizeWidth = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("height", ""))

                If (liNodes.Count = 4) Then
                    Dim myPicture1 As String = ""
                    Try
                        myPicture1 = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try

                    If (myPicture1.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture1
                    ElseIf (myPicture1.Contains("freepic.gif") OrElse myPicture1.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture1
                    End If

                ElseIf (liNodes.Count >= 5) Then
                    Dim myPicture1 As String = ""
                    Dim myPicture2 As String = ""
                    Try
                        myPicture1 = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try
                    Try
                        myPicture2 = liNodes(4).SelectSingleNode("img").GetAttributeValue("src", "")
                    Catch ex As Exception
                        'Ignore
                    End Try
                    If (myPicture1.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture1
                    ElseIf (myPicture1.Contains("freepic.gif") OrElse myPicture1.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture1
                    End If
                    If (myPicture2.Contains("icon_ships24.gif")) Then
                        product.ShipsImg = myPicture2
                    ElseIf (myPicture2.Contains("freepic.gif") OrElse myPicture2.Contains("frees.gif")) Then
                        product.FreeShippingImg = myPicture2
                    End If
                End If
                listProducts.Add(product)
            Next
        Catch ex As Exception
            LogText(ex.ToString())
        End Try
        Return listProducts
    End Function

    ''' <summary>
    ''' 获取Hot Sales中的信息，并添加到Products表，同时生成Products_Issue表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <remarks></remarks>
    Private Sub GetHotSales(ByVal siteId As Integer, ByVal IssueId As Integer)
        Dim ulNodes As HtmlNodeCollection = hd.GetElementbyId("main0").SelectNodes("ul")
        Dim nowTime As DateTime = Now
        Dim listProductId As New List(Of Integer)
        For i As Integer = 0 To ulNodes.Count - 1
            Dim liNodes As HtmlNodeCollection = ulNodes(i).SelectNodes("li")
            For Each li In liNodes
                Dim j As Integer = 2
                Dim divNodes As HtmlNodeCollection = li.SelectNodes("div")
                Dim product As String = divNodes(1).SelectSingleNode("a").InnerText
                Dim url As String = ParseLink(divNodes(1).SelectSingleNode("a").GetAttributeValue("href", "")).Trim()
                Dim freeShipImg As String = ""
                Dim shipsImg As String = ""

                '2013/3/28新加，判断是否存在免运费层
                Dim price As Decimal '=Decimal.Parse(divNodes(j).SelectSingleNode("span/span/span[2]").InnerText.Trim())
                Dim currency As String '=divNodes(j).SelectSingleNode("span/span/span[1]").InnerHtml.Trim()
                If (divNodes(j).InnerText.Trim().Contains("USD")) Then
                    price = Decimal.Parse(divNodes(j).SelectSingleNode("span/span/span[2]").InnerText.Trim())
                    currency = divNodes(j).SelectSingleNode("span/span/span[1]").InnerHtml.Trim()
                Else '2013/4/11新增，Hot Sales处多增加2层，获取数据失败
                    If (divNodes(j + 1).InnerText.Trim().Contains("USD")) Then '一个产品的div有4层
                        price = Decimal.Parse(divNodes(j + 1).SelectSingleNode("span/span/span[2]").InnerText.Trim())
                        currency = divNodes(j + 1).SelectSingleNode("span/span/span[1]").InnerHtml.Trim()
                        Dim myImg As String = divNodes(j).SelectSingleNode("img").GetAttributeValue("src", "")
                        If (myImg.Contains("freepic.gif")) Then
                            freeShipImg = myImg
                        ElseIf (myImg.Contains("icon_ships24.gif")) Then
                            freeShipImg = myImg
                        End If
                    Else '一个产品的div有5层
                        price = Decimal.Parse(divNodes(j + 2).SelectSingleNode("span/span/span[2]").InnerText.Trim())
                        currency = divNodes(j + 2).SelectSingleNode("span/span/span[1]").InnerHtml.Trim()
                        Dim divImg1 As String = divNodes(j).SelectSingleNode("img").GetAttributeValue("src", "")
                        Dim divImg2 As String = divNodes(j + 1).SelectSingleNode("img").GetAttributeValue("src", "")
                        If (divImg1.Contains("freepic.gif")) Then
                            freeShipImg = divImg1
                            shipsImg = divImg2
                        ElseIf (divImg1.Contains("icon_ships24.gif")) Then
                            shipsImg = divImg1
                            freeShipImg = divImg2
                        End If
                    End If
                End If

                '2013/4/8添加
                Dim pictureUrl As String = divNodes(0).SelectSingleNode("a/img").GetAttributeValue("src", "")
                If (divNodes(0).SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                    pictureUrl = divNodes(0).SelectSingleNode("a/img").GetAttributeValue("data-src", "")
                End If

                Dim lastUpdate As DateTime = nowTime
                Dim description As String = divNodes(1).SelectSingleNode("a").InnerText
                Dim pictureAlt As String = divNodes(0).SelectSingleNode("a/img").GetAttributeValue("alt", "")
                Dim sizeWidth As Integer = 150
                Dim sizeHeight As Integer = 150
                Dim categoryId As Integer = InsertCategory(url, siteId, IssueId)
                Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, nowTime, description, siteId, currency, pictureAlt, sizeWidth, sizeHeight, categoryId, listProduct, freeShipImg, shipsImg)
                listProductId.Add(returnId)
            Next
        Next
        InsertProductIssue(siteId, IssueId, "CA", listProductId)
    End Sub

    ''' <summary>
    ''' 获取Men's Shoes中的信息，并添加到Products表，同时生成Products_Issue表中的数据
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <remarks></remarks>
    Public Sub GetMenShoes(ByVal siteId As Integer, ByVal IssueId As Integer)
        Dim doc As HtmlDocument = GetHtmlDocByUrl(menShoesUrl, 30000)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
        Dim queryMenShoes = (From m In efContext.Categories
                          Where m.Category1 = "Men's Shoes" AndAlso m.SiteID = siteId
                          Select m).First()
        Dim categoryId As Integer = queryMenShoes.CategoryID
        Dim listProductId As New List(Of Integer)
        Dim nowTime As DateTime = Now
        For i As Integer = 0 To 8
            Dim liNodes As HtmlNodeCollection = divNodes(i).SelectNodes("ul/li")
            Dim imgNode As HtmlNode = liNodes(0).SelectSingleNode("a/img")
            Dim linkNode As HtmlNode = liNodes(1).SelectSingleNode("h4/a")
            Dim currency As String = liNodes(2).SelectSingleNode("span[2]").InnerHtml.Trim()
            Dim price As Decimal = Decimal.Parse(liNodes(2).SelectSingleNode("span[3]").InnerText.Trim())
            Dim product As String = linkNode.InnerText
            Dim url As String = ParseLink(linkNode.GetAttributeValue("href", "")).Trim()
            Dim freeShipImg As String = ""
            Dim shipImg As String = ""
            If (liNodes.Count = 4) Then
                Dim myImg As String = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                If (myImg.Contains("icon_ships24.gif")) Then
                    shipImg = myImg
                ElseIf (myImg.Contains("freepic.gif") OrElse myImg.Contains("frees.gif")) Then
                    freeShipImg = myImg
                End If
            End If
            If (liNodes.Count >= 5) Then
                Dim img1 As String = liNodes(3).SelectSingleNode("img").GetAttributeValue("src", "")
                Dim img2 As String = liNodes(4).SelectSingleNode("img").GetAttributeValue("src", "")
                If (img1.Contains("freepic.gif") OrElse img1.Contains("frees.gif")) Then
                    freeShipImg = img1
                ElseIf (img1.Contains("icon_ships24.gif")) Then
                    shipImg = img1
                End If
                If (img2.Contains("freepic.gif") OrElse img2.Contains("frees.gif")) Then
                    freeShipImg = img2
                ElseIf (img2.Contains("icon_ships24.gif")) Then
                    shipImg = img2
                End If
            End If
            '2013/4/8添加
            Dim pictureUrl As String = imgNode.GetAttributeValue("src", "")
            If (imgNode.GetAttributeValue("src", "").Contains("loading.gif")) Then
                pictureUrl = imgNode.GetAttributeValue("data-src", "")
            End If

            Dim lastUpdate As DateTime = nowTime
            Dim description As String = linkNode.InnerText
            Dim pictureAlt As String = imgNode.GetAttributeValue("alt", "")
            Dim sizeWidth As Integer = -1
            Dim sizeHeight As Integer = -1
            Dim returnId As Integer = InsertProduct(product, url, price, pictureUrl, lastUpdate, description, siteId, currency, pictureAlt, sizeWidth, sizeHeight, categoryId, listProduct, freeShipImg, shipImg)
            listProductId.Add(returnId)
        Next
        InsertProductIssue(siteId, IssueId, "CA", listProductId)
    End Sub

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
    Public Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                              ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal sizeWidth As Integer, _
                              ByVal sizeHeight As Integer, ByVal categoryId As Integer, ByVal list As List(Of Product), _
                              ByVal freeShipingImg As String, ByVal shipsImg As String) As Integer
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.SingleOrDefault
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
        product.SizeWidth = sizeWidth
        product.SizeHeight = sizeHeight
        product.FreeShippingImg = freeShipingImg
        product.ShipsImg = shipsImg
        If (JudgeProduct(product.Prodouct, product.Url, product.Price, product.PictureUrl, product.SiteID, product.Currency, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = Enumerable.FirstOrDefault(efContext.Products, Function(m) m.Url = url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            'Dim returnId As Integer = updateProduct.ProdouctID
            'efContext.SaveChanges()

            '2013/4/17新增，以防止在填充模板时，产品丢失
            Dim updateCategory = updateProduct.Categories
            If Not (updateCategory.Contains(category)) Then
                category.Products.Add(updateProduct) '在查询分类时，添加一条关系在ProductCategory表中
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
    End Function

    Public Function InsertProduct(ByVal myProduct As String, ByVal url As String, ByVal price As Decimal, ByVal discount As Decimal, ByVal pictureUrl As String, ByVal lastUpdate As DateTime, _
                              ByVal description As String, ByVal siteId As Integer, ByVal currency As String, ByVal pictureAlt As String, ByVal sizeWidth As Integer, _
                              ByVal sizeHeight As Integer, ByVal categoryId As Integer, ByVal list As List(Of Product), _
                              ByVal freeShipingImg As String, ByVal shipsImg As String) As Integer
        url = url.Trim()
        Dim queryCategory = From c In efContext.Categories Where c.CategoryID = categoryId Select c
        Dim category As Category = queryCategory.FirstOrDefault()
        Dim product As New Product()
        product.Prodouct = myProduct
        product.Url = url
        product.Price = price
        product.Discount = discount
        product.PictureUrl = pictureUrl
        product.LastUpdate = lastUpdate
        product.Description = description
        product.SiteID = siteId
        product.Currency = currency
        product.PictureAlt = pictureAlt
        product.SizeWidth = sizeWidth
        product.SizeHeight = sizeHeight
        product.FreeShippingImg = freeShipingImg
        product.ShipsImg = shipsImg
        If (JudgeProduct(product.Prodouct, product.Url, product.Price, product.PictureUrl, product.SiteID, product.Currency, list)) Then
            product.Categories.Add(category)
            efContext.AddToProducts(product)
            efContext.SaveChanges()
            Return product.ProdouctID
        Else
            Dim updateProduct = efContext.Products.FirstOrDefault(Function(m) m.Url = url)
            updateProduct.Prodouct = product.Prodouct
            updateProduct.Price = product.Price
            updateProduct.Discount = discount
            updateProduct.PictureUrl = product.PictureUrl
            updateProduct.Description = description
            updateProduct.PictureAlt = product.PictureAlt
            updateProduct.Currency = product.Currency
            updateProduct.LastUpdate = lastUpdate
            updateProduct.FreeShippingImg = freeShipingImg
            updateProduct.ShipsImg = shipsImg
            'Dim returnId As Integer = updateProduct.ProdouctID
            'efContext.SaveChanges()

            '2013/4/17新增，以防止在填充模板时，产品丢失
            Dim updateCategory = updateProduct.Categories
            If Not (updateCategory.Contains(category)) Then
                category.Products.Add(updateProduct) '在查询分类时，添加一条关系在ProductCategory表中
            End If
            efContext.SaveChanges()

            Return updateProduct.ProdouctID
        End If
    End Function

    ''' <summary>
    ''' 根据产品的名称，插入分类名到Categories表和ProductCategory表中
    ''' 说明：该方法后续使用，Spread暂时不支持。
    ''' </summary>
    ''' <param name="productUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns>返回最后一个分类的CategoryID号</returns>
    ''' <remarks></remarks>
    Public Function InsertCategory2(ByVal productUrl As String, ByVal siteId As Integer) As Integer
        Dim queryCategory = efContext.Categories
        Dim listCategory As New List(Of Category)
        For Each q In queryCategory
            listCategory.Add(New Category With {.Category1 = q.Category1, .SiteID = q.SiteID, .Url = q.Url})
        Next
        Dim doc As HtmlDocument = GetHtmlDocByUrl(productUrl, 3000)
        Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//li[@class='home']/div")
        Dim sFirstClass As String = divNodes(0).SelectSingleNode("a/span").InnerText.Trim()
        Dim query = From c In efContext.Categories Where c.Category1 = sFirstClass AndAlso c.SiteID = siteId Select c.CategoryID
        Dim parentId As Integer = Integer.Parse(query.Single())
        For i As Integer = 1 To divNodes.Count - 1
            Dim category As New Category()
            category.ParentID = parentId
            category.Category1 = divNodes(i).SelectSingleNode("a/span").InnerText.Trim()
            category.SiteID = siteId
            category.LastUpdate = DateTime.Now.ToString()
            category.Description = divNodes(i).SelectSingleNode("a/span").InnerText.Trim()
            category.Url = ParseLink(divNodes(i).SelectSingleNode("a").GetAttributeValue("href", ""))
            If (JudgeCategory(category.Category1, siteId, category.Url, listCategory)) Then
                If (i = divNodes.Count - 1) Then
                    efContext.AddToCategories(category)
                    efContext.SaveChanges()
                    Return category.CategoryID
                Else
                    efContext.AddToCategories(category)
                    efContext.SaveChanges()
                End If
            Else
                Dim returnQuery = From c In efContext.Categories Where c.ParentID = category.ParentID And c.Category1 = category.Category1 And c.Url = category.Url Select c.CategoryID
                Dim returnVal As Integer = returnQuery.Single()
                Return returnVal
            End If
        Next
    End Function

    ''' <summary>
    ''' 当前Spread只支持最顶层的分类追踪,每期更新LastUpdate也会不相同，故加一个限制条件
    ''' </summary>
    ''' <param name="productUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertCategory(ByVal productUrl As String, ByVal siteId As Integer, ByVal IssueId As Integer) As Integer
        Dim doc As HtmlDocument = GetHtmlDocByUrl(productUrl, 60000)

        '2013/4/16修改，DressLily获取分类名
        Dim sFirstClass As String = ""
        Try
            Dim divNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//li[@class='home']/div")
            sFirstClass = divNodes(0).SelectSingleNode("a/span").InnerText.Trim()
        Catch ex As Exception
            Dim category As String = doc.DocumentNode.SelectSingleNode("//li[@class='home']").InnerText
            sFirstClass = System.Text.RegularExpressions.Regex.Split(category, "&raquo;", System.Text.RegularExpressions.RegexOptions.IgnoreCase)(1).Trim()
        End Try

        Dim query = From c In efContext.Categories
                    Where c.Category1 = sFirstClass AndAlso c.SiteID = siteId
                    Select c.CategoryID
        Dim returnId As Integer = query.FirstOrDefault()
        Return returnId
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
    Private Function JudgeProduct(ByVal product As String, ByVal url As String, ByVal price As Decimal, ByVal pictureUrl As String, ByVal siteId As Integer, ByVal currency As String, ByVal list As List(Of Product)) As Boolean
        For Each li In list
            If (li.Url.Trim() = url.Trim()) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 添加数据到Products_Issue表中
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="section"></param>
    ''' <remarks></remarks>
    Private Sub InsertProductIssue(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal section As String, ByVal listProductId As List(Of Integer))
        '2013/4/11新增，Products_Issue表中的数据重复
        Dim listId As New HashSet(Of Integer)
        Dim queryProductIssueId = From iss In efContext.Products_Issue
                                Where iss.IssueID = IssueId AndAlso iss.SiteId = siteId AndAlso iss.SectionID = section
                                Select iss
        For Each p In queryProductIssueId
            listId.Add(p.ProductId)
        Next

        For Each li In listProductId
            Dim pIssue As New Products_Issue()
            pIssue.ProductId = li
            pIssue.SiteId = siteId
            pIssue.IssueID = IssueId
            pIssue.SectionID = section
            'If (JudgeProductIssue(pIssue.ProductId, pIssue.SiteId, pIssue.IssueDate, listProductIssue)) Then
            If Not (listId.Contains(li)) Then '2013/4/11新增
                efContext.AddToProducts_Issue(pIssue)
            End If
            'End If
        Next
        efContext.SaveChanges()
    End Sub

    ''' <summary>
    ''' 将Products表中需要匹配列的信息添加到list中
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetProductList(ByVal siteId As Integer) As List(Of Product)
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


    ''' <summary>
    ''' 判断将要插入Categories表中URL是否重复，若数据重复，则返回false
    ''' </summary>
    ''' <param name="category"></param>
    ''' <param name="siteId"></param>
    ''' <param name="Url"></param>
    ''' <param name="list"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function JudgeCategory(ByVal category As String, ByVal siteId As Integer, ByVal Url As String, ByVal list As List(Of Category)) As Boolean
        For Each li In list
            If (li.Url = Url) Then
                Return False
            End If
        Next
        Return True
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
    Private Function JudgeAds(ByVal ad As String, ByVal pictureUrl As String, ByVal url As String, ByVal list As List(Of Ad)) As Boolean
        For Each li In list
            'If (li.Ad = ad And li.PictureUrl = pictureUrl And li.Url = url) Then
            '2013/3/19，根据URL不同，判断是需要更新还是添加一条新纪录
            If (li.Url.Trim() = url.Trim()) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 转换链接，使链接成为完整链接
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ParseLink(ByVal str As String) As String
        If (Not str.Contains("http")) Then
            Return (htmlUrl + str).Trim()
        Else
            Return str.Trim()
        End If
    End Function

    ''' <summary>
    ''' 根据url去获取HtmlDocument，并通过设置timeout的时间去设置加载网页的时间
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="iTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetHtmlDocByUrl(ByRef url As String, ByVal iTime As Integer) As HtmlDocument
        Dim myWebRequest As WebRequest = WebRequest.Create(url)
        Dim myWebResponse As WebResponse
        Try
            myWebRequest.Timeout = iTime
            myWebResponse = myWebRequest.GetResponse() 'The operation has timed out
        Catch ex As Exception
            myWebRequest.Timeout = 120000 '2013/4/3添加120000，time out处理
            myWebResponse = myWebRequest.GetResponse()
            Try '2013/4/3添加
                myWebRequest.Timeout = 120000
                myWebResponse = myWebRequest.GetResponse()
            Catch ex1 As Exception
                LogText(ex1.ToString() + "URL:")
            End Try
        End Try

        Dim receiveStream As Stream = myWebResponse.GetResponseStream()
        Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim readStream As New StreamReader(receiveStream, encode)
        Dim hd As New HtmlDocument()
        hd.Load(receiveStream, encode)
        Return hd
    End Function

    ''' <summary>
    ''' 更新Issues表中的Subject
    ''' </summary>
    ''' <param name="issueId"></param>
    ''' <remarks></remarks>
    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String)
        '2013/05/22，不再使用Promotion作为Subject，开始
        'Dim querySubject = From ad In efContext.Ads
        '                 Join iss In efContext.Ads_Issue On ad.AdID Equals iss.AdId
        '                 Where iss.IssueID = issueId And ad.Type = "P"
        '                 Select ad.Ad1
        'Dim subject As String = querySubject.Single()
        'Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        'myIssue.Subject = "Hi [FIRSTNAME]," & subject & " ★ Weekly Deals" '
        'myIssue.Subject = subject & " 【Weekly Deals】"
        '2013/05/22，不再使用Promotion作为Subject，结束

        Dim queryLastSubject = From iss In efContext.Issues
                             Where iss.IssueID < issueId AndAlso iss.SiteID = siteId AndAlso iss.PlanType = planType
                             Order By iss.IssueID Descending
                             Select iss
        Dim listLastSubject As New HashSet(Of String)
        Dim counter As Integer = 0
        For Each query In queryLastSubject
            If Not (String.IsNullOrEmpty(query.Subject)) AndAlso counter < 10 Then
                listLastSubject.Add(query.Subject)
                counter = counter + 1
            End If
            If (counter >= 10) Then
                Exit For
            End If
        Next
        Dim querySubject = From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = issueId AndAlso pi.SectionID = "NE" AndAlso pi.SiteId = siteId
                         Select p.Prodouct
        Dim issueSubject As String = ""
        For Each mySubject In querySubject
            Dim subject As String = ""
            subject = "Hi [FIRSTNAME]," & mySubject & " ★ Weekly Deals"
            If Not (listLastSubject.Contains(subject)) Then
                issueSubject = subject
                Exit For
            End If
        Next
        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = issueSubject
        efContext.SaveChanges()
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
    Public Sub CreateContacts(ByVal siteId As Integer, ByVal planType As String, ByVal spreaLogin As String, _
                              ByVal appId As String, ByVal splitContactCount As Integer, _
                              ByVal issueId As Integer)
        Dim listName As String
        If (planType = "HO") Then
            listName = "Open"
        ElseIf (planType = "HN") Then
            listName = "NotOpen"  '表示没开启邮件的用户
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
    Private Sub InsertContactList(ByVal spreaLogin As String, ByVal appId As String, ByVal siteId As Integer, _
                                  ByVal planType As String, ByVal issueId As Integer, ByVal listName As String, _
                                  ByVal splitContactCount As Integer)
        Dim querySubscriber As New QuerySubscriber()
        If (planType = "HO") Then
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
            Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()
            mySpread.Timeout = 1200000
            mySpread.SearchContacts(spreaLogin, appId, criteria, topN, saveAsList, forceCreate)

            Dim contactName() As String = saveAsList.Split(";")

            For j As Integer = 0 To splitContactCount - 1
                Dim splitList As New SplitContactList()
                splitList.ShopName = "DressLily"
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
            contactList.SendingStatus = "draft" '2013/12/31 added，Dresslily创建成草稿状态
            efContext.AddToContactLists_Issue(contactList)
            efContext.SaveChanges()
        Else 'splitContactCount=0或1
            If (planType = "HO") Then
                saveAsList = "Opens" & planType & Now.ToString("yyyyMMdd") 'Now.ToString("yyyyMMdd") '
            ElseIf (planType = "HN") Then
                saveAsList = "NotOpen" & planType & Now.ToString("yyyyMMdd") 'Now.ToString("yyyyMMdd") '表示没开启邮件的用户
            End If
            Dim forceCreate As Boolean = True
            Dim mySpread As SpreadWebReference.SpreadWebService = New SpreadWebReference.SpreadWebService()

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
                contactList.SendingStatus = "draft" '2013/12/31 added，Dresslily创建成草稿状态
                efContext.AddToContactLists_Issue(contactList)
            End If

            Dim contact1 As New ContactLists_Issue()
            contact1.IssueID = issueId
            contact1.ContactList = "Opens"
            contact1.SendingStatus = "draft" '2013/12/31 added，Dresslily创建成草稿状态
            efContext.AddToContactLists_Issue(contact1)

            efContext.SaveChanges()
        End If
    End Sub

    ''' <summary>
    ''' 2013/4/11新增，单独获取页面的信息，而非首页，并插入到Products表和Products_Issue表中
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="maxCount"></param>
    ''' <param name="categoryName"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="nowTime"></param>
    ''' <remarks></remarks>
    Public Sub GetProductAndProductIssue(ByVal pageUrl As String, ByVal maxCount As Integer, _
                                         ByVal categoryName As String, ByVal siteId As Integer, _
                                         ByVal issueId As Integer, ByVal nowTime As DateTime)
        Dim listProductUrl As New HashSet(Of String)
        Dim queryProductUrl = From u In efContext.Products
                            Where u.SiteID = siteId
                            Select u
        For Each q In queryProductUrl
            listProductUrl.Add(q.Url)
        Next

        Dim helper As New EFHelper()
        Dim insertProducts As New List(Of Product)
        Dim updateProducts As New List(Of Product)
        Dim listProductId As New List(Of Integer)
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
            For counter As Integer = 0 To maxCount - 1
                Dim liNodes As HtmlNodeCollection = productDivs(counter).SelectNodes("ul/li")
                Dim liNode1 As HtmlNode = liNodes(0)
                Dim liNode2 As HtmlNode = liNodes(1)
                Dim liNode3 As HtmlNode = liNodes(2)
                Dim product As New Product()
                product.Prodouct = liNode2.SelectSingleNode("h4/a").InnerText
                product.Url = ParseLink(liNode2.SelectSingleNode("h4/a").GetAttributeValue("href", "")).Trim()
                product.Price = Double.Parse(liNode3.SelectNodes("span")(2).InnerText.Trim())

                '特殊图片处理
                product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "") '2013/4/7添加
                If (liNode1.SelectSingleNode("a/img").GetAttributeValue("src", "").Contains("loading.gif")) Then
                    product.PictureUrl = liNode1.SelectSingleNode("a/img").GetAttributeValue("data-src", "")
                End If

                product.LastUpdate = nowTime
                product.Description = liNode2.SelectSingleNode("h4/a").InnerText
                product.SiteID = siteId
                product.Currency = liNode3.SelectNodes("span")(1).InnerText.Trim()
                product.PictureAlt = liNode1.SelectSingleNode("a/img").GetAttributeValue("alt", "")
                product.SizeHeight = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("width", ""))
                product.SizeWidth = Integer.Parse(liNode1.SelectSingleNode("a/img").GetAttributeValue("height", ""))
                If Not (listProductUrl.Contains(product.Url)) Then
                    insertProducts.Add(product)
                Else
                    updateProducts.Add(product)
                End If
            Next
            If Not (insertProducts.Count = 0) Then
                listProductId.AddRange(helper.InsertListProduct(insertProducts, categoryName, siteId))
            End If
            If Not (updateProducts.Count = 0) Then
                listProductId.AddRange(helper.UpdateListProduct(updateProducts, categoryName, siteId))
            End If
            helper.InsertProductIssue(listProductId, siteId, issueId, "CA")
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

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

    Sub LogText(ByVal Ex As String)
        Try
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
        Catch ex1 As Exception
            'ignore
        End Try
    End Sub
End Class
