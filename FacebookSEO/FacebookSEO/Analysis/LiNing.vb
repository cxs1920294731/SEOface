
Imports HtmlAgilityPack
Imports System.Math
Imports System.Text.RegularExpressions

Public Class LiNing

    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）
    Private domin As String = "http://www.e-lining.com/"

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        InsertCategories(siteId)
        Dim bannerCategoryName As String
        'Dim contactlistCount As Integer
        If planType.Trim.Contains("HO") Then

            ''bannerCategoryName = "promotion"
            GetFirstPromotionBanner(siteId, IssueID, "paobuxie", planType)
            '' http://www.e-lining.com/
            ''需要获取首页第一个Banner，由三张图片组成，每个图片作为一个Ads，分别插入到Ads表，Type为"B"，获取和插入数据库的顺序必须为：左->上->下
            ''然后插入到Ads，返回AdsID，
            ''根据类别名称（promotion）查询Categoryies表获取到对应的Categoryid，插入到AdsCategory表
            ''最好确定好AdsID后插入IssueID和AdsID到Ads_Issue

            ''获取首页第二个Banner，由2张图片组成
            'bannerCategoryName = "paobuxie"
            'GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)


            GetHotSellerProducts(siteId, IssueID, "paobuxie", planType)
            ''http://www.e-lining.com/, categoryName="hotseller"
            ''需要获取首页上跑步鞋、休闲鞋、篮球鞋、服装类别下的各一个产品，产品三个礼拜没不能重复(以产品URL作为唯一判断依据),
            ''重复判断方法：根据IssueID查询Issues表，拿出IssueDate做判断，再结合Products_Issue/Products表做判断
            ''然后插入到Products，返回ProductID，
            ''根据类别名称（hotseller）查询Categoryies表获取到对应的Categoryid，插入到ProductCategory表
            ''最好确定好ProductID后插入IssueID和ProductID到Products_Issue



            '主推4个分类,每个分类获取3个产品
            GetHotOrNewProducts(siteId, IssueID, "hot", planType)
            ''Add Subject
            Dim querySubject = From p In efContext.Products
                       Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                       Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                       Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "Hi [FIRSTNAME]你好,本季热销推荐：" & querySubject.FirstOrDefault.ToString() & " ...")

        ElseIf (planType.Trim.Contains("HA")) Then
            bannerCategoryName = "paobuxie"
            GetPromotionBanner(siteId, IssueID, bannerCategoryName, planType)
            'GetWeeklyDealsOrNewArri(siteId, IssueID, "NE", planType)
            '主推4个分类,每个分类获取3个产品

            GetHotSellerProducts(siteId, IssueID, "hotseller", planType)
            ''GetHotSellerProducts(siteId, IssueID, "hotseller", planType),
            ''http://www.e-lining.com/, categoryName="hotseller"，只获取产品名称（描述）中含有“新品字眼”的产品
            ''需要获取首页上跑步鞋、休闲鞋、篮球鞋、服装类别下的各一个产品，产品三个礼拜没不能重复(以产品URL作为唯一判断依据),
            ''重复判断方法：根据IssueID查询Issues表，拿出IssueDate做判断，再结合Products_Issue/Products表做判断
            ''然后插入到Products，返回ProductID，
            ''根据类别名称（hotseller）查询Categoryies表获取到对应的Categoryid，插入到ProductCategory表
            ''最好确定好ProductID后插入IssueID和ProductID到Products_Issue


            GetHotOrNewProducts(siteId, IssueID, "new", planType)
            ''Add Subject
            Dim querySubject = From p In efContext.Products
                       Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                       Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                       Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "让改变发生，新品推荐:" & querySubject.FirstOrDefault.ToString() & " ...")

        End If
    End Sub

    ''' <summary>
    ''' 往数据库中categories插入数据
    ''' </summary>
    ''' <param name="siteID">siteid </param>
    ''' <remarks></remarks>
    Public Sub InsertCategories(ByVal siteID As Integer)
        Try
            Dim queryCategory = From c In efContext.Categories
                                Where c.SiteID = siteID
                                Select c
            Dim listCategoryName As New HashSet(Of String)
            For Each q In queryCategory
                listCategoryName.Add(q.Category1)
            Next
            Dim nowTime As DateTime = Now
            '添加11个分类 
            Dim paobuxie As New Category()
            Dim paobufu As New Category()
            Dim lanqiuxie As New Category()
            Dim lanqiufu As New Category()
            Dim lanqiuAcc As New Category()
            Dim yundong As New Category()
            Dim yundongfu As New Category()
            Dim yundongAcc As New Category()
            Dim xunlianxie As New Category()
            Dim xunlianfu As New Category()
            Dim xunlianAcc As New Category()
            If Not listCategoryName.Contains("paobuxie") Then
                paobuxie.Category1 = "paobuxie"
                paobuxie.SiteID = siteID
                paobuxie.Url = "http://www.e-lining.com/shop/searchkey-flunkt45a6ibehn3asuk-15-update-desc-image-1-0.html"
                paobuxie.LastUpdate = nowTime
                paobuxie.Description = "paobuxie"
                paobuxie.ParentID = -1
                efContext.AddToCategories(paobuxie)
                'Else
                '    paobuxie = efContext.Categories.FirstOrDefault(Function(m) m.Category1 = "paobuxie" And m.SiteID = siteID)
                '    paobuxie.Url = "http://www.e-lining.com/shop/searchkey-flunkt45a6ibehn3asuk-15-update-desc-image-1-0.html"
            End If
            If Not listCategoryName.Contains("paobufu") Then
                paobufu.Category1 = "paobufu"
                paobufu.SiteID = siteID
                paobufu.Url = "http://www.e-lining.com/shop/searchkey-flunkt45a6ibehn3asuk-16-update-desc-image-1-0.html"
                paobufu.LastUpdate = nowTime
                paobufu.Description = "paobufu"
                paobufu.ParentID = -1
                efContext.AddToCategories(paobufu)
            End If
            If Not listCategoryName.Contains("lanqiuxie") Then
                lanqiuxie.Category1 = "lanqiuxie"
                lanqiuxie.SiteID = siteID
                lanqiuxie.Url = "http://www.e-lining.com/shop/searchkey-flunisn6gpibehtoa1tk-15-update-desc-image-1-0.html"
                lanqiuxie.LastUpdate = nowTime
                lanqiuxie.Description = "lanqiuxie"
                lanqiuxie.ParentID = -1
                efContext.AddToCategories(lanqiuxie)
            End If
            If Not listCategoryName.Contains("lanqiufu") Then
                lanqiufu.Category1 = "lanqiufu"
                lanqiufu.SiteID = siteID
                lanqiufu.Url = "http://www.e-lining.com/shop/searchkey-flunisn6gpibehtoa1tk-16-update-desc-image-1-0.html"
                lanqiufu.LastUpdate = nowTime
                lanqiufu.Description = "lanqiufu"
                lanqiufu.ParentID = -1
                efContext.AddToCategories(lanqiufu)
            End If
            If Not listCategoryName.Contains("lanqiuAcc") Then
                lanqiuAcc.Category1 = "lanqiuAcc"
                lanqiuAcc.SiteID = siteID
                lanqiuAcc.Url = "http://www.e-lining.com/shop/searchkey-flunisn6gpibehtoa1tk-17-update-desc-image-1-0.html"
                lanqiuAcc.LastUpdate = nowTime
                lanqiuAcc.Description = "lanqiuAcc"
                lanqiuAcc.ParentID = -1
                efContext.AddToCategories(lanqiuAcc)
            End If
            If Not listCategoryName.Contains("yundong") Then
                yundong.Category1 = "yundong"
                yundong.SiteID = siteID
                yundong.Url = "http://www.e-lining.com/shop/searchkey-flunktn5a2ibehdmg5w6sxdraosqgr4patm4ask-15-update-desc-image-1-0.html"
                yundong.LastUpdate = nowTime
                yundong.Description = "yundong"
                yundong.ParentID = -1
                efContext.AddToCategories(yundong)
            End If
            If Not listCategoryName.Contains("yundongfu") Then
                yundongfu.Category1 = "yundongfu"
                yundongfu.SiteID = siteID
                yundongfu.Url = "http://www.e-lining.com/shop/searchkey-flunktn5a2ibehdmg5w6sxdraosqgr4patm4ask-16-update-desc-image-1-0.html"
                yundongfu.LastUpdate = nowTime
                yundongfu.Description = "yundongfu"
                yundongfu.ParentID = -1
                efContext.AddToCategories(yundongfu)
            End If
            If Not listCategoryName.Contains("yundongAcc") Then
                yundongAcc.Category1 = "yundongAcc"
                yundongAcc.SiteID = siteID
                yundongAcc.Url = "http://www.e-lining.com/shop/searchkey-flunktn5a2ibehdmg5w6sxdraosqgr4patm4ask-17-update-desc-image-1-0.html"
                yundongAcc.LastUpdate = nowTime
                yundongAcc.Description = "yundongAcc"
                yundongAcc.ParentID = -1
                efContext.AddToCategories(yundongAcc)
            End If
            If Not listCategoryName.Contains("xunlianxie") Then
                xunlianxie.Category1 = "xunlianxie"
                xunlianxie.SiteID = siteID
                xunlianxie.Url = "http://www.e-lining.com/shop/searchkey-flunksn2glibehv6ggtk-15-update-desc-image-1-0.html"
                xunlianxie.LastUpdate = nowTime
                xunlianxie.Description = "xunlianxie"
                xunlianxie.ParentID = -1
                efContext.AddToCategories(xunlianxie)
            End If
            If Not listCategoryName.Contains("xunlianfu") Then
                xunlianfu.Category1 = "xunlianfu"
                xunlianfu.SiteID = siteID
                xunlianfu.Url = "http://www.e-lining.com/shop/searchkey-flunksn2glibehv6ggtk-16-update-desc-image-1-0.html"
                xunlianfu.LastUpdate = nowTime
                xunlianfu.Description = "xunlianfu"
                xunlianfu.ParentID = -1
                efContext.AddToCategories(xunlianfu)
            End If
            If Not listCategoryName.Contains("xunlianAcc") Then
                xunlianAcc.Category1 = "xunlianAcc"
                xunlianAcc.SiteID = siteID
                xunlianAcc.Url = "http://www.e-lining.com/shop/searchkey-flunksn2glibehv6ggtk-17-update-desc-image-1-0.html"
                xunlianAcc.LastUpdate = nowTime
                xunlianAcc.Description = "xunlianAcc"
                xunlianAcc.ParentID = -1
                efContext.AddToCategories(xunlianAcc)
            End If

            efContext.SaveChanges()
        Catch ex As Exception
            efHelper.LogText(ex.Message.ToString)
        End Try
    End Sub

    Public Sub GetFirstPromotionBanner(ByVal siteId As Integer, ByVal IssueID As Integer, ByVal bannerCategoryName As String, ByVal planType As String)
        Dim pageUrl As String = "http://www.e-lining.com/"
        'Dim pageDoc As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000)
        'Dim imageNode As HtmlNode = pageDoc.DocumentNode.SelectSingleNode("//div[@class='indexContent']/table/tbody/tr/td/table/tbody/tr")
        'Dim aNodes As HtmlNodeCollection = imageNode.SelectNodes("a")
        'Dim ad As New Ad
        'Dim adlist As New List(Of Ad)
        'For Each node In aNodes
        '    ad.Url = node.GetAttributeValue("href", "").ToString.Trim
        '    ad.PictureUrl = node.SelectSingleNode("img").GetAttributeValue("src", "").ToString.Trim
        '    ad.SiteID = siteId
        '    If bannerCategoryName = "promotion" Then
        '        ad.Type = "P"
        '    End If
        '    ad.Lastupdate = DateTime.Now

        'Next


        Dim updateTime As DateTime = Now
        'Dim pageUrl As String = "http://www.e-lining.com/"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息

        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, bannerCategoryName)
        Dim listBigAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
        InsertAdsIssue(siteId, IssueID, listBigAdId, iProIssueCount)



        Dim bannerCount As Integer = 3
        'Dim categoryId As Integer = efHelper.GetCategoryId(siteId, bannerCategoryName)
        Dim listAdId As List(Of Integer) = GetListAdids_FirstBanner(pageUrl, siteId, categoryId, DateTime.Now, planType)
        InsertAdsIssue(siteId, IssueID, listAdId, bannerCount)
    End Sub


    Public Sub GetPromotionBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)

        Dim updateTime As DateTime = Now
        Dim pageUrl As String = "http://www.e-lining.com/"
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息

        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
        Dim listAdId As List(Of Integer) = GetListAdids(pageUrl, siteId, categoryId, updateTime, planType)
        InsertAdsIssue(siteId, Issueid, listAdId, iProIssueCount)

        iProIssueCount = 2
        Dim listLittleAdId As List(Of Integer) = GetLittleListAdids(pageUrl, siteId, categoryId, updateTime, planType)
        InsertAdsIssue(siteId, Issueid, listLittleAdId, iProIssueCount)
    End Sub


    Public Function GetListAdids_FirstBanner(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-18).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim leftbanners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='indexContent']/table/tbody/tr/td/table/tbody/tr/td")  '/div/div/a/
        Dim rightbanners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='indexContent']/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td")
        'Dim pageDoc As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000)
        'Dim leftimageNode As HtmlNode = pageDoc.DocumentNode.SelectSingleNode("//div[@class='indexContent']/table/tbody/tr/td/table/tbody/tr/td/a")

        'Dim aleftNodes As HtmlNodeCollection = leftimageNode.SelectNodes("a")
        Dim ad As New Ad
        Dim adlist As New List(Of Ad)
        'For Each node In aNodes
        '    ad.Url = node.GetAttributeValue("href", "").ToString.Trim
        '    ad.PictureUrl = node.SelectSingleNode("img").GetAttributeValue("src", "").ToString.Trim
        '    ad.SiteID = siteid
        '    If bannerCategoryName = "promotion" Then
        '        ad.Type = "P"
        '    End If
        '    ad.Lastupdate = DateTime.Now

        'Next
        Dim count As Integer = 1
        For Each ba In leftbanners
            Dim SingleAd As Ad = New Ad()
            If (ba.SelectSingleNode("a/img")) IsNot Nothing AndAlso count = 2 Then
                Try
                    SingleAd.Ad1 = ba.SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    SingleAd.Description = ba.SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                Try
                    SingleAd.Url = ba.SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim
                    SingleAd.Url = efHelper.addDominForUrl(domin, SingleAd.Url)
                    SingleAd.PictureUrl = ba.SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.Trim
                    SingleAd.PictureUrl = efHelper.addDominForUrl(domin, SingleAd.PictureUrl)

                    SingleAd.SiteID = siteid
                    SingleAd.Lastupdate = nowTime
                    SingleAd.Type = "B"
                    Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                    listAd = efHelper.GetListAd(siteid)
                    'If (Not efHelper.isAdSent(siteid, SingleAd.Url, startDate, endDate, planType) AndAlso SingleAd.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                    adid = efHelper.InsertAd(SingleAd, listAd, categroyId)
                    listAdIds.Add(adid)
                    'End If
                Catch ex As Exception

                End Try
            End If
            count = count + 1
        Next
        For Each ba In rightbanners
            If (ba.SelectSingleNode("a/img")) IsNot Nothing Then
                Dim SingleAd As Ad = New Ad()
                Try
                    SingleAd.Ad1 = ba.SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    SingleAd.Description = ba.SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                SingleAd.Url = ba.SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim
                SingleAd.Url = efHelper.addDominForUrl(domin, SingleAd.Url)
                SingleAd.PictureUrl = ba.SelectSingleNode("a/img").GetAttributeValue("src", "").ToString.Trim
                SingleAd.PictureUrl = efHelper.addDominForUrl(domin, SingleAd.PictureUrl)

                SingleAd.SiteID = siteid
                SingleAd.Lastupdate = nowTime
                SingleAd.Type = "B"
                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)
                If (Not efHelper.isAdSent(siteid, SingleAd.Url, startDate, endDate, planType) AndAlso SingleAd.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                    adid = efHelper.InsertAd(SingleAd, listAd, categroyId)
                    listAdIds.Add(adid)
                End If
            End If
        Next



        Return listAdIds
    End Function


    Public Function GetListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listAdIds As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-18).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@id='swiper-container']/div/div")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Try
                Ad.Ad1 = banners(counter).SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Ad.Description = banners(counter).SelectSingleNode("a/img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
            Catch ex As Exception
                'ignore
            End Try
            Ad.Url = banners(counter).SelectSingleNode("a").GetAttributeValue("href", "").ToString.Trim
            Ad.Url = efHelper.addDominForUrl(domin, Ad.Url)
            Ad.PictureUrl = banners(counter).SelectSingleNode("a/img").GetAttributeValue("orginalsrc", "").ToString.Trim
            Ad.PictureUrl = efHelper.addDominForUrl(domin, Ad.PictureUrl)

            Ad.SiteID = siteid
            Ad.Lastupdate = nowTime

            Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
            listAd = efHelper.GetListAd(siteid)
            If (Not efHelper.isAdSent(siteid, Ad.Url, startDate, endDate, planType) AndAlso Ad.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                adid = efHelper.InsertAd(Ad, listAd, categroyId)
                listAdIds.Add(adid)
            End If
        Next

        Return listAdIds
    End Function

    Public Function GetLittleListAdids(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String) As List(Of Integer)
        Dim listLittleAdId As New List(Of Integer)
        Dim adid As Integer
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-18).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='banner']/ul/li")  '/div/div/a/
        For counter As Integer = 0 To banners.Count - 1 '循环明星系列、新品系列
            Dim myBanners As HtmlNodeCollection = banners(counter).SelectNodes("div/div/a")
            For Each bannerItem In myBanners  '循环某个系列下的banner
                Dim Ad As Ad = New Ad()
                Try
                    Ad.Ad1 = bannerItem.SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    Ad.Description = bannerItem.SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                Ad.Url = bannerItem.GetAttributeValue("href", "").ToString.Trim
                Ad.Url = efHelper.addDominForUrl(domin, Ad.Url)
                Ad.PictureUrl = bannerItem.SelectSingleNode("img").GetAttributeValue("src", "").ToString.Trim
                Ad.PictureUrl = efHelper.addDominForUrl(domin, Ad.PictureUrl)

                Ad.Type = "B"
                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                Dim listAd As New List(Of Ad) '获取本站的的Ads表中的所有ad
                listAd = efHelper.GetListAd(siteid)
                If (Not efHelper.isAdSent(siteid, Ad.Url, startDate, endDate, planType) AndAlso Ad.PictureUrl <> "") Then '控制一个月内获取的Ad不重复：
                    adid = efHelper.InsertAd(Ad, listAd, categroyId)
                    listLittleAdId.Add(adid)
                    Exit For
                End If
            Next
        Next
        If (listLittleAdId.Count = 1) Then
            listLittleAdId.RemoveAt(0)
        End If
        Return listLittleAdId
    End Function


    Public Sub GetHotOrNewProducts(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal HotORNew As String, ByVal planType As String)
        Dim paobuxie As String = ""
        Dim paobufu As String = ""
        Dim lanqiuxie As String = ""
        Dim lanqiufu As String = ""
        Dim lanqiuAcc As String = ""
        Dim yundong As String = ""
        Dim yundongfu As String = ""
        Dim yundongAcc As String = ""
        Dim xunlianxie As String = ""
        Dim xunlianfu As String = ""
        Dim xunlianAcc As String = ""

        Try
            Dim queryURL = From c In efContext.Categories
                           Where c.SiteID = siteId
                           Select c
            For Each q In queryURL
                Select Case q.Category1.Trim()
                    Case "paobuxie"
                        If (HotORNew = "hot") Then
                            paobuxie = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            paobuxie = q.Url.Trim() '& "2.html"
                        End If
                    Case "paobufu"
                        If (HotORNew = "hot") Then
                            paobufu = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            paobufu = q.Url.Trim() '& "2.html"
                        End If
                    Case "lanqiuxie"
                        If (HotORNew = "hot") Then
                            lanqiuxie = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            lanqiuxie = q.Url.Trim() '& "2.html"
                        End If
                    Case "lanqiufu"
                        If (HotORNew = "hot") Then
                            lanqiufu = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            lanqiufu = q.Url.Trim() '& "2.html"
                        End If
                    Case "lanqiuAcc"
                        If (HotORNew = "hot") Then
                            lanqiuAcc = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            lanqiuAcc = q.Url.Trim() '& "2.html"
                        End If
                    Case "yundong"
                        If (HotORNew = "hot") Then
                            yundong = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            yundong = q.Url.Trim() '& "2.html"
                        End If
                    Case "yundongfu"
                        If (HotORNew = "hot") Then
                            yundongfu = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            yundongfu = q.Url.Trim() '& "2.html"
                        End If
                    Case "yundongAcc"
                        If (HotORNew = "hot") Then
                            yundongAcc = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            yundongAcc = q.Url.Trim() '& "2.html"
                        End If
                    Case "xunlianxie"
                        If (HotORNew = "hot") Then
                            xunlianxie = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            xunlianxie = q.Url.Trim() '& "2.html"
                        End If
                    Case "xunlianfu"
                        If (HotORNew = "hot") Then
                            xunlianfu = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            xunlianfu = q.Url.Trim() '& "2.html"
                        End If
                    Case "xunlianAcc"
                        If (HotORNew = "hot") Then
                            xunlianAcc = q.Url.Trim().Replace("update-desc", "sale-desc")
                        Else
                            xunlianAcc = q.Url.Trim() '& "2.html"
                        End If
                End Select
            Next
            'Dim maxProductCount As Integer = 34
            Dim iProIssueCount As Integer
            Dim updateTime As DateTime = Now
            iProIssueCount = 2
            GetProducts(siteId, Issueid, paobuxie, "paobuxie", iProIssueCount, updateTime, planType)

            iProIssueCount = 1
            GetProducts(siteId, Issueid, paobufu, "paobufu", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, lanqiuxie, "lanqiuxie", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, lanqiufu, "lanqiufu", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, lanqiuAcc, "lanqiuAcc", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, yundong, "yundong", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, yundongfu, "yundongfu", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, yundongAcc, "yundongAcc", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, xunlianxie, "xunlianxie", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, xunlianfu, "xunlianfu", iProIssueCount, updateTime, planType)

            GetProducts(siteId, Issueid, xunlianAcc, "xunlianAcc", iProIssueCount, updateTime, planType)
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueID"></param>
    ''' <param name="CategoryName"></param>
    ''' <param name="planType"></param>
    ''' <remarks></remarks>
    Public Sub GetHotSellerProducts(ByVal siteId As Integer, ByVal IssueID As Long, ByVal CategoryName As String, ByVal planType As String)
        Dim ajax_hotGoodsUrl As String = "http://www.e-lining.com/shop/comm_tplfun/ajax_hotGoods.php"
        Dim dataStr As String = efHelper.GetHtmlDocument(ajax_hotGoodsUrl, 60000).DocumentNode.InnerHtml
        Dim dataStrs As String() = dataStr.Split(New String() {"}]}"}, StringSplitOptions.RemoveEmptyEntries)
        Dim runShoeStr As String = dataStrs(0)
        Dim lifeShoeStr As String = dataStrs(1)
        Dim basketballShoeStr As String = dataStrs(2)
        Dim clothesStr As String = dataStrs(3)

        Dim sectionid As String = ""
        Select Case planType
            Case "HO"
                sectionid = "DA"
            Case "HA"
                sectionid = "NE"
        End Select

        Dim liProductId As New List(Of Integer)
        Dim runshoeProductid As Integer = InsertProductsByStr(runShoeStr, IssueID, siteId, "paobuxie", sectionid)
        Dim lifeShoeProductid As Integer = InsertProductsByStr(lifeShoeStr, IssueID, siteId, "xunlianxie", sectionid)
        Dim baskeballProductid As Integer = InsertProductsByStr(basketballShoeStr, IssueID, siteId, "lanqiuxie", sectionid)
        Dim clothesProductid As Integer = InsertProductsByStr(clothesStr, IssueID, siteId, "xunlianfu", sectionid)
        If (runshoeProductid <> 0) Then
            liProductId.Add(runshoeProductid)

        End If
        If (lifeShoeProductid <> 0) Then
            liProductId.Add(lifeShoeProductid)
        End If
        If (baskeballProductid <> 0) Then
            liProductId.Add(baskeballProductid)
        End If
        If (clothesProductid <> 0) Then
            liProductId.Add(clothesProductid)
        End If



        InsertProductsIssue(siteId, IssueID, sectionid, liProductId, liProductId.Count)

    End Sub

    Public Function InsertProductsByStr(ByVal str As String, ByVal issueId As Integer, ByVal siteid As Integer, ByVal categoryName As String, ByVal sectionid As String) As Integer
        Dim aProduct As New Product
        Dim productStrs As String() = GetproductStrs(str, "},")
        Dim productuid As String
        Dim productname As String = ""
        Dim i As Integer = 0
        Dim url As String = ""
        For i = 0 To 9
            productuid = GetMiddleStr(productStrs(i), """uid"":""", """,""")
            url = "http://www.e-lining.com/shop/goods-" & productuid & ".html"
            If (sectionid = "NE") Then
                productname = Regex.Unescape(GetMiddleStr(productStrs(i), """goodsName"":""", ""","""))
                If productname.Contains("新品") Then
                    If (efHelper.JudgeOneProduct(url, siteid, issueId, 21)) Then
                        Exit For
                    End If
                End If
            Else
                If (efHelper.JudgeOneProduct(url, siteid, issueId, 21)) Then
                    Exit For
                End If
            End If
        Next

        If efHelper.JudgeOneProductOfSite(url, siteid) <> -1 Then
            Return efHelper.JudgeOneProductOfSite(url, siteid)
        End If
        If i <= 9 Then
            aProduct.Prodouct = Regex.Unescape(GetMiddleStr(productStrs(i), """goodsName"":""", ""","""))
            aProduct.Url = url
            'Dim aProductDoc As HtmlDocument = efHelper.GetHtmlDocument(url, 60000)
            'Dim imgnode As HtmlNode = aProductDoc.DocumentNode.SelectSingleNode("//div[@class='imgshowbox']/div/a/img")
            aProduct.PictureUrl = "http://cdn.e-lining.com/postsystem/docroot/images/goods/" & GetMiddleStr(productStrs(i), """goodsThumbPic"":""", """,""").Replace("\", "")
            aProduct.PictureAlt = aProduct.Prodouct
            aProduct.Price = Double.Parse(GetMiddleStr(productStrs(i), """marketPrice"":""", ""","""))
            aProduct.Discount = Double.Parse(GetMiddleStr(productStrs(i), """salePrice"":""", ""","""))
            aProduct.LastUpdate = DateTime.Now
            aProduct.Description = aProduct.Prodouct
            aProduct.SiteID = siteid
            aProduct.Currency = "￥ "
            Dim productid As Long = efHelper.InsertSingleProduct(aProduct, categoryName, siteid)
            Return productid
            'efHelper.ins()
            'Dim categoryid = efHelper.GetCategoryId(siteId, CategoryName)
            'efHelper .i
        Else
            Return 0
        End If
    End Function

    Public Function GetproductStrs(ByVal OriStr As String, ByVal SplitStr As String) As String()
        Return OriStr.Split(New String() {SplitStr}, StringSplitOptions.RemoveEmptyEntries)
    End Function


    ''' <summary>
    ''' 获取两字符串中间的字符串
    ''' </summary>
    ''' <param name="OriStr"></param>
    ''' <param name="headStr"></param>
    ''' <param name="endStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMiddleStr(ByVal OriStr As String, ByVal headStr As String, ByVal endStr As String) As String
        Dim mc As MatchCollection
        Dim r As Regex = New Regex(("(?<=" _
                        + (headStr + (").*?(?=" _
                        + (endStr + ")")))), (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
        mc = r.Matches(OriStr)
        If (mc.Count > 0) Then
            Return mc(0).Value
        Else
            Return ""
        End If
    End Function

    Private Sub GetProducts(ByVal siteId As Integer, ByVal IssueId As Integer, ByVal categoryUrl As String, ByVal categoryName As String, ByVal iProIssueCount As Integer,
                            ByVal updateTime As DateTime, ByVal planType As String)
        Try
            Dim listProducts As List(Of Product) = GetListProducts(categoryUrl, siteId, updateTime, categoryName, planType, iProIssueCount)

            Dim listProduct As New List(Of Product)
            listProduct = efHelper.GetProductList(siteId)
            Dim listProductId As New List(Of Integer)

            Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
            efHelper.LogText("siteID:" & siteId & "categoryName:" & categoryName & "ID:" & categoryId)

            For Each li In listProducts
                Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listProduct)

                listProductId.Add(returnId)
                If (listProductId.Count = iProIssueCount) Then
                    Exit For
                End If
            Next
            InsertProductsIssue(siteId, IssueId, "CA", listProductId, iProIssueCount)
        Catch ex As Exception
            efHelper.LogText(ex.ToString())
        End Try
    End Sub

    Public Function GetListProducts(ByVal pageUrl As String, ByVal siteId As Integer, ByVal nowTime As DateTime, ByVal categoryName As String, ByVal planType As String,
                                  ByVal iProIssueCount As Integer) As List(Of Product)
        Dim listProducts As New List(Of Product)
        Dim helper As New EFHelper
        Dim endDate As String = Now.AddDays(1).ToShortDateString()
        Dim startDate As String = Now.AddDays(-proNoRepeatSpan).ToShortDateString()
        Try
            Dim hd As HtmlDocument = helper.GetHtmlDocument(pageUrl, 60000)
            Dim productDivs As HtmlNodeCollection
            productDivs = hd.DocumentNode.SelectNodes("//div[@id='s_l_b']/table/tbody/tr/td/div")
            For counter = 0 To productDivs.Count - 1
                Dim product As New Product()
                Dim productDetails As HtmlNode = productDivs(counter)

                product.Prodouct = productDetails.SelectSingleNode("div[@class='itemdescontent']/div[@class='description']/a").InnerText

                ''限定product-name的长度，既图片下方的产品文字描述，使得邮件样式更整齐
                'If (product.Prodouct.Length > 70) Then
                '    product.Prodouct = product.Prodouct.Substring(0, 70) & "..."
                'End If
                product.Url = efHelper.addDominForUrl(domin, productDetails.SelectSingleNode("div[@class='itemdescontent']/div[@class='description']/a").GetAttributeValue("href", "").Trim())
                If (listProductUrl.Contains(product.Url)) Then
                    Continue For
                Else
                    listProductUrl.Add(product.Url)
                End If
                Try
                    product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='itemdescontent']/div[@class='price pricediv']/div[@class='price offerprice']").InnerText.Replace("￥", "").Trim())
                Catch ex As Exception

                End Try
                'Try
                '    product.Discount = Double.Parse(productDetails.SelectSingleNode("div[@class='p_show_price']/div[@class='p_show_price_n']").InnerText.Replace("$", "").Trim())
                'Catch
                '    product.Discount = Double.Parse(productDetails.SelectSingleNode("div/div[3]/p[@class='p_price_n red']").InnerText.Replace("$", "").Trim())
                'End Try
                product.Currency = "￥"

                '特殊图片处理 pList2_mg
                Dim productImage As HtmlNode = productDetails.SelectSingleNode("div[1]/div")
                product.PictureUrl = efHelper.addDominForUrl(domin, productImage.SelectSingleNode("a/img").GetAttributeValue("src", ""))

                product.LastUpdate = nowTime
                product.Description = product.Prodouct
                product.SiteID = siteId
                product.PictureAlt = product.Prodouct
                If (Not efHelper.IsProductSent(siteId, product.Url, startDate, endDate, planType)) Then '控制新品不重复：
                    listProducts.Add(product)
                End If
                If (listProducts.Count = 2 * iProIssueCount) Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            Throw New Exception("failed when get products from the pageUrl:" & pageUrl & " categoryName:" & categoryName & " error:" & ex.Message.ToString)
        End Try
        Return listProducts
    End Function

    Public Sub AddIssueSubject(ByVal issueId As Integer, ByVal siteId As Integer, ByVal planType As String, ByVal categoryName As String, ByVal spreadLogin As String)
        Dim addNum As Integer
        Dim preSubject As String
        If (planType.Contains("HO")) Then
            preSubject = "Hi [FIRSTNAME]你好,本季热销推荐：+[某个产品名称] " 'Sammydress Deals For Jan.2014.Vol.01
        ElseIf (planType.Contains("HA")) Then
            preSubject = "让改变发生，新品推荐: " 'Sammydress New Arrivals For Jan.2014.Vol.01
        End If
        Dim nowYear As String = Date.Now.ToString("yyyy")

        '对于个性化邮件，有很多HO1,HO2,HO3。。。，对于他们的标题的计数，应根据plantype=‘HO’/‘HA’来计算，
        If (planType.Trim = "HO" OrElse planType.Trim = "HA" OrElse planType.Trim.Contains("HP")) Then
            If (spreadLogin.ToLower.Trim = "lihuojian@hofan.cn") Then
                If (planType.Trim.Contains("HP")) Then
                    addNum = 1
                Else
                    addNum = 8
                End If
            Else
                addNum = 1
            End If
        Else
            addNum = 0
        End If

        If (planType.Contains("HO") OrElse planType.Contains("HA")) Then
            planType = planType.Substring(0, 2)
        End If
        Dim query = From i In efContext.Issues
                     Where Year(i.IssueDate) = nowYear AndAlso i.Subject <> "" AndAlso i.SentStatus = "ES" And i.SiteID = siteId And i.PlanType = planType
                     Select i
        Dim subject As String = preSubject & DateTime.Now.ToString("MMM.yyyy") & ".Vol." & (query.Count + addNum).ToString.PadLeft(2, "0")

        Dim myIssue = efContext.Issues.Single(Function(m) m.IssueID = issueId)
        myIssue.Subject = subject
        efContext.SaveChanges()
    End Sub

    Public Sub InsertAdsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal listAdId As List(Of Integer), ByVal iProIssueCount As Integer)
        Dim i As Integer = 0
        For Each li In listAdId
            If i < iProIssueCount Then
                Dim aIssue As New Ads_Issue
                aIssue.AdId = li
                aIssue.SiteId = siteId
                aIssue.IssueID = issueId
                efContext.AddToAds_Issue(aIssue)
                i = i + 1
            End If
            If (i >= iProIssueCount) Then
                Exit For
            End If
        Next
        efContext.SaveChanges()
    End Sub

    Private Sub InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
                                    ByVal iProIssueCount As Integer)

        Dim queryProductId = (From pro In efContext.Products_Issue
                           Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId
                           Select pro.ProductId).ToList() '同一个产品分属于不同的类别，避免同一个issue获取到相同的产品
        Dim i As Integer = 0
        For Each li In listProductId
            If i < iProIssueCount AndAlso Not (queryProductId.Contains(li)) Then
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

    Private Function sortTopWeeklyDealProducts(ByRef listProducts As List(Of Product), ByVal topCount As Integer)
        '交换排序，按照产品价格product.Discount从大到小排序
        Dim maxProduct As New Product()
        Dim temp As New Product()
        Dim k As Integer
        Try
            '只排好价格product.Discount最高的前topCount个产品，其余的不关心
            For i As Integer = 0 To topCount - 1
                maxProduct = listProducts(i)
                k = i
                For j As Integer = i + 1 To listProducts.Count - 1
                    If (listProducts(j).Discount > maxProduct.Discount) Then
                        k = j
                        maxProduct = listProducts(k)
                    End If
                Next
                If Not (k = i) Then
                    temp = listProducts(i)
                    listProducts(i) = maxProduct
                    listProducts(k) = temp
                End If
            Next
        Catch ex As Exception
            Throw New Exception("error occured when sort products by discount ")
        End Try
    End Function

    Private Function getCategoryName(ByVal Url As String) As String
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(Url, 60000)
        Dim productDivs As HtmlNode = hd.DocumentNode.SelectSingleNode("//div[@class='breadcrumbs']/ul/li[@class='home']") 'hd.DocumentNode.SelectNodes("//div[@id='g_pro']/div")
        Dim dfdfdf As String = productDivs.InnerText.Replace("&amp;", "#").Replace("&gt;", "@")
        Dim dfa As String() = dfdfdf.Split("@")
        Dim categoryName As String = dfa(1).Replace("#", "&").Trim()
        Return categoryName
    End Function


End Class
