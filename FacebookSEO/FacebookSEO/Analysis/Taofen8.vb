Imports HtmlAgilityPack
Imports System.Math
Imports System.Text.RegularExpressions

Public Class Taofen8
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private listProductUrl As New List(Of String)
    Private proNoRepeatSpan As Integer = 50 '产品不重复时间长（天为单位）
    Private domin As String = "http://www.taofen8.com/"


    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        If planType = "HO" Then
            GetCategory(siteId)
            ''GetBanner()每期获取一张Banner，逐期轮流抓取，顺时针抓取；
            ''GetHotProducts()获取超“超高返”部分的8个产品，一个月内去重 Section插入'DA',插入产品表的时候，需要计算出“价格差（负数）”插入产品表的Description字段，
            ''以及计算出“折扣”插入产品表的PictureAlt字段（插入前，先格式化为3个字符）

            ''GetSubject()生成当期的邮件标题，更新到Issue表里的Subject字段，Subject标题的生成规则如下：
            ''标题内容为：淘粉吧专题咨询，会员专享：[第一个产品的产品名称]
            GetBanner(siteId, IssueID, "热销", planType)
            GetHotProducts(siteId, IssueID, "热销", planType)

            Dim querySubject = From p In efContext.Products
           Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
           Where pi.IssueID = IssueID AndAlso pi.SectionID = "DA" AndAlso pi.SiteId = siteId
           Select p.Prodouct
            efHelper.InsertIssueSubject(IssueID, "淘粉吧专题咨询，会员专享：" & querySubject.FirstOrDefault.ToString() & " ...")

        End If
    End Sub

    Public Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "热销"
        myCategory.Description = "热销"
        myCategory.Url = ""
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Public Sub GetBanner(ByVal siteId As Integer, ByVal Issueid As Integer, ByVal categoryName As String, ByVal planType As String)
        '请求banner，拿到banner imgurl 判断重复，重复则拿下一个，以此类推，直到当获取的bannernode中的banner的imgurl都重复时，则取上一个
        'issue的banner，返回此banner的下一个banner
        Dim pageUrl As String = "http://www.taofen8.com/"
        Dim categoryId As Integer = efHelper.GetCategoryId(siteId, categoryName)
        Dim iProIssueCount As Integer = 1 '每次只获取一个Banner信息
        Dim time As DateTime = DateTime.Now
        Dim adid As Long = GetBannerid(pageUrl, siteId, categoryId, time, planType, Issueid)
        efHelper.InsertSingleAdsIssue(adid, siteId, Issueid)


    End Sub

    Public Function GetBannerid(ByVal pageUrl As String, ByVal siteid As Integer, ByVal categroyId As Integer, ByVal nowTime As DateTime, ByVal planType As String, ByVal issueid As Long) As Integer
        Dim adid As Integer
        Dim flagadid As Integer
        Dim endDate As String = Now
        Dim startDate As String = Now.AddDays(-18).ToShortDateString()
        'Dim listAds As New List(Of Ad)
        Dim hd As HtmlDocument = efHelper.GetHtmlDocument(pageUrl, 60000) 'slideBox mt10 mb10
        Dim banners As HtmlNodeCollection = hd.DocumentNode.SelectNodes("//div[@class='wrapper']/div/a")  '/div/div/a/
        'Dim listAdidNew As New List(Of Integer)
        For counter As Integer = 0 To banners.Count - 1
            Dim Ad As Ad = New Ad()
            Dim adpicurl As String = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString.Trim
            flagadid = efHelper.isADexist(adpicurl, siteid)
            If (flagadid = -1) Then
                Try
                    Ad.Ad1 = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                    Ad.Description = banners(counter).SelectSingleNode("img").GetAttributeValue("alt", "").ToString 'img的alt属性现还没有加入2013-01-09
                Catch ex As Exception
                    'ignore
                End Try
                Ad.Url = banners(counter).GetAttributeValue("href", "").ToString.Trim
                Ad.Url = efHelper.addDominForUrl(domin, Ad.Url)
                Ad.PictureUrl = banners(counter).SelectSingleNode("img").GetAttributeValue("src", "").ToString.Trim
                Ad.PictureUrl = efHelper.addDominForUrl(domin, Ad.PictureUrl)

                Ad.SiteID = siteid
                Ad.Lastupdate = nowTime

                If Not (efHelper.isAdSent(siteid, Ad, startDate, endDate)) Or counter = banners.Count - 1 Then '如果是最后一个banner了，就算重复也要发，否则导致程序后面发生错误！
                    adid = efHelper.InsertSingleAd(Ad, categroyId)
                    Return adid
                End If

                'listAdidNew.Add(adid)
            Else
                Dim existAd As Ad = (From banner In efContext.Ads
                                     Where banner.AdID = flagadid Select banner).FirstOrDefault

                If Not (efHelper.isAdSent(siteid, existAd, startDate, endDate)) Then
                    Return flagadid
                End If
            End If

            'If (Not efHelper.isAdSentInRecentSixIssue(siteid, Ad.Url, planType) AndAlso Ad.PictureUrl <> "") Then '控制内获取的Ad在最近6次发送的issue不重复不重复：
            '    adid = efHelper.JudgeOneAd(Ad.Url, siteid) = -1
            '    If (adid = -1) Then
            '        adid = efHelper.InsertAd(Ad, listAd, categroyId)
            '    End If
            '    Exit For
            'End If
        Next

    End Function

    Public Sub GetHotProducts(ByVal siteid As Integer, ByVal issueid As Long, ByVal categoryName As String, ByVal plantype As String)
        ''DEBUG #6.7 
        Dim pageUrl As String = "http://www.taofen8.com/"
        'Dim pageUrl As String = "http://www.taofen8.com/pptList"
        Dim doc As HtmlDocument
        doc = efHelper.GetHtmlDocument(pageUrl, 60000)
        Dim producturl As String = ""
        ''DEBUG #6.7 
        Dim productNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//ul[@class='tf8_sp_ul']/li")
        'Dim productNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[@class='ppt_list']/ul/li")
        Dim infoNodes As HtmlNodeCollection
        'Dim priceNodes As HtmlNodeCollection

        Dim i As Integer = 0
        Dim sectionid As String = ""
        Dim liProductId As New List(Of Integer)
        Dim pid As Integer
        Select Case plantype
            Case "HO"
                sectionid = "DA"
            Case "HA"
                sectionid = "NE"
        End Select
        For Each node In productNodes
            infoNodes = node.SelectNodes("a")
            'priceNodes = node.SelectNodes("div")
            producturl = "http://www.taofen8.com/item/" & infoNodes(0).GetAttributeValue("id", "").Replace("item_", "") & "?pageName=index_cgf_all"
            If Not (efHelper.IsProductSent(siteid, producturl, DateTime.Now.AddDays(-30), DateTime.Now)) Then
                Dim item As New Product
                item.Prodouct = System.Web.HttpUtility.HtmlDecode(infoNodes(2).InnerHtml.ToString.Trim)
                item.Url = producturl
                'item.Price = Double.Parse(priceNodes(priceNodes.Count - 2).SelectSingleNode("div/del").InnerText.ToString.Trim)
                item.Discount = Double.Parse(node.SelectSingleNode("div[@class='tf8-index-1']/div/span").InnerText.ToString.Trim)
                item.PictureUrl = infoNodes(1).SelectSingleNode("img").GetAttributeValue("original", "").ToString.Trim
                item.LastUpdate = DateTime.Now
                item.SiteID = siteid
                item.Currency = "￥ "
                'item.Price = node.SelectSingleNode("div[@class='tf8_tis']/div[@class='tf8_shop_span']").InnerText.ToString.Trim
                item.Description = node.SelectSingleNode("div[@class='tf8_tis']/div[@class='tf8_shop_span']").InnerText.ToString.Trim     '价格差
                item.PictureAlt = node.SelectSingleNode("div[@class='tf8_tis']/div[@class='change_price']").InnerText.ToString.Trim  '到手价
                Try
                    pid = efHelper.InsertSingleProduct(item, categoryName, siteid)
                Catch ex As Exception
                    EFHelper.LogText(ex.ToString)
                End Try
                If pid > 0 Then
                    liProductId.Add(pid)
                    i = i + 1
                    If (i >= 12) Then
                        Exit For
                    End If
                End If


            End If

        Next
        efHelper.InsertProductsIssue(siteid, issueid, sectionid, liProductId, liProductId.Count)
    End Sub


End Class
