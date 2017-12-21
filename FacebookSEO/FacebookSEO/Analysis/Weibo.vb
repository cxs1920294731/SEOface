Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Net
Imports System.Configuration
Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq
Imports System.Text

Public Class Weibo
    Private efHelper As EFHelper = New EFHelper()
    Public siteUrl As String
    Private _filepath As String
    Private _fileName As String


    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        siteUrl = url
        _filepath = ConfigurationManager.AppSettings("imgFilePath").ToString.Trim
        _fileName = ConfigurationManager.AppSettings("imgFileName").ToString.Trim

        efHelper.LogText("weibo start!")
        'GetCategory(siteId, url)
        GetHOTProducts(siteId, planType, IssueID, "k11")
    End Sub

    Private Sub GetHOTProducts(ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer, ByVal categoryName As String)


        Dim iproIssueCount As Integer = 5 '在模板中要填充产品的个数
        Dim listPostId As New List(Of String)

        Dim listProduct As New List(Of Product)
        '获取到post详细内容拼成一个product
        listProduct = GetPosts(siteID, siteUrl)

        Dim listQueryProduct As New List(Of Product)
        listQueryProduct = efHelper.GetProductList(siteID)
        'Dim listProductId As New List(Of Integer)
        Dim categoryId As Integer = efHelper.GetCategoryId(siteID, categoryName)

        For Each li In listProduct
            Dim returnId As Integer = efHelper.InsertProduct(li, Now, categoryId, listQueryProduct)
            'listProductId.Add(returnId)
        Next
    End Sub

    Public Function GetPosts(ByVal siteid As Integer, ByVal fromUrl As String) As List(Of Product)
        Dim listProd As New List(Of Product)
        Dim cookie As String = ConfigurationManager.AppSettings("WBcookie").ToString.Trim
        Dim htmlstring As String = efHelper.GetHtmlStringByUrlSinaWB(fromUrl, cookie)
        Dim wbDetail As String() = {"<div class=\""WB_detail\"">"}
        Dim wbpost As String() = htmlstring.Split(wbDetail, StringSplitOptions.RemoveEmptyEntries)
        If (wbpost.Count >= 2) Then
            For i As Integer = 1 To wbpost.Count - 1
                Dim postString As String = wbpost(i).Trim
                postString = postString.Substring(0, postString.Length - 7)
                postString = postString.Replace("\r", "").Replace("\n", "").Replace("\t", "").Replace("\", "")
                Dim postDoc As New HtmlDocument()
                postDoc.LoadHtml(postString)

                Dim myProd As New Product
                myProd.Description = postDoc.DocumentNode.SelectSingleNode("//div[@class='WB_text W_f14']").InnerText.Trim
                Dim myRegex As New Regex("#[\w\W]*?#")
                Dim m As Match = myRegex.Match(myProd.Description)
                If (m.Success) Then
                    myProd.PictureAlt = m.Value.Trim
                    myProd.Description = myProd.Description.Replace(myProd.PictureAlt, "")
                End If
                myProd.ExpiredDate = postDoc.DocumentNode.SelectSingleNode("//div[@class='WB_from S_txt2']/a[1]").GetAttributeValue("title", "").Trim
                myProd.Url = efHelper.addDominForUrl("http://weibo.com/", postDoc.DocumentNode.SelectSingleNode("//div[@class='WB_from S_txt2']/a[1]").GetAttributeValue("href", "").Trim)

                Dim imgNodes As HtmlNodeCollection = postDoc.DocumentNode.SelectNodes("//li[@class='WB_pic S_bg1 bigcursor']/img")
                If (imgNodes Is Nothing) Then
                    imgNodes = postDoc.DocumentNode.SelectNodes("//li[@class='WB_video WB_video_a S_bg1']/img")
                End If
                If (imgNodes Is Nothing) Then
                    Continue For
                End If
                Dim imgsrc As String = imgNodes(0).GetAttributeValue("src", "").ToString.Trim.Replace("thumbnail", "bmiddle").Replace("square", "bmiddle")

                myProd.PictureUrl = efHelper.DownloadImage(imgsrc, _filepath, _fileName, siteid)
                myProd.Prodouct = myProd.PictureUrl
                If Not (String.IsNullOrEmpty(myProd.PictureUrl)) Then
                    myProd.SiteID = siteid
                    myProd.LastUpdate = Now
                    myProd.Currency = "WB"
                    listProd.Add(myProd)
                End If
            Next
        Else
            efHelper.LogText("request reslult does not contain ""WB_detail""")
            efHelper.LogText(htmlstring)
        End If
        Return listProd
    End Function
End Class
