
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Net
Imports System.Configuration
Imports Newtonsoft.Json.Linq

Public Class K11forFBSeo
    Private efHelper As EFHelper = New EFHelper()
    Public siteUrl As String
    Private _filepath As String
    Private _fileName As String

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String,
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)
        siteUrl = url
        _filepath = ConfigurationManager.AppSettings("imgFilePath").ToString.Trim
        _fileName = ConfigurationManager.AppSettings("imgFileName").ToString.Trim
        Common.LogText("K11 start")
        EFHelper.LogText("k11 start!")
        'GetCategory(siteId, url)
        Dim accessToken As String = efHelper.GetLongTimeToken(1)
        Dim fbPageName As String = GetfbPageName(url.Trim)
        EFHelper.LogText("start fbpagename:" & fbPageName & "  fbsiteid:" & siteId)
        Dim maxLimit As Integer = 50 '250 
        Dim listProduct As New List(Of Product)
        listProduct = FetchfbPosts(fbPageName, maxLimit, accessToken, siteId)
        Common.LogText(listProduct.Count.ToString + "收录的条数")
        efHelper.insertK11Products(listProduct, "groupbuyer", "CA", planType, maxLimit, siteId, IssueID)
        'efHelper.insertProducts(listProduct, "k11", "CA", planType, maxLimit, siteId, IssueID)
    End Sub
    Private Shared Sub GetCategory(ByVal siteId As Integer, ByVal siteUrl As String)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "k11"
        myCategory.Description = "k11"
        myCategory.Url = siteUrl
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Private Function GetfbPageName(ByVal fblink As String) As String
        Dim fbPageName As String
        If (fblink.ToLower.Contains("timeline")) Then
            fblink = fblink.Replace("/timeline", "")
        End If
        Dim index As Integer = fblink.LastIndexOf("/")
        Dim endIndex As Integer = fblink.IndexOf("?")
        If (endIndex > 0) Then
            fbPageName = fblink.Substring(index + 1, endIndex - index - 1)
        Else
            fbPageName = fblink.Substring(index + 1)
        End If
        Return fbPageName
    End Function

    ''' <summary>
    ''' 获取fb的一个page上发布的post，并返回List(Of Product)
    ''' </summary>
    ''' <param name="fbPageName">fb的page名</param>
    ''' <param name="limit">指定获取前几个post（按从新到旧）</param>
    ''' <param name="accessToken"></param>
    ''' <param name="siteID"></param>
    ''' <param name="planType"></param>
    ''' <param name="issueID"></param>
    ''' <param name="categoryName">产品所属分类名</param>
    ''' <param name="section">"CA","DA","NE"等</param>
    ''' <returns>List(Of Product)</returns>
    ''' <remarks></remarks>
    Public Function FetchfbPosts(ByVal fbPageName As String, ByVal limit As Integer, ByVal accessToken As String, ByVal siteID As Integer) As List(Of Product)
        Dim listProduct As New List(Of Product)
        Dim postJarr As JArray = efHelper.GetfbPosts(fbPageName, limit, accessToken)
        Common.LogText("返回的json数据" + postJarr.ToString())
        Common.LogText(postJarr.Count.ToString())
        For i As Integer = 0 To postJarr.Count - 1
            Dim item As JObject = postJarr(i)
            Try
                Dim typeAndLink As JObject = efHelper.GetfbTypeAndLink(item("id").ToString.Trim, accessToken)
                Dim type As String = typeAndLink("type").ToString.Trim()

                If (type = "photo") Then
                    'If (1) Then 
                    Dim myPro As New Product()

                    myPro.Description = item("message").ToString.Trim()
                    item.TryGetValue("created_time", myPro.ExpiredDate)
                    'myPro.Url = item("link").ToString.Trim
                    myPro.Url = typeAndLink("link").ToString.Trim

                    '不需要post文中的链接，减少外链
                    'Dim rmatch As Match = Regex.Match(myPro.Description, "http:\S+\s?")
                    'If (rmatch.Success) Then
                    '    myPro.Prodouct = rmatch.Value.Trim 'post文中的超链接（“立即购买”按钮）
                    'Else
                    '    myPro.Prodouct = myPro.Url
                    'End If

                    If (myPro.Description.Contains(vbLf)) Then
                        myPro.PictureAlt = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                        myPro.Description = myPro.Description.Remove(0, myPro.Description.IndexOf(vbLf))
                    End If
                    myPro.Description = myPro.Description.Replace(vbLf, "</br>")

                    myPro.SiteID = siteID
                    myPro.LastUpdate = DateTime.Now
                    myPro.Currency = ""

                    Dim photoid As String = item("id").ToString.Trim
                    If photoid IsNot Nothing Then
                        myPro.FreeShippingImg = photoid '借此FreeShippingImg位用于标志id,避免重复
                    End If
                    Dim photoJson As JObject = efHelper.GetfbPhotobyid(photoid, accessToken)
                    myPro.PictureUrl = DownloadImage(photoJson("full_picture").ToString.Trim, _filepath, _fileName, siteID) '高清大图
                    Dim littleImageUrl As String
                    'Dim littleImage As JArray = photoJson("images")

                    'Dim index As Integer = 2 '获取处于第3个位置的图片，其像素大小适中
                    'If (littleImage.Count - 1 < 2) Then
                    'Index = littleImage.Count - 1
                    'End If
                    'littleImageUrl = littleImage(index)("source").ToString.Trim()
                    littleImageUrl = photoJson("full_picture").ToString.Trim 'by --Roy
                    myPro.Prodouct = DownloadImage(littleImageUrl, _filepath, _fileName, siteID) '小图
                    If (Not String.IsNullOrEmpty(myPro.PictureUrl) AndAlso Not String.IsNullOrEmpty(myPro.Prodouct)) Then '如没有获取到图片，则不收录此post
                        listProduct.Add(myPro)
                    End If
                End If
            Catch ex As Exception
                'NEXT
            End Try
        Next
        Return listProduct
    End Function


    Public Function GetPost(ByRef listPostId As List(Of String), ByVal siteid As Integer, ByVal iProIssueCount As Integer) As List(Of Product)
        Dim postUrl As String
        Dim categoryDoc As String
        Dim postCreateTime As Date
        Dim listProduct As New List(Of Product)
        For Each postId In listPostId
            efHelper.LogText("start request postid " & postId)
            postUrl = "https://graph.facebook.com/" & postId
            Try
                'categoryDoc = GetHtmlString(postUrl, siteUrl).Replace("\n", "[mylittleN]")
            Catch ex As Exception
                Continue For
            End Try
            Dim jss As New JavaScriptSerializer()
            Dim json = jss.DeserializeObject(categoryDoc)

            postCreateTime = json("created_time").ToString.Remove(json("created_time").ToString.ToUpper.IndexOf("T"))

            Dim product As New Product
            product.ExpiredDate = postCreateTime
            '当发布的post是一个视频文件而不是图片时，其描述的字段名是"description"而不是"name"
            '所以当捕捉到KeyNotFoundException时，跳过此post

            Try
                product.Description = Regex.Unescape(json("name").ToString.Trim).Replace("[mylittleN]", " </br>")
                product.Url = json("link").ToString.Trim 'product.Url阅读更多按钮的链接
            Catch ex As KeyNotFoundException
                Try
                    product.Description = Regex.Unescape(json("description").ToString.Trim).Replace("[mylittleN]", " </br>")
                Catch ex1 As Exception
                    Continue For
                End Try
            End Try

            If (product.Description.Contains("</br>")) Then
                product.PictureAlt = product.Description.Remove(product.Description.IndexOf("</br>"))
                product.Description = product.Description.Remove(0, product.Description.IndexOf("</br>"))
            End If
            product.Currency = ""
            product.PictureUrl = DownloadImage(json("source").ToString.Trim, _filepath, _fileName, siteid) '高清大图
            Dim littleImageUrl As String
            Dim littleImage As JArray = JObject.Parse(categoryDoc)("images")
            Dim index As Integer = 2
            If (littleImage.Count - 1 < 2) Then
                index = littleImage.Count - 1
            End If
            littleImageUrl = littleImage(index)("source").ToString.Trim() '数组中的最后一张图是倒数第二小的
            product.Prodouct = DownloadImage(littleImageUrl, _filepath, _fileName, siteid) '小图
            If (Not String.IsNullOrEmpty(product.PictureUrl) AndAlso Not String.IsNullOrEmpty(product.Prodouct)) Then '如没有获取到图片，则不收录此post
                product.SiteID = siteid
                product.LastUpdate = Now
                listProduct.Add(product)
            End If
        Next
        Return listProduct
    End Function



    Private Function DownloadImage(ByVal imageUrl As String, ByVal filePath As String, ByVal fileName As String, ByVal siteid As Integer) As String
        Dim client As New WebClient()
#If DEBUG Then
        '+debug# proxy
        'Dim myProxy As New WebProxy("127.0.0.1:1080", True)
        'client.Proxy = myProxy
#End If
        Dim myProxy As New WebProxy("127.0.0.1:1080", True)
        client.Proxy = myProxy
        Dim index As Integer = imageUrl.LastIndexOf("/")
        Dim imageName As String 'imageurl中第一个参数值+name作为imagename
        imageName = imageUrl.Substring(index + 1)
        Dim paramIndex As Integer = imageName.IndexOf("?")
        If (paramIndex > 0) Then
            'Dim params As String() = imageName.Split("=")
            imageName = imageName.Substring(0, paramIndex)
        End If
        imageName = Now.Day & "_" & Now.Minute & "_" & Now.Second & "_" & Now.Millisecond & "_" & imageName

        If Not (filePath.EndsWith("\")) Then
            filePath = filePath & "\"
        End If
        Dim datetime As String = Now.Year & "_" & Now.Month
        '以店铺名年月划分文件夹
        filePath = filePath & fileName & "\" & siteid & "_" & datetime
        If Not IO.Directory.Exists(filePath) Then IO.Directory.CreateDirectory(filePath)

        filePath = filePath & "\" & imageName
        Try
            client.DownloadFile(imageUrl, filePath)
            Return "/" & fileName & "/" & siteid & "_" & datetime & "/" & imageName
        Catch ex As Exception
            EFHelper.Log(ex)
            Return ""
        End Try
    End Function
End Class
