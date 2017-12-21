Imports System.Text.RegularExpressions
Imports Analysis
Imports Newtonsoft.Json.Linq

<DllType("fb")>
Public Class FbDllType
    Inherits ProductStart

    Protected pictureUrl As String = "https://graph.facebook.com/v2.10/{0}?access_token={1}&fields=full_picture,permalink_url&format=json"
    Protected limitUrl As String = "https://graph.facebook.com/v2.3/{0}?access_token={1}&fields=posts.limit({2})&format=json"
    Protected tokenUrl As String = "https://graph.facebook.com/v2.3/oauth/access_token?grant_type=fb_exchange_token&client_id=874014232612712&client_secret=01e0ec489cb4a17546e2f0674242db69&fb_exchange_token={0}"
    Protected sinceUrl As String = "https://graph.facebook.com/v2.3/{0}/posts?access_token={1}&since={2}&format=json"
    Protected attachmentsUrl As String = "https://graph.facebook.com/v2.3/{0}/attachments?access_token={1}&format=json"

    Protected efHelper As New EFHelper
    Protected efContext As New EmailAlerterEntities()

    Protected Overridable Property FbPage As String
    Protected Overridable Property Limit As String

    Public Sub New(issues As Integer)
        MyBase.New(issues)
    End Sub

    Public Overrides Sub Start(list As Subscriptions)



    End Sub

    Public Overridable Sub FetchCategoryProduct(ByVal siteid As Integer, ByVal planType As String, ByVal issueId As Integer)
        Dim listProPath As List(Of ProductPath) = (From proPath In efContext.ProductPaths
                                                   Where proPath.siteId = siteid AndAlso proPath.planType = planType
                                                   Select proPath).ToList()
        Dim listProduct As List(Of Product) = FetchfbPosts(FbPage, Limit, UpdateToken(), siteid)
        For Each item As ProductPath In listProPath
            Dim cate As Category = (From c In efContext.Categories
                                    Where c.SiteID = siteid AndAlso c.CategoryID = item.prodcate
                                    Select c).FirstOrDefault()

            If cate.Category1.ToLower = "hotclick" Then
                '往期热门，后面从数据库中提取，不需要爬站点
                Continue For
            End If

            Dim mylistProduct As List(Of Product) = GetListProducts(listProduct, cate, item)
            Dim listProductId As List(Of Integer) = efHelper.insertProducts(mylistProduct, cate.Category1, "CA", planType, item.prodDisplayCount, siteid, issueId)
            'InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount)
            ProductIssueDAL.InsertProductsIssue(siteid, issueId, "CA", listProductId, item.prodDisplayCount, cate.CategoryID)

        Next
    End Sub
    Public Overridable Function GetListProducts(listProduct As List(Of Product), cate As Category, item As ProductPath) As List(Of Product)
        Return New List(Of Product)
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
    Public Overridable Function FetchfbPosts(ByVal fbPageName As String, ByVal limit As Integer, ByVal accessToken As String, ByVal siteID As Integer) As List(Of Product)
        Dim listProduct As New List(Of Product)
        Dim postJarr As JArray = GetfbPosts(fbPageName, limit, accessToken)
        For i As Integer = 0 To postJarr.Count - 1
            Dim item As JObject = postJarr(i)
            Dim myPro As New Product()

            OnFetching(item, myPro)

            If Not (item.TryGetValue("message", myPro.Description)) Then
                Continue For
            End If

            Dim photoid As String = ""
            'If Not (item.TryGetValue("object_id", photoid)) Then
            '    Continue For
            'End If

            If Not (item.TryGetValue("id", photoid)) Then
                Continue For
            End If

            Dim photoJson As JObject = GetfbPhotobyid(photoid, accessToken)
            If Not (photoJson.TryGetValue("full_picture", myPro.PictureUrl)) Then
                Continue For
            End If

            item.TryGetValue("created_time", myPro.ExpiredDate)
            photoJson.TryGetValue("permalink_url", myPro.Url)

            myPro.Description = efHelper.AddfbPosttextHyperlink(myPro.Description)
            If (myPro.Description.Contains(vbLf)) Then
                myPro.PictureAlt = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                myPro.Description = myPro.Description.Remove(0, myPro.Description.IndexOf(vbLf))
            End If
            myPro.Description = myPro.Description.Replace(vbLf, "</br>")
            myPro.Description = myPro.Description
            myPro.SiteID = siteID
            myPro.LastUpdate = DateTime.Now

            OnFetched(item, myPro)

            If Not (efHelper.IsProductSent(siteID, myPro.Url, Now.AddDays(-100), Now)) Then
                listProduct.Add(myPro)
            End If
        Next
        Return listProduct
    End Function
    Public Overridable Sub OnFetching(ByRef item As JObject, product As Product)

    End Sub
    Public Overridable Sub OnFetched(ByRef item As JObject, product As Product)

    End Sub

    Protected Function UpdateToken() As String
        EFHelper.UpdateFbToken()
        Return efHelper.GetLongTimeToken(1)
    End Function
    ''' <summary>
    ''' 根据id获取一条photo的数据
    ''' </summary>
    ''' <param name="id">相片的id</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以jObject返回相片数据</returns>
    ''' <remarks></remarks>
    Public Function GetfbPhotobyid(ByVal id As String, ByVal accessToken As String) As JObject
        Dim requestUrl As String = String.Format(pictureUrl, id, accessToken)
        Dim photoStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")

        Dim photoJson As JObject = JObject.Parse(photoStr)
        Return photoJson
    End Function
    ''' <summary>
    ''' 根据id获取所有attachments的数据
    ''' </summary>
    ''' <param name="id">相片的id</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以jObject返回相片数据</returns> 
    ''' <remarks>method=GET&path=ID/attachments&version=v2.7</remarks>
    Public Function GetAttachmentsbyid(ByVal id As String, ByVal accessToken As String) As JObject
        Dim requestUrl As String = String.Format(attachmentsUrl, id, accessToken)
        Dim imageStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")
        Dim imageJson As JObject = JObject.Parse(imageStr)
        Return imageJson
    End Function
    ''' <summary>
    ''' 获取一个fb页面的post的Json数据
    ''' </summary>
    ''' <param name="fbPageName">fb页面名字</param>
    ''' <param name="limit">int，获取几个最新的post</param>
    ''' <param name="accessToken">token</param>
    ''' <returns>以Jarry返回post数据</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetfbPosts(ByVal fbPageName As String, ByVal limit As Integer, ByVal accessToken As String) As JArray
        Dim requestUrl As String = "https://graph.facebook.com/v2.3/" & fbPageName & "?access_token=" & accessToken & "&fields=posts.limit(" & limit & ")&format=json"

        Dim postsStr As String = EFHelper.GetHtmlStringByUrl(requestUrl, "", "", "")

        Dim postsJson As JObject = JObject.Parse(postsStr)
        Dim postsJArr As JArray = postsJson("posts")("data")
        Return postsJArr
    End Function
End Class
