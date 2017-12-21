
Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports System.Web.Script.Serialization
Public Class TravelExpert
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()
    Private subjectString As String

    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        Dim accessToken As String = efHelper.GetLongTimeToken(1)
        GetCategory(siteId)
        GetHotProducts(siteId, planType, IssueID, "TravelExpert", accessToken)

        Dim querySubject = (From p In efContext.Products
                     Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                     Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                     Select p.PictureAlt).FirstOrDefault
        If Not (querySubject Is Nothing) Then
            subjectString = querySubject
            efHelper.InsertIssueSubject(IssueID, subjectString)
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "TravelExpert"
        myCategory.Description = "TravelExpert"
        myCategory.Url = "https://www.facebook.com/TravelExpertHKG/timeline"
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Public Sub GetHotProducts(ByVal siteID As Integer, ByVal planType As String, ByVal issueID As Integer, ByVal categoryName As String, ByVal accessToken As String)
        Dim listProduct As New List(Of Product)
        Dim postJarr As JArray = efHelper.GetfbPosts("TravelExpertHKG", 5, accessToken)
        For i As Integer = 0 To postJarr.Count - 1
            Dim item As JObject = postJarr(i)
            Dim type As String = item("type").ToString.Trim()
            If (type = "photo") Then
                Dim myPro As New Product()

                If (item.TryGetValue("message", myPro.Description)) Then
                    If (myPro.Description.Contains("限時優惠")) Then
                        myPro.Url = item("link").ToString.Trim
                        Dim photoid As String = item("object_id").ToString.Trim
                        Dim photoJson As JObject = efHelper.GetfbPhotobyid(photoid, accessToken)
                        myPro.PictureUrl = photoJson("full_picture").ToString.Trim
                        'facebook分享链接
                        myPro.Prodouct = "http://www.facebook.com/sharer.php?s=100&p[url]=" & myPro.Url & "&p[images][0]=" & myPro.PictureUrl

                        myPro.Description = efHelper.AddfbPosttextHyperlink(myPro.Description)
                        If (myPro.Description.Contains(vbLf)) Then
                            myPro.PictureAlt = myPro.Description.Remove(myPro.Description.IndexOf(vbLf))
                            myPro.Description = myPro.Description.Remove(0, myPro.Description.IndexOf(vbLf))
                        End If
                        myPro.Description = myPro.Description.Replace(vbLf, "</br>")
                        If Not (efHelper.IsProductSent(siteID, myPro.Url, Now.AddDays(-100), Now)) Then
                            listProduct.Add(myPro)
                        End If
                    End If
                End If
            End If
        Next
        Dim listProudctid As New List(Of Integer)
        listProudctid = efHelper.insertProducts(listProduct, categoryName, "CA", planType, 1, siteID, issueID)
        efHelper.InsertProductsIssue(siteID, issueID, "CA", listProudctid, 1)
    End Sub

End Class
