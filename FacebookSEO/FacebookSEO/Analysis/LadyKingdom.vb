Imports System.Configuration
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.Web.Script.Serialization


Public Class LadyKingdom
    Private efContext As New EmailAlerterEntities()
    Private efHelper As New EFHelper()


    Public Sub Start(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                ByVal splitContactCount As Integer, ByVal spreadLogin As String, ByVal appId As String, ByVal url As String, ByVal categories As String)

        If planType.Trim = "HO" Then '热卖
            GetCategory(siteId)

            Dim accessToken As String = efHelper.GetLongTimeToken(1)
            Dim listProduct As List(Of Product) = efHelper.FetchfbPosts("LadyKingdomHK", 7, accessToken, siteId)
            ShortenPostDescription(listProduct)
            Dim listProudctid As List(Of Integer)
            listProudctid = efHelper.insertProducts(listProduct, "ladyKingdom", "CA", planType, 5, siteId, IssueID)
            efHelper.InsertProductsIssue(siteId, IssueID, "CA", listProudctid, 5)

            Dim querySubject As String = (From p In efContext.Products
                         Join pi In efContext.Products_Issue On p.ProdouctID Equals pi.ProductId
                         Where pi.IssueID = IssueID AndAlso pi.SectionID = "CA" AndAlso pi.SiteId = siteId
                         Select p.PictureAlt).FirstOrDefault()
            efHelper.InsertIssueSubject(IssueID, "Lady Kingdom 週三美容快遞:" & querySubject & " ...")
        End If
    End Sub

    Private Shared Sub GetCategory(ByVal siteId As Integer)
        Dim lastUpdate As DateTime = Now
        Dim helper As New EFHelper
        Dim myCategory As New Category
        myCategory.Category1 = "ladyKingdom"
        myCategory.Description = "ladyKingdom"
        myCategory.Url = "https://www.facebook.com/LadyKingdomHK/timeline"  'changed at 20141009,by dora, "www." -> "zh-hk."
        myCategory.SiteID = siteId
        myCategory.LastUpdate = lastUpdate
        helper.InsertOrUpdateCate(myCategory, siteId)
    End Sub

    Public Function ShortenPostDescription(ByRef listPost As List(Of Product))
        For Each item In listPost
            If (item.Description.Length > 150) Then
                item.Description = item.Description.Substring(0, 150) & "..."
            End If
        Next
    End Function
End Class

