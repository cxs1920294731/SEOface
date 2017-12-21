Imports Analysis
Imports Newtonsoft.Json.Linq

Public Class TokenManage
    Inherits System.Web.UI.Page
    Dim entity As New FaceBookForSEOEntities()
    Dim emailAlerterEntity As New EmailAlerterEntitiesForRefreshFBToken
    Dim efHelper As New EFHelper()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'tokenId=1-->facebook的token信息，tokenId=2-->新浪微博的token信息
        Dim myToken As Token = (From t In entity.Tokens
                                Where t.id = 1
                                Select t).FirstOrDefault()
        txtLongTokenExpriedin.Text = myToken.longTokenExpireDate
    End Sub



    Protected Sub btnGetNewLongtoken_Click(sender As Object, e As EventArgs) Handles btnGetNewLongtoken.Click
        If Not (String.IsNullOrEmpty(txtShortToken.Text)) Then
            Dim requestUrl As String = "https://graph.facebook.com/v2.3/oauth/access_token?grant_type=fb_exchange_token&client_id=874014232612712&client_secret=01e0ec489cb4a17546e2f0674242db69&fb_exchange_token="
            requestUrl = requestUrl & txtShortToken.Text.Trim
            Dim resultStr As String
            Try
                resultStr = EFHelper.GetHtmlStringByUrl(requestUrl, 5) ''GetHtmlStringByUrl(requestUrl, "", "", Encoding.UTF8) ''modified by alex.
            Catch ex As Exception
                resultStr = EFHelper.GetHtmlStringByUrl(requestUrl, 5) ''GetHtmlStringByUrl(requestUrl, "", "", Encoding.UTF8) ''modified by alex.
            End Try

            Dim tokenJobject As JObject = JObject.Parse(resultStr)
            Dim longToken As String = tokenJobject("access_token").ToString.Trim
            lbllongTimeToken.Text = longToken
            Dim myToken As Token = (From t In entity.Tokens
                               Where t.id = 1
                               Select t).FirstOrDefault()
            myToken.longTokenExpireDate = Date.Now.AddDays(60)
            myToken.longTimeToken = longToken
            entity.SaveChanges()
            Common.LogText("k11 facebook success!")
            Response.Write("k11 facebook success!")
            Dim myEmailalerterToken As EmailAlerterToken = (From t In emailAlerterEntity.EmailAlerterTokens
                               Where t.id = 1
                               Select t).FirstOrDefault()
            myEmailalerterToken.longTokenExpireDate = Date.Now.AddDays(60)
            myEmailalerterToken.longTimeToken = longToken
            emailAlerterEntity.SaveChanges()
            Response.Write("emailalerter success!")
        End If
    End Sub
End Class