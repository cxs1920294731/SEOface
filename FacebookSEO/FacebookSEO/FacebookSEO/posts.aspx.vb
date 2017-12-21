Public Class posts
    Inherits System.Web.UI.Page
    Dim entity As New EmailAlerterEntities()
    Dim i As Integer = 0
    Public siteid As Integer
    Public head As String = ""
    Public header As String = ""
    Public footer As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim hostname As String = Request.ServerVariables("server_name")
        If (hostname.ToLower().Contains("k11")) Then
            siteid = 1
        Else
            siteid = 74 '本地测试环境
        End If
        'head = (From t In entity.Templates
        '   Where t.SiteID = siteid And t.PlanType = "HA1"
        '   Select t.Contents).FirstOrDefault()

        'header = (From t In entity.Templates
        '        Where t.SiteID = siteid And t.PlanType = "HA2"
        '        Select t.Contents).FirstOrDefault()

        'footer = (From t In entity.Templates
        '        Where t.SiteID = siteid And t.PlanType = "HA3"
        '        Select t.Contents).FirstOrDefault()
        Dim products As List(Of Product) = (From p In entity.Products
                                            Where p.SiteID = siteid Order By p.ProdouctID Descending
                                            Select p).ToList()
        ReItems.DataSource = products
        ReItems.DataBind()
    End Sub

    Protected Sub ReItems_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles ReItems.ItemDataBound
        If (i Mod 3 = 0) Then
            e.Item.Controls.Add(New LiteralControl("</TR><TR>"))
        End If
        i = i + 1
    End Sub

End Class