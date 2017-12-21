Imports System.Net.Mail

Public Class Post1
    Inherits System.Web.UI.Page
    Dim entity As New FaceBookForSEOEntities()
    Public product As New Product()

    Public sitename As String
    Public shopUrl As String = ""

    Public header As String = ""
    Public footer As String = ""

    Public prePostUrl As String = ""
    Public nextPostUrl As String = ""
    Public pageTitle As String = ""

    Public shopsite As AutomationSite = New AutomationSite()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim siteid As Integer = Request.QueryString("siteid")
        Dim productID As String = Request.QueryString("id")
        If (String.IsNullOrEmpty(productID)) Then
            Response.Redirect("Default.aspx")
        End If
        product = (From p In entity.Products
                   Where (p.ProdouctID = productID)
                   Select p).FirstOrDefault()

        Dim siteid As Integer = product.SiteID
        shopsite = (From autosite In entity.AutomationSites
                    Where (autosite.siteid = siteid)
                    Select autosite).FirstOrDefault()

        Dim hostname As String = Request.ServerVariables("server_name")
        Dim msiteid As Integer '店铺其隶属于的商场siteid,即parentid
        Common.LogText("hostName:" & hostname)
        If (hostname.ToLower().Contains("k11")) Then
            msiteid = Integer.Parse(ConfigurationManager.AppSettings("k11ID").ToString().Trim())
        Else
            msiteid = 168 '本地测试环境
        End If

        Dim pageTitlek11 As String = "<a href ='/' style='color:#3e3e3e; text-decoration:none;font-weight :bold ;'> 香港groupbuyer動態</a>  "
        Dim pageTitlek11Sc As String = "<a href ='/tst' style='color:#3e3e3e; text-decoration:none;font-weight :bold ;'> 香港groupbuyer动态</a>  "
        Dim pageTitleAllshop As String = "|&nbsp; <a href='/AllShops'   style='color:#3e3e3e;text-decoration:none;font-weight :bolder ;font-size: 18px;'>商店一覽</a> &gt; &nbsp; "
        Dim pageTitleAllshopSc As String = "|&nbsp; <a href='/tst/AllShops'   style='color:#3e3e3e;text-decoration:none;font-weight :bolder ;font-size: 18px;'>商店一览</a> &gt; &nbsp; "
        Dim pageTitleShopName As String = "&gt; &nbsp;<a href='[URL]'  style='color:#3e3e3e; text-decoration:none;font-weight :bold ;' > [SITENAME] </a>&nbsp;  "
        Dim pageTitleSpan As String = "&gt; &nbsp;<span  style='color:#3e3e3e;font-weight :bold ;' > [SPANNAME] </a>&nbsp;  "

        Dim currency As String = ""
        If (product.Currency.Trim = "WB") Then
            header = (From t In entity.Templates
                   Where t.SiteID = msiteid And t.TemplateName = "HeaderSC"
                   Select t.Contents).FirstOrDefault()

            footer = (From t In entity.Templates
                    Where t.SiteID = msiteid And t.TemplateName = "FooterSC"
                    Select t.Contents).FirstOrDefault()
            

            currency = product.Currency.Trim
            sitename = shopsite.SiteNameSc
            pageTitle = pageTitlek11Sc
        Else
            header = (From t In entity.Templates
                    Where t.SiteID = msiteid And t.TemplateName = "Header"
                    Select t.Contents).FirstOrDefault()

            footer = (From t In entity.Templates
                    Where t.SiteID = msiteid And t.TemplateName = "Footer"
                    Select t.Contents).FirstOrDefault()
            

            sitename = shopsite.SiteName
            pageTitle = pageTitlek11
        End If
        pageTitle = pageTitle & pageTitleShopName
        pageTitle = pageTitle.Replace("[SITENAME]", sitename).Replace("[URL]", UrlValid.getUrlValid(shopsite.k11SiteUrl))
        If Not (String.IsNullOrEmpty(product.PictureAlt)) Then
            If (product.PictureAlt.Length > 15) Then
                Page.Title = product.PictureAlt.Substring(0, 14)
            Else
                Page.Title = product.PictureAlt
            End If
        End If
        If Not (String.IsNullOrEmpty(product.Description)) Then
            If (product.Description.Length > 60) Then
                Page.MetaDescription = product.Description.Substring(0, 59)
            Else
                Page.Title = product.Description
            End If
        End If

        Dim minProID As Integer = (From p In entity.Products
                               Where p.SiteID = siteid And p.Currency.Equals(currency)
                               Select p.ProdouctID).Min()
        Dim maxProID As Integer = (From p In entity.Products
                                   Where p.SiteID = siteid And p.Currency.Equals(currency)
                           Select p.ProdouctID).Max()
        '上一个post的信息
        'Dim preProID As Integer = productID - 1
        Dim preProduct As New Product
        If (productID = minProID) Then
            preProduct = (From p In entity.Products
                          Where p.ProdouctID = maxProID
                          Select p).FirstOrDefault()
        Else
            preProduct = (From p In entity.Products
                         Where (p.ProdouctID < productID And p.SiteID = siteid And p.Currency.Equals(currency))
                         Order By p.ProdouctID Descending
                         Select p).FirstOrDefault()
        End If
        prePostUrl = "/" & UrlValid.getUrlValid(sitename) & "/post/" & preProduct.ProdouctID & ".aspx"
        'hiddenPrePostUrl.Value = prePostUrl
        '下一个post的信息
        'Dim nextProID As Integer = productID + 1
        Dim nextProduct As New Product
        If (productID = maxProID) Then
            nextProduct = (From p In entity.Products
                        Where (p.ProdouctID = minProID)
                        Select p).FirstOrDefault()
        Else
            nextProduct = (From p In entity.Products
                        Where (p.ProdouctID > productID And p.SiteID = siteid And p.Currency.Equals(currency))
                        Order By p.ProdouctID
                        Select p).FirstOrDefault()
        End If
        nextPostUrl = "/" & UrlValid.getUrlValid(sitename) & "/post/" & nextProduct.ProdouctID & ".aspx"
    End Sub

End Class