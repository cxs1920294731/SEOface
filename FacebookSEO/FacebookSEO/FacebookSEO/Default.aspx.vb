Imports Enumerable = System.Linq.Enumerable
Imports System.Configuration

Public Class _Default1
    Inherits System.Web.UI.Page
    Public entity As New FaceBookForSEOEntities()

    Public siteid As String
    Public msiteid As Integer '店铺其隶属于的商场siteid,即parentid

    Public header As String = ""
    Public footer As String = ""
    Dim pagesize As Integer

    Public pageTitle As String = ""

    Dim languageType As String = ""
    Public prodType As String = ""

    Public k11Description As String = ""

    Public parentCateTags As List(Of CateTag)
    Public chlidCateTags As List(Of CateTag)

    Public homePage As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim dfd As String() = Request.UserLanguages
        '获取主机名称
        Dim hostname As String = Request.ServerVariables("server_name")
        Common.LogText("hostName:" & hostname)
        If (hostname.ToLower().Contains("k11")) Then
            msiteid = Integer.Parse(ConfigurationManager.AppSettings("k11ID").ToString().Trim())
        Else
            msiteid = 168 '本地测试环境
        End If
        If Not (Request.QueryString("languageType") Is Nothing) Then
            languageType = Request.QueryString("languageType")
        End If

        If (languageType = "tst") Then
            prodType = "WB"
        End If
        '获取从url传的参数
        '传的是店铺id
        Dim pSiteid As String = Request.QueryString("siteid")
        Dim shopName As String = Request.QueryString("sitename")
        '标签ID
        Dim cateId As String = Request.QueryString("cateId")
        '关键词
        Dim keyWordId As String = Server.UrlDecode(Request.QueryString("keyWordId"))
        Dim KeyWord As String

        Dim templateName As String
        Dim k11siteUrl As String

        Dim defaultPageTitle As String = " GroupBuyer"
        Dim defaultPageTitleSc As String = "  GroupBuyer"
        Dim mateDescriptionSc As String = "香港Group Buyer不断为您争取最新，最抵最多的消费优惠，竭力把团购建立一种香港人·的生活习惯。"
        Dim mateDescription As String = "香港Group Buyer不斷為您爭取最新，最抵最多的消費優惠，竭力把團購建立壹種香港人·的生活習慣。"

        Dim pageTitlek11 As String = "<a href ='/' style='color:#3e3e3e; text-decoration:none;font-weight :bold ;'> 香港groupbuyer動態</a>  "
        Dim pageTitlek11Sc As String = "<a href ='/tst' style='color:#3e3e3e; text-decoration:none;font-weight :bold ;'> 香港groupbuyer动态</a>  "
        Dim pageTitleAllshop As String = "|&nbsp; <a href='/AllShops'   style='color:#3e3e3e;text-decoration:none;font-weight :bolder ;font-size: 18px;'>商店一覽</a> &gt; &nbsp; "
        Dim pageTitleAllshopSc As String = "|&nbsp; <a href='/tst/AllShops'   style='color:#3e3e3e;text-decoration:none;font-weight :bolder ;font-size: 18px;'>商店一览</a> &gt; &nbsp; "
        Dim pageTitleShopName As String = "&gt; &nbsp;<a href='[URL]'  style='color:#3e3e3e; text-decoration:none;font-weight :bold ;' > [SITENAME] </a>&nbsp;  "
        Dim pageTitleSpan As String = "&gt; &nbsp;<span  style='color:#3e3e3e;font-weight :bold ;' > [SPANNAME] </a>&nbsp;  "

        If (String.IsNullOrEmpty(pSiteid)) Then
            k11siteUrl = "#"
            If Not (String.IsNullOrEmpty(cateId)) Then
                Dim cateTag As CateTag = (From cate In entity.CateTags
                                          Where cate.ID = cateId
                                          Select cate).FirstOrDefault()
                Dim parentCateTag As CateTag = (From cate In entity.CateTags
                                                Where cate.ID = cateTag.ParentID
                                                Select cate).FirstOrDefault()
                If (cateTag.HasShopID Is Nothing) Then
                    siteid = "20" '20在数据库中是个不存在的店铺id
                Else
                    siteid = cateTag.HasShopID.Trim
                End If
                If (prodType = "WB") Then
                    AspNetPager1.EnableUrlRewriting = True
                    AspNetPager1.UrlRewritePattern = "/ctag/tst/" & cateTag.CateNameSC.Trim() & "/" & cateId & "{0}.aspx"
                    shopName = cateTag.CateNameSC
                    pageTitle = pageTitlek11Sc & pageTitleSpan.Replace("[SPANNAME]", parentCateTag.CateNameSC) & pageTitleSpan.Replace("[SPANNAME]", cateTag.CateNameSC)
                    templateName = "PageTitleSC"
                    Page.Title = parentCateTag.CateNameSC.Trim & "-" & cateTag.CateNameSC.Trim & defaultPageTitleSc
                    Page.MetaDescription = Page.Title & mateDescriptionSc
                Else

                    AspNetPager1.EnableUrlRewriting = True
                    AspNetPager1.UrlRewritePattern = "/ctag/" & cateTag.CateName.Trim() & "/" & cateId & "{0}.aspx"
                    shopName = cateTag.CateName
                    pageTitle = pageTitlek11 & pageTitleSpan.Replace("[SPANNAME]", parentCateTag.CateName) & pageTitleSpan.Replace("[SPANNAME]", cateTag.CateName)
                    templateName = "PageTitle"
                    Page.Title = parentCateTag.CateName.Trim & "-" & cateTag.CateName.Trim & defaultPageTitle
                    Page.MetaDescription = Page.Title & mateDescription
                End If
            ElseIf Not (String.IsNullOrEmpty(keyWordId)) Then

                KeyWord = (From k In entity.KeyWords
                           Where k.ID = keyWordId
                           Select k.KeyWord1).FirstOrDefault
                siteid = "-2"
                If (prodType = "WB") Then
                    AspNetPager1.EnableUrlRewriting = True
                    AspNetPager1.UrlRewritePattern = "/kw/tst/" & KeyWord & "/{0}.aspx"
                    shopName = KeyWord
                    pageTitle = pageTitlek11Sc & pageTitleSpan.Replace("[SPANNAME]", "热门话题") & pageTitleSpan.Replace("[SPANNAME]", KeyWord)
                    templateName = "PageTitleSC"
                    Page.Title = KeyWord & defaultPageTitleSc
                    Page.MetaDescription = Page.Title & mateDescriptionSc
                Else
                    AspNetPager1.EnableUrlRewriting = True
                    AspNetPager1.UrlRewritePattern = "/ctag/" & KeyWord & "/{0}.aspx"
                    shopName = KeyWord
                    pageTitle = pageTitlek11 & pageTitleSpan.Replace("[SPANNAME]", "熱門話題") & pageTitleSpan.Replace("[SPANNAME]", KeyWord)
                    templateName = "PageTitle"
                    Page.Title = KeyWord & defaultPageTitle
                    Page.MetaDescription = Page.Title & mateDescription
                End If
            Else
                If (prodType = "WB") Then
                    If (String.IsNullOrEmpty(shopName)) Then
                        siteid = "0"
                        AspNetPager1.EnableUrlRewriting = True
                        AspNetPager1.UrlRewritePattern = "/tst/{0}.aspx"
                        Page.Title = defaultPageTitleSc
                        Page.MetaDescription = mateDescriptionSc

                    ElseIf (shopName = "AllShops") Then
                        siteid = "-1"
                        AspNetPager1.EnableUrlRewriting = True
                        AspNetPager1.UrlRewritePattern = "/tst/" & shopName.Trim() & "/{0}.aspx"
                        Page.Title = "商店一览" & " " & defaultPageTitleSc
                        Page.MetaDescription = mateDescriptionSc

                    End If
                    pageTitle = pageTitlek11Sc & pageTitleAllshopSc
                    templateName = "PageTitleSC"
                Else
                    If (String.IsNullOrEmpty(shopName)) Then
                        siteid = "0"
                        AspNetPager1.EnableUrlRewriting = True
                        AspNetPager1.UrlRewritePattern = "/{0}"
                        Page.Title = defaultPageTitle
                        Page.MetaDescription = mateDescription
                    ElseIf (shopName = "AllShops") Then
                        siteid = "-1"
                        AspNetPager1.EnableUrlRewriting = True
                        AspNetPager1.UrlRewritePattern = "/" & shopName.Trim() & "/{0}"

                        Page.Title = "商店一覽" & " " & defaultPageTitle
                        Page.MetaDescription = mateDescription

                    End If
                    pageTitle = pageTitlek11 & pageTitleAllshop
                    templateName = "PageTitle"
                End If
                shopName = ""
            End If
        Else
            'Dim pSite As Product = (From pro In entity.Products Where pro.ProdouctID = pSiteid).FirstOrDefault()
            siteid = Integer.Parse(pSiteid)
            Dim shopsite As AutomationSite = (From autosite In entity.AutomationSites
                                             Where autosite.siteid = siteid
                                             Select autosite).FirstOrDefault()
            If (prodType = "WB") Then
                pageTitle = pageTitlek11Sc & pageTitleShopName
                templateName = "PageTitleSC"
                shopName = shopsite.SiteNameSc
                k11siteUrl = shopsite.k11SiteUrlSC
                AspNetPager1.EnableUrlRewriting = True
                AspNetPager1.UrlRewritePattern = "/tst/" & shopName.Trim() & "/" & siteid & "/{0}.aspx"
                Page.Title = shopName & " " & defaultPageTitleSc
                Page.MetaDescription = Page.Title & mateDescriptionSc

            Else
                pageTitle = pageTitlek11 & pageTitleShopName
                templateName = "PageTitle"
                shopName = shopsite.SiteName
                k11siteUrl = shopsite.k11SiteUrl
                AspNetPager1.EnableUrlRewriting = True
                AspNetPager1.UrlRewritePattern = "/" & shopName.Trim() & "/" & siteid & "/{0}.aspx"
                Page.Title = shopName & " " & defaultPageTitle
                Page.MetaDescription = Page.Title & mateDescription
            End If
            pageTitle = pageTitle.Replace("[SITENAME]", shopName).Replace("[URL]", UrlValid.getUrlValid(k11siteUrl))
        End If
        If 1 Then
            Dim s As Integer
            s = 3
        End If
        If (prodType = "WB") Then
            header = (From t In entity.Templates
                      Where t.SiteID = msiteid And t.TemplateName = "HeaderSC"
                      Select t.Contents).FirstOrDefault()

            footer = (From t In entity.Templates
                      Where t.SiteID = msiteid And t.TemplateName = "FooterSC"
                      Select t.Contents).FirstOrDefault()
            k11Description = (From t In entity.Templates
                              Where t.SiteID = msiteid And t.TemplateName = "k11DescriptionSC"
                              Select t.Contents).FirstOrDefault()
        Else
            header = (From t In entity.Templates
               Where t.SiteID = msiteid And t.TemplateName = "Header"
               Select t.Contents).FirstOrDefault()
            footer = (From t In entity.Templates
                      Where t.SiteID = msiteid And t.TemplateName = "Footer"
                      Select t.Contents).FirstOrDefault()
            k11Description = (From t In entity.Templates
                    Where t.SiteID = msiteid And t.TemplateName = "k11Description"
                    Select t.Contents).FirstOrDefault()
        End If

        Dim totalProducts As Integer
        If (siteid = "0") Then 'k11 ,首页
            '以下是获取k11店铺post数，将首页改版为k11尖沙咀介绍后不再获取此内容。
            'If (prodType = "") Then
            '    totalProducts = (From p In entity.Products
            '                   Where p.SiteID = msiteid
            '                   Select p).Count()
            'Else
            '    totalProducts = (From p In entity.Products
            '                       Where p.SiteID = msiteid And p.Currency = prodType
            '                       Select p).Count()
            'End If
            totalProducts = 5
            If (prodType = "") Then
                homePage = (From t In entity.Templates
                            Where t.SiteID = msiteid And t.TemplateName = "HomePage"
                            Select t.Contents).FirstOrDefault()
            Else
                homePage = (From t In entity.Templates
                            Where t.SiteID = msiteid And t.TemplateName = "HomePageSC"
                            Select t.Contents).FirstOrDefault()
            End If

        ElseIf (siteid = "-1") Then '店铺一览，所有店铺
            If (prodType = "") Then
                totalProducts = (From p In entity.Products
                                 Select p).Count()
            Else
                totalProducts = (From p In entity.Products
                                 Where p.Currency = prodType
                                 Select p).Count()
            End If
        ElseIf (siteid = "-2") Then '关键词查询数据
            totalProducts = (From p In entity.Products
                             Where p.PictureAlt.Contains(KeyWord) OrElse p.Description.Contains(KeyWord)
                             Select p).Count()
        Else
            Dim arrSiteid As String() = siteid.Split(",")
            If (prodType = "") Then
                totalProducts = (From p In entity.Products
                                 Where arrSiteid.Contains(p.SiteID)
                                 Select p).Count()
            Else
                totalProducts = (From p In entity.Products
                                 Where arrSiteid.Contains(p.SiteID) And p.Currency = prodType
                                 Select p).Count()
            End If
        End If

        pagesize = AspNetPager1.PageSize
        AspNetPager1.RecordCount = totalProducts
        parentCateTags = (From c In entity.CateTags
                          Where c.ParentID = 0
                          Select c).ToList()
    End Sub

    Protected Sub AspNetPager1_PageChanged(sender As Object, e As EventArgs) Handles AspNetPager1.PageChanged
        Dim pageindex As Integer = AspNetPager1.CurrentPageIndex
        'Dim pageindex As Integer = 2
        Dim products As New List(Of GetPostsFacebookSeo_Result)
        Dim keyWordid As String = Request.QueryString("keyWordId")
        'pageindex = Val(Request.QueryString("pageIndex"))
        Dim KeyWord As String = (From k In entity.KeyWords
                                         Where k.ID = keyWordId
                                         Select k.KeyWord1).FirstOrDefault

        If (siteid = "0") Then 'siteid = 0 是k11首页的情况，仅获取8个post
            products = entity.GetPostsFacebookSeo(msiteid, prodType, "").Skip((pageindex - 1) * pagesize).Take(5).ToList()
        Else
            products = entity.GetPostsFacebookSeo(siteid, prodType, KeyWord).Skip((pageindex - 1) * pagesize).Take(pagesize).ToList()
        End If

        ReItems.DataSource = products
        ReItems.DataBind()
    End Sub

    Public Function isAutoPlanExist(ByVal siteid As Integer, ByVal planType As String) As Boolean
        Dim myAutoplan As AutomationPlan = (From a In entity.AutomationPlans
                                            Where a.SiteID = siteid AndAlso a.PlanType = planType
                                            Select a).FirstOrDefault()
        If (myAutoplan Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function


End Class