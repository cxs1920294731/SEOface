Imports System.Configuration
Imports Analysis
Imports HtmlAgilityPack
Imports System.Text
Imports System.Text.RegularExpressions


Public Class SEmailTmall
    Dim myanalysisEfhelper As New EFHelper()

    Public Function GetTestHtmlDemo(queryParams As QueryParams, ByVal iProCount As Integer, ByRef firstProduct As String) As String

        Dim cates As List(Of Category) = queryParams.Categories.ToList()
        If (cates.Count <= 0) Then
            Return ""
        End If
        Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
        Dim template As String = queryParams.Template

        If (template.Contains("[BEGIN_CATEGORIES_P]")) Then
            Dim BeginStrLen = "[BEGIN_CATEGORIES_P]".Length
            Dim EndStrLen = "[END_CATEGORIES_P]".Length
            Dim cateStartIndex As Integer = template.IndexOf("[BEGIN_CATEGORIES_P]")
            Dim cateEndIndex As Integer = template.IndexOf("[END_CATEGORIES_P]")
            Dim oldCategory As String = template.Substring(cateStartIndex, cateEndIndex - cateStartIndex + EndStrLen)
            Dim categoryContent As String = template.Substring(cateStartIndex + BeginStrLen, cateEndIndex - cateStartIndex - BeginStrLen)
            Dim cateblock As StringBuilder = New StringBuilder()


            Dim lastindexcateStartIndex As Integer = template.LastIndexOf("[BEGIN_CATEGORIES_P]")
            If (cateStartIndex <> lastindexcateStartIndex Or cates(0).Url.Contains("测试")) Then
                '第一个[BEGIN_CATEGORIES_P]标签与最后一个此标签不是同一个位置，说明有多个"[BEGIN_CATEGORIES_P]""[END_CATEGORIES_P]"替换块
                '此时以[BEGIN_CATEGORIES_P]标签为准循环
                Dim cateIndex As Int32 = 0
                Dim itemCategory As Category
                Do
                    itemCategory = cates(cateIndex)

                    If (itemCategory.Category1 = "每周上新") Then
                        itemCategory.Url = myanalysisEfhelper.GetXinpinUrlandSubject(itemCategory.Url)(0)
                    End If
                    Dim cname As String = itemCategory.Category1
                    Dim curl As String = itemCategory.Url

                    Dim productlist As List(Of Product) = GetProductsByUrl(curl, iProCount)
                    If (cateIndex = 0) Then
                        firstProduct = productlist(0).Prodouct.Trim
                    End If
                    Dim newTempContent As String = InsertProductOfOthers(categoryContent, lengthProduct, productlist)
                    newTempContent = newTempContent.Replace("[CATENAME]", cname).Replace("[CATEURL]", curl)
                    template = template.Remove(cateStartIndex, cateEndIndex - cateStartIndex + EndStrLen)
                    template = template.Insert(cateStartIndex, newTempContent)

                    cateIndex = cateIndex + 1
                    cateStartIndex = template.IndexOf("[BEGIN_CATEGORIES_P]")
                    cateEndIndex = template.IndexOf("[END_CATEGORIES_P]")
                    If (cateEndIndex - cateStartIndex - BeginStrLen <= 0) Then
                        cateIndex = cates.Count
                    Else
                        categoryContent = template.Substring(cateStartIndex + BeginStrLen, cateEndIndex - cateStartIndex - BeginStrLen)
                    End If
                Loop While (template.Contains("[BEGIN_CATEGORIES_P]") AndAlso cateIndex < cates.Count)
                While (template.Contains("[BEGIN_CATEGORIES_P]"))
                    Dim productStartIndex As Integer = template.IndexOf("[BEGIN_CATEGORIES_P]")
                    Dim productEndIndex As Integer = template.IndexOf("[END_CATEGORIES_P]")
                    Dim lenProductEnd As Integer = "[END_CATEGORIES_P]".Length '2013/06/26 added

                    template = template.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                    template = template.Insert(productStartIndex, "")
                End While
            Else
                For Each category As Category In cates
                    'For i As Integer = 0 To catescount - 1
                    If Not String.IsNullOrEmpty(category.Category1) Then
                        If (category.Category1 = "每周上新") Then
                            category.Url = myanalysisEfhelper.GetXinpinUrlandSubject(category.Url)(0)
                        End If
                        'TODO 填充数据
                        Dim cname As String = category.Category1
                        Dim curl As String = category.Url

                        Dim productlist As List(Of Product) = GetProductsByUrl(curl, iProCount)
                        firstProduct = productlist(0).Prodouct.Trim
                        cateblock.Append(InsertProductOfOthers(categoryContent, lengthProduct, productlist).ToString().Replace("[CATENAME]", cname).Replace("[CATEURL]", curl))
                    End If
                Next

                template = template.Replace(oldCategory, cateblock.ToString())
            End If
        End If
        Return template
    End Function


    Public Function GetHtmlDemo(queryParams As QueryParams, ByVal iProCount As Integer, ByRef firstProduct As String, ByVal eBusinessSite As EBusinessSite) As String

        Dim cates As List(Of Category) = queryParams.Categories.ToList()
        If (cates.Count <= 0) Then
            Return ""
        End If
        Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
        Dim template As String = queryParams.Template

        If (template.Contains("[BEGIN_CATEGORIES_P]")) Then
            Dim BeginStrLen = "[BEGIN_CATEGORIES_P]".Length
            Dim EndStrLen = "[END_CATEGORIES_P]".Length
            Dim cateStartIndex As Integer = template.IndexOf("[BEGIN_CATEGORIES_P]")
            Dim cateEndIndex As Integer = template.IndexOf("[END_CATEGORIES_P]")
            Dim oldCategory As String = template.Substring(cateStartIndex, cateEndIndex - cateStartIndex + EndStrLen)
            Dim categoryContent As String = template.Substring(cateStartIndex + BeginStrLen, cateEndIndex - cateStartIndex - BeginStrLen)
            Dim cateblock As StringBuilder = New StringBuilder()


            Dim lastindexcateStartIndex As Integer = template.LastIndexOf("[BEGIN_CATEGORIES_P]")
            If (cateStartIndex <> lastindexcateStartIndex Or cates(0).Url.Contains("测试")) Then
                '第一个[BEGIN_CATEGORIES_P]标签与最后一个此标签不是同一个位置，说明有多个"[BEGIN_CATEGORIES_P]""[END_CATEGORIES_P]"替换块
                '此时以[BEGIN_CATEGORIES_P]标签为准循环
                Dim cateIndex As Int32 = 0
                Dim itemCategory As Category
                Do
                    itemCategory = cates(cateIndex)

                    'If (itemCategory.Category1 = "每周上新") Then
                    '    itemCategory.Url = myanalysisEfhelper.GetXinpinUrlandSubject(itemCategory.Url)(0)
                    'End If
                    Dim cname As String = itemCategory.Category1
                    Dim curl As String = itemCategory.Url

                    Dim productlist As List(Of Product) = eBusinessSite.GetProductList(curl, 0).Take(iProCount).ToList()
                    If (cateIndex = 0) Then
                        firstProduct = productlist(0).Prodouct.Trim
                    End If
                    Dim newTempContent As String = InsertProductOfOthers(categoryContent, lengthProduct, productlist)
                    newTempContent = newTempContent.Replace("[CATENAME]", cname).Replace("[CATEURL]", curl)
                    template = template.Remove(cateStartIndex, cateEndIndex - cateStartIndex + EndStrLen)
                    template = template.Insert(cateStartIndex, newTempContent)

                    cateIndex = cateIndex + 1
                    cateStartIndex = template.IndexOf("[BEGIN_CATEGORIES_P]")
                    cateEndIndex = template.IndexOf("[END_CATEGORIES_P]")
                    If (cateEndIndex - cateStartIndex - BeginStrLen <= 0) Then
                        cateIndex = cates.Count
                    Else
                        categoryContent = template.Substring(cateStartIndex + BeginStrLen, cateEndIndex - cateStartIndex - BeginStrLen)
                    End If
                Loop While (template.Contains("[BEGIN_CATEGORIES_P]") AndAlso cateIndex < cates.Count)
                While (template.Contains("[BEGIN_CATEGORIES_P]"))
                    Dim productStartIndex As Integer = template.IndexOf("[BEGIN_CATEGORIES_P]")
                    Dim productEndIndex As Integer = template.IndexOf("[END_CATEGORIES_P]")
                    Dim lenProductEnd As Integer = "[END_CATEGORIES_P]".Length '2013/06/26 added

                    template = template.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                    template = template.Insert(productStartIndex, "")
                End While
            Else
                For Each category As Category In cates
                    'For i As Integer = 0 To catescount - 1
                    If Not String.IsNullOrEmpty(category.Category1) Then
                        'If (category.Category1 = "每周上新") Then
                        '    category.Url = myanalysisEfhelper.GetXinpinUrlandSubject(category.Url)(0)
                        'End If
                        'TODO 填充数据
                        Dim cname As String = category.Category1
                        Dim curl As String = category.Url

                        Dim productlist As List(Of Product) = eBusinessSite.GetProductList(curl, 0).Take(iProCount).ToList()
                        firstProduct = productlist(0).Prodouct.Trim
                        cateblock.Append(InsertProductOfOthers(categoryContent, lengthProduct, productlist).ToString().Replace("[CATENAME]", cname).Replace("[CATEURL]", curl))
                    End If
                Next

                template = template.Replace(oldCategory, cateblock.ToString())
            End If
        End If
        Return template
    End Function

    ''' <summary>
    ''' 通过Url插入产品列表
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="IssueId"></param>
    ''' <param name="count"></param>
    ''' <param name="category"></param>
    ''' <param name="cateType"></param>
    ''' <param name="doc"></param>
    ''' <remarks></remarks>
    Private Function GetProductsByUrl(ByVal curl As String, ByVal count As Integer)

        'TODO 根据出入的参数生成url，获得产品信息

        Dim productlist As New List(Of Product)
        Dim nowTime As DateTime = DateTime.Now
        If (curl.ToLower.Contains("测试")) Then
            Dim index As Integer = curl.IndexOf("=")
            Dim orderType As String = curl.Substring(index + 1)
            Dim prod As New Product
            prod.Url = "http://rspread.cn/"
            prod.Prodouct = "Spread邮件营销软件及服务——" & orderType
            prod.PictureUrl = "http://app.rspread.com/Spread5/SpreaderFiles/9278/files/upload/Spread-Logo-130.jpg"
            prod.PictureAlt = "Spread邮件营销软件及服务" & orderType
            prod.Price = 5000
            prod.Discount = 5000
            'prod.Rating = rating
            'prod.TotalComment = comment
            'prod.SalesAmount = sales
            prod.Currency = "￥"
            For i As Integer = count To 1 Step -1
                productlist.Add(prod)
            Next
        Else

            productlist = myanalysisEfhelper.GetAsynProductList(curl, 0).Take(count).ToList()
            If (productlist.Count <= 0) Then
                Dim myProduct As New Product
                myProduct.Prodouct = "请求太多，天猫很忙，天猫说需要5分钟后再试/(ㄒoㄒ)/~~。"
                productlist.Add(myProduct)
            End If
        End If
        Return productlist
    End Function



    ''' <summary>
    ''' 返回所有Section != "Category"（Sections表）的产品填充的模板
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="issueId"></param>
    ''' <param name="sectionId"></param>
    ''' <param name="lengthProduct"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertProductOfOthers(ByVal oldContent As String, ByVal lengthProduct As Integer, ByVal productlist As List(Of Product))
        Dim queryNewArrival As List(Of Product) = productlist
        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")
        For Each q In queryNewArrival
            If (productStartIndex > -1) Then
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
                If (newProduct.Contains("[URL]")) Then
                    newProduct = newProduct.Replace("[URL]", q.Url)
                End If
                'If (newProduct.Contains("[CATEGORY_ID]")) Then
                '    newProduct = newProduct.Replace("[CATEGORY_ID]", q.Categories.First.CategoryID.ToString())
                'End If
                If (newProduct.Contains("[PICTURE_SRC]")) Then
                    newProduct = newProduct.Replace("[PICTURE_SRC]", q.PictureUrl)
                End If
                If (newProduct.Contains("[PICTURE_ALT]")) Then
                    newProduct = newProduct.Replace("[PICTURE_ALT]", q.PictureAlt)
                End If
                If (newProduct.Contains("[PRODUCT_DESCRIPTION]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_DESCRIPTION]", q.Description)
                End If
                If (newProduct.Contains("[PRODUCT_NAME]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_NAME]", q.Prodouct)
                End If
                If (newProduct.Contains("[PRODUCT_PRICE]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_PRICE]", q.Currency + " " + String.Format("{0:0.00}", q.Price))
                End If
                If (newProduct.Contains("[PRODUCT_MONEY]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", q.Discount))
                End If
                If (newProduct.Contains("[PRODUCT_SALES]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_SALES]", q.Sales)
                End If
                If (newProduct.Contains("[PRODUCT_TBSCORE]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_TBSCORE]", String.Format("{0:0.0}", q.TbScore))
                End If
                If (newProduct.Contains("[PRODUCT_TBCOMMENT]")) Then
                    newProduct = newProduct.Replace("[PRODUCT_TBCOMMENT]", q.TbComment)
                End If

                oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
                oldContent = oldContent.Insert(productStartIndex, newProduct)
            End If

            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增，否则报错
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            Else
                Exit For
            End If
        Next
        While (oldContent.Contains("[BEGIN_PRODUCT]"))
            oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增判断，否则报错
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            End If
        End While
        Return oldContent
    End Function


End Class
