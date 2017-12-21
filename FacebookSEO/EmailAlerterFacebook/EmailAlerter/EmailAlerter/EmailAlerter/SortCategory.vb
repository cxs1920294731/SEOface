Imports System.Linq
Imports System.Text
Imports HtmlAgilityPack
Imports System.Data.SqlClient

Public Class SortCategory
    Private Shared efContext As New FaceBookForSEOEntities()

    ''' <summary>
    ''' 创建邮件模板,包括获取填充好数据的模板、添加link_category、添加追踪代码，分类促发
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSpreadTemplate(ByVal siteId As Integer, ByVal issueId As Integer, ByVal templateId As Integer, _
                                       ByVal siteName As String, ByVal planType As String, _
                                       ByVal compaignName As String, ByVal dllType As String, ByVal Url As String, _
                                       ByVal splitCategory As String, ByVal UrlSpecialCode As String, ByVal logoUrl As String) As String
        Try
            Dim queryCate = From c In efContext.AutoCategories
                        Where c.SiteID = siteId
                        Select c
            Dim listCategory As New List(Of AutoCategory)
            For Each q As AutoCategory In queryCate
                Dim cate As New AutoCategory()
                cate.Category1 = q.Category1
                cate.CategoryID = q.CategoryID
                listCategory.Add(cate)
            Next
            Dim SpreadTemplate As String = efContext.AutoTemplates.Single(Function(t) t.Tid = templateId).Contents

            SpreadTemplate = FillSpreadTemplate(SpreadTemplate, siteId, issueId, planType, splitCategory, Url)

            SpreadTemplate = HandleCategory(SpreadTemplate, siteId)

            If (SpreadTemplate.Contains("[SHOP_NAME]") AndAlso Not String.IsNullOrEmpty(siteName)) Then
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", siteName)
            End If
            If (SpreadTemplate.Contains("[SHOP_URL]") AndAlso Not String.IsNullOrEmpty(Url)) Then
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", Url)
            End If
            SpreadTemplate = SpreadTemplate.Replace("[SHOP_LOGOIMGURL]", logoUrl)

            If (String.IsNullOrEmpty(logoUrl) AndAlso SpreadTemplate.Contains("[BEGIN_LOGO]")) Then '当logo不存在时用店铺名替代logo的位置
                Dim index As Integer = SpreadTemplate.IndexOf("[BEGIN_LOGO]")
                Dim endIndex As Integer = SpreadTemplate.IndexOf("[END_LOGO]")
                Dim endLength As Integer = "[END_LOGO]".Length
                Dim replaceString As String = SpreadTemplate.Substring(index, endLength + endIndex - index)
                SpreadTemplate = SpreadTemplate.Replace(replaceString, siteName)
            End If
            SpreadTemplate = SpreadTemplate.Replace("[BEGIN_LOGO]", "").Replace("[END_LOGO]", "")

            '2014/04/24 added,begin
            If Not (String.IsNullOrEmpty(UrlSpecialCode)) Then  '半固定追踪代码，即不是很特殊的追踪代码
                SpreadTemplate = SpecialCode.AddSpecialCode(UrlSpecialCode, SpreadTemplate)
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", siteName)
                SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", Url)

                Return SpreadTemplate
            End If
            'end
            '2013/05/07新增，客户在没有要求加追踪代码的时候，不加追踪代码
            Select Case dllType
                Case "focalprice"
                    '请在此处添加FocalPrice的追踪代码
                    'URL?utm_source=RSpread&utm_medium=EM&utm_campaign=DM_1350ES_MH0953W
                    SpreadTemplate = SpecialCode.AddSpecialCodeForFocalPrice(SpreadTemplate, "ES") 'ES:代表FocalPrice的意大利站点
                Case "sammydress", "ahappydeal", "oasap", "dresslilynew"
                    If (planType.Trim.Contains("HP")) Then
                        Dim codeType As String = "?utm_source=mail_spread&utm_medium=mail&utm_campaign=special." & DateTime.Now.ToString("yyyyMMdd")
                        SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                    Else
                        Dim codeType As String = "?utm_source=mail_spread&utm_medium=mail&utm_campaign=regular." & DateTime.Now.ToString("yyyyMMdd")
                        SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
                    End If
                Case "eachbuyer"
                    Dim codeType As String = "?utm_source=affiliate&utm_medium=EDM&utm_campaign=spead_mail_special" & DateTime.Now.ToString("yyyyMMdd")
                    SpreadTemplate = SpecialCode.AddSpecialCode(codeType, SpreadTemplate)
            End Select
            'For taobao/tmall add shopName
            'SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", senderName)
            'SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", Url)

            Return SpreadTemplate
        Catch ex As Exception
            Using efContext2 As New FaceBookForSEOEntities '防止缓存,Issues表的发送状态位Error Template
                Dim queryIssue As AutoIssue = efContext2.AutoIssues.Where(Function(iss) iss.IssueID = issueId).SingleOrDefault()
                queryIssue.SentStatus = "ET"  'ET-Error Template
                efContext2.SaveChanges()
            End Using

            Common.LogText(ex.ToString())
            'a.LogText(ex.InnerException.Message)
            Throw New Exception(ex.ToString()) 'throw the exception to the upper layer
        End Try
    End Function



    ''' <summary>
    ''' 使用数据库中的数据填充模板，返回填充好的模板
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <param name="siteId"></param>
    ''' <param name="issueId"></param>
    ''' <param name="planType"></param>
    ''' <param name="splitCategory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FillSpreadTemplate(ByVal SpreadTemplate As String, ByVal siteId As Integer, ByVal issueId As Integer, _
                                              ByVal planType As String, ByVal splitCategory As String, ByVal shopUrl As String) As String
        '分类促发的邮件模板中包含大标签[BEGIN_SORT_CATEGORY]
        If (SpreadTemplate.Contains("[BEGIN_SORT_CATEGORY]")) Then
            Dim iBeginSortCate As Integer = "[BEGIN_SORT_CATEGORY]".Length
            Dim iEndSortCate As Integer = "[END_SORT_CATEGORY]".Length
            Dim iStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_SORT_CATEGORY]")
            Dim iEndIndex As Integer = SpreadTemplate.IndexOf("[END_SORT_CATEGORY]")
            Dim oldContent As String = SpreadTemplate.Substring(iStartIndex, iEndIndex - iStartIndex + iEndSortCate)
            Dim newContent As String = SpreadTemplate.Substring(iStartIndex + iBeginSortCate, iEndIndex - iStartIndex - iBeginSortCate)

            Dim categorys As String()
            If (splitCategory.Contains("^")) Then '分类促发有多个分类
                categorys = splitCategory.Split("^")
            Else '分类促发只有1个分类
                categorys = New String() {splitCategory}
            End If
            Dim i As Integer = 1
            For Each cate As String In categorys
                Dim labelBeginCategoryI As String = "[BEGIN_CATEGORY_" & i.ToString() & "]"
                Dim labelEndCategoryI As String = "[END_CATEGORY_" & i.ToString() & "]"
                If (newContent.Contains(labelBeginCategoryI)) Then

                    Dim category As AutoCategory
                    category = efContext.AutoCategories.FirstOrDefault(Function(m) m.Category1 = cate AndAlso m.SiteID = siteId AndAlso m.Url.Contains(shopUrl))
                    If (cate.Trim.EndsWith("2")) Then 'dresslily 2.0 的categoryname比较特别，为了在显示时不显示出结尾的2
                        newContent = newContent.Replace("[CATEGORY_NAME_" & i.ToString() & "]", cate.Replace("2", ""))
                    Else
                        newContent = newContent.Replace("[CATEGORY_NAME_" & i.ToString() & "]", category.Category1)
                    End If
                    newContent = newContent.Replace("[CATEGORY_URL_" & i.ToString() & "]", category.Url)
                    newContent = newContent.Replace("[CATEGORY_ID_" & i.ToString() & "]", category.CategoryID)

                    Dim iBeginCategory As Integer = labelBeginCategoryI.Length
                    Dim iEndCategory As Integer = labelEndCategoryI.Length
                    Dim iBeginCateIndex As Integer = newContent.IndexOf(labelBeginCategoryI)
                    Dim iEndCateIndex As Integer = newContent.IndexOf(labelEndCategoryI)
                    Dim oldCategoryContent As String = newContent.Substring(iBeginCateIndex, iEndCateIndex - iBeginCateIndex + iEndCategory)
                    Dim newCategoryContent As String = newContent.Substring(iBeginCateIndex + iBeginCategory, iEndCateIndex - iBeginCateIndex - iBeginCategory)
                    Dim products As New List(Of AutoProduct)
                    Dim cateProducts As New List(Of AutoProduct)
                    Dim categoriesForaPro As New List(Of AutoCategory)
                    products = (From pIssue In efContext.AutoProducts_Issue
                             Join p In efContext.AutoProducts On pIssue.ProductId Equals p.ProdouctID
                            Where pIssue.IssueID = issueId AndAlso pIssue.SectionID = "CA"
                            Select p).ToList()

                    For Each p As AutoProduct In products '找出某个分类的产品
                        categoriesForaPro = p.Categories.ToList
                        For Each cateee As AutoCategory In categoriesForaPro
                            Common.LogText(p.ProdouctID & "  " & cateee.Category1)
                            If (cateee.Category1 = cate.Trim()) Then
                                cateProducts.Add(p)
                                Common.LogText(cate & "  " & p.ProdouctID & "  " & p.Url)
                            End If
                        Next
                    Next

                    For Each p As AutoProduct In cateProducts '填充模板

                        If (newCategoryContent.Contains("[BEGIN_PRODUCT]")) Then
                            Common.LogText(cate & "  " & p.ProdouctID & "  " & p.Url)
                            Dim productStartIndex As Integer = newCategoryContent.IndexOf("[BEGIN_PRODUCT]")
                            Dim productEndIndex As Integer = newCategoryContent.IndexOf("[END_PRODUCT]")
                            Dim lenProductEnd As Integer = "[END_PRODUCT]".Length '2013/06/26 added
                            Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
                            Dim oldProduct As String = newCategoryContent.Substring(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                            Dim newProduct As String = newCategoryContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                            If (newProduct.Contains("[URL]")) Then
                                newProduct = newProduct.Replace("[URL]", p.Url)
                            End If
                            If (newProduct.Contains("[CATEGORY_ID]")) Then
                                newProduct = newProduct.Replace("[CATEGORY_ID]", p.Categories.First.CategoryID.ToString())
                            End If
                            If (newProduct.Contains("[PICTURE_SRC]")) Then
                                newProduct = newProduct.Replace("[PICTURE_SRC]", p.PictureUrl)
                            End If
                            If (newProduct.Contains("[PICTURE_ALT]")) Then
                                newProduct = newProduct.Replace("[PICTURE_ALT]", p.PictureAlt)
                            End If
                            If (newProduct.Contains("[PRODUCT_NAME]")) Then
                                newProduct = newProduct.Replace("[PRODUCT_NAME]", p.Prodouct)
                            End If
                            If (newProduct.Contains("[PRODUCT_PRICE]")) Then
                                If (p.Price Is Nothing) Then  '原价获取不到，只获取到折扣价，则把原价删除掉
                                    newProduct = newProduct.Replace("[PRODUCT_PRICE]", "")
                                Else
                                    newProduct = newProduct.Replace("[PRODUCT_PRICE]", p.Currency + " " + String.Format("{0:0.00}", p.Price))
                                End If
                            End If
                            If (newProduct.Contains("[PRODUCT_MONEY]")) Then
                                newProduct = newProduct.Replace("[PRODUCT_MONEY]", String.Format("{0:0.00}", p.Discount))
                            End If
                            If (newProduct.Contains("[PRODUCT_SALES]")) Then
                                newProduct = newProduct.Replace("[PRODUCT_SALES]", p.Sales)
                            End If

                            '2013/05/21,add free shiping pictures and ships pictures in template,begin
                            If (newProduct.Contains("[FREESHIPPING]")) Then
                                If (String.IsNullOrEmpty(p.FreeShippingImg)) Then
                                    newProduct = newProduct.Replace("[FREESHIPPING]", "")
                                Else
                                    Dim addStrPicture As String = "<tr><td align=""center""><img alt="""" src=""" & p.FreeShippingImg & """ style=""border-width: 0px; border-style: solid; display: block;"" /></td></tr>"
                                    newProduct = newProduct.Replace("[FREESHIPPING]", addStrPicture)
                                End If
                            End If
                            If (newProduct.Contains("[SHIPS]")) Then
                                If (String.IsNullOrEmpty(p.ShipsImg)) Then
                                    newProduct = newProduct.Replace("[SHIPS]", "")
                                Else
                                    Dim addStrPicture As String = "<tr><td align=""center""><img alt="""" src=""" & p.ShipsImg & """ style=""border-width: 0px; border-style: solid; display: block;"" /></td></tr>"
                                    newProduct = newProduct.Replace("[SHIPS]", addStrPicture)
                                End If
                            End If
                            '2013/05/21,add free shiping pictures and ships pictures in template,end

                            newCategoryContent = newCategoryContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                            newCategoryContent = newCategoryContent.Insert(productStartIndex, newProduct)
                        Else
                            Exit For
                        End If
                    Next
                    While (newCategoryContent.Contains("[BEGIN_PRODUCT]"))
                        Dim productStartIndex As Integer = newCategoryContent.IndexOf("[BEGIN_PRODUCT]")
                        Dim productEndIndex As Integer = newCategoryContent.IndexOf("[END_PRODUCT]")
                        Dim lenProductEnd As Integer = "[END_PRODUCT]".Length '2013/06/26 added
                        'Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
                        'Dim oldProduct As String = newCategoryContent.Substring(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                        'Dim newProduct As String = newCategoryContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                        newCategoryContent = newCategoryContent.Remove(productStartIndex, productEndIndex - productStartIndex + lenProductEnd)
                        newCategoryContent = newCategoryContent.Insert(productStartIndex, "")

                    End While
                    newContent = newContent.Replace(oldCategoryContent, newCategoryContent)
                End If
                i = i + 1
            Next
            SpreadTemplate = SpreadTemplate.Replace(oldContent, newContent)
        ElseIf (SpreadTemplate.Contains("[BEGIN_CATEGORIES_P]")) Then '公共淘宝填充模板
            Dim BeginStrLen = "[BEGIN_CATEGORIES_P]".Length
            Dim EndStrLen = "[END_CATEGORIES_P]".Length
            Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
            Dim cateStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_CATEGORIES_P]")
            Dim cateEndIndex As Integer = SpreadTemplate.IndexOf("[END_CATEGORIES_P]")
            Dim oldCategory As String = SpreadTemplate.Substring(cateStartIndex, cateEndIndex - cateStartIndex + EndStrLen)
            Dim categoryContent As String = SpreadTemplate.Substring(cateStartIndex + BeginStrLen, cateEndIndex - cateStartIndex - BeginStrLen)
            Dim cateblock As StringBuilder = New StringBuilder()
            Dim arrCategorys As String()
            If (splitCategory.Contains("^")) Then '分类促发有多个分类
                arrCategorys = splitCategory.Split("^")
            Else '分类促发只有1个分类
                arrCategorys = New String() {splitCategory}
            End If
            For Each strCategory As String In arrCategorys
                'For i As Integer = 0 To catescount - 1
                If Not String.IsNullOrEmpty(strCategory) Then
                    Dim category As AutoCategory = efContext.AutoCategories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = strCategory AndAlso c.Url.Contains(shopUrl)).SingleOrDefault()
                    Dim sql As String = "select * from Products p " &
                                        "inner join Products_Issue pi on p.ProdouctID=pi.ProductId " &
                                        "inner join ProductCategory pc on pc.ProductID=p.ProdouctID " &
                                        "inner join Categories c on pc.CategoryID=c.CategoryID " &
                                        "where pi.IssueID=@IssueID and c.Category=@Category and c.SiteID=@SiteID"
                    Dim sqlParam As SqlParameter()
                    sqlParam = New SqlParameter() {New SqlParameter With {.ParameterName = "IssueID", .Value = issueId.ToString()}, _
                                                   New SqlParameter With {.ParameterName = "Category", .Value = strCategory.Trim()}, _
                                                   New SqlParameter With {.ParameterName = "SiteID", .Value = siteId.ToString()}}
                    Dim productlist As List(Of AutoProduct) = efContext.ExecuteStoreQuery(Of AutoProduct)(sql, sqlParam).ToList()
                    cateblock.Append(InsertProductOfOthers(categoryContent, lengthProduct, productlist).ToString().Replace("[CATENAME]", category.Description).Replace("[CATEURL]", category.Url))
                End If
                'Next
            Next
            SpreadTemplate = SpreadTemplate.Replace(oldCategory, cateblock.ToString())
        End If
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 返回所有Section != "Category"（Sections表）的产品填充的模板
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="lengthProduct"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function InsertProductOfOthers(ByVal oldContent As String, ByVal lengthProduct As Integer, ByVal productlist As List(Of AutoProduct))
        Dim queryNewArrival As List(Of AutoProduct) = productlist
        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")
        For Each q As AutoProduct In queryNewArrival
            If (productStartIndex > -1) Then
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
                If (newProduct.Contains("[URL]")) Then
                    newProduct = newProduct.Replace("[URL]", q.Url)
                End If
                If (newProduct.Contains("[CATEGORY_ID]")) Then
                    Dim myProduct As AutoProduct = efContext.AutoProducts.Where(Function(prod) prod.ProdouctID = q.ProdouctID).Single()
                    newProduct = newProduct.Replace("[CATEGORY_ID]", myProduct.Categories.FirstOrDefault.CategoryID.ToString)
                End If
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
                    newProduct = newProduct.Replace("[PRODUCT_NAME]", q.Description)
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

    ''' <summary>
    ''' 用CategoryID(表：AutoCategories)替换模板中的link_category的值标签
    ''' （如,[CATEGORY_ID_BAGS]用553替换）
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function HandleCategory(ByVal SpreadTemplate As String, ByVal siteId As Integer) As String
        Dim queryCategory = From c In efContext.AutoCategories Where c.SiteID = siteId Select c
        For Each q As AutoCategory In queryCategory
            Dim cateModifyStr As String = q.Category1.ToUpper().Replace(" ", "_") 'Women's Dresses替换成WOMEN'S_DRESSES
            Dim replaceStr As String = "[CATEGORY_ID_" & cateModifyStr & "]"
            SpreadTemplate = SpreadTemplate.Replace(replaceStr, q.CategoryID.ToString())
        Next
        Return SpreadTemplate
    End Function

    'Public Shared Sub LogText(ByVal Ex As String)
    '    Try
    '        '----------------------------------------------------
    '        System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub
End Class
