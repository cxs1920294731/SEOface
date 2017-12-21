Imports System.Text.RegularExpressions

Public Class GenTemplate
    Private Shared efContext As New EmailAlerterEntities()

    ''' <summary>
    ''' 创建邮件模板
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' Private Function GetSpreadTemplate(ByVal siteId As Integer, ByVal issueId As Integer) As String
    Public Shared Function GetSpreadTemplate(ByVal siteId As Integer, ByVal issueId As Integer, ByVal templateId As Integer, ByVal senderName As String, ByVal planType As String, ByVal shopUrl As String) As String
        Dim queryCate = From c In efContext.Categories
                        Where c.SiteID = siteId
                        Select c
        Dim listCategory As New List(Of Category)
        For Each q In queryCate
            Dim cate As New Category()
            cate.Category1 = q.Category1
            cate.CategoryID = q.CategoryID
            listCategory.Add(cate)
        Next
        Dim queryTemplate = From t In efContext.Templates Where t.Tid = templateId Select t.Contents
        Dim SpreadTemplate As String = queryTemplate.Single()

        'Promotion大分类数据填充，单个Promtion
        If (SpreadTemplate.Contains("[BEGIN_PROMOTION]")) Then
            ''2013/4/12新加
            Dim queryPromotion = From adIss In efContext.Ads_Issue
                               Where adIss.SiteId = siteId AndAlso adIss.IssueID = issueId
                               Select adIss
            Dim promotionCount = queryPromotion.Count

            For i As Integer = 0 To promotionCount - 1
                Dim beginLen As Integer = "[BEGIN_PROMOTION]".Length
                Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_PROMOTION]")
                Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_PROMOTION]")
                Dim oldPromotion As String = SpreadTemplate.Substring(startProIndex, endProIndex - startProIndex + beginLen)
                Dim promotion As String = SpreadTemplate.Substring(startProIndex + beginLen, endProIndex - startProIndex - beginLen)
                Dim newPromotion As String = InsertPromotionAndAd(promotion, issueId, siteId, i)
                SpreadTemplate = SpreadTemplate.Replace(oldPromotion, newPromotion)
            Next

        End If


        ''Promotion大分类数据填充，2个或2个以上Promtion，2013/4/12新添加
        'If (SpreadTemplate.Contains("[BEGIN_ADS_PROMOTION]")) Then
        '    Dim beginLen As Integer = "[BEGIN_ADS_PROMOTION]".Length
        '    Dim startProIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_ADS_PROMOTION]")
        '    Dim endProIndex As Integer = SpreadTemplate.IndexOf("[END_ADS_PROMOTION]")
        '    Dim oldPromotion As String = SpreadTemplate.Substring(startProIndex, endProIndex - startProIndex + beginLen)
        '    Dim promotion As String = SpreadTemplate.Substring(startProIndex + beginLen, endProIndex - startProIndex - beginLen)
        'End If



        'NewArrival大分类数据填充
        If (SpreadTemplate.Contains("[BEGIN_NEW_ARRIVALS]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_NEW_ARRIVALS]", "[END_NEW_ARRIVALS]", SpreadTemplate, "NE")
        End If
        'Daily Deals大分类数据填充
        If (SpreadTemplate.Contains("[BEGIN_DAILY_DEALS]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_DAILY_DEALS]", "[END_DAILY_DEALS]", SpreadTemplate, "DA")
        End If

        '2013/4/12新增，fever-print
        If (SpreadTemplate.Contains("[BEGIN_BEST_ARRIVAL]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_BEST_ARRIVAL]", "[END_BEST_ARRIVAL]", SpreadTemplate, "DA")
        End If
        If (SpreadTemplate.Contains("[BEGIN_BEST_SELLER]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_BEST_SELLER]", "[END_BEST_SELLER]", SpreadTemplate, "BE")
        End If


        'taobaoCategory-2013/3/26
        If (SpreadTemplate.Contains("[BEGIN_BESTSELLER]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_BESTSELLER]", "[END_BESTSELLER]", SpreadTemplate, "BE")
        End If
        If (SpreadTemplate.Contains("[BEGIN_DEALS]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_DEALS]", "[END_DEALS]", SpreadTemplate, "DE")
        End If
        If (SpreadTemplate.Contains("[BEGIN_HOT_KEEP]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_HOT_KEEP]", "[END_HOT_KEEP]", SpreadTemplate, "HK")
        End If
        If (SpreadTemplate.Contains("[BEGIN_MOST_POPULAR]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_MOST_POPULAR]", "[END_MOST_POPULAR]", SpreadTemplate, "MP")
        End If
        If (SpreadTemplate.Contains("[BEGIN_PROMOTION_PRODUCT]")) Then
            SpreadTemplate = FillRootCategory(siteId, issueId, "[BEGIN_PROMOTION_PRODUCT]", "[END_PROMOTION_PRODUCT]", SpreadTemplate, "PO")
        End If


        '大Category分类数据填充
        If (SpreadTemplate.Contains("[BEGIN_CATEGORIES]")) Then
            Dim categoryLen = "[BEGIN_CATEGORIES]".Length
            Dim cateStartIndex As Integer = SpreadTemplate.IndexOf("[BEGIN_CATEGORIES]")
            Dim cateEndIndex As Integer = SpreadTemplate.IndexOf("[END_CATEGORIES]")
            Dim oldCategory As String = SpreadTemplate.Substring(cateStartIndex, cateEndIndex - cateStartIndex + categoryLen)
            Dim categoryContent As String = SpreadTemplate.Substring(cateStartIndex + categoryLen, cateEndIndex - cateStartIndex - categoryLen)


            'Dim childLen = "[BEGIN_CATEGORY_".Length  '2013/3/25 新加
            Dim childStartIndex As Integer = categoryContent.IndexOf("[BEGIN_CATEGORY_")
            Dim childEndIndex As Integer = categoryContent.IndexOf("[END_CATEGORY_")
            Dim childCategory As String = categoryContent.Substring(childStartIndex, childEndIndex - childStartIndex)
            While (categoryContent.Contains("[BEGIN_CATEGORY_") AndAlso childCategory.Contains("[BEGIN_CATEGORY_"))
                Dim beginIndex As Integer = childCategory.IndexOf("[")
                Dim endIndex As Integer = childCategory.IndexOf("]")
                Dim labelName As String = childCategory.Substring(beginIndex + 1, endIndex - beginIndex - 1)
                Dim splitStr = Regex.Split(labelName, "BEGIN_CATEGORY_", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                Dim searchStr = splitStr(1).Replace("_", " ")
                Dim beginStr = "[BEGIN_CATEGORY_" + splitStr(1) + "]"
                Dim endStr = "[END_CATEGORY_" + splitStr(1) + "]"
                categoryContent = FillRootCategory(siteId, issueId, beginStr, endStr, categoryContent, searchStr)
                If (categoryContent.Contains("[BEGIN_CATEGORY_")) Then
                    childStartIndex = categoryContent.IndexOf("[BEGIN_CATEGORY_")
                    childEndIndex = categoryContent.IndexOf("[END_CATEGORY_")
                    childCategory = categoryContent.Substring(childStartIndex, childEndIndex - childStartIndex)
                End If
            End While

            SpreadTemplate = SpreadTemplate.Replace(oldCategory, categoryContent)
        End If
        SpreadTemplate = HandleCategory(SpreadTemplate, siteId)
        SpreadTemplate = SpreadTemplate.Replace("[SHOP_NAME]", senderName)
        SpreadTemplate = SpreadTemplate.Replace("[SHOP_URL]", shopUrl)
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 读取Ads表中Type="P"的一条记录，并填充相应的Promotion模板
    ''' </summary>
    ''' <param name="promotion"></param>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <remarks></remarks>
    Private Shared Function InsertPromotionAndAd(ByVal promotion As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal counter As Integer) As String
        Dim queryPromotion = (From ad In efContext.Ads _
                             Join issue In efContext.Ads_Issue _
                             On ad.AdID Equals issue.AdId
                             Order By issue.IssueID Descending
                             Where issue.IssueID = issueId And siteId = siteId And ad.Type = "P"
                             Select ad).Skip(counter).Take(1)
        Dim myAd As Ad = queryPromotion.Single()
        promotion = promotion.Replace("[URL]", myAd.Url)
        promotion = promotion.Replace("[PICTURE_SRC]", myAd.PictureUrl)
        promotion = promotion.Replace("[PICTURE_ALT]", myAd.Description)
        Return promotion
    End Function

    ''' <summary>
    ''' 处理含有产品的大分类，返回填充好产品的模板
    ''' </summary>
    ''' <param name="beginStr">大分类的开始标签，[NEW_...]</param>
    ''' <param name="endStr">大分类的结束标签，[END_...]</param>
    ''' <param name="SpreadTemplate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function FillRootCategory(ByVal siteId As Integer, ByVal issueId As Integer, ByVal beginStr As String, ByVal endStr As String, ByVal SpreadTemplate As String, ByVal searchStr As String) As String
        Dim strLen As Integer = beginStr.Length
        Dim startIndex As Integer = SpreadTemplate.IndexOf(beginStr)
        Dim endIndex As Integer = SpreadTemplate.IndexOf(endStr)
        Dim oldTemplateStr As String = SpreadTemplate.Substring(startIndex, endIndex - startIndex + strLen)
        Dim TemplateStr As String = SpreadTemplate.Substring(startIndex + strLen, endIndex - startIndex - strLen)
        Dim newTemplateStr As String = InsertProducts(TemplateStr, issueId, siteId, beginStr, searchStr)
        SpreadTemplate = SpreadTemplate.Replace(oldTemplateStr, newTemplateStr)
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 在大分类中填充产品信息
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="issueId"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function InsertProducts(ByVal oldContent As String, ByVal issueId As Integer, ByVal siteId As Integer, ByVal beginStr As String, ByVal searchStr As String) As String
        Dim lengthProduct As Integer = "[BEGIN_PRODUCT]".Length
        'NewArrival分类产品填充
        If (beginStr.Contains("[BEGIN_NEW_ARRIVALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "NE", lengthProduct, siteId)

        ElseIf (beginStr.Contains("[BEGIN_DAILY_DEALS]")) Then   'Daily Deals分类产品填充
            oldContent = InsertProductOfOthers(oldContent, issueId, "DA", lengthProduct, siteId)

            '2013/4/12,fever-print新增
        ElseIf (beginStr.Contains("BEGIN_BEST_ARRIVAL")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "BR", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_BEST_SELLER]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "BE", lengthProduct, siteId)


            'taobaoProduct-2013/3/26
        ElseIf (beginStr.Contains("[BEGIN_BESTSELLER]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "BE", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_DEALS]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "DE", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_HOT_KEEP]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "HK", lengthProduct, siteId)
        ElseIf (beginStr.Contains("[BEGIN_MOST_POPULAR]")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "MP", lengthProduct, siteId)
        ElseIf (beginStr.Contains("BEGIN_PROMOTION_PRODUCT")) Then
            oldContent = InsertProductOfOthers(oldContent, issueId, "PO", lengthProduct, siteId)


        ElseIf (beginStr.Contains("[BEGIN_CATEGORY_")) Then '2013/3/21新加，待测试
            oldContent = InsertProductOfCategory(oldContent, searchStr, lengthProduct, siteId, issueId)
        End If
        Return oldContent
    End Function


    ''' <summary>
    ''' 所有Section="Category"（Sections表）的产品填充
    ''' </summary>
    ''' <param name="oldContent"></param>
    ''' <param name="searchString">大类的名称(Categories表)</param>
    ''' <param name="lengthProduct">"[BEGIN_PRODUCT]"字符串的长度</param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function InsertProductOfCategory(ByVal oldContent As String, ByVal searchString As String, ByVal lengthProduct As Integer, ByVal siteId As Integer, ByVal issueId As Integer) As String
        Dim queryProduct = From p In efContext.Products
                         Join pi In efContext.Products_Issue
                         On p.ProdouctID Equals pi.ProductId
                         Where pi.SectionID = "CA" AndAlso pi.SiteId = siteId AndAlso pi.IssueID = issueId
                         Select p

        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")

        For Each q As Product In queryProduct
            Dim categoryName = q.Categories.ToList() '2013/4/15
            Dim cName As New HashSet(Of String)
            For Each c As Category In categoryName
                cName.Add(c.Category1.ToUpper)
            Next

            'If (productStartIndex > -1 And q.Categories.First.Category1.ToUpper = searchString) Then
            If (productStartIndex > -1 And cName.Contains(searchString)) Then
                'Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + length)
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lengthProduct)

                '2013/3/25 新加
                If (newProduct.Contains("[URL]")) Then
                    newProduct = newProduct.Replace("[URL]", q.Url)
                End If
                If (newProduct.Contains("[CATEGORY_ID]")) Then
                    newProduct = newProduct.Replace("[CATEGORY_ID]", q.Categories.First.CategoryID.ToString())
                End If
                If (newProduct.Contains("[PICTURE_SRC]")) Then
                    newProduct = newProduct.Replace("[PICTURE_SRC]", q.PictureUrl)
                End If
                If (newProduct.Contains("[PICTURE_ALT]")) Then
                    newProduct = newProduct.Replace("[PICTURE_ALT]", q.PictureAlt)
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

                oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
                oldContent = oldContent.Insert(productStartIndex, newProduct)
                If (oldContent.Contains("[BEGIN_PRODUCT]")) Then '2013/3/27新增
                    productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                    productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
                Else
                    Exit For
                End If
            End If
        Next
        While (oldContent.Contains("[BEGIN_PRODUCT]"))
            oldContent = oldContent.Remove(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
            If (oldContent.Contains("[BEGIN_PRODUCT]")) Then
                productStartIndex = oldContent.IndexOf("[BEGIN_PRODUCT]", productStartIndex)
                productEndIndex = oldContent.IndexOf("[END_PRODUCT]", productEndIndex)
            End If
        End While
        Return oldContent
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
    Private Shared Function InsertProductOfOthers(ByVal oldContent As String, ByVal issueId As Integer, ByVal sectionId As String, ByVal lengthProduct As Integer, ByVal siteId As Integer)
        Dim queryNewArrival = From issue In efContext.Products_Issue _
                                  Join product In efContext.Products
                                  On issue.ProductId Equals product.ProdouctID
                                  Where issue.SectionID = sectionId And issue.SiteId = siteId And issue.IssueID = issueId
                                  Select product
        Dim productStartIndex As Integer = oldContent.IndexOf("[BEGIN_PRODUCT]")
        Dim productEndIndex As Integer = oldContent.IndexOf("[END_PRODUCT]")
        For Each q In queryNewArrival
            If (productStartIndex > -1) Then
                Dim newProduct As String = oldContent.Substring(productStartIndex + lengthProduct, productEndIndex - productStartIndex - lengthProduct)
                Dim oldProduct As String = oldContent.Substring(productStartIndex, productEndIndex - productStartIndex + lengthProduct)
                If (newProduct.Contains("[URL]")) Then
                    newProduct = newProduct.Replace("[URL]", q.Url)
                End If
                If (newProduct.Contains("[CATEGORY_ID]")) Then
                    newProduct = newProduct.Replace("[CATEGORY_ID]", q.Categories.First.CategoryID.ToString())
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

    ''' <summary>
    ''' 用CategoryID(表：Categories)替换模板中的link_category的值标签
    ''' （如,[CATEGORY_ID_BAGS]用553替换）
    ''' </summary>
    ''' <param name="SpreadTemplate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function HandleCategory(ByVal SpreadTemplate As String, ByVal siteId As Integer) As String
        Dim queryCategory = From c In efContext.Categories Where c.SiteID = siteId
        For Each q In queryCategory
            Dim cateModifyStr As String = q.Category1.ToUpper().Replace(" ", "_") 'Women's Dresses替换成WOMEN'S_DRESSES
            Dim replaceStr As String = "[CATEGORY_ID_" & cateModifyStr & "]"
            SpreadTemplate = SpreadTemplate.Replace(replaceStr, q.CategoryID.ToString())
        Next
        Return SpreadTemplate
    End Function

    ''' <summary>
    ''' 将数据写入到Issues表中
    ''' </summary>
    ''' <param name="nowTime"></param>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <returns>返回插入单条数据的IssueId</returns>
    ''' <remarks></remarks>
    Public Shared Function InsertIssue(ByVal nowTime As DateTime, ByVal siteId As Integer, ByVal planType As String) As Integer
        Dim issue As New Issue()
        issue.IssueDate = nowTime
        issue.Subject = ""
        issue.SiteID = siteId
        issue.PlanType = planType
        efContext.AddToIssues(issue)
        Try
            efContext.SaveChanges()
        Catch ex As Exception
            'LogText(ex.ToString())
        End Try
        Return issue.IssueID
    End Function

End Class
