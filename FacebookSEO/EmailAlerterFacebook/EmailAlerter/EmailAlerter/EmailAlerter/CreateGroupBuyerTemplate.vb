Imports System.Linq

Public Class CreateGroupBuyerTemplate
    Private Shared efContext As New FaceBookForSEOEntities

    Public Shared Function GetTemplate(ByVal templateId As Integer, ByVal issueId As Integer, ByVal GACode As String) As String
        Dim template As String = efContext.AutoTemplates.Where(Function(t) t.Tid = templateId).Single().Contents
        Dim listProducts As List(Of AutoProduct) = (From p In efContext.AutoProducts
                Join pi In efContext.AutoProducts_Issue On p.ProdouctID Equals pi.ProductId
                Where (String.IsNullOrEmpty(p.LastUpdate) OrElse p.LastUpdate > Now) AndAlso pi.IssueID = issueId
                Order By p.Sales Descending
                Select p).ToList()



        Dim arrTemplate() As String = template.Split("^")
        'FTemplate = Templates(0)
        's = FTemplate.Replace(vbCrLf, "")
        'NLeftTemplate = Templates(1)
        'NRightTemplate = Templates(2)
        Dim FTemplate As String = arrTemplate(0)
        Dim NLeftTemplate As String = arrTemplate(1)
        Dim NRightTemplate As String = arrTemplate(2)

        '2014/04/01增加三个一行的
        Dim LeftThreeTemplate As String = arrTemplate(3)
        Dim MiddleThreeTemplate As String = arrTemplate(4)
        Dim RightThreeTemplate As String = arrTemplate(5)

        Dim sign As Integer = 0
        Dim MainDealsHtml As String = ""
        Dim counter As Integer = 0

        Dim randomIndex As New List(Of Integer)
        GetRandom(randomIndex, listProducts.Count) '产品的填充按随机顺序
        For Each index As Integer In randomIndex
            Dim li As AutoProduct = listProducts.ElementAt(index)
            If (counter >= 10) Then '2014/04/01 added,前五行每行放两个产品
                Exit For
            Else
                Dim listCategorys As List(Of AutoCategory) = li.Categories.ToList()
                For Each cate As AutoCategory In listCategorys
                    'If (cate.Category1 = "-5") Then --由dora注释掉，因新增了beauty专题，所以将此限定去掉20140926
                    If (sign = 0) Then
                        Dim NLdeal As String = ""
                        If (li.Discount Is Nothing) Then
                            NLdeal = String.Format(NLeftTemplate, li.Url,
                                                   li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                         "http://c.groupbuyermail.com"), li.Prodouct,
                                                   "", "", li.Url)
                        ElseIf (li.Price Is Nothing And li.Discount IsNot Nothing) Then
                            Dim discount As Double = li.Discount
                            NLdeal = String.Format(NLeftTemplate, li.Url,
                                                   li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                         "http://c.groupbuyermail.com"), li.Prodouct,
                                                   li.Currency & Format("{0:#.00}", discount), "", li.Url)
                        Else
                            Dim price As Double = li.Price
                            Dim discount As Double = li.Discount
                            NLdeal = String.Format(NLeftTemplate, li.Url,
                                                   li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                         "http://c.groupbuyermail.com"), li.Prodouct,
                                                   li.Currency & Format("{0:#.00}", discount),
                                                   li.Currency & Format("{0:#.00}", price), li.Url)
                        End If
                        'NLdeal = String.Format(NLeftTemplate, li.Url, li.PictureUrl, li.Prodouct, li.Discount, li.Price, li.Url)
                        MainDealsHtml = MainDealsHtml & NLdeal
                        sign = 1
                    ElseIf (sign = 1) Then
                        Dim NLdeal2 As String
                        '= String.Format(NRightTemplate, li.Url, li.PictureUrl, li.Prodouct, li.Discount, li.Price, li.Url)
                        If (li.Discount Is Nothing) Then
                            NLdeal2 = String.Format(NRightTemplate, li.Url,
                                                    li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                          "http://c.groupbuyermail.com"),
                                                    li.Prodouct, "", "", li.Url)
                        ElseIf (li.Price Is Nothing AndAlso li.Discount IsNot Nothing) Then
                            Dim discount As Double = li.Discount
                            NLdeal2 = String.Format(NRightTemplate, li.Url,
                                                    li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                          "http://c.groupbuyermail.com"),
                                                    li.Prodouct, li.Currency & Format("{0:#.00}", discount), "",
                                                    li.Url)
                        Else
                            Dim price As Double = li.Price
                            Dim discount As Double = li.Discount
                            NLdeal2 = String.Format(NRightTemplate, li.Url,
                                                    li.PictureUrl.Replace("http://www.groupbuyer.com.hk",
                                                                          "http://c.groupbuyermail.com"),
                                                    li.Prodouct, li.Currency & Format("{0:#.00}", discount),
                                                    li.Currency & Format("{0:#.00}", price), li.Url)
                        End If
                        MainDealsHtml = MainDealsHtml & NLdeal2
                        sign = 0
                    End If
                    counter = counter + 1
                    Exit For
                    '只要把分类为5的产品填充好了即可
                    'End If --由dora注释掉，因新增了beauty专题，所以将此限定去掉20140926
                Next
            End If
        Next

        If (counter Mod 2 = 1) Then
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" class=""mobile_hidden"" width=""13""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table width=""315"" align=""left"" border=""0"" cellspacing=""0"" valign=""top"" class=""mobile_hidden""><tbody><tr><td><table align=""left"" width=""315"" height=""333"" border=""0"" cellspacing=""0"" cellpadding=""5"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><tbody><tr><td style=""width:317px;height:333px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""305"" height=""367px"" border=""0"" alt=""""></a></td></tr></tbody></table></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""700"" style=""clear: both;"" class=""mobile_hidden""><tbody><tr><td style=""height: 16px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""1"" height=""16"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "</td></tr></table></td></tr>"
        End If
        Dim spreadTemplate As String = FTemplate.Replace("[ALL_NEW_DEALS]", MainDealsHtml)

        '2014/04/01增加三个一行的
        Dim counter2 As Integer = 0
        Dim sign2 As Integer = 0
        MainDealsHtml = ""
        If (listProducts.Count > 10) Then
            For i As Integer = 10 To randomIndex.Count - 1
                Dim jiindex As Integer = randomIndex(i)
                Dim listCategorys As List(Of AutoCategory) = listProducts(jiindex).Categories.ToList()
                For Each cate As AutoCategory In listCategorys
                    'If (cate.Category1 = "-5") Then --由dora注释掉，因新增了beauty专题，所以将此限定去掉20140926
                    If (sign2 = 0) Then
                        Dim NThreeDeal As String = ""
                        If (listProducts(jiindex).Discount Is Nothing) Then
                            NThreeDeal = String.Format(LeftThreeTemplate, listProducts(jiindex).Url,
                                                       listProducts(jiindex).PictureUrl.Replace(
                                                           "http://www.groupbuyer.com.hk",
                                                           "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                       "", "", listProducts(jiindex).Url)
                        ElseIf (listProducts(jiindex).Price Is Nothing And listProducts(jiindex).Discount IsNot Nothing) Then
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal = String.Format(LeftThreeTemplate, listProducts(jiindex).Url,
                                                       listProducts(jiindex).PictureUrl.Replace(
                                                           "http://www.groupbuyer.com.hk",
                                                           "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                       listProducts(jiindex).Currency & Format("{0:#.00}", discount), "",
                                                       listProducts(jiindex).Url)
                        Else
                            Dim price As Double = listProducts(jiindex).Price
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal = String.Format(LeftThreeTemplate, listProducts(jiindex).Url,
                                                       listProducts(jiindex).PictureUrl.Replace(
                                                           "http://www.groupbuyer.com.hk",
                                                           "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                       listProducts(jiindex).Currency & Format("{0:#.00}", discount),
                                                       listProducts(jiindex).Currency & Format("{0:#.00}", price),
                                                       listProducts(jiindex).Url)
                        End If
                        MainDealsHtml = MainDealsHtml & NThreeDeal
                        sign2 = 1
                    ElseIf (sign2 = 1) Then
                        Dim NThreeDeal2 As String = ""
                        If (listProducts(jiindex).Discount Is Nothing) Then
                            NThreeDeal2 = String.Format(MiddleThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        "", "", listProducts(jiindex).Url)
                        ElseIf (listProducts(jiindex).Price Is Nothing And listProducts(jiindex).Discount IsNot Nothing) Then
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal2 = String.Format(MiddleThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", discount), "",
                                                        listProducts(jiindex).Url)
                        Else
                            Dim price As Double = listProducts(jiindex).Price
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal2 = String.Format(MiddleThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", discount),
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", price),
                                                        listProducts(jiindex).Url)
                        End If
                        MainDealsHtml = MainDealsHtml & NThreeDeal2
                        sign2 = 2
                    ElseIf (sign2 = 2) Then
                        Dim NThreeDeal3 As String = ""
                        If (listProducts(jiindex).Discount Is Nothing) Then
                            NThreeDeal3 = String.Format(RightThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        "", "", listProducts(jiindex).Url)
                        ElseIf (listProducts(jiindex).Price Is Nothing And listProducts(jiindex).Discount IsNot Nothing) Then
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal3 = String.Format(RightThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", discount), "",
                                                        listProducts(jiindex).Url)
                        Else
                            Dim price As Double = listProducts(jiindex).Price
                            Dim discount As Double = listProducts(jiindex).Discount
                            NThreeDeal3 = String.Format(RightThreeTemplate, listProducts(jiindex).Url,
                                                        listProducts(jiindex).PictureUrl.Replace(
                                                            "http://www.groupbuyer.com.hk",
                                                            "http://c.groupbuyermail.com"), listProducts(jiindex).Prodouct,
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", discount),
                                                        listProducts(jiindex).Currency & Format("{0:#.00}", price),
                                                        listProducts(jiindex).Url)
                        End If
                        MainDealsHtml = MainDealsHtml & NThreeDeal3
                        sign2 = 0
                    End If
                    counter2 = counter2 + 1
                    Exit For
                    '只要把分类为5的产品填充好了即可
                    'End If --由dora注释掉，因新增了beauty专题，所以将此限定去掉20140926
                Next
            Next
        End If
        If counter2 Mod 3 = 1 Then
            'OtherDealsHtml = OtherDealsHtml & "<td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""16""></td><td width=""204"" style=""border: 1px solid #999999; border-collapse: collapse;"" valign=""top""><a href=""http://www.groupbuyer.com.hk/zh/special"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_2.jpg"" width=""194"" height=""250"" border=""0"" alt="""" /></a></td><td width=""25""></td></tr>"
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289px;""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289px"" border=""0"" alt="""" /></a></td></tr></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "</td></tr>"
        ElseIf counter2 Mod 3 = 2 Then
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" cellpadding=""0"" cellspacing=""0"" border=""0"" width=""13"" class=""mobile_hidden""><tbody><tr><td style=""height: 1px;""><img alt="""" style=""display: block; margin: 0px;"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""13"" height=""1"" /></td></tr></tbody></table><table align=""left"" width=""208"" border=""0"" cellspacing=""0"" cellpadding=""0""  valign=""top"" style=""border: 1px solid #999999; border-collapse: collapse;"" class=""mobile_hidden""><tr><td style=""height:289""><a href=""http://www.groupbuyer.com.hk/zh/hot"" target=""_blank""><img style=""display:block"" src=""http://app.rspread.com/spreaderfiles/6819/image/more_1.jpg"" width=""194"" height=""289"" border=""0"" alt="""" /></a></td></tr></table>"
            MainDealsHtml = MainDealsHtml &
                            "<table align=""left"" width=""20"" border=""0"" cellspacing=""0"" cellpadding=""0""><tbody><tr><td class=""mobile_hidden""><img alt="""" style=""display: block; margin-left: 0px; margin-right: 0px;"" id=""rw-img-25"" src=""http://app.rspread.com/spreaderfiles/16577/182833/output/img/trans.gif"" width=""20"" height=""1"" /></td></tr></tbody></table>"
            MainDealsHtml = MainDealsHtml & "</td></tr>"
        End If
        spreadTemplate = spreadTemplate.Replace("[ALL_OLD_DEALS]", MainDealsHtml)
        spreadTemplate = SpecialCode.AddSpecialCode(GACode, spreadTemplate)
        Return spreadTemplate
    End Function

    ''' <summary>
    ''' 生成一个不重复的随机数的list。从0开始至totalcount-1
    ''' </summary>
    ''' <param name="randomIndex"></param>
    ''' <param name="totalCount"></param>
    ''' <remarks></remarks>
    Private Shared Sub GetRandom(ByRef randomIndex As List(Of Integer), ByVal totalCount As Integer)
        Dim oRand As New Random(Now.Millisecond)
        Dim oTemp As Integer = -1
        Do Until randomIndex.Count = totalCount
            oTemp = oRand.Next(0, totalCount)
            If Not randomIndex.Contains(oTemp) Then randomIndex.Add(oTemp)
        Loop
    End Sub
End Class
