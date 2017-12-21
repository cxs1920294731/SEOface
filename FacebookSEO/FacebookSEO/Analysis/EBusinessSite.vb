Imports System.Text.RegularExpressions

Public MustInherit Class EBusinessSite
    Public Property SortedType As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public Property ShopType As String
    Public Property ShopUrl As String
    Public Property ShopName As String
    Public Property DefaultSubject As String
    Public Property DefaultTriggerSubject As String
    Public Property TemplateType As String
    Public Property BannerImgRegex As String


    ''' <summary>
    ''' 获取统一dom结构的店铺页面上的产品
    ''' </summary>
    ''' <param name="pageUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function GetProductList(ByVal pageUrl As String, ByVal siteId As Integer) As List(Of Product)

    ''' <summary>
    ''' 获取店铺的名称及logo。其中list(0)-->shopName , list(1)-->Logo
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function GetShopNameandLogo() As String()


    Public MustOverride Function GetSortedTypeURl(ByVal cateName As String) As String

    Public MustOverride Function StandardlizeShopURl(ByVal shopUrl As String) As String

    ''' <summary>
    ''' 返回排序好的一个复合字符串，'(1；2；5；3；4；6$$h1；h2；h3；h4；h5;$$s1,s2,s3,s4,s5）
    ''' 第一部分是排好的索引，第二部分是原始图源串，第三部分为原始图大小
    ''' 用第一部分的索引，取第二第三部分的信息
    ''' </summary>
    ''' <param name="bannerFromUrl"></param>
    ''' <returns></returns>
    Public Function MatchBannerImgs(ByVal bannerFromUrl As String) As String
        If Not (String.IsNullOrEmpty(bannerFromUrl)) Then
            Dim hrefs As String = ""
            Dim sizeStr As String = ""
            Dim txtDocHtml As String = EFHelper.GetHtmlStringByUrlAli(bannerFromUrl)
            Dim mCollection As MatchCollection = Regex.Matches(txtDocHtml, BannerImgRegex)
            For i As Integer = 0 To mCollection.Count - 1
                Dim href As String = mCollection(i).Groups(1).Value
                hrefs = hrefs & EFHelper.AddHttpForAli(href.Trim) & ";"
            Next
            If String.IsNullOrEmpty(hrefs) Then
                Return ""
            End If
            Dim result As Dictionary(Of Integer, Integer) = SortBanner(hrefs.TrimEnd(";"), sizeStr)
            Dim indexOrders As String
            For Each key As Integer In result.Keys
                indexOrders = indexOrders & key & ";"
            Next
            'Return bannerSrcs.TrimEnd(";")
            Dim hybirdStr As String = indexOrders.TrimEnd(";") & "$$" & hrefs.TrimEnd(";") & "$$" & sizeStr.TrimEnd(";") '(1；2；5；3；4；6$$h1；h2；h3；h4；h5；h6）
            Return hybirdStr
        End If
    End Function


    ''' <summary>
    ''' 对一串href分析并排序，好banner排最前，且从大到小
    ''' </summary>
    ''' <param name="hrefs">以；分隔的url</param>
    ''' <returns>返回字典，key为序号，value为尺寸</returns>
    ''' <remarks></remarks>
    Public Function SortBanner(ByVal hrefs As String, ByRef sizeStr As String) As Dictionary(Of Integer, Integer)
        Dim bannerArray As String()
        Dim goodBannerDict As Dictionary(Of Integer, Integer) = New Dictionary(Of Integer, Integer)
        Dim badBannerDict As Dictionary(Of Integer, Integer) = New Dictionary(Of Integer, Integer)
        Dim size As Integer
        bannerArray = hrefs.Split(";")
        For index = 0 To bannerArray.Length - 1
            Try
                Dim imgRequst As System.Net.WebRequest = System.Net.WebRequest.Create(bannerArray(index))
                Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(imgRequst.GetResponse().GetResponseStream())
                size = image.Size.Height * image.Size.Width
                sizeStr = sizeStr & image.Size.Width & "*" & image.Size.Height & ";"
            Catch ex As Exception
                size = 0
                sizeStr = sizeStr & "0*0;"
            End Try
            If IsGoodBanner(bannerArray(index)) Then
                goodBannerDict.Add(index, size)
            Else
                badBannerDict.Add(index, size)
            End If
        Next
        'sort by size and combine into one dictionary
        goodBannerDict = (From entry In goodBannerDict Order By entry.Value Descending Select entry).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        badBannerDict = (From entry In badBannerDict Order By entry.Value Descending Select entry).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        Dim combineDict As Dictionary(Of Integer, Integer) = goodBannerDict.Union(badBannerDict).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        Return combineDict
    End Function


    ''' <summary>
    '''  callback of generate  banner thumbnail to scan every pixel of a picture
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ThumbnailCallback() As Boolean
        'LogText("GetThumbnailImageAbort,exit and try next banner")
        Return False
    End Function


    ''' <summary>
    ''' 判断是否为符合条件的banner
    ''' </summary>
    ''' <param name="imgurl">传入url</param>
    ''' <returns>匹配则返回true</returns>
    ''' <remarks></remarks>
    Public Function IsGoodBanner(ByVal imgurl As String) As Boolean
        Dim height As Double = 0
        Dim width As Double = 0
        Dim abortCallback As New System.Drawing.Bitmap.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        Dim imgRequst As System.Net.WebRequest = System.Net.WebRequest.Create(imgurl)
        Dim originalImage As System.Drawing.Image
        Try
            originalImage = System.Drawing.Image.FromStream(imgRequst.GetResponse().GetResponseStream())
        Catch ex As Exception
            Return False
        End Try
        Dim thumbnails As System.Drawing.Bitmap = originalImage.GetThumbnailImage(100, 80, abortCallback, IntPtr.Zero)
        Dim color As System.Drawing.Color
        Dim colorDictionary As Dictionary(Of System.Drawing.Color, Integer) = New Dictionary(Of System.Drawing.Color, Integer)
        'originalImage.Save(System.AppDomain.CurrentDomain.BaseDirectory.ToString + "banneroriginal")
        'thumbnails.Save(System.AppDomain.CurrentDomain.BaseDirectory.ToString + "bannerthumnail")
        Dim watch As Stopwatch = Stopwatch.StartNew()
        For y As Integer = 0 To thumbnails.Height - 1   'Count all pixs with dictionary
            For x As Integer = 0 To thumbnails.Width - 1
                color = thumbnails.GetPixel(x, y)
                If Not colorDictionary.ContainsKey(color) Then
                    colorDictionary.Add(color, 1)
                Else
                    colorDictionary.Item(color) = colorDictionary.Item(color) + 1
                End If
            Next
        Next
        watch.Stop()
        Dim processTime As Long = watch.ElapsedMilliseconds
        'sort colorDictionary by value to get max pixel 
        colorDictionary = (From entry In colorDictionary Order By entry.Value Descending Select entry).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        Dim maxPixelRate As Double = colorDictionary.Values(0) / (thumbnails.Width * thumbnails.Height)  '单一像素最大占比，通常为白色像素最大占比
        height = originalImage.Size.Height
        width = originalImage.Size.Width

        If (height > 300 And width > 700 And height / width > 0.26 And height / width < 0.74) Then  'Level1 size filter
            If (colorDictionary.Count > 10 And maxPixelRate < 0.85) Then  'level2 pixel filter, simple color or wide range pure color are not good banner
                Return True  '找到合适bannerindex，返回结果
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Public Function BindDefaultTemplates(ByVal spreadAccount As String, ByVal ddlTemplate As Global.System.Web.UI.WebControls.DropDownList)
        Dim ef As New EFHelper()
        Dim commTemps As List(Of Template) = ef.GetCommonTemplates(TemplateType)
        Dim specialTemps As New List(Of Template)
        If Not (String.IsNullOrEmpty(spreadAccount)) Then
            specialTemps = ef.GetTemplates(spreadAccount)
            commTemps = commTemps.Union(specialTemps).ToList()
        End If
        ddlTemplate.DataSource = commTemps
        ddlTemplate.DataTextField = "TemplateName"
        ddlTemplate.DataValueField = "Tid"
        ddlTemplate.DataBind()
    End Function


    '    cateName = cateName.Trim
    '    If (ShopUrl.EndsWith("/")) Then
    '        ShopUrl = ShopUrl.Substring(0, ShopUrl.Length - 1)
    '    End If
    'Dim sortedcaterul = SortedType(cateName)
    '    If Not (sortedcaterul.StartsWith("/")) Then
    '        sortedcaterul = "/" & sortedcaterul
    '    End If
    '    Return ShopUrl & sortedcaterul

End Class
