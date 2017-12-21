Imports Newtonsoft.Json.Linq
Imports Analysis

Public Class PublicMethod
    '以下两个成员同意存在于/js/smartEmail.js文件，其值需要与/js/smartEmail.js文件中一致,以/js/smartEmail.js文件中为准
    Public Shared Property SplitChar As String = ";"     '//用于分隔一个产品分类的name及url
    Public Shared Property UrlSplitChar As String = ","  '//用于分隔一个产品分类的name及url
    Public Shared Property AutoPlanCategoriesSplitchar As String = "^"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="splitChar"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Array2string(ByRef arr As String(), ByVal splitChar As Char) As String
        Dim arrString As String = ""
        For Each item In arr
            arrString = item & splitChar & arrString
        Next
        arrString = arrString.TrimEnd(splitChar).Trim
        Return arrString
    End Function

    ''' <summary>
    ''' 规范化自动化计划发送间隔的天数
    ''' </summary>
    ''' <param name="intervalDay"></param>
    ''' <remarks></remarks>
    Public Shared Sub AutoPlanInterval(ByRef intervalDay As Integer)
        Select Case intervalDay
            Case 0 To 6
                intervalDay = 1
            Case 7 To 13
                intervalDay = 10
            Case 14 To 20
                intervalDay = 17
            Case 21 To 200
                intervalDay = 24
            Case Else
                intervalDay = 1
        End Select
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="timeLimit"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' TimeSpan: Send time limit of campaign. format like: [ { '1': ['01:00-02:00','01:00-02:00'] }, { '2': ['01:00-02:00','01:00-02:00'] } ]
    ''' '0' means every day, '1': Monday, '2': Tuesday,'3': Wednesday, '4': Thursday,'5': Friday, '6': Saturday, '7': Sunday
    ''' 
    ''' </remarks>
    Public Shared Function DeserializeTimeLimit(ByVal timeLimit As String) As List(Of String)
        Dim activeDays As String() = timeLimit.TrimStart("[").TrimEnd("]").Split(",")
        Dim activeSendDays As New List(Of String)
        For Each item In activeDays
            If Not String.IsNullOrEmpty(item) Then
                Dim activeweekday As String = item.Split(":")(0).Split("'")(1)

                Select Case activeweekday
                    Case 6, 7
                        activeSendDays.Add(activeweekday)
                    Case 5
                        activeSendDays.Add(8) '代表工作日
                    Case Else

                End Select
            End If

        Next
        Return activeSendDays
    End Function

    Public Shared Function SerializeTimeLimit(ByVal activeSendDays As List(Of Integer)) As String
        Dim timeLimit As String = ""
        For Each item In activeSendDays
            Select Case item
                Case 8
                    If (String.IsNullOrEmpty(timeLimit)) Then
                        timeLimit = "{'1': ['00:00-23:00']},{'2': ['00:00-23:00']},{'3': ['00:00-23:00']},{'4': ['00:00-23:00']},{'5': ['00:00-23:00']}"
                    Else
                        timeLimit = timeLimit & "," & "{'1': ['00:00-23:00']},{'2': ['00:00-23:00']},{'3': ['00:00-23:00']},{'4': ['00:00-23:00']},{'5': ['00:00-23:00']}"
                    End If
                Case 6, 7
                    If (String.IsNullOrEmpty(timeLimit)) Then
                        timeLimit = "{'" & item & "': ['00:00-23:00']}"
                    Else
                        timeLimit = timeLimit & "," & "{'" & item & "': ['00:00-23:00']}"
                    End If
            End Select
        Next
        Return "[" & timeLimit & "]"
    End Function

    ''' <summary>
    ''' 从页面的tbContentItems的内容整合成list cate
    ''' </summary>
    ''' <param name="tbContentItems"></param>
    ''' <param name="hasBanner"></param>
    ''' <param name="categories"></param>
    ''' <param name="shopUrl"></param>
    ''' <param name="siteId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCateListFromPage(ByVal tbContentItems As String, ByRef hasBanner As Boolean, ByRef categories As String, ByRef firstCatename As String,
                                               ByVal shopUrl As String, ByVal siteId As Integer, ByVal ebusinessSite As EBusinessSite) As List(Of Category)
        Dim i As Integer = 0
        Dim cates As New List(Of Category)
        Dim cateStrs As String = tbContentItems
        Dim cateStrArray As String() = cateStrs.Split(SplitChar)
        For Each item As String In cateStrArray
            Dim cate As New Category
            If Not (String.IsNullOrEmpty(item)) Then
                cate.Category1 = item.Split(UrlSplitChar)(0).Trim()
                If (cate.Category1 = "焦点图") Then
                    hasBanner = True
                Else
                    categories = categories & cate.Category1 & AutoPlanCategoriesSplitchar
                    If (i = 0) Then
                        firstCatename = cate.Category1
                    End If
                    i = i + 1
                    cate.Url = ebusinessSite.GetSortedTypeURl(cate.Category1)
                    If (String.IsNullOrEmpty(cate.Url)) Then
                        cate.Url = item.Split(UrlSplitChar)(1).Trim
                    End If
                    cate.Description = cate.Category1
                    cate.SiteID = siteId
                    cate.LastUpdate = DateTime.Now
                    cates.Add(cate)
                End If
            End If
        Next
        categories = categories.TrimEnd(AutoPlanCategoriesSplitchar)
        Return cates
    End Function



    ''' <summary>
    ''' 返回所有副模板组成的tid字符串，如115,25
    ''' </summary>
    ''' <param name="tbSubTemplate"></param>
    ''' <returns></returns>
    Public Shared Function GetMultiTemplataIdFromListBox(ByVal tbSubTemplate As String) As String
        If String.IsNullOrEmpty(tbSubTemplate) Then
            Return ""
        Else
            Dim subTemplateArray() = tbSubTemplate.TrimEnd(";").Split(";")  ' Tmall_Red,115;HomePage,25;
            Dim returnStr = ""
            For Each Template As String In subTemplateArray
                returnStr = returnStr & Template.Split(",")(1) & ","
            Next
            Return returnStr.TrimEnd(",")
        End If

    End Function

    Public Shared Function CreateEBusinessSiteInstance(ByVal shopurltext As String) As EBusinessSite
        shopurltext = shopurltext.ToLower.Trim()
        Dim ebusinessSite As EBusinessSite
        If (shopurltext.Contains("tmall.com") Or shopurltext.Contains("taobao.com")) Then
            ebusinessSite = New TmallTaobaoSite(shopurltext)
        ElseIf (shopurltext.Contains("ebay.com")) Then
            ebusinessSite = New EbaySite(shopurltext)
        Else
            ebusinessSite = New TmallTaobaoSite(shopurltext)
        End If
        Return ebusinessSite
    End Function


    Public Shared Function SubjectHandle(ByVal subject As String, ByVal firstProduct As String, ByVal catename As String)
        subject = subject.Replace("[FIRST_PRODUCT]", firstProduct).Replace("[VOL_NUMBER]", "01").Replace("[CATE_NAME]", catename)
        Return subject
    End Function

End Class


