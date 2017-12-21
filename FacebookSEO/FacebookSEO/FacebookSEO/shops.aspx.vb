Imports System.Net


Public Class shops
    Inherits System.Web.UI.Page
    Dim entity As New FaceBookForSEOEntities()
    Public shops As List(Of AutomationSite)
    'Public Mallsiteid As Integer


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim wbProducts As List(Of Product) = (From p In entity.Products
                                               Select p).ToList()
        Dim type As String
        For Each item As Product In wbProducts
            Dim listKeyWord As List(Of String) = FilterKeyword(item.PictureAlt & item.Description)
            If (item.Currency.Trim = "WB") Then
                type = "tst"
            Else
                type = "fb"
            End If
            InsertKeyWord(listKeyWord, type, item.SiteID)
        Next
    End Sub

    ''' <summary>
    ''' 匹配一串文本中的“#** ” 或“#***#”类型的文字，这些文字组提取出来作为k11Seo关键字用
    ''' </summary>
    ''' <param name="postText"></param>
    ''' <returns>返回文字组list</returns>
    ''' <remarks></remarks>
    Public Function FilterKeyword(ByVal postText As String) As List(Of String)
        Dim listKeyWrod As New List(Of String)
        Dim regexp As String = "#([\w]+?)(?:#|\s)"
        Dim mCollection As MatchCollection = Regex.Matches(postText, regexp)
        Dim keyWord As String
        For i = 0 To mCollection.Count - 1
            keyWord = mCollection.Item(i).Value.Replace("#", "").Trim
            If Not (listKeyWrod.Contains(keyWord)) Then
                listKeyWrod.Add(keyWord)
            End If
        Next
        Return listKeyWrod
    End Function

    ''' <summary>
    ''' 将获取的一串关键字保存至数据库,来源于微博的type='tst',来源于facebook则type='fb'
    ''' </summary>
    ''' <param name="keyWords"></param>
    ''' <param name="type">tst或“”</param>
    ''' <remarks>点击链接url均保存在字段KeyURl</remarks>
    Public Sub InsertKeyWord(ByVal keyWords As List(Of String), ByVal type As String, ByVal siteid As Integer)
        For Each aKey As String In keyWords
            Dim existKey As KeyWord = (From k In entity.KeyWords
                                       Where k.KeyWord1 = aKey
                                       Select k).FirstOrDefault()
            If (existKey Is Nothing) Then
                Dim newKeyWord As New KeyWord()
                newKeyWord.KeyWord1 = aKey
                newKeyWord.Type = type
                newKeyWord.Siteid = siteid
                If (type = "tst") Then
                    newKeyWord.KeyUrl = "/kw/tst/" & aKey
                Else
                    newKeyWord.KeyUrl = "/kw/" & aKey
                End If
                entity.KeyWords.AddObject(newKeyWord)
            Else
                existKey.Siteid = siteid
            End If
        Next
        entity.SaveChanges()
    End Sub

    Public Function updateWBDescription()
        Dim listproduct As List(Of Product) = (From p In entity.Products
                                               Where p.Currency = "WB" And p.Description.Equals(Nothing)
                                                     Order By p.ProdouctID Descending
                                                     Select p).ToList()
        For Each myprod As Product In listproduct

            Dim myRegex As New Regex("#[\w\W]*?#")
            Dim Description As String = myprod.PictureAlt
            Dim m As Match = myRegex.Match(Description)
            If (m.Success) Then
                myprod.PictureAlt = m.Value.Trim
                myprod.Description = Description.Replace(myprod.PictureAlt, "")
            Else
                myprod.Description = myprod.PictureAlt
                myprod.PictureAlt = ""
            End If
        Next
        entity.SaveChanges()
    End Function

    Public Function updateImgSource()

        Dim listproduct As List(Of Product) = (From p In entity.Products
                                                         Where p.PictureUrl.Contains("http")
                                                         Order By p.ProdouctID Descending
                                                         Select p).ToList()

        For Each item In listproduct
            Dim imgsrc As String = DownloadImage(item.PictureUrl, "D:\WebSites\FacebookSEO\", "fbImage", item.SiteID)
            If Not (String.IsNullOrEmpty(imgsrc)) Then
                item.PictureUrl = imgsrc
                item.Prodouct = imgsrc
            End If
            entity.SaveChanges()
        Next
    End Function

    Public Function DownloadImage(ByVal imageUrl As String, ByVal filePath As String, ByVal fileName As String, ByVal siteid As Integer) As String
        Dim client As New WebClient()

        Dim index As Integer = imageUrl.LastIndexOf("/")
        Dim imageName As String 'imageurl中第一个参数值+name作为imagename
        imageName = imageUrl.Substring(index + 1)
        Dim paramIndex As Integer = imageName.IndexOf("?")
        If (paramIndex > 0) Then
            'Dim params As String() = imageName.Split("=")
            imageName = imageName.Substring(0, paramIndex)
        End If
        imageName = Now.Day & "_" & Now.Minute & "_" & Now.Second & "_" & Now.Millisecond & "_" & imageName

        If Not (filePath.EndsWith("\")) Then
            filePath = filePath & "\"
        End If
        Dim datetime As String = Now.Year & "_" & Now.Month
        '以店铺名年月划分文件夹
        filePath = filePath & fileName & "\" & siteid & "_" & datetime
        If Not IO.Directory.Exists(filePath) Then IO.Directory.CreateDirectory(filePath)

        filePath = filePath & "\" & imageName
        Try
            client.DownloadFile(imageUrl, filePath)
            Return "\" & fileName & "\" & siteid & "_" & datetime & "\" & imageName
        Catch ex As Exception
            Common.LogText(ex.StackTrace.ToString)
            Return ""
        End Try
    End Function
End Class