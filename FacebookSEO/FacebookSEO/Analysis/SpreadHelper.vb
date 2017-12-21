Imports System.IO
Imports System.IO.Packaging
Imports Analysis.SpreadWebReference
Imports System.Text
Imports System.Configuration

Public Class SpreadHelper

    Private Shared myspread As Service = New Service()

    Private Shared Function ConfigSpreadWebServiceUrl()
        myspread.Url = ConfigurationManager.AppSettings("SpreadWebServiceURl").ToString.Trim
    End Function

    ''' <summary>
    ''' 添加发件人地址，创建成功返回TRUE。如果senderemail已存在或不合法，则会返回false
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="apiKey"></param>
    ''' <param name="senderEmail"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddSenderEmail(ByVal spreadAccount As String, ByVal apiKey As String, ByVal senderEmail As String) As Boolean
        Try
            ConfigSpreadWebServiceUrl()
            Return myspread.AddSenderEmail(spreadAccount, apiKey, senderEmail)
        Catch ex As Exception
            EFHelper.Log(ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取spread账号的ApiKey
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="password"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getSpreadApiKey(ByVal spreadAccount As String, ByVal password As String) As String
        ConfigSpreadWebServiceUrl()
        Dim apiKey As New Guid
        Try
            apiKey = Guid.Parse(myspread.GetAPIKey(spreadAccount, password))
            Return apiKey.ToString.Trim.ToUpper()
        Catch ex As Exception
            Return ""
        End Try
    End Function


    Public Shared Function IsApikeyMatched(ByVal spreadAccount As String, ByVal apikey As String) As Boolean
        ConfigSpreadWebServiceUrl()
        spreadAccount = spreadAccount.Trim
        apikey = apikey.Trim
        If Not (String.IsNullOrEmpty(spreadAccount) OrElse String.IsNullOrEmpty(apikey)) Then
            Try
                myspread.GetAccessToken(spreadAccount, apikey)
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 调用spreadAPI uplodeZipFile（）方法上传一个campaign的内容
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="apiKey"></param>
    ''' <param name="contentsByte"></param>
    ''' <param name="camId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UpdateCampaignContentWithZipFile(ByVal spreadAccount As String, ByVal apiKey As String, ByVal contentsByte() As Byte, ByVal camId As Integer)
        ConfigSpreadWebServiceUrl()
        Return myspread.UplodeZipFile(spreadAccount, apiKey, contentsByte, camId)
    End Function

    ''' <summary>
    ''' 将一个文件压缩成".zip"格式压缩包
    ''' </summary>
    ''' <param name="zipFileName"></param>
    ''' <param name="fileToAdd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddFileToZip(ByVal zipFileName As String, ByVal fileToAdd As String)
        Using zip As Package = Packaging.Package.Open(zipFileName, FileMode.Create)
            Dim destFilename As String = ".\" & Path.GetFileName(fileToAdd)
            Dim uri As Uri = PackUriHelper.CreatePartUri(New Uri(destFilename, UriKind.Relative))
            If (zip.PartExists(uri)) Then
                zip.DeletePart(uri)
            End If
            Dim part As PackagePart = zip.CreatePart(uri, "", CompressionOption.Normal)
            Using fileStream As FileStream = New FileStream(fileToAdd, FileMode.Open, FileAccess.Read)
                Using dest As Stream = part.GetStream()
                    CopyStream(fileStream, dest)
                End Using
            End Using
        End Using
    End Function

    Private Shared Function CopyStream(ByVal inputStream As FileStream, ByVal outputStream As Stream)
        Dim bufferSize As Long = IIf(inputStream.Length < 5000, inputStream.Length, 5000)
        Dim buffer() As Byte = New Byte(bufferSize) {}
        Dim bytesRead As Integer = 0
        Dim bytesWritten As Long = 0
        bytesRead = inputStream.Read(buffer, 0, buffer.Length)
        While (Not bytesRead = 0)
            outputStream.Write(buffer, 0, bytesRead)
            bytesWritten += bufferSize
            bytesRead = inputStream.Read(buffer, 0, buffer.Length)
        End While
    End Function

    ''' <summary>
    ''' 创建一个草稿内容的.htm文件，文件命名规则为siteid + ampaignID ＋ yyyyMMddHHmm.htm，创建成功返回文件名，否则返回空
    ''' </summary>
    ''' <param name="contents">文件内容</param>
    ''' <param name="siteName"></param>
    ''' <param name="campaignId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateFile(ByVal contents As String, ByVal filePath As String, ByVal shopName As String, ByVal campaignId As Integer) As String
        Dim streamW As StreamWriter
        Dim fileS As FileStream
        Try
            Dim fileName As String = shopName & "_CamID_" & campaignId & "_" & Date.Now.ToString("yyyyMMddHHmm") & ".htm"
            fileS = New FileStream(filePath & fileName, FileMode.Create)

            streamW = New StreamWriter(fileS, Encoding.GetEncoding("utf-8"))
            streamW.Write(contents)
            streamW.Flush()
            Return fileName
        Catch ex As Exception
            EFHelper.Log(ex)
            Return ""
        Finally
            streamW.Close()
            fileS.Close()
        End Try
    End Function


    ''' <summary>
    ''' 返回账号下状态不是删除的所有名单，返回结果各名单以逗号分隔串联组成一个string
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <returns>String:"contact list 20120203,contact list 20120204,contact list 20120205"</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSubscription2String(ByVal spreadAccount As String) As String
        ConfigSpreadWebServiceUrl()
        Dim activeResult As String = myspread.getSubscriptions2String(spreadAccount, SEmailEfHelper.getApiKey(spreadAccount), SubscriptionStatus.Active).Trim
        Dim invisibleResult As String = myspread.getSubscriptions2String(spreadAccount, SEmailEfHelper.getApiKey(spreadAccount), SubscriptionStatus.Invisible)

        If String.IsNullOrEmpty(activeResult) And String.IsNullOrEmpty(invisibleResult) Then
            Return ""
        ElseIf String.IsNullOrEmpty(activeResult) And Not String.IsNullOrEmpty(invisibleResult) Then
            Return invisibleResult
        ElseIf Not String.IsNullOrEmpty(activeResult) And String.IsNullOrEmpty(invisibleResult) Then
            Return activeResult
        ElseIf Not String.IsNullOrEmpty(activeResult) And Not String.IsNullOrEmpty(invisibleResult) Then
            Return activeResult & "," & invisibleResult
        End If

    End Function

    ''' <summary>
    ''' 获取名单下有效的Subscription数量
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="password"></param>
    ''' <param name="cate"></param>
    ''' <returns></returns>
    Public Shared Function GetSubscriptionCount(ByVal spreadAccount As String, ByVal password As String, ByVal cate As String) As Integer
        Try
            Return myspread.getActiveSubscribersByContactList(spreadAccount, password, cate)
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ''' <summary>
    ''' 获取发送名单
    ''' </summary>
    ''' <param name="cate"></param>
    ''' <returns></returns>
    Public Shared Function GetFilterCategory(ByVal spreadAccount As String, ByVal password As String, ByVal cates As String(), ByVal isNetEase As Boolean) As String

        Return myspread.ExcludeContactsByNetEase(spreadAccount, password, cates, isNetEase)
    End Function

    Public Shared Function GetNtesLevelCount(account As String, ByVal password As String) As Integer
        Return myspread.GetNetEaseRank(account, password)
    End Function
End Class
