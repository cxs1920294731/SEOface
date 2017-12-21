Imports System.Net.Mail

Public Class NotificationEmail

    Private Shared emailAlerterGroup As String = "autoedm@reasonable.cn"

    Private Shared Function sendEmailWithSMTP(ByVal toEmail As String, ByVal subject As String, ByVal mailBody As String)
        Common.LogText("notification toemail:" & toEmail)
        Try

            Dim smtpClient As New SmtpClient()
            smtpClient.Timeout = 120000
            Dim fromEmail As New MailAddress("autoedm@reasonables.com", "思齐智能邮件提醒")
            'Dim fromEmail As New MailAddress("hluo@reasonables.com", "思齐智能邮件提醒测试")
            Dim toEmaila As New MailAddress(toEmail)
            Dim message As New MailMessage(fromEmail, toEmaila)
            message.Subject = subject
            message.Body = mailBody
            message.IsBodyHtml = True
            smtpClient.Send(message)
        Catch ex As Exception
            Common.LogException(ex)
        End Try

    End Function



    'Private Shared mySpread As New SpreadService.Service
    ''' <summary>
    ''' 当自动化出现错误时，发送电子邮件到组别邮箱emailalerter@reasonables.com中
    ''' </summary>
    ''' <param name="errorMsg"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentErrorEmail(ByVal errorMsg As String)
        sendEmailWithSMTP(emailAlerterGroup, "自动化错误信息", errorMsg)
    End Sub


    ''' <summary>
    ''' 当自动化出现错误时，发送电子邮件到组别邮箱emailalerter@reasonables.com中
    ''' </summary>
    ''' <param name="errorMsg"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentNotificationEmail(ByVal toEmail As String, ByVal subject As String, ByVal mailBody As String)
        sendEmailWithSMTP(toEmail, subject, mailBody)
    End Sub


    ''' <summary>
    ''' 当自动化邮件开始发送，或者创建草稿状态时，发送电子邮件到指定邮箱中
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="content"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentStartEmail(ByVal subject As String, ByVal content As String, ByVal receiverEmail As String)

        sendEmailWithSMTP(receiverEmail.Trim, subject, content)
    End Sub

    ''' <summary>
    ''' 当自动化邮件开始发送，或者创建草稿状态时，发送电子邮件到指定邮箱中
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="content"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentStartEmail(ByVal subject As String, ByVal content As String, ByVal lists() As String, ByVal senderMail As String, ByVal senderName As String,
                                     ByVal campaignName As String, ByVal receiverEmail As String, ByVal isPublished As String)
        Dim myList As String = ""
        'subject = subject & ""
        For Each list As String In lists
            myList = myList & list & ","
        Next
        content = content & vbCrLf & "发送对象：" & myList & vbCrLf & "发件人Email：" & senderMail & vbCrLf & "发件人：" & senderName & vbCrLf & "Campaign名：" & campaignName & vbCrLf & "是否Public to Newsletter Archive：" & isPublished & vbCrLf

        sendEmailWithSMTP(receiverEmail.Trim, subject, content)
    End Sub

    ''' <summary>
    ''' 当自动化邮件开始发送，或者创建草稿状态时，发送电子邮件到组别邮箱emailalerter@reasonables.com中
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="content"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentStartEmail(ByVal subject As String, ByVal content As String)
        'mySpread.Timeout = 120000

        sendEmailWithSMTP(emailAlerterGroup, subject, content)
    End Sub

    ''' <summary>
    ''' 当自动化邮件开始发送，或者创建草稿状态时，发送电子邮件到组别邮箱emailalerter@reasonables.com中
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="content"></param>
    ''' <remarks></remarks>
    Public Shared Sub SentStartEmail(ByVal subject As String, ByVal content As String, ByVal lists() As String, ByVal senderMail As String, ByVal senderName As String,
                                     ByVal campaignName As String, ByVal isPublished As String)
        'mySpread.Timeout = 120000
        Dim myList As String = ""
        'subject = subject & ""
        For Each list As String In lists
            myList = myList & list & ","
        Next
        content = content & vbCrLf & "发送对象：" & myList & vbCrLf & "发件人Email：" & senderMail & vbCrLf & "发件人：" & senderName & vbCrLf & "Campaign名：" & campaignName & vbCrLf & "是否Public to Newsletter Archive：" & isPublished & vbCrLf

        sendEmailWithSMTP(emailAlerterGroup, subject, content)
    End Sub
End Class

