Imports System.Net.Mail

Public Class NotificationEmail

    Private Shared emailAlerterGroup As String = "autoedm@reasonable.cn"

    Private Shared Function sendEmailWithSMTP(ByVal toEmail As String, ByVal subject As String, ByVal mailBody As String)
        'Common.LogText("notification toemail:" & toEmail)
        'Dim smtpClient As New SmtpClient()
        'SmtpClient.Timeout = 120000
        'Dim fromEmail As New MailAddress("autoedm@reasonables.com", "K11 FacebookSEO")
        'Dim toEmaila As New MailAddress(toEmail)

        'Dim message As New MailMessage(fromEmail, toEmaila)
        'message.Subject = subject
        'message.Body = mailBody
        'message.IsBodyHtml = True
        '/*暂且不发送提醒邮件，提醒邮件太多，导致无法专心查看已上线自动化的运行情况*/
        'Try
        '    smtpClient.Send(message)
        'Catch ex As Exception
        '    Common.LogException(ex)
        'End Try
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
        'mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "思齐智能化邮件提醒", "emailalerter@reasonables.com", subject, content)
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
        'mySpread.Send("gtang@reasonables.com", "8A6EEB47-B789-4A70-83E3-8F0BAE78B5E4", "autoedm@reasonables.com", "思齐智能化邮件提醒", "emailalerter@reasonables.com", subject, content)
        sendEmailWithSMTP(emailAlerterGroup, subject, content)
    End Sub
End Class

