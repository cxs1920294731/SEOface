

Public Class Common
    Shared LogCache As New ArrayList
    Const N As Integer = 1

    Public Shared Sub LogText(ByVal LogText As String)
        '2013/08/08 added, 发送错误日志到制定的邮箱组
        If (LogText.Contains("Exception")) Then
            NotificationEmail.SentErrorEmail(LogText)
        End If
        ' Cached write log
        LogText = String.Format("At {0} ({1}), {2}", Now.ToString, LogCache.Count, LogText & ControlChars.NewLine)
        LogCache.Add(LogText)
        If LogCache.Count >= N Then 'cannot load configuration manager
            Dim Folder As String = ""
            'Folder = System.Web.HttpContext.Current.Server.MapPath("~/App_Data")
            Folder = System.AppDomain.CurrentDomain.BaseDirectory.ToString + "App_Data\Send"
            'Folder = System.Net.System.Reflection.Assembly.GetExecutingAssembly.Location
            If Not IO.Directory.Exists(Folder) Then IO.Directory.CreateDirectory(Folder)
            Dim LogFile As String = Folder & "\" & Now.Year & "-" & Now.Month & "-" & Now.Day & ".log"
            SyncLock (LogCache.SyncRoot)
                Try
                    System.IO.File.AppendAllText(LogFile, String.Join("", CType(LogCache.ToArray(System.Type.GetType("System.String")), String())))
                    'System.IO.File.AppendAllText(LogFile, LogText)
                    LogCache.Clear()
                Catch ex As Exception 'ignore
                    System.IO.File.AppendAllText(LogFile, "Error, cache size = " & LogCache.Count)
                End Try
            End SyncLock
        End If
    End Sub

    Public Shared Sub LogException(ByVal ex As Exception)

        Dim LogMsg As String = ex.Message & ControlChars.NewLine & ex.StackTrace & ControlChars.NewLine
        LogText(LogMsg)
    End Sub
End Class
