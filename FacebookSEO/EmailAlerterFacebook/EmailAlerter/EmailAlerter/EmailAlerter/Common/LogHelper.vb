Imports System
Imports log4net


''' <summary>
''' 日志记录辅助类
''' </summary>
''' <remarks></remarks>
    Public Class LogHelper
    'Private Shared ReadOnly ErrorLog As ILog = LogManager.GetLogger("logerror")
    'Private Shared ReadOnly InfoLog As ILog = LogManager.GetLogger("loginfo")

    ''' <summary>
    ''' 将指定信息写入日志
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Shared Sub WriteLog(ByVal message As String)

        'If InfoLog.IsInfoEnabled Then
        '    InfoLog.Info(message:=message)
        'End If
    End Sub

    ''' <summary>
    ''' 将指定信息和异常写入日志
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="exc"></param>
    ''' <remarks></remarks>
    Public Shared Sub WriteLog(ByVal message As String, ByVal exc As Exception)

        'If ErrorLog.IsErrorEnabled Then
        '    ErrorLog.Error(message, exc)
        'End If
    End Sub
End Class
