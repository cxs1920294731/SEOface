
Public Module Exceptions
    Public Sub LogException(ByVal exc As Exception)
        LogHelper.WriteLog(exc.Message, exc)
    End Sub
End Module

