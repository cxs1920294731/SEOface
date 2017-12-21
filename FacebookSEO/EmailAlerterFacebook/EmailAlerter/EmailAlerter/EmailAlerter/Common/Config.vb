Imports System.Web.Configuration


Public NotInheritable Class Config
    Private Const FORMAT_APPKEYNOTEXISTS As String = "web.config,[{0}] value not exists or emtpy"
    Private Shared ReadOnly configKeyLostList As List(Of String) = New List(Of String)()

    ''' <summary>
    ''' Get the specified connection string
    ''' </summary>
    ''' <param name="name">name of the connection string to return</param>
    ''' <returns>The Connection string</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Robin] 2012-4-9 created
    ''' </history>
    Public Shared Function GetConnectionString(ByVal name As String) As String

        Dim connectionString As String = ""

        If name <> "" Then
            connectionString = WebConfigurationManager.ConnectionStrings(name).ConnectionString
        End If

        If connectionString = "" Then
            If name <> "" Then
                connectionString = Config.GetSetting(name)
            End If
        End If

        Return connectionString
    End Function

    Public Shared Function GetSetting(ByVal setting As String) As String
        Return WebConfigurationManager.AppSettings(setting)
    End Function

    'Public Shared Function GetConfigValueLogWhenNotExists(configKey As String) As String
    '    Dim valueInConfig As String = Config.GetSetting(configKey)
    '    If String.IsNullOrEmpty(valueInConfig) Then
    '        If Not configKeyLostList.Contains(configKey) Then
    '            Dim message As String = String.Format(FORMAT_APPKEYNOTEXISTS, configKey)
    '            LogHelper.WriteLog(message, New Exception(message))
    '            configKeyLostList.Add(configKey)
    '        End If
    '    End If
    '    Return valueInConfig
    'End Function
End Class

