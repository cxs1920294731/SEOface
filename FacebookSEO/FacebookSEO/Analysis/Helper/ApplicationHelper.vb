
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class ApplicationHelper
    Public Shared Function Run(file As String, ParamArray args As Object()) As String
        Dim process As New Process()
        process.StartInfo.FileName = file
        process.StartInfo.UseShellExecute = False
        process.StartInfo.CreateNoWindow = True
        process.StartInfo.RedirectStandardInput = True
        process.StartInfo.Arguments = String.Join(" ", args)
        process.StartInfo.RedirectStandardOutput = True
        '重定向输出流 
        process.StartInfo.RedirectStandardError = True
        '重定向错误流 
        process.Start()
        Dim output As String = process.StandardOutput.ReadToEnd()
        process.WaitForExit()
        process.Close()
        Return output
    End Function
End Class