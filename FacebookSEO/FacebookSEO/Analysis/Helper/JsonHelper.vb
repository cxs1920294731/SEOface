
Imports Newtonsoft.Json
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class JsonHelper
    ''' <summary>
    ''' Convert An Object to Json String
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns></returns>
    Public Shared Function ToJsonString(data As Object) As String
        Return JsonConvert.SerializeObject(data)
    End Function
    ''' <summary>
    ''' Convert a Json String To An Object
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="json"></param>
    ''' <returns></returns>
    Public Shared Function ParseFromJson(Of T)(json As String) As T
        Return JsonConvert.DeserializeObject(Of T)(json)
    End Function
End Class