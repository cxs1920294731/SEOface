Namespace IAnalysis
    Public Interface AutoInterface
        ''' <summary>
        ''' Start()方法接口
        ''' </summary>
        ''' <param name="IssueID"></param>
        ''' <param name="siteId"></param>
        ''' <param name="planType"></param>
        ''' <param name="splitContactCount"></param>
        ''' <param name="spreadLogin"></param>
        ''' <param name="appId"></param>
        ''' <param name="url"></param>
        ''' <param name="nowTime"></param>
        ''' <remarks></remarks>
        Sub IStart(ByVal IssueID As Integer, ByVal siteId As Integer, ByVal planType As String, _
                         ByVal splitContactCount As Integer, ByVal spreadLogin As String, _
                         ByVal appId As String, ByVal url As String, ByVal nowTime As DateTime)
    End Interface
End Namespace
