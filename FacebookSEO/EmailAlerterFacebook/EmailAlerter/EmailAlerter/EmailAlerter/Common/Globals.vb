
Public Class Globals
    Public Shared Function GetTotalRecords(ByRef dr As IDataReader) As Integer

        Dim total As Integer = 0

        If dr.Read Then
            Try
                total = Convert.ToInt32(dr("TotalRecords"))
            Catch ex As Exception
                total = - 1
            End Try
        End If

        Return total
    End Function
End Class
