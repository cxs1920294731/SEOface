Imports System.Data.SqlClient

Public Class ProductCategoryDAL
    Private Shared efContext As New EmailAlerterEntities()
    Public Shared Function GetRelatedCategory(ByVal productId As Integer) As List(Of Integer)

        Dim categoryIds As New List(Of Integer)
        Try

            Dim sql As String = "select CategoryID from ProductCategory pc where pc.ProductID=@productId"
            Dim sqlParam As SqlParameter()
            sqlParam = New SqlParameter() {
                                           New SqlParameter With {.ParameterName = "ProductID", .Value = productId.ToString()}
            }

            categoryIds = efContext.ExecuteStoreQuery(Of Integer)(sql, sqlParam).ToList()
            'productList = efContext.ExecuteStoreQuery(Of AutoProduct)(sql, sqlParam).ToList()

        Catch ex As Exception
            Common.LogText("GetRelatedCategory()-->" & ex.ToString)
        End Try

        Return categoryIds
    End Function

End Class
