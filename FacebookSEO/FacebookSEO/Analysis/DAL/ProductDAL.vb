Imports System.Data.SqlClient

Public Class ProductDAL
    Private Shared efContext As New EmailAlerterEntities()

    Public Shared Function GetProducts(ByVal issueid As Integer, ByVal categoryName As String, ByVal siteid As Integer) As List(Of Product)
        Dim productList As List(Of Product) = New List(Of Product)
        Try
            categoryName = categoryName.Trim
            Dim sql As String = "select * from Products p " &
                                                "inner join Products_Issue pi on p.ProdouctID=pi.ProductId " &
                                                "inner join ProductCategory pc on pc.ProductID=p.ProdouctID " &
                                                "inner join Categories c on pc.CategoryID=c.CategoryID " &
                                                "where pi.IssueID=@IssueID and c.Category=@Category and c.SiteID=@SiteID order by p.LastUpdate"
            Dim sqlParam As SqlParameter()
            sqlParam = New SqlParameter() {New SqlParameter With {.ParameterName = "IssueID", .Value = issueid.ToString()},
                                           New SqlParameter With {.ParameterName = "Category", .Value = categoryName.Trim()},
                                           New SqlParameter With {.ParameterName = "SiteID", .Value = siteid.ToString()}}
            productList = efContext.ExecuteStoreQuery(Of Product)(sql, sqlParam).ToList()

        Catch ex As Exception
            Common.LogText("GetProducts()-->" & ex.ToString)
        End Try

        Return productList
    End Function


End Class
