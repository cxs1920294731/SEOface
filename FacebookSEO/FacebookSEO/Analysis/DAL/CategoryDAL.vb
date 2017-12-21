Public Class CategoryDAL


    Private Shared efContext As New EmailAlerterEntities()


    ''' <summary>
    ''' 根据某个分类的名字获取某个账户下面的分类对象
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCategory(ByVal siteId As Integer, ByVal categoryName As String) As Category

        Dim category As New Category
        Try
            category = efContext.Categories.Where(Function(c) c.SiteID = siteId AndAlso c.Category1 = categoryName).SingleOrDefault()
        Catch ex As Exception
            Common.LogText(String.Format("GetCategory()--> siteid:{0}, categoryName={1} -->{2} ", siteId, categoryName, ex.StackTrace))
        End Try

        Return category


    End Function




End Class
