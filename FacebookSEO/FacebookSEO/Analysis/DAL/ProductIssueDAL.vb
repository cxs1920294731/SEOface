Public Class ProductIssueDAL

    Private Shared efContext As New EmailAlerterEntities()

    Public Shared Function InsertProductsIssue(ByVal siteId As Integer, ByVal issueId As Integer, ByVal sectionId As String, ByVal listProductId As List(Of Integer),
                                ByVal iProIssueCount As Integer, ByVal categoryId As Integer) As Integer
        Dim addedCount As Integer = 0
        Try

            '这个issue下这个category的记录
            Dim existed = (From pro In efContext.Products_Issue
                           Where pro.IssueID = issueId AndAlso pro.SiteId = siteId AndAlso pro.SectionID = sectionId 'AndAlso pro.category =categoryId 
                           Select pro.ProductId).ToList()


            For Each li In listProductId

                If (existed.Contains(li)) Then
                    Common.LogText(String.Format(" duplicate product in the same issue ,jump ---where issueid={0},siteId={1},productid={2},categoryid={3}", issueId, siteId, li, categoryId))
                    Continue For
                End If


                If addedCount < iProIssueCount Then
                    Try
                        Dim pIssue As New Products_Issue
                        pIssue.ProductId = li
                        pIssue.SiteId = siteId
                        pIssue.IssueID = issueId
                        pIssue.SectionID = sectionId
                        pIssue.CategoryId = categoryId
                        efContext.Products_Issue.AddObject(pIssue)
                        efContext.SaveChanges()
                        addedCount = addedCount + 1
                    Catch ex As Exception
                        Common.LogText("InsertProductsIssue()-->" & ex.ToString)
                        Continue For
                    End Try

                End If


                If (addedCount >= iProIssueCount) Then
                    Exit For
                End If

            Next

        Catch ex As Exception
            Common.LogText("InsertProductsIssue()-->" & ex.ToString)

        End Try

        Return addedCount

    End Function

End Class
