Imports EmailAlerter.Service1
Imports EmailAlerter.GroupBuyer2

Public Class MyComparer
    Implements IComparer(Of EmailCampaignElementDeal)
    Public Favorite As String
    Public name As String
    Public Function Compare(ByVal x As EmailCampaignElementDeal, ByVal y As EmailCampaignElementDeal) As Integer Implements System.Collections.Generic.IComparer(Of EmailCampaignElementDeal).Compare
        Try
            If x.DealCategory.Length > 0 And y.DealCategory.Length > 0 Then
                Dim xCaregory() = x.DealCategory.Split("-")
                Dim yCaregory() = y.DealCategory.Split("-")
                If Array.IndexOf(xCaregory, Favorite) <> -1 And Array.IndexOf(yCaregory, Favorite) = -1 Then
                    Return -1
                ElseIf Array.IndexOf(xCaregory, Favorite) = -1 And Array.IndexOf(yCaregory, Favorite) <> -1 Then
                    Return 1
                ElseIf Array.IndexOf(xCaregory, Favorite) <> -1 And Array.IndexOf(yCaregory, Favorite) <> -1 Then
                    If x.Index > y.Index Then
                        Return 1
                    Else : Return -1
                    End If
                ElseIf Array.IndexOf(xCaregory, Favorite) = -1 And Array.IndexOf(yCaregory, Favorite) = -1 Then
                    If x.Index > y.Index Then
                        Return 1
                    Else : Return -1
                    End If
                End If
            ElseIf x.DealCategory.Length = 0 And y.DealCategory.Length > 0 Then
                Dim yCaregory() = y.DealCategory.Split("-")
                If Array.IndexOf(yCaregory, Favorite) <> -1 Then
                    Return 1
                Else
                    If x.Index > y.Index Then
                        Return 1
                    Else : Return -1
                    End If
                End If
            ElseIf x.DealCategory.Length > 0 And y.DealCategory.Length = 0 Then
                Dim xCaregory() = x.DealCategory.Split("-")
                If Array.IndexOf(xCaregory, Favorite) <> -1 Then
                    Return -1
                Else
                    If x.Index > y.Index Then
                        Return 1
                    Else : Return -1
                    End If
                End If
            End If
        Catch ex As Exception
            System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & "1.log", ex.Message & "-----" & DateTime.Now & Environment.NewLine())
        End Try



        'If y.DealCategory.Contains(Favorite) And (Not x.DealCategory.Contains(Favorite)) Then
        '    Return 1
        'ElseIf x.DealCategory.Contains(Favorite) And (Not y.DealCategory.Contains(Favorite)) Then
        '    Return -1
        'ElseIf (Not x.DealCategory.Contains(Favorite)) And (Not y.DealCategory.Contains(Favorite)) Then
        '    Return -1
        'Else
        '    Return 0
        'End If

    End Function
    'Sub Log(ByVal Ex As Exception)
    '    Try
    '        LogText(Now & ", " & Ex.Message & Environment.NewLine() & Ex.StackTrace & Environment.NewLine())
    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub

    'Sub LogText(ByVal Ex As String)
    '    Try
    '        System.IO.File.AppendAllText(System.Reflection.Assembly.GetExecutingAssembly.Location & Now.Year & "-" & Now.Month & ".log", Now & ", " & Ex & Environment.NewLine())
    '    Catch ex1 As Exception
    '        'ignore
    '    End Try
    'End Sub
End Class
