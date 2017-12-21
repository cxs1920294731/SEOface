
Imports Analysis

<Serializable()> _
Public Class QueryParams
    Public Sub New()
    End Sub

    Public Property Categories As List(Of Category)
    Public Property Template As String
End Class

<Serializable()> _
<Flags()> _
Public Enum TaobaoSortedType
    hotsell_desc '
    newOn_desc 'new arrival
    price_asc
    hotkeep_desc '
    coefp_desc '人气
End Enum

<Serializable()> _
<Flags()> _
Public Enum TmallSortedType
    hotsell_desc '
    newOn_desc 'new arrival
    price
    hotkeep_desc '
    koubei '人气
End Enum
