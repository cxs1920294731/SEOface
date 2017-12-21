Public Interface IHydratable
    Property KeyID() As Integer

    Sub Fill(ByVal dr As IDataReader)
End Interface
