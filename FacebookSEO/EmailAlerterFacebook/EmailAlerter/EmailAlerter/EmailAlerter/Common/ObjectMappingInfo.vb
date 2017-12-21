Imports System.Collections.Generic
Imports System.Reflection


<Serializable()>
Public Class ObjectMappingInfo

#Region "Private Members"

    Private _CacheByProperty As String
    Private _CacheTimeOutMultiplier As Integer
    Private _ColumnNames As Dictionary(Of String, String)
    Private _Properties As Dictionary(Of String, PropertyInfo)
    Private _ObjectType As String
    Private _TableName As String
    Private _PrimaryKey As String

    Private Const RootCacheKey As String = "ObjectCache_"

#End Region

#Region "Constructors"

    '''-----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new ObjectMappingInfo Object
    ''' </summary>
    ''' <history>
    '''     [cnurse]	01/12/2008	created
    ''' </history>
    '''-----------------------------------------------------------------------------
    Public Sub New()
        _Properties = New Dictionary(Of String, PropertyInfo)
        _ColumnNames = New Dictionary(Of String, String)
    End Sub

#End Region

#Region "Public Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CacheKey gets the root value of the key used to identify the cached collection 
    ''' in the ASP.NET Cache.
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CacheKey() As String
        Get
            Dim _CacheKey As String = RootCacheKey + TableName + "_"
            If Not String.IsNullOrEmpty(CacheByProperty) Then
                _CacheKey &= CacheByProperty + "_"
            End If
            Return _CacheKey
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CacheByProperty gets and sets the property that is used to cache collections
    ''' of the object.  For example: Modules are cached by the "TabId" proeprty.  Tabs 
    ''' are cached by the PortalId property.
    ''' </summary>
    ''' <remarks>If empty, a collection of all the instances of the object is cached.</remarks>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property CacheByProperty() As String
        Get
            Return _CacheByProperty
        End Get
        Set(ByVal value As String)
            _CacheByProperty = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CacheTimeOutMultiplier gets and sets the multiplier used to determine how long
    ''' the cached collection should be cached.  It is multiplied by the Performance
    ''' Setting - which in turn can be modified by the Host Account.
    ''' </summary>
    ''' <remarks>Defaults to 20.</remarks>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property CacheTimeOutMultiplier() As Integer
        Get
            Return _CacheTimeOutMultiplier
        End Get
        Set(ByVal value As Integer)
            _CacheTimeOutMultiplier = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ColumnNames gets a dictionary of Database Column Names for the Object
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/02/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ColumnNames() As Dictionary(Of String, String)
        Get
            Return _ColumnNames
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ObjectType gets and sets the type of the object
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ObjectType() As String
        Get
            Return _ObjectType
        End Get
        Set(ByVal value As String)
            _ObjectType = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' PrimaryKey gets and sets the property of the object that corresponds to the
    ''' primary key in the database
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property PrimaryKey() As String
        Get
            Return _PrimaryKey
        End Get
        Set(ByVal value As String)
            _PrimaryKey = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Properties gets a dictionary of Properties for the Object
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Properties() As Dictionary(Of String, PropertyInfo)
        Get
            Return _Properties
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' TableName gets and sets the name of the database table that is used to
    ''' persist the object.
    ''' </summary>
    ''' <history>
    ''' 	[cnurse]	12/01/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property TableName() As String
        Get
            Return _TableName
        End Get
        Set(ByVal value As String)
            _TableName = value
        End Set
    End Property

#End Region
End Class

