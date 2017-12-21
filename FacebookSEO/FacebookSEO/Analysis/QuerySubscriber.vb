Imports Newtonsoft.Json

<Serializable()> _
Public Class QuerySubscriber
    ''' <summary>
    ''' 查询指定国家或地区的收件人
    ''' </summary>
    ''' <remarks></remarks>
    Private _countrylist As String()
    Public Property CountryList() As String()
        Get
            Return _countrylist
        End Get
        Set(ByVal value As String())
            _countrylist = value
        End Set
    End Property

    ''' <summary>
    ''' 查询指定性别的收件人
    ''' </summary>
    ''' <remarks></remarks>
    Private _gender As Gender
    Public Property Gender() As Gender
        Get
            Return _gender
        End Get
        Set(ByVal value As Gender)
            _gender = value
        End Set
    End Property

    ''' <summary>
    ''' 查询指定年龄范围的收件人
    ''' </summary>
    ''' <remarks></remarks>
    Private _agerange As ValueRange(Of Integer)
    Public Property AgeRange() As ValueRange(Of Integer)
        Get
            Return _agerange
        End Get
        Set(ByVal value As ValueRange(Of Integer))
            _agerange = value
        End Set
    End Property

    ''' <summary>
    ''' 选择收件人的策略,如收件人之前是否打开过类似邮件
    ''' </summary>
    ''' <remarks></remarks>
    Private _strategy As ChooseStrategy
    Public Property Strategy() As ChooseStrategy
        Get
            Return _strategy
        End Get
        Set(ByVal value As ChooseStrategy)
            _strategy = value
        End Set
    End Property



    Private _favorite As String
    Public Property Favorite() As String
        Get
            Return _favorite
        End Get
        Set(ByVal value As String)
            _favorite = value
        End Set
    End Property

    'Gary start
    '2013-06-09,更新可以根据时间段进行搜索联系人,时间格式为DateTime的String：yyyy-mm-dd 或者 yyyy-mm-dd hh:mm:ss,时间段分以下几种格式：
    '1，startDate有赋值，endDate有赋值，代表startDate到endDate
    '2,startDate没赋值，endDate有赋值，代表开账号日期到endDate
    '3,startDate有赋值，endDate没赋值，代表startDate到当前日期
    '4，startDate和endDate都没赋值，代表开账号日期到当前日期

    Private _StartDate As String
    Public Property StartDate() As String
        Get
            Return _StartDate
        End Get
        Set(ByVal value As String)
            _StartDate = value
        End Set
    End Property

    Private _EndDate As String
    Public Property EndDate() As String
        Get
            Return _EndDate
        End Get
        Set(ByVal value As String)
            _EndDate = value
        End Set
    End Property
    'Gary end

    'Private _contactLists As String
    'Public Property ContactLists() As String
    '    Get
    '        Return _contactLists
    '    End Get
    '    Set(ByVal value As String)
    '        _contactLists = value
    '    End Set
    'End Property

    ''' <summary>
    ''' [序列化]将查询对象转换为json字符串，作为SpreadAPI的参数传递
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToJsonString() As String
        Return JsonConvert.SerializeObject(Me)
    End Function

    ''' <summary>
    ''' [反序列化]将传递过来的json字符串反序列化为查询对象
    ''' </summary>
    ''' <param name="JsonString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FromJsonString(ByVal JsonString As String) As QuerySubscriber
        Return TryCast(JsonConvert.DeserializeObject(JsonString, GetType(QuerySubscriber)), QuerySubscriber)
    End Function
End Class

<Serializable()> _
Public Class ValueRange(Of T)

    Private _minvalue As T
    Private _maxvalue As T

    Public Property MinValue() As T
        Get
            Return _minvalue
        End Get
        Set(ByVal value As T)
            _minvalue = value
        End Set
    End Property

    Public Property MaxValue() As T
        Get
            Return _maxvalue
        End Get
        Set(ByVal value As T)
            _maxvalue = value
        End Set
    End Property

    Sub New()

    End Sub

End Class

<Serializable()> _
Public Enum Gender
    Male = 1
    Female = 2
    All = 3
End Enum

<Serializable()> _
<Flags()> _
Public Enum ChooseStrategy
    Open = 1 'All open
    Random = 2
    OpenRandom = 3
    Favorite = 4 'Category
    OpenExcludeCategory = 5 'Open Not Category
    NotOpen = 6 'Not Open
End Enum

