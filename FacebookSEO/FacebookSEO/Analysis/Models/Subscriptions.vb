Public Class Subscriptions

#Region "Private Members"

    Private _siteId As Integer
    Private _planType As String
    Private _url As String
    Private _startAt As DateTime
    Private _intervalDay As Double
    Private _weekDays As String
    Private _senderName As String
    Private _senderEmail As String
    Private _spreadLoginEmail As String
    Private _appId As String
    Private _templateId As Integer
    Private _multiTemplate As String '多模板设计
    Private _splitContactList As Integer
    Private _categories As String
    Private _siteName As String
    Private _dllType As String
    Private _siteUrl As String
    Private _volumn As Integer
    '新功能属性
    Private _contactList As String
    Private _excludeList As String
    Private _campaignStatus As String
    Private _scheduleTimeInteval As Double
    Private _listType As String
    Private _searchAPIType As String
    Private _searchStartDayInterval As Integer
    Private _searchEndDayInterval As Integer
    Private _urlSpecialCode As String
    Private _limitQuantity As Long
    Private _timeLimit As String
    Private _sellerEmail As String
    Private _logoUrl As String
    Private _displayCount As Integer
    Private _bannerFromUrl As String
    Private _bannerRegex As String
    Private _bannerIndex As Integer
    'NFC功能属性
    Private _triggerForNFC As Integer
    Private _subjectForNFC As String

#End Region

#Region "Public Properties"

    Public Property LogoUrl() As String
        Get
            Return _logoUrl
        End Get
        Set(ByVal value As String)
            _logoUrl = value
        End Set
    End Property

    Public Property TimeLimit() As String
        Get
            Return _timeLimit
        End Get
        Set(ByVal value As String)
            _timeLimit = value
        End Set
    End Property

    Public Property SellerEmail() As String
        Get
            Return _sellerEmail
        End Get
        Set(ByVal value As String)
            _sellerEmail = value
        End Set
    End Property

    Public Property LimitQuantity() As Long
        Get
            Return _limitQuantity
        End Get
        Set(ByVal value As Long)
            _limitQuantity = value
        End Set
    End Property

    Public Property ContactList() As String
        Get
            Return _contactList
        End Get
        Set(ByVal value As String)
            _contactList = value
        End Set
    End Property

    Public Property ExcludeList() As String
        Get
            Return _excludeList
        End Get
        Set(ByVal value As String)
            _excludeList = value
        End Set
    End Property

    Public Property CampaignStatus() As String
        Get
            Return _campaignStatus
        End Get
        Set(ByVal value As String)
            _campaignStatus = value
        End Set
    End Property

    Public Property ScheduleTimeInteval() As Double
        Get
            Return _scheduleTimeInteval
        End Get
        Set(ByVal value As Double)
            _scheduleTimeInteval = value
        End Set
    End Property

    Public Property ListType() As String
        Get
            Return _listType
        End Get
        Set(ByVal value As String)
            _listType = value
        End Set
    End Property

    Public Property SearchAPIType() As String
        Get
            Return _searchAPIType
        End Get
        Set(ByVal value As String)
            _searchAPIType = value
        End Set
    End Property

    Public Property SearchStartDayInterval() As Integer
        Get
            Return _searchStartDayInterval
        End Get
        Set(ByVal value As Integer)
            _searchStartDayInterval = value
        End Set
    End Property

    Public Property SearchEndDayInterval() As Integer
        Get
            Return _searchEndDayInterval
        End Get
        Set(ByVal value As Integer)
            _searchEndDayInterval = value
        End Set
    End Property

    Public Property UrlSpecialCode() As String
        Get
            Return _urlSpecialCode
        End Get
        Set(ByVal value As String)
            _urlSpecialCode = value
        End Set
    End Property

    Public Property SiteId() As Integer
        Get
            Return _siteId
        End Get
        Set(ByVal value As Integer)
            _siteId = value
        End Set
    End Property

    Public Property PlanType() As String
        Get
            Return _planType
        End Get
        Set(ByVal value As String)
            _planType = value
        End Set
    End Property

    Public Property Url() As String
        Get
            Return _url
        End Get
        Set(ByVal value As String)
            If String.IsNullOrEmpty(value) Then
                _url = ""
            Else
                _url = value
            End If
        End Set
    End Property

    Public Property StartAt() As DateTime
        Get
            Return _startAt
        End Get
        Set(ByVal value As DateTime)
            If String.IsNullOrEmpty(value) Then
                _startAt = DateTime.Parse("1900-01-01 00:01:00")
            Else
                _startAt = value
            End If
        End Set
    End Property

    Public Property IntervalDay() As Double
        Get
            Return _intervalDay
        End Get
        Set(ByVal value As Double)
            _intervalDay = value
        End Set
    End Property

    Public Property WeekDays() As String
        Get
            Return _weekDays
        End Get
        Set(ByVal value As String)
            _weekDays = value
        End Set
    End Property

    Public Property SenderName() As String
        Get
            Return _senderName
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _senderName = ""
            Else
                _senderName = value
            End If
        End Set
    End Property

    Public Property SenderEmail() As String
        Get
            Return _senderEmail
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _senderEmail = SpreadLoginEmail
            Else
                _senderEmail = value
            End If
        End Set
    End Property

    Public Property SpreadLoginEmail() As String
        Get
            Return _spreadLoginEmail
        End Get
        Set(ByVal value As String)
            _spreadLoginEmail = value
        End Set
    End Property

    Public Property AppId() As String
        Get
            Return _appId
        End Get
        Set(ByVal value As String)
            _appId = value
        End Set
    End Property

    Public Property TemplateId() As Integer
        Get
            Return _templateId
        End Get
        Set(ByVal value As Integer)
            _templateId = value
        End Set
    End Property

    Public Property MultiTemplate() As String
        Get
            Return _multiTemplate
        End Get
        Set(ByVal value As String)
            _multiTemplate = value
        End Set
    End Property

    Public Property SplitContactList() As Integer
        Get
            Return _splitContactList
        End Get
        Set(ByVal value As Integer)
            _splitContactList = value
        End Set
    End Property

    Public Property Categories() As String
        Get
            Return _categories
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _categories = ""
            Else
                _categories = value
            End If
        End Set
    End Property

    Public Property SiteName() As String
        Get
            Return _siteName
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _siteName = ""
            Else
                _siteName = value
            End If
        End Set
    End Property

    Public Property DllType() As String
        Get
            Return _dllType
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _dllType = ""
            Else
                _dllType = value
            End If
        End Set
    End Property

    Public Property SiteUrl As String
        Get
            Return _siteUrl
        End Get
        Set(ByVal value As String)
            If (String.IsNullOrEmpty(value)) Then
                _siteUrl = ""
            Else
                _siteUrl = value
            End If
        End Set
    End Property

    Public Property Volumn As Integer
        Get
            Return _volumn
        End Get
        Set(ByVal value As Integer)
            _volumn = value
        End Set
    End Property

    Public Property TriggerForNFC() As Integer
        Get
            Return _triggerForNFC
        End Get
        Set(ByVal value As Integer)
            _triggerForNFC = value
        End Set
    End Property

    Public Property SubjectForNFC() As String
        Get
            Return _subjectForNFC
        End Get
        Set(ByVal value As String)
            _subjectForNFC = value
        End Set
    End Property

    Public Property DisplayCount() As Integer
        Get
            Return _displayCount
        End Get
        Set(ByVal value As Integer)
            _displayCount = value
        End Set
    End Property

    Public Property BannerFromUrl() As String
        Get
            Return _bannerFromUrl
        End Get
        Set(ByVal value As String)
            _bannerFromUrl = value
        End Set
    End Property

    Public Property BannerRegex() As String
        Get
            Return _bannerRegex
        End Get
        Set(ByVal value As String)
            _bannerRegex = value
        End Set
    End Property

    Public Property BannerIndex() As Integer
        Get
            Return _bannerIndex
        End Get
        Set(ByVal value As Integer)
            _bannerIndex = value
        End Set
    End Property
#End Region
End Class
