Imports Analysis
Imports Enumerable = System.Linq.Enumerable

Public Class SEmailEfHelper
#Region "automationSites表"
    Shared efEntity As New EmailAlerterEntities()

    ''' <summary>
    ''' 判断指定spreadAccount是否存在于autosites表中，如存在返回默认siteid，否则返回-1
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SpreadAccountExist(ByVal spreadAccount As String) As Integer
        spreadAccount = spreadAccount.Trim
        Dim existSites As List(Of AutomationSite) = (From autosite In efEntity.AutomationSites
                                                      Where autosite.SpreadLogin = spreadAccount).ToList()
        If existSites.Count > 0 Then
            Return existSites(0).siteid
        Else
            Return -1
        End If
    End Function

    ''' <summary>
    ''' 判断指定shopName-spreadAccount 键值对 是否存在于autosites表中，如存在返回siteid（同一个shopName只有一条），否则返回-1
    ''' </summary>
    ''' <param name="shopName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function shopNameExist(ByVal shopName As String, ByVal spreadAccount As String) As Integer
        shopName = shopName.Trim
        spreadAccount = spreadAccount.Trim
        Dim existSites As List(Of AutomationSite) = (From autosite In efEntity.AutomationSites
                                                      Where autosite.SiteName = shopName And autosite.SpreadLogin = spreadAccount).ToList()
        If existSites.Count > 0 Then
            Return existSites(0).siteid
        Else
            Return -1
        End If
    End Function


    ''' <summary>
    ''' 使用传入的参数spread account 及apikey不做任何判断创建一条autosites记录，返回siteID
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="spreadApiKey"></param>
    ''' <returns>siteID</returns>
    ''' <remarks></remarks>
    Public Shared Function AddNewSpreadAccount(ByVal spreadAccount As String, ByVal spreadApiKey As String) As Integer
        spreadAccount = spreadAccount.Trim
        spreadApiKey = spreadApiKey.Trim
        Dim newItem As New AutomationSite
        newItem.SpreadLogin = spreadAccount
        newItem.AppID = spreadApiKey
        newItem.CreateTime = DateTime.Now.ToString()
        efEntity.AutomationSites.AddObject(newItem)
        efEntity.SaveChanges()
        Return newItem.siteid
    End Function

    ''' <summary>
    ''' 添加一条新的autosites记录，其中appid根据spreadAccount获取，dlltype根据siteURL确定；返回值：siteid
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="siteName"></param>
    ''' <param name="siteUrl"></param>
    ''' <param name="logoImgUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function addNewSite(ByVal spreadAccount As String, ByVal siteName As String, ByVal siteUrl As String, ByVal logoImgUrl As String, ByVal dlltypeFlag As Boolean) As Integer
        Dim newSiteItem As New AutomationSite
        newSiteItem.SiteName = siteName.Trim
        newSiteItem.SpreadLogin = spreadAccount
        newSiteItem.AppID = getApiKey(spreadAccount)
        newSiteItem.DllType = defineDlltype(siteUrl, dlltypeFlag)

        newSiteItem.SiteUrl = siteUrl
        newSiteItem.Description = "From SmartMailWeb User Interface"
        newSiteItem.LogoImgUrl = logoImgUrl
        newSiteItem.CreateTime = Now
        efEntity.AutomationSites.AddObject(newSiteItem)
        efEntity.SaveChanges()
        Return newSiteItem.siteid
    End Function

    ''' <summary>
    ''' 更新一条autosite记录，目前仅支持更新logoImgUrl字段
    ''' </summary>
    ''' <param name="spreadaccount"></param>
    ''' <param name="sitename"></param>
    ''' <param name="logoImgUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Sub updateAutoSite(ByVal spreadaccount As String, ByVal sitename As String, ByVal logoImgUrl As String)
        Dim siteItem As AutomationSite = Enumerable.FirstOrDefault(efEntity.AutomationSites, Function(m) m.SpreadLogin = spreadaccount AndAlso m.SiteName = sitename)
        'siteItem.SiteUrl = siteUrl.Trim
        'siteItem.DllType = defineDlltype(siteUrl)
        siteItem.LogoImgUrl = logoImgUrl
        efEntity.SaveChanges()
    End Sub

    ''' <summary>
    ''' 通过spreadAccount从autosite表中获取apikey
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getApiKey(ByVal spreadAccount As String) As String
        If String.IsNullOrEmpty(spreadAccount) Then
            Return ""
        End If
        spreadAccount = spreadAccount.Trim
        Dim apiKey As String = (From autosite In efEntity.AutomationSites
                                Where autosite.SpreadLogin = spreadAccount
                                Select autosite.AppID).FirstOrDefault()
        Return apiKey
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="siteUrl"></param>
    ''' <param name="dlltypeflag">true:根据店铺shopurl确定类型，false：全部判断为indepwebsite</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function defineDlltype(ByVal siteUrl As String, ByVal dlltypeflag As Boolean) As String
        Dim dllType As String
        If (dlltypeflag) Then
            If (siteUrl.ToLower.Contains("taobao")) Then
                dllType = "taobao"
            ElseIf (siteUrl.ToLower.Contains("tmall")) Then
                dllType = "tmall"
            ElseIf (siteUrl.ToLower.Contains("aliexpress")) Then
                dllType = "aliexpress"
            ElseIf (siteUrl.ToLower.Contains("ebay")) Then
                dllType = "ebay"
            Else
                dllType = "indepwebsite"
            End If
        Else
            dllType = "indepwebsite"
        End If

        Return dllType
    End Function

    ''' <summary>
    ''' indepwebsite独立站流程且是淘宝天猫店铺的域名为https://,添加在url及图片src前面
    ''' </summary>
    ''' <param name="siteUrl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDomain(ByVal siteUrl As String, ByVal dlltype As String) As String
        If (dlltype.Trim = "indepwebsite" AndAlso (siteUrl.ToLower.Contains("taobao") OrElse siteUrl.ToLower().Contains("tmall"))) Then
            Return "https://" '淘宝天猫店铺的域名为https://,添加在url及图片src前面
        Else
            Return siteUrl
        End If
    End Function

    Public Shared Function InsertOrUpdateAutoSite(ByVal spreadAccount As String, ByVal spreadApiKey As String) As Integer
        spreadAccount = spreadAccount.Trim
        spreadApiKey = spreadApiKey.Trim
        If (SpreadAccountExist(spreadAccount) > 0) Then
            Return SpreadAccountExist(spreadAccount)
        Else
            Dim newItem As New AutomationSite
            newItem.SpreadLogin = spreadAccount
            newItem.AppID = spreadApiKey
            efEntity.AutomationSites.AddObject(newItem)
            efEntity.SaveChanges()
            Return newItem.siteid
        End If
    End Function


#End Region

#Region "automationPlan表"
    Public Shared Function GetAutoPlanInfos(ByVal spreadAccount As String, ByVal keyWord As String, ByVal status As String) As List(Of AutoPlanInfo)
        spreadAccount = spreadAccount.Trim
        Dim listAutoPlanInfo As New List(Of AutoPlanInfo)
        listAutoPlanInfo = efEntity.GetAutoPlanInfos(spreadAccount, keyWord.Trim, status.Trim).ToList()
        Return listAutoPlanInfo
    End Function

    Public Shared Function GetAutoPlanInfo(ByVal siteId As Integer, ByVal planType As String) As AutoPlanInfo
        planType = planType.Trim

        Dim autoPlanInfo As AutoPlanInfo
        autoPlanInfo = efEntity.GetOneAutoPlanInfo(siteId, planType).FirstOrDefault()
        Return autoPlanInfo
    End Function


    Public Shared Function GetAutoMationSite(ByVal spreadLogin As String, ByVal siteId As Integer) As AutomationSite

        Dim autoSite As AutomationSite = (From ats As AutomationSite In efEntity.AutomationSites Select ats Where ats.siteid = siteId AndAlso ats.SpreadLogin = spreadLogin).FirstOrDefault()
        Return autoSite

    End Function



    Public Shared Function UpdateAutoMationSite(ByVal spreadLogin As String, ByVal siteId As Integer, ByVal NewSite As AutomationSite) As Boolean
        Dim isDone As Boolean = False
        Try
            Dim autoSite As AutomationSite = (From ats As AutomationSite In efEntity.AutomationSites Select ats Where ats.siteid = siteId AndAlso ats.SpreadLogin = spreadLogin).FirstOrDefault()
            autoSite.AppID = NewSite.AppID
            autoSite.CreateTime = NewSite.CreateTime
            autoSite.Description = NewSite.Description
            autoSite.DllType = NewSite.DllType
            autoSite.LogoImgUrl = NewSite.LogoImgUrl
            autoSite.siteid = NewSite.siteid
            autoSite.SiteName = NewSite.SiteName
            autoSite.SiteUrl = NewSite.SiteUrl
            autoSite.SpreadLogin = NewSite.SpreadLogin
            autoSite.volumn = NewSite.volumn
            efEntity.SaveChanges()
            isDone = True
        Catch ex As Exception
            Common.LogText("UpdateAutoMationSite()-->" & ex.ToString)

        End Try

        Return isDone




    End Function

    Public Shared Function UpdateAutoPlan(ByVal spreadAccount As String, ByVal autoInfo As AutomationPlan, ByVal shopType As String) As Integer
        Try
            Dim apiKey As String = getApiKey(spreadAccount)
            SpreadHelper.AddSenderEmail(spreadAccount, apiKey, autoInfo.SenderEmail)
            Dim myAutoInfo As AutomationPlan = efEntity.AutomationPlans.FirstOrDefault(Function(m) m.SiteID = autoInfo.SiteID And m.PlanType = autoInfo.PlanType)

            myAutoInfo.ContactList = autoInfo.ContactList

            myAutoInfo.IntervalDay = autoInfo.IntervalDay

            myAutoInfo.SenderEmail = autoInfo.SenderEmail
            myAutoInfo.SenderName = autoInfo.SenderName
            myAutoInfo.StartAt = autoInfo.StartAt
            myAutoInfo.Status = autoInfo.Status

            If (shopType.Trim.ToLower.Contains("customized")) Then
                myAutoInfo.TemplateID = myAutoInfo.TemplateID '对于定制化客户的templateID不做改变
            Else
                myAutoInfo.TemplateID = autoInfo.TemplateID '非定制化客户支持模板的修改
            End If
            myAutoInfo.SubjectForNFC = autoInfo.SubjectForNFC
            myAutoInfo.WeekDays = autoInfo.WeekDays
            myAutoInfo.LastModifyTime = DateTime.Now

            efEntity.SaveChanges()
            Return 1
        Catch ex As Exception
            Common.LogException(ex)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' 添加一条autoPlan记录,创建成功返回1，出现异常返回-1，记录已存在返回-2。如果此记录对应的senderEmail不存在则会直接创建
    ''' </summary>
    ''' <param name="spreadAccount"></param>
    ''' <param name="apiKey"></param>
    ''' <param name="newAtuoplanItem"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function addNewAutoPlan(ByVal spreadAccount As String, ByVal apiKey As String, ByVal newAtuoplanItem As AutomationPlan) As Integer
        Try
            SpreadHelper.AddSenderEmail(spreadAccount, apiKey, newAtuoplanItem.SenderEmail)
            Dim updateautoplan As AutomationPlan = AutoPlanExist(newAtuoplanItem.SiteID, newAtuoplanItem.PlanType)
            If (updateautoplan Is Nothing) Then
                efEntity.AutomationPlans.AddObject(newAtuoplanItem)
                efEntity.SaveChanges()
            Else
                updateautoplan.StartAt = newAtuoplanItem.StartAt
                updateautoplan.WeekDays = newAtuoplanItem.WeekDays
                updateautoplan.IntervalDay = newAtuoplanItem.IntervalDay
                updateautoplan.SenderName = newAtuoplanItem.SenderName
                updateautoplan.SenderEmail = newAtuoplanItem.SenderEmail
                updateautoplan.ContactList = newAtuoplanItem.ContactList
                'updateautoplan.Status = newAtuoplanItem.Status
                updateautoplan.SplitContactCount = newAtuoplanItem.SplitContactCount
                updateautoplan.Categories = newAtuoplanItem.Categories

                updateautoplan.displayCount = newAtuoplanItem.displayCount

                updateautoplan.bannerFromUrl = newAtuoplanItem.bannerFromUrl
                updateautoplan.bannerIndex = newAtuoplanItem.bannerIndex
                updateautoplan.bannerRegex = newAtuoplanItem.bannerRegex

                updateautoplan.ContactType = newAtuoplanItem.ContactType
                updateautoplan.ScheduleTimeInteval = newAtuoplanItem.ScheduleTimeInteval
                updateautoplan.LimitQuantity = newAtuoplanItem.LimitQuantity
                updateautoplan.TimeLimit = newAtuoplanItem.TimeLimit
                updateautoplan.IsAssociateSite = newAtuoplanItem.IsAssociateSite
                updateautoplan.ExcludeList = newAtuoplanItem.ExcludeList
                updateautoplan.MultiTemplate = newAtuoplanItem.MultiTemplate
                updateautoplan.TemplateID = newAtuoplanItem.TemplateID
                updateautoplan.CampaignStatus = newAtuoplanItem.CampaignStatus
                updateautoplan.ShopType = newAtuoplanItem.ShopType
                updateautoplan.CreateTime = newAtuoplanItem.CreateTime
                updateautoplan.URL = newAtuoplanItem.URL
                updateautoplan.SellerEmail = newAtuoplanItem.SellerEmail
                updateautoplan.SubjectForNFC = newAtuoplanItem.SubjectForNFC
                updateautoplan.UrlSpecialCode = newAtuoplanItem.UrlSpecialCode
                efEntity.SaveChanges()
            End If
            Return 1
        Catch ex As Exception
            Common.LogException(ex)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' siteid 及plantype的autoplan记录是否已存在，存在返回TRUE，不存在false
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="planType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AutoPlanExist(ByVal siteId As Integer, ByVal planType As String) As AutomationPlan
        Dim autoInfo As AutomationPlan = (From ai In efEntity.AutomationPlans
                                          Where ai.SiteID = siteId AndAlso ai.PlanType = planType
                                          Select ai).FirstOrDefault()
        Return autoInfo
    End Function

    ''' <summary>
    ''' 根据siteid和categories 字符串查询是否已存在HP类型的点击计划，如有返回HP*（*代表数字）,否则返回空字符""
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <param name="categories"></param>
    ''' <returns></returns>
    Public Shared Function TriggerPlanExit(ByVal siteid As Integer, ByVal categories As String) As String
        Dim planList = (From ai In efEntity.AutomationPlans
                        Where ai.SiteID = siteid AndAlso ai.Categories = categories AndAlso ai.PlanType.StartsWith("HP")
                        Select ai).ToList()
        For Each plan As AutomationPlan In planList
            If plan.Categories = categories Then
                Return plan.PlanType
            End If
        Next
        Return ""
    End Function

    ''' <summary>
    ''' siteid 及categories的autoplan记录是否已存在，存在返回其true，不存在false
    ''' </summary>
    ''' <param name="siteId"></param>
    ''' <param name="categories"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AutoPlanCategoriesExist(ByVal siteId As Integer, ByVal categories As String) As Boolean
        categories = categories.Trim
        Dim autoInfo As AutomationPlan = (From ai In efEntity.AutomationPlans
                                          Where ai.SiteID = siteId AndAlso ai.Categories = categories
                                          Select ai).FirstOrDefault()
        If (autoInfo Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function updateAutoPlanStatus(ByVal siteId As Integer, ByVal planType As String, ByVal status As String) As Boolean
        Dim updateAutoPlan As AutomationPlan = efEntity.AutomationPlans.FirstOrDefault(Function(m) m.SiteID = siteId And m.PlanType = planType)
        updateAutoPlan.Status = status
        efEntity.SaveChanges()
    End Function

    ''' <summary>
    ''' 获取到一个目前没有使用的plantype，仅供HO,HA类型
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPlantype(ByVal siteid As Integer) As String
        Dim allPlantype As List(Of String) = (From p In efEntity.PlanTypes
                                               Where p.PlanType1.Contains("HO") Or p.PlanType1.Contains("HA")
                                               Select p.PlanType1).ToList()
        Dim plantypes As List(Of String) = (From a In efEntity.AutomationPlans
                                  Where a.SiteID = siteid
                                  Select a.PlanType).ToList()
        For Each item As String In allPlantype
            If Not (plantypes.Contains(item)) Then
                Return item
            End If
        Next
        Return ""
    End Function

    ''' <summary>
    ''' 获取到一个目前没有使用的plantype,根据是否是触发返回HP组别还是{"HO","HA"}组别
    ''' </summary>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPlantype(ByVal siteid As Integer, ByVal isTrigger As Boolean) As String
        Dim allPlantype As List(Of String)
        If (isTrigger) Then '触发类型
            allPlantype = (From p In efEntity.PlanTypes
                       Where p.PlanType1.Contains("HP")
                       Select p.PlanType1).ToList()
        Else '基本类型
            allPlantype = (From p In efEntity.PlanTypes
                       Where p.PlanType1.Contains("HO") Or p.PlanType1.Contains("HA")
                       Select p.PlanType1).ToList()
        End If

        Dim plantypes As List(Of String) = (From a In efEntity.AutomationPlans
                                  Where a.SiteID = siteid
                                  Select a.PlanType).ToList()
        For Each item As String In allPlantype
            If Not (plantypes.Contains(item)) Then
                Return item
            End If
        Next
        Return ""
    End Function
#End Region


#Region "category表"

    ''' <summary>
    ''' 将category数据整理进数据库。如果cateName存在，则更新此category。如果cateName不存在则添加至数据库
    ''' </summary>
    ''' <param name="listCategory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrUpdateCategory(ByVal listCategory As List(Of Category)) As String
        If (listCategory.Count > 0) Then

            Dim siteid As Integer = listCategory(0).SiteID
            Dim listExistCategoryName As List(Of String) = (From c In efEntity.Categories
                                                          Where c.SiteID = siteid
                                                          Select c.Category1).ToList()
            For Each item As Category In listCategory
                If Not (listExistCategoryName.Contains(item.Category1)) Then
                    efEntity.Categories.AddObject(item)
                Else
                    Dim updateCate As Category = (From c In efEntity.Categories
                                                  Where c.SiteID = item.SiteID AndAlso c.Category1 = item.Category1
                                                  Select c).FirstOrDefault()
                    updateCate.Url = item.Url
                    updateCate.LastUpdate = DateTime.Now
                    efEntity.SaveChanges()
                End If
            Next
            efEntity.SaveChanges()
        End If
    End Function

    ''' <summary>
    ''' cates的参数规范cate1^cate2...
    ''' </summary>
    ''' <param name="cates"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCategories(ByVal cates As String, ByVal siteid As Integer) As List(Of Category)
        Dim arrCate As String() = cates.Split("^")
        Dim listCate As List(Of Category) = New List(Of Category)
        For Each item In arrCate
            Dim acate As Category = (From c In efEntity.Categories
                         Where c.SiteID = siteid AndAlso c.Category1 = item
                         Select c).FirstOrDefault()
            If Not (acate Is Nothing) Then
                listCate.Add(acate)
            End If
        Next
        Return listCate
    End Function


    ''' <summary>
    ''' cates的参数规范cate1^cate2...返回新品，新品；焦点图，banner；格式的string
    ''' </summary>
    ''' <param name="cates"></param>
    ''' <param name="siteid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCategoryItems(ByVal cates As String, ByVal siteid As Integer) As String
        Dim arrCate As String() = cates.Split("^")
        Dim cateItems As String = ""
        Dim defaultCates As String() = {"新品", "热销", "低价", "收藏", "人气"}
        For Each item In arrCate
            Dim acate As Category = (From c In efEntity.Categories
                         Where c.SiteID = siteid AndAlso c.Category1 = item
                         Select c).FirstOrDefault()
            If Not (acate Is Nothing) Then
                If (defaultCates.Contains(acate.Category1.Trim)) Then
                    cateItems = cateItems & acate.Category1.Trim & "," & acate.Category1.Trim & ";"
                Else
                    cateItems = cateItems & acate.Category1.Trim & "," & acate.Url.Trim & ";"
                End If
            End If
        Next
        Return cateItems
    End Function

#End Region


End Class
