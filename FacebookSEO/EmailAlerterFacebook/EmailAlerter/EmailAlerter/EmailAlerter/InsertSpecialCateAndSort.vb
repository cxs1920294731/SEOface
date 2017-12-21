Imports EmailAlerter.GroupBuyer2
Imports System.Linq

Public Class InsertSpecialCateAndSort
    ''' <summary>
    ''' 2014/02/08新需求：
    ''' 增加新需求1：将当天的Foods分类（分类Num: 5）的Best selling的一个Deals放置在所有的分类邮件里的第一位，
    ''' Best Selling的判断根据RSS里的"noofPurchased"标签判断；
    ''' 增加新需求2：以前根据Spread点击获取最受欢迎的某一个Deal，改为根据"noofPurchased"获取；
    ''' </summary>
    ''' <param name="LoginEmail"></param>
    ''' <param name="AppID"></param>
    ''' <param name="EmailTemplateElements"></param>
    ''' <param name="categoryId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertSpecialCateAndSort(ByVal LoginEmail As String, ByVal AppID As String,
                                                    ByVal EmailTemplateElements As EmailCampaignElementDeal(),
                                                    ByVal categoryId As String) As EmailCampaignElementDeal()
        Dim listElements As List(Of EmailCampaignElementDeal) =
                New List(Of EmailCampaignElementDeal)(EmailTemplateElements)
        listElements = listElements.OrderByDescending(Function(element) element.noofPurchased).ToList()
        Dim newElement As New EmailCampaignElementDeal
        'categoryId = "-" & categoryId  '获取到的分类格式：-2或者-14-16
        Dim counter As Integer = 0
        For Each li As EmailCampaignElementDeal In listElements
            If (DateTime.Parse(li.PubDate).AddDays(1).ToString("yyyyMMdd") = DateTime.Now.ToString("yyyyMMdd")) Then
                Dim arrCategoryId As String() = li.DealCategory.Split("-")
                '获取到的分类格式：-2或者-14-16
                Dim flag As Boolean = False
                For Each arr As String In arrCategoryId
                    If (arr = categoryId) Then '找到主推类目的产品
                        flag = True
                        Exit For
                    End If
                Next
                If (flag = True) Then
                    newElement = li
                    Exit For
                End If

            End If
            counter = counter + 1
        Next
        If Not (String.IsNullOrEmpty(newElement.DealCategory)) Then '查看新对象是否含产品对象
            listElements.RemoveAt(counter)
            listElements.Insert(0, newElement)
        End If
        EmailTemplateElements = listElements.ToArray()
        Return EmailTemplateElements
    End Function

    ''' <summary>
    ''' list的第一个产品位置不变，把当前分类的产品提到第一个产品后面几个位置
    ''' </summary>
    ''' <param name="categoryId"></param>
    ''' <param name="listEmailCampaignElementDeal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertCurrentCateProduct(ByVal categoryId As String,
                                                    ByVal listEmailCampaignElementDeal As _
                                                       List(Of EmailCampaignElementDeal)) _
        As List(Of EmailCampaignElementDeal)
        Dim listCurrentCateProd As New List(Of EmailCampaignElementDeal)
        Dim listNotCurrenCateProd As New List(Of EmailCampaignElementDeal)
        '先把第一个Food产品从传过来的list移动新的list中
        listNotCurrenCateProd.Add(listEmailCampaignElementDeal(0))
        listEmailCampaignElementDeal.RemoveAt(0)

        For Each element As EmailCampaignElementDeal In listEmailCampaignElementDeal
            Dim arrCategoryId As String() = element.DealCategory.Split("-")
            '获取到的分类格式：-2或者-14-16
            Dim flag As Boolean = False
            For Each arr As String In arrCategoryId
                If (arr = categoryId) Then '找到主推类目的产品
                    flag = True
                    Exit For
                End If
            Next
            If (flag = True) Then
                listCurrentCateProd.Add(element)
                'Continue For
            Else
                listNotCurrenCateProd.Add(element)
            End If
        Next
        For Each li As EmailCampaignElementDeal In listCurrentCateProd
            listNotCurrenCateProd.Insert(1, li)
        Next
        Return listNotCurrenCateProd
    End Function
End Class
