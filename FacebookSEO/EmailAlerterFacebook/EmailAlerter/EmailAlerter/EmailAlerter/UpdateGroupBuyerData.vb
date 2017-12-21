Imports EmailAlerter.GroupBuyer2
Imports System.Linq

Public Class UpdateGroupBuyerData
    Private Shared efContext As New FaceBookForSEOEntities

    Public Shared Sub Start(ByVal siteId As Integer, ByVal elementDeals() As EmailCampaignElementDeal)
        GetCategorys(siteId)
        For Each element As EmailCampaignElementDeal In elementDeals
            Dim myProduct As New AutoProduct
            myProduct.Prodouct = element.Subject
            myProduct.Url = element.Link.Substring(0, element.Link.IndexOf("""")).Trim()
            myProduct.Price =
                Double.Parse(If(element.Price.Contains("HK$"), element.Price.Replace("HK$", ""), element.Price))
            myProduct.Discount =
                Double.Parse(
                    If _
                                (element.DiscountPrice.Contains("HK$"), element.DiscountPrice.Replace("HK$", ""),
                                 element.DiscountPrice))
            myProduct.Sales = element.noofPurchased
            myProduct.PictureUrl = element.ProductImage
            myProduct.LastUpdate = element.PubDate
            myProduct.ExpiredDate = element.ExpiredDate
            myProduct.SiteID = siteId
            myProduct.Currency = "HK$"
            Try
                Dim queryProduct As AutoProduct =
                        efContext.AutoProducts.Where(
                            Function(p) p.Url = myProduct.Url AndAlso p.SiteID = myProduct.SiteID).Single()
                Dim arrCategory() As String = element.DealCategory.Split("-")
                For Each cate As String In arrCategory
                    If Not (String.IsNullOrEmpty(cate)) Then
                        Dim myCate As String = "-" & cate
                        Try
                            Dim queryCategory As AutoCategory =
                                    efContext.AutoCategories.Where(
                                        Function(c) c.Category1 = myCate AndAlso c.SiteID = siteId).Single()
                            queryProduct.Categories.Add(queryCategory)
                            efContext.SaveChanges()
                        Catch ex As Exception
                            'Ignore
                        End Try
                    End If
                Next
                queryProduct.Prodouct = myProduct.Prodouct
                queryProduct.Price = myProduct.Price
                queryProduct.Discount = myProduct.Discount
                queryProduct.Sales = myProduct.Sales
                queryProduct.LastUpdate = myProduct.LastUpdate
                queryProduct.ExpiredDate = myProduct.ExpiredDate
            Catch ex As Exception
                Dim arrCategory() As String = element.DealCategory.Split("-")
                For Each cate As String In arrCategory
                    If Not (String.IsNullOrEmpty(cate)) Then
                        Dim myCate As String = "-" & cate
                        Try
                            Dim queryCategory As AutoCategory =
                                    efContext.AutoCategories.Where(
                                        Function(c) c.Category1 = myCate AndAlso c.SiteID = siteId).Single()
                            myProduct.Categories.Add(queryCategory)
                        Catch ex1 As Exception
                            'Ignore
                        End Try
                    End If
                Next
                efContext.AddToAutoProducts(myProduct)
            End Try
            efContext.SaveChanges()
        Next
    End Sub

    Private Shared Sub GetCategorys(ByVal siteId As Integer)
        '1,2,3,4,5,6,7,8,9,10,13,14,16,17,18,19
        Dim cates() As GroupCate = New GroupCate() {New GroupCate With {.cName = "-1", .desc = "wine"},
                                                    New GroupCate With {.cName = "-2", .desc = "beauty"},
                                                    New GroupCate With {.cName = "-3", .desc = "movie"},
                                                    New GroupCate With {.cName = "-4", .desc = "deposit"},
                                                    New GroupCate With {.cName = "-5", .desc = "food"},
                                                    New GroupCate With {.cName = "-6", .desc = "product"},
                                                    New GroupCate With {.cName = "-7", .desc = "course/family"},
                                                    New GroupCate With {.cName = "-8", .desc = "travel"},
                                                    New GroupCate With {.cName = "-9", .desc = "baby"},
                                                    New GroupCate With {.cName = "-10", .desc = "computer"},
                                                    New GroupCate With {.cName = "-13", .desc = "yuenlong"},
                                                    New GroupCate With {.cName = "-14", .desc = "hitech"},
                                                    New GroupCate With {.cName = "-16", .desc = "fashion"},
                                                    New GroupCate With {.cName = "-17", .desc = "entertainment"},
                                                    New GroupCate With {.cName = "-18", .desc = "special"},
                                                    New GroupCate With {.cName = "-19", .desc = "Beauty product"}}
        For Each cate As GroupCate In cates
            Dim myCategory As New AutoCategory
            myCategory.Category1 = cate.cName
            myCategory.SiteID = siteId
            myCategory.LastUpdate = Now
            myCategory.Description = cate.desc
            Try
                Dim queryCate As AutoCategory =
                        efContext.AutoCategories.Where(
                            Function(c) c.Category1 = myCategory.Category1 AndAlso c.SiteID = siteId).Single()
            Catch ex As Exception
                efContext.AddToAutoCategories(myCategory)
            End Try
        Next
        efContext.SaveChanges()
    End Sub
End Class

Public Class GroupCate
    Public cName As String
    Public desc As String
End Class