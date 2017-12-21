Imports System
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions

Public Class EbaySite
    Inherits EBusinessSite


    Public Sub New(ByVal myshopUrl As String)
        MyBase.ShopUrl = myshopUrl.ToLower()
        MyBase.ShopType = "ebay"

        SortedType.Add("Best Match", "/m.html?_fcid=1&_sop=12")
        SortedType.Add("ending soonest", "/m.html?_fcid=1&_sop=1")
        SortedType.Add("newly listed", "/m.html?_fcid=1&_sop=10")
        SortedType.Add("lowest first", "/m.html?_fcid=1&_sop=15")
     
        DefaultSubject = "Hi [FIRSTNAME],New Arrivals From " & ShopName & ": [FIRST_PRODUCT] ..."
        DefaultTriggerSubject = "Hi [FIRSTNAME], Weekly Special [CATE_NAME] From " & ShopName & " Vol.[VOL_NUMBER]"
        TemplateType = "E"
        BannerImgRegex = "src=""([\w\W]*?(?:.bmp|.jpg|.jpeg|.png|.gif))"
    End Sub

    Public Overrides Function GetProductList(pageUrl As String, siteId As Integer) As System.Collections.Generic.List(Of Product)
        Dim productList As New List(Of Product)
        Dim efhelper As New EFHelper()
        Dim ebayCookie As String = "cid=nAa1OFr3LHR4uEKm%232037790891; shs=BAQAAAVBZWLl4AAaAAVUADlf8a7U2NjgxNTU4MjYwMDAsM1S3HfpptQ67LqScQB5Nkod+BZP6; JSESSIONID=7882EAD6BBF60AEC31C4580FF35F7DA5; _gat=1; _ga=GA1.2.614618620.1443524997; npii=btrm/svid%3D17670303016369686957fe03b9^tguid/6ea2f5a914c0a5698382d742fff7d5ff57fe03b9^cguid/6ea3052d14c0a7e38647dfd7fbeb43b157fe03b9^; ds1=ats/1444624437495; ns1=BAQAAAVBZWLl4AAaAANgAVVf+B/hjODF8NjAxXjE0NDQ2MzM5NzYzMzFeYkM1c056STJOUT09XjBeM3wyfDV8NHw3XjFeMl40XjNeMTJeMTJeMl4xXjFeMF4wXjBeMV42NDQyNDU5MDc1AKUAGlf+B/gxNDE4NDQyODQzLzA7MTQxODQ0NDU4MS8wO91IKS7NmFQ+dvhasDQ/XLrTsckL; cssg=6ea2f5a914c0a5698382d742fff7d5ff; s=BAQAAAVBZWLl4AAWAABIAClYeJfh0ZXN0Q29va2llAAMAAVYeJfgwAPQAIlYeJfgkMiRzWDIzbkdpRSRwa0ZxT2ZzZDRjTTJWWko3Y0hNQlAuAUUACFf+B/g1NjFiMzQ5NgFlAAJWHiX4IzIABgABVh4l+DAA+AAgVh4l+DZlYTJmNWE5MTRjMGE1Njk4MzgyZDc0MmZmZjdkNWZmAAwAClYeJfgxNDE4NDQ0NTgxAD0AB1YeJfhsLmw3MjY1AO4AZFYeJfgzBmh0dHA6Ly93d3cuZWJheS5jb20vc2NoL2Rpc2NvdW50ZWRzdW5nbGFzc2VzMDA3L20uaHRtbD9fbmt3PSZfYXJtcnM9MSZfaXBnPSZfZnJvbT0jaXRlbTFlOWYyMGQ3MDEHAS4AFVYeJfg1NjFiMmE2YS4wLjguOS4tMS44My4RTDQiYqF7ZmuDOEDzsyZDHCaJgg**; nonsession=BAQAAAVBZWLl4AAaAAEAAB1f+B/hsLmw3MjY1AWQAA1f+B/gjOGEABAAHV/xrtWwubDcyNjUACAAcVkRheDE0NDQ3MjgzODV4MTMxNTE4NzQ4NDE3eDB4MlkAqgABV/4H+DAAygAgX4LV+DZlYTJmNWE5MTRjMGE1Njk4MzgyZDc0MmZmZjdkNWZmAMsAAVYc24A5ABAAB1f+B/hsLmw3MjY1ADMAClf+B/g1MTAwMDAsQ0hOAPMAIlf+B/gkMiRzWDIzbkdpRSRwa0ZxT2ZzZDRjTTJWWko3Y0hNQlAuAJoACFYd2zVsLmw3MjY1eACcADhX/gf4blkrc0haMlByQm1kajZ3Vm5ZK3NFWjJQckEyZGo2QUFrWXFsREpTSHFRMmRqNng5blkrc2VRPT0AnQAIV/4H+DAwMDAwMDAx2JOemLlF1nLWiVqvJmDWVuIfXIQ*; lucky9=3498288; ds2=asotr/b7piCzQMzzzz^ssts/1444729985758^; dp1=bexpt/0001444622962235570bcc32^a1p/0561e25f8^bl/CN59df3b78^fm/4.3.25642bfbd^kms/in59df3b78^pcid/203779089157fe07f8^mpc/0%7C20156446178^pbf/%23284004a0002000008180c200000457fe07f8^tzo/-1e059df3b81^exc/0%3A0%3A2%3A256446178^u1p/bC5sNzI2NQ**57fe07f8^u1f/dora57fe07f8^idm/1561e21a2^; ebay=%5Epsi%3DAiSsHhYg*%5EsfLMD%3D0%5Esbf%3D%2341000000c00120008162124%5Ecos%3D1%5Ecv%3D15555%5Esin%3Din%5Ejs%3D1%5Edv%3D561b58d8%5E"
        Dim prodHtmlDouctment As HtmlDocument = efhelper.GetHtmlDocument(pageUrl, ebayCookie, "", "utf-8")
        Dim prodNodeList As HtmlNodeCollection = prodHtmlDouctment.DocumentNode.SelectNodes("//ul[@id='GalleryViewInner']/li/div/div")
        Dim index As Integer
        For i As Integer = 0 To prodNodeList.Count - 1
            Dim itmeProdNode As HtmlNode = prodNodeList(i)
            Dim myProduct As New Product
            Dim productName As String = itmeProdNode.SelectSingleNode("div[@class='gvtitle']").InnerText

            myProduct.Prodouct = productName
            myProduct.Url = itmeProdNode.SelectSingleNode("div[@class='img l-shad lftd']/div/div/div/a").GetAttributeValue("href", "").Trim()
            If (myProduct.Url.Contains("?")) Then
                index = myProduct.Url.IndexOf("?")
                myProduct.Url = myProduct.Url.Remove(index)
            End If
            Try
                myProduct.Discount = itmeProdNode.SelectSingleNode("div[@class='prices']/div[@class='bin']/div[1]/span[@class='amt']/span").InnerText.Trim().Replace("$", "")
            Catch ex As Exception
                Continue For
            End Try
            'Dim discount As String = itmeProdNode.SelectSingleNode("div[@class='prices']/div[@class='bin']/div[1]/span[@class='amt']/span").InnerText.Trim().Replace("$", "")
            'Dim prodDiscount As Decimal
            'If (Decimal.TryParse(discount, prodDiscount)) Then
            '    myProduct.Discount = prodDiscount
            'Else
            '    Continue For
            'End If
            myProduct.PictureUrl = itmeProdNode.SelectSingleNode("div[@class='img l-shad lftd']/div/div/div/a/img").GetAttributeValue("src", "").Trim()
            myProduct.Currency = "$"
            myProduct.PictureUrl = myProduct.PictureUrl
            myProduct.Description = productName
            myProduct.LastUpdate = DateTime.Now
            myProduct.SiteID = siteId
            productList.Add(myProduct)
        Next
        Return productList
    End Function

    Public Overrides Function GetShopNameandLogo() As String()
        Dim efhelper As New EFHelper()
        Dim htmldocument As HtmlDocument = efhelper.GetHtmlDocument(ShopUrl, "", "", "")
        Dim resultString(2) As String
        resultString(0) = htmldocument.DocumentNode.SelectSingleNode("//a[@class='mbg-id']").InnerText.Trim.Replace("User ID&nbsp;", "") ''shopName
        resultString(1) = htmldocument.DocumentNode.SelectSingleNode("//img[@class='prof_img img']").GetAttributeValue("src", "") ''shopLOGO
        If resultString(1) = "http://assets.alicdn.com/s.gif" Then '识别错误
            resultString(1) = ""
        End If
        Return resultString
    End Function

    Public Overrides Function GetSortedTypeURl(ByVal cateName As String) As String
        cateName = cateName.Trim
        If (ShopUrl.EndsWith("/")) Then
            ShopUrl = ShopUrl.Substring(0, ShopUrl.Length - 1)
        End If
        Dim resultUrl As String
        resultUrl = ShopUrl.Replace("/usr", "/sch")

        Dim sortedcaterul As String
        If (SortedType.TryGetValue(cateName, sortedcaterul)) Then
            If Not (sortedcaterul.StartsWith("/")) Then
                sortedcaterul = "/" & sortedcaterul
            End If
            Return resultUrl & sortedcaterul
        Else
            Return ""
        End If
    End Function

    Public Overrides Function StandardlizeShopURl(ByVal shopUrl As String) As String

        Return shopUrl
    End Function
End Class
