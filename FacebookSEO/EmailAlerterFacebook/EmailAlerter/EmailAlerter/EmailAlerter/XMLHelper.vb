Imports System.Xml

''' <summary>
''' XMLHelper XML文档操作管理器
''' </summary>
Public Class XMLHelper
    '
    ' TODO: 在此处添加构造函数逻辑
    '
    Public Sub New()
    End Sub


#Region "XML文档节点查询和读取"
    ''' <summary>
    ''' 选择匹配XPath表达式的第一个节点XmlNode.
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名")</param>
    ''' <returns>返回XmlNode</returns>
    Public Shared Function GetXmlNodeByXpath(ByVal xmlFileName As String, ByVal xpath As String) As XmlNode
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            Return xmlNode
        Catch ex As Exception
            'throw ex; //这里可以定义你自己的异常处理
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 选择匹配XPath表达式的节点列表XmlNodeList.
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名")</param>
    ''' <returns>返回XmlNodeList</returns>
    Public Shared Function GetXmlNodeListByXpath(ByVal xmlFileName As String, ByVal xpath As String) As XmlNodeList
        Dim xmlDoc As New XmlDocument()

        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNodeList As XmlNodeList = xmlDoc.SelectNodes(xpath)
            Return xmlNodeList
        Catch ex As Exception
            'throw ex; //这里可以定义你自己的异常处理
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 选择匹配XPath表达式的节点列表XmlNodeList下的loc节点的url
    ''' </summary>
    ''' <param name="xmlFileName"></param>
    ''' <param name="xpath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetUrlListByXpath(ByVal xmlFileName As String, ByVal xpath As String) As List(Of String)
        Dim listUrl As New List(Of String)
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNodeList As XmlNodeList = xmlDoc.SelectNodes(xpath)
            For Each node As XmlNode In xmlNodeList
                Dim url As String = node.SelectSingleNode("loc").InnerText
                listUrl.Add(url)
            Next
        Catch ex As Exception
            'throw ex; //这里可以定义你自己的异常处理
        End Try
        Return listUrl
    End Function

    ''' <summary>
    ''' 选择匹配XPath表达式的第一个节点的匹配xmlAttributeName的属性XmlAttribute.
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名</param>
    ''' <param name="xmlAttributeName">要匹配xmlAttributeName的属性名称</param>
    ''' <returns>返回xmlAttributeName</returns>
    Public Shared Function GetXmlAttribute(ByVal xmlFileName As String, ByVal xpath As String, ByVal xmlAttributeName As String) As XmlAttribute
        Dim content As String = String.Empty
        Dim xmlDoc As New XmlDocument()
        Dim xmlAttribute As XmlAttribute = Nothing
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If xmlNode IsNot Nothing Then
                If xmlNode.Attributes.Count > 0 Then
                    xmlAttribute = xmlNode.Attributes(xmlAttributeName)
                End If
            End If
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return xmlAttribute
    End Function
#End Region

#Region "XML文档创建和节点或属性的添加、修改"
    ''' <summary>
    ''' 创建一个XML文档
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="rootNodeName">XML文档根节点名称(须指定一个根节点名称)</param>
    ''' <param name="version">XML文档版本号(必须为:"1.0")</param>
    ''' <param name="encoding">XML文档编码方式</param>
    ''' <param name="standalone">该值必须是"yes"或"no",如果为null,Save方法不在XML声明上写出独立属性</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function CreateXmlDocument(ByVal xmlFileName As String, ByVal rootNodeName As String, ByVal version As String, ByVal encoding As String, ByVal standalone As String) As Boolean
        Dim isSuccess As Boolean = False
        Try
            Dim xmlDoc As New XmlDocument()
            Dim xmlDeclaration As XmlDeclaration = xmlDoc.CreateXmlDeclaration(version, encoding, standalone)
            Dim root As XmlNode = xmlDoc.CreateElement(rootNodeName)
            xmlDoc.AppendChild(xmlDeclaration)
            xmlDoc.AppendChild(root)
            xmlDoc.Save(xmlFileName)
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 依据匹配XPath表达式的第一个节点来创建它的子节点(如果此节点已存在则追加一个新的同名节点)
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名")</param>
    ''' <param name="xmlNodeName">要匹配xmlNodeName的节点名称</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function CreateXmlNodeByXPath(ByVal xmlFileName As String, ByVal xpath As String, ByVal xmlNodeName As String) As Boolean
        'Public Shared Function CreateXmlNodeByXPath(ByVal xmlFileName As String, ByVal xpath As String, ByVal xmlNodeName As String, ByVal innerText As String, ByVal xmlAttributeName As String, ByVal value As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            'Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If xmlNode IsNot Nothing Then
                '存不存在此节点都创建
                Dim subElement As XmlElement = xmlDoc.CreateElement(xmlNodeName)
                subElement.InnerText = ""

                ''如果属性和值参数都不为空则在此新节点上新增属性o d
                'If Not String.IsNullOrEmpty(xmlAttributeName) AndAlso Not String.IsNullOrEmpty(value) Then
                '    Dim xmlAttribute As XmlAttribute = xmlDoc.CreateAttribute(xmlAttributeName)
                '    xmlAttribute.Value = value
                '    subElement.Attributes.Append(xmlAttribute)
                'End If

                xmlNode.AppendChild(subElement)
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 依据匹配XPath表达式的第一个节点来创建或更新它的子节点(如果节点存在则更新,不存在则创建)
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名</param>
    ''' <param name="xmlNodeName">要匹配xmlNodeName的节点名称</param>
    ''' <param name="innerText">节点文本值</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function CreateOrUpdateXmlNodeByXPath(ByVal xmlFileName As String, ByVal xpath As String, ByVal path As String, _
                                                        ByVal xmlNodeName As String, ByVal innerText As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim isExistsNode As Boolean = False
        '标识节点是否存在
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode
            For Each node As XmlNode In xmlDoc.SelectNodes(path) '
                Dim nullNode As XmlNode = node.SelectSingleNode(xmlNodeName)
                If (nullNode Is Nothing) Then
                    xmlNode = node
                    Exit For
                End If
            Next
            '= xmlDoc.SelectSingleNode(xpath)
            innerText = EscapeXMLValue(innerText)
            If xmlNode IsNot Nothing Then
                '遍历xpath节点下的所有子节点()
                For Each node As XmlNode In xmlNode.ChildNodes
                    If node.Name.ToLower() = xmlNodeName.ToLower() Then
                        '存在此节点则更新
                        node.InnerXml = innerText
                        isExistsNode = True
                        Exit For
                    End If
                Next
                If Not isExistsNode Then
                    '不存在此节点则创建
                    Dim subElement As XmlElement = xmlDoc.CreateElement(xmlNodeName)
                    subElement.InnerXml = innerText
                    xmlNode.AppendChild(subElement)
                End If
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 依据匹配XPath表达式的第一个节点来创建或更新它的属性(如果属性存在则更新,不存在则创建)
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名</param>
    ''' <param name="xmlAttributeName">要匹配xmlAttributeName的属性名称</param>
    ''' <param name="value">属性值</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function CreateOrUpdateXmlAttributeByXPath(ByVal xmlFileName As String, ByVal xpath As String, ByVal xmlAttributeName As String, ByVal value As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim isExistsAttribute As Boolean = False
        '标识属性是否存在
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If xmlNode IsNot Nothing Then
                '遍历xpath节点中的所有属性
                For Each attribute As XmlAttribute In xmlNode.Attributes
                    If attribute.Name.ToLower() = xmlAttributeName.ToLower() Then
                        '节点中存在此属性则更新
                        attribute.Value = value
                        isExistsAttribute = True
                        Exit For
                    End If
                Next
                If Not isExistsAttribute Then
                    '节点中不存在此属性则创建
                    Dim xmlAttribute As XmlAttribute = xmlDoc.CreateAttribute(xmlAttributeName)
                    xmlAttribute.Value = value
                    xmlNode.Attributes.Append(xmlAttribute)
                End If
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function
#End Region

#Region "XML文档节点或属性的删除"
    ''' <summary>
    ''' 删除匹配XPath表达式的第一个节点所有子元素(子节点)
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名"或者"//节点名")</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function DeleteXmlNodeByXPath(ByVal xmlFileName As String, ByVal xpath As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If xmlNode IsNot Nothing Then
                '删除节点
                'xmlNode.ParentNode.RemoveChild(xmlNode)
                xmlNode.RemoveAll()  '删除xpath下的所有子节点
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 删除匹配XPath表达式的第一个节点中的匹配参数xmlAttributeName的属性
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名</param>
    ''' <param name="xmlAttributeName">要删除的xmlAttributeName的属性名称</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function DeleteXmlAttributeByXPath(ByVal xmlFileName As String, ByVal xpath As String, ByVal xmlAttributeName As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim isExistsAttribute As Boolean = False
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            Dim xmlAttribute As XmlAttribute = Nothing
            If xmlNode IsNot Nothing Then
                '遍历xpath节点中的所有属性
                For Each attribute As XmlAttribute In xmlNode.Attributes
                    If attribute.Name.ToLower() = xmlAttributeName.ToLower() Then
                        '节点中存在此属性
                        xmlAttribute = attribute
                        isExistsAttribute = True
                        Exit For
                    End If
                Next
                If isExistsAttribute Then
                    '删除节点中的属性
                    xmlNode.Attributes.Remove(xmlAttribute)
                End If
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function

    ''' <summary>
    ''' 删除匹配XPath表达式的第一个节点中的所有属性
    ''' </summary>
    ''' <param name="xmlFileName">XML文档完全文件名(包含物理路径)</param>
    ''' <param name="xpath">要匹配的XPath表达式(例如:"//节点名//子节点名</param>
    ''' <returns>成功返回true,失败返回false</returns>
    Public Shared Function DeleteAllXmlAttributeByXPath(ByVal xmlFileName As String, ByVal xpath As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.Load(xmlFileName)
            '加载XML文档
            Dim xmlNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If xmlNode IsNot Nothing Then
                '遍历xpath节点中的所有属性
                xmlNode.Attributes.RemoveAll()
            End If
            xmlDoc.Save(xmlFileName)
            '保存到XML文档
            isSuccess = True
        Catch ex As Exception
            '这里可以定义你自己的异常处理
            Throw ex
        End Try
        Return isSuccess
    End Function
#End Region


#Region "XML属性值规范转化"
 
    Public Shared Function UnescapeXMLValue(ByVal xmlString As String) As String
        If (String.IsNullOrEmpty(xmlString)) Then
            Return ""
        End If
        Return xmlString.Replace("&apos;", "'").Replace("&quot;", """").Replace("&gt;", ">").Replace("&lt;", "<").Replace("&amp;", "&")
    End Function

    ''' <summary>
    '''将不合法的xml属性值合法化，例如单引号'替换成&apos;
    ''' </summary>
    ''' <param name="xmlString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function EscapeXMLValue(ByVal xmlString As String) As String
        If (String.IsNullOrEmpty(xmlString)) Then
            Return ""
        End If
        Return xmlString.Replace("'", "&apos;").Replace("""", "&quot;").Replace(">", "&gt;").Replace("<", "&lt;").Replace("&", "&amp;")
    End Function

#End Region
End Class
