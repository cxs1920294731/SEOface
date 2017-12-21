Imports System.Collections.Generic
Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Serialization
Imports System.Net


''' -----------------------------------------------------------------------------
''' <summary>
''' The XmlUtils class provides Shared/Static methods for manipulating xml files
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[cnurse]	11/08/2004	created
''' </history>
''' -----------------------------------------------------------------------------
    Public Class XmlUtils
    Public Shared Sub AppendElement(ByRef objDoc As XmlDocument, ByVal objNode As XmlNode, ByVal attName As String,
                                    ByVal attValue As String, ByVal includeIfEmpty As Boolean)
        AppendElement(objDoc, objNode, attName, attValue, includeIfEmpty, False)
    End Sub

    Public Shared Sub AppendElement(ByRef objDoc As XmlDocument, ByVal objNode As XmlNode, ByVal attName As String,
                                    ByVal attValue As String, ByVal includeIfEmpty As Boolean, ByVal CDATA As Boolean)
        If attValue = "" And Not includeIfEmpty Then
            Exit Sub
        End If
        If CDATA Then
            objNode.AppendChild(CreateCDataElement(objDoc, attName, attValue))
        Else
            objNode.AppendChild(CreateElement(objDoc, attName, attValue))
        End If
    End Sub

    Public Shared Function CreateAttribute(ByVal objDoc As XmlDocument, ByVal attName As String,
                                           ByVal attValue As String) As XmlAttribute
        Dim attribute As XmlAttribute = objDoc.CreateAttribute(attName)
        attribute.Value = attValue
        Return attribute
    End Function

    Public Shared Sub CreateAttribute(ByVal objDoc As XmlDocument, ByVal objNode As XmlNode, ByVal attName As String,
                                      ByVal attValue As String)
        Dim attribute As XmlAttribute = objDoc.CreateAttribute(attName)
        attribute.Value = attValue
        objNode.Attributes.Append(attribute)
    End Sub

    Public Shared Function CreateElement(ByVal objDoc As XmlDocument, ByVal NodeName As String) As XmlElement
        Return objDoc.CreateElement(NodeName)
    End Function

    Public Shared Function CreateElement(ByVal objDoc As XmlDocument, ByVal NodeName As String,
                                         ByVal NodeValue As String) As XmlElement
        Dim element As XmlElement = objDoc.CreateElement(NodeName)
        element.InnerText = NodeValue
        Return element
    End Function

    Public Shared Function CreateCDataElement(ByVal objDoc As XmlDocument, ByVal NodeName As String,
                                              ByVal NodeValue As String) As XmlElement
        Dim element As XmlElement = objDoc.CreateElement(NodeName)
        element.AppendChild(objDoc.CreateCDataSection(NodeValue))
        Return element
    End Function

    Public Shared Function Deserialize(ByVal xmlObject As String, ByVal type As System.Type) As Object
        Dim ser As XmlSerializer = New XmlSerializer(type)
        Dim sr As New StringReader(xmlObject)
        Dim obj As Object = ser.Deserialize(sr)
        sr.Close()
        Return obj
    End Function

    Public Shared Function DeSerializeDictionary (Of TValue)(ByVal objStream As Stream, ByVal rootname As String) _
        As Dictionary(Of Integer, TValue)
        Dim objDictionary As Dictionary(Of Integer, TValue) = Nothing

        Dim xmlDoc As New XmlDocument
        xmlDoc.Load(objStream)

        objDictionary = New Dictionary(Of Integer, TValue)

        For Each xmlItem As XmlElement In xmlDoc.SelectNodes(rootname + "/item")
            Dim key As Integer = Convert.ToInt32(xmlItem.GetAttribute("key"))
            Dim typeName As String = xmlItem.GetAttribute("type")

            Dim objValue As TValue = Activator.CreateInstance (Of TValue)()

            'Create the XmlSerializer
            Dim xser As New XmlSerializer(objValue.GetType)

            'A reader is needed to read the XML document.
            Dim reader As New XmlTextReader(New StringReader(xmlItem.InnerXml))

            ' Use the Deserialize method to restore the object's state, and store it
            ' in the Hashtable
            objDictionary.Add(key, CType(xser.Deserialize(reader), TValue))
        Next

        Return objDictionary
    End Function

    Public Shared Function DeSerializeHashtable(ByVal xmlSource As String, ByVal rootname As String) As Hashtable
        Dim objHashTable As Hashtable
        If xmlSource <> "" Then
            objHashTable = New Hashtable

            Dim xmlDoc As New XmlDocument
            xmlDoc.LoadXml(xmlSource)

            For Each xmlItem As XmlElement In xmlDoc.SelectNodes(rootname + "/item")
                Dim key As String = xmlItem.GetAttribute("key")
                Dim typeName As String = xmlItem.GetAttribute("type")

                'Create the XmlSerializer
                Dim xser As New XmlSerializer(Type.GetType(typeName))

                'A reader is needed to read the XML document.
                Dim reader As New XmlTextReader(New StringReader(xmlItem.InnerXml))

                ' Use the Deserialize method to restore the object's state, and store it
                ' in the Hashtable
                objHashTable.Add(key, xser.Deserialize(reader))
            Next
        Else
            objHashTable = New Hashtable
        End If
        Return objHashTable
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of an attribute
    ''' </summary>
    ''' <param name="nav">Parent XPathNavigator</param>
    ''' <param name="AttributeName">Thename of the Attribute</param>
    ''' <returns></returns>
    ''' <history>
    ''' 	[cnurse]	05/14/2008	created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetAttributeValue(ByVal nav As XPathNavigator, ByVal AttributeName As String) As String
        Return nav.GetAttribute(AttributeName, "")
    End Function

    Public Shared Function GetAttributeValueAsInteger(ByVal nav As XPathNavigator, ByVal AttributeName As String,
                                                      ByVal DefaultValue As Integer) As Integer
        Dim intValue As Integer = DefaultValue

        Dim strValue As String = GetAttributeValue(nav, AttributeName)
        If Not String.IsNullOrEmpty(strValue) Then
            intValue = Convert.ToInt32(strValue)
        End If

        Return intValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of a node
    ''' </summary>
    ''' <param name="nav">Parent XPathNavigator</param>
    ''' <param name="path">The Xpath expression to the value</param>
    ''' <returns></returns>
    ''' <history>
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValue(ByVal nav As XPathNavigator, ByVal path As String) As String
        Dim strValue As String = Null.NullString

        Dim elementNav As XPathNavigator = nav.SelectSingleNode(path)
        If elementNav IsNot Nothing Then
            strValue = elementNav.Value
        End If

        Return strValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <param name="DefaultValue">Default value to return</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Created
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValue(ByVal objNode As XmlNode, ByVal NodeName As String,
                                        Optional ByVal DefaultValue As String = "") As String
        Dim strValue As String = DefaultValue

        If (objNode.Item(NodeName) IsNot Nothing) Then
            strValue = objNode.Item(NodeName).InnerText

            If strValue = "" And DefaultValue <> "" Then
                strValue = DefaultValue
            End If
        End If

        Return strValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value (False) will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueBoolean(ByVal objNode As XmlNode, ByVal NodeName As String) As Boolean
        Return GetNodeValueBoolean(objNode, NodeName, False)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <param name="DefaultValue">Default value to return</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueBoolean(ByVal objNode As XmlNode, ByVal NodeName As String,
                                               ByVal DefaultValue As Boolean) As Boolean
        Dim bValue As Boolean = DefaultValue

        If (objNode.Item(NodeName) IsNot Nothing) Then
            Dim strValue As String = objNode.Item(NodeName).InnerText

            If Not String.IsNullOrEmpty(strValue) Then
                bValue = Convert.ToBoolean(strValue)
            End If
        End If

        Return bValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <param name="DefaultValue">Default value to return</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueDate(ByVal objNode As XmlNode, ByVal NodeName As String,
                                            ByVal DefaultValue As DateTime) As DateTime
        Dim dateValue As DateTime = DefaultValue

        If (objNode.Item(NodeName) IsNot Nothing) Then
            Dim strValue As String = objNode.Item(NodeName).InnerText

            If Not String.IsNullOrEmpty(strValue) Then

                dateValue = Convert.ToDateTime(strValue)
                If dateValue.Date.Equals(Null.NullDate.Date) Then
                    dateValue = Null.NullDate
                End If
            End If
        End If

        Return dateValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value (0) will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueInt(ByVal objNode As XmlNode, ByVal NodeName As String) As Integer
        Return GetNodeValueInt(objNode, NodeName, 0)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <param name="DefaultValue">Default value to return</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueInt(ByVal objNode As XmlNode, ByVal NodeName As String,
                                           ByVal DefaultValue As Integer) As Integer
        Dim intValue As Integer = DefaultValue

        If (objNode.Item(NodeName) IsNot Nothing) Then
            Dim strValue As String = objNode.Item(NodeName).InnerText

            If Not String.IsNullOrEmpty(strValue) Then
                intValue = Convert.ToInt32(strValue)
            End If
        End If

        Return intValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value (0) will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueSingle(ByVal objNode As XmlNode, ByVal NodeName As String) As Single
        Return GetNodeValueSingle(objNode, NodeName, 0)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the value of node
    ''' </summary>
    ''' <param name="objNode">Parent node</param>
    ''' <param name="NodeName">Child node to look for</param>
    ''' <param name="DefaultValue">Default value to return</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the node does not exist or it causes any error the default value will be returned.
    ''' </remarks>
    ''' <history>
    ''' 	[VMasanas]	09/09/2004	Added new method to return converted values
    ''' 	[cnurse]	11/08/2004	moved from PortalController and made Public Shared
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetNodeValueSingle(ByVal objNode As XmlNode, ByVal NodeName As String,
                                              ByVal DefaultValue As Single) As Single
        Dim sValue As Single = DefaultValue

        If (objNode.Item(NodeName) IsNot Nothing) Then
            Dim strValue As String = objNode.Item(NodeName).InnerText

            If Not String.IsNullOrEmpty(strValue) Then
                sValue = Convert.ToSingle(strValue)
            End If
        End If

        Return sValue
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets an XmlWriterSettings object
    ''' </summary>
    ''' <param name="conformance">Conformance Level</param>
    ''' <returns>An XmlWriterSettings</returns>
    ''' <history>
    ''' 	[cnurse]	08/22/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetXmlWriterSettings(ByVal conformance As ConformanceLevel) As XmlWriterSettings
        Dim settings As New XmlWriterSettings()
        settings.ConformanceLevel = conformance
        settings.OmitXmlDeclaration = True
        settings.Indent = True

        Return settings
    End Function

    Public Shared Sub UpdateAttribute(ByVal node As XmlNode, ByVal attName As String, ByVal attValue As String)
        If Not node Is Nothing Then
            Dim attrib As XmlAttribute = node.Attributes(attName)
            attrib.InnerText = attValue
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Xml Encodes HTML
    ''' </summary>
    ''' <param name="HTML">The HTML to encode</param>
    ''' <returns></returns>
    ''' <history>
    '''		[cnurse]	09/29/2005	moved from Globals
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function XMLEncode(ByVal HTML As String) As String
        Return "<![CDATA[" & HTML & "]]>"
    End Function

    Public Shared Sub XSLTransform(ByVal doc As XmlDocument, ByRef writer As StreamWriter, ByVal xsltUrl As String)

        Dim xslt As Xsl.XslCompiledTransform = New Xsl.XslCompiledTransform
        xslt.Load(xsltUrl)
        'Transform the file.
        xslt.Transform(doc, Nothing, writer)
    End Sub
End Class

