Imports System.Collections

Public Class IXMLDOMNode
    Public text As String
    Public Property childNodes As IXMLDOMNodeList
End Class

Public Class IXMLDOMNodeList : Implements IEnumerable
    Public Property length As Long

    Default Public Property Item(index As Long) As IXMLDOMNode
        Get : End Get
        Set(value As IXMLDOMNode) : End Set
    End Property

    Public Function nextNode() As IXMLDOMNode
    End Function

    Public Sub Reset()
    End Sub
End Class

Namespace MSXML2
    Public Class DOMDocument60
        Public Function load(xmlSource As String) As Boolean
        End Function

        Public Function getElementsByTagName(tagName As String) As IXMLDOMNodeList
        End Function

        Public Function SelectNodes(xpath As String) As IXMLDOMNodeList
        End Function

        Public Function SelectSingleNode(xpath As String) As IXMLDOMNode
        End Function
    End Class
End Namespace