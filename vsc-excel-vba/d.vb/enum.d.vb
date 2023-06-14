Public Module ExcelVBAEnum
    Public Enum XlFindLookIn
        xlComments
        xlCommentsThreaded
        xlFormulas
        xlValues
    End Enum
    Public Const xlComments
    Public Const xlCommentsThreaded
    Public Const xlFormulas
    Public Const xlValues
    
    Public Enum XlLookAt
        xlWhole
        xlPart
    End Enum
    Public Const xlWhole
    Public Const xlPart

    Public Enum XlDirection
        xlDown
        xlToLeft
        xlToRight
        xlUp
    End Enum
    Public Const xlDown
    Public Const xlToLeft
    Public Const xlToRight
    Public Const xlUp
End Module