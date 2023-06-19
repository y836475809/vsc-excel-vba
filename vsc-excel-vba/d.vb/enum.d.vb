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

    Public Enum XlAutoFilterOperator
        xlAnd
        xlOr
    End Enum
    Public Const xlAnd
    Public Const xlOr 

    Public Enum XlCellType
        lCellTypeAllFormatConditions
        xlCellTypeAllValidation
        xlCellTypeBlanks
        xlCellTypeComments
        xlCellTypeConstants
        xlCellTypeFormulas
        xlCellTypeLastCell
        xlCellTypeSameFormatConditions
        xlCellTypeSameValidation
        xlCellTypeVisible
    End Enum
    Public Const lCellTypeAllFormatConditions
    Public Const xlCellTypeAllValidation
    Public Const xlCellTypeBlanks
    Public Const xlCellTypeComments
    Public Const xlCellTypeConstants
    Public Const xlCellTypeFormulas
    Public Const xlCellTypeLastCell
    Public Const xlCellTypeSameFormatConditions
    Public Const xlCellTypeSameValidation
    Public Const xlCellTypeVisible
End Module