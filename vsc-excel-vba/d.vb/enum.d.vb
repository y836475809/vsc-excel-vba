Public Module ExcelVBAEnum
    Public Enum XlFindLookIn
        xlComments = -4144
        xlCommentsThreaded = -4184
        xlFormulas = -4123
        xlValues = -4163
    End Enum
    Public Const xlComments = -4144
    Public Const xlCommentsThreaded = -4184
    Public Const xlFormulas = -4123
    Public Const xlValues = -4163
    
    Public Enum XlLookAt
        xlWhole = 2
        xlPart = 1
    End Enum
    Public Const xlWhole = 2
    Public Const xlPart = 1

    Public Enum XlDirection
        xlDown = -4121
        xlToLeft = -4159
        xlToRight = -4161
        xlUp = -4162
    End Enum
    Public Const xlDown = -4121
    Public Const xlToLeft = -4159
    Public Const xlToRight = -4161
    Public Const xlUp = -4162

    Public Enum XlAutoFilterOperator
        xlAnd = 1
        xlBottom10Items = 4
        xlBottom10Percent = 6
        xlFilterCellColor = 8
        xlFilterDynamic = 11
        xlFilterFontColor = 9
        xlFilterIcon = 10
        xlFilterValues = 7
        xlOr = 2
        xlTop10Items = 3
        xlTop10Percent = 5
    End Enum
    Public Const xlAnd = 1
    Public Const xlBottom10Items = 4
    Public Const xlBottom10Percent = 6
    Public Const xlFilterCellColor = 8
    Public Const xlFilterDynamic = 11
    Public Const xlFilterFontColor = 9
    Public Const xlFilterIcon = 10
    Public Const xlFilterValues = 7
    Public Const xlOr = 2
    Public Const xlTop10Items = 3
    Public Const xlTop10Percent = 5 

    Public Enum XlCellType
        lCellTypeAllFormatConditions = -4172
        xlCellTypeAllValidation = -4174
        xlCellTypeBlanks = 4
        xlCellTypeComments = -4144
        xlCellTypeConstants = 2
        xlCellTypeFormulas = -4123
        xlCellTypeLastCell = 11
        xlCellTypeSameFormatConditions = -4173
        xlCellTypeSameValidation = -4175
        xlCellTypeVisible = 12
    End Enum
    Public Const lCellTypeAllFormatConditions = -4172
    Public Const xlCellTypeAllValidation = -4174
    Public Const xlCellTypeBlanks = 4
    Public Const xlCellTypeComments = -4144
    Public Const xlCellTypeConstants = 2
    Public Const xlCellTypeFormulas = -4123
    Public Const xlCellTypeLastCell = 11
    Public Const xlCellTypeSameFormatConditions = -4173
    Public Const xlCellTypeSameValidation = -4175
    Public Const xlCellTypeVisible = 12

    Public Enum XlPlatform
        xlMacintosh = 1
        xlMSDOS = 3
        xlWindows = 2
    End Enum
    Public Const xlMacintosh = 1
    Public Const xlMSDOS = 3
    Public Const xlWindows = 2

    Public Enum XlCorruptLoad
        xlExtractData = 2
        xlNormalLoad = 0
        xlRepairFile = 1
    End Enum
    Public Const xlExtractData = 2
    Public Const xlNormalLoad = 0
    Public Const xlRepairFile = 1
End Module